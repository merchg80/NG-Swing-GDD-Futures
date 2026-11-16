"""Synthetic H-reference pricing and source refresh regressions."""
from pathlib import Path
from io import BytesIO
import tempfile,unittest
import numpy as np
import pandas as pd
import deal_economics as econ
from nymex_reference import MODES,henry_hub
from ice_pricing_sources import ordered,select_newest


def prices(count=48,asof='2026-11-16'):
    rows=[]
    for i,month in enumerate(pd.date_range('2026-12-01',periods=count,freq='MS')):
        for code,value in [('H',3.+i%2),('HXS',-.2 if i%2==0 else .1),('SHS',-.1 if i%2==0 else .2)]:
            rows.append(dict(trade_date=pd.Timestamp(asof),hub='Synthetic hub',product='Synthetic '+code,
                strip=month,contract=code,settlement_price=value))
    return pd.DataFrame(rows)


SELECTIONS=[('Synthetic hub','Synthetic '+c,c) for c in ['H','HXS','SHS']]


class NymexEconomicsTests(unittest.TestCase):
    def test_modeled_and_manual_h_adders_use_same_monthly_usage_and_supply_cost(self):
        p=prices(2);c=econ.curve(p,p.trade_date.max(),SELECTIONS,pd.date_range('2026-12-01',periods=2,freq='MS'))
        first,target=econ.calculate(c,[100,300],.4,.1,.3,capital_hurdle=0,pricing_mode='NYMEX + adder',transport_pass_through=True)
        self.assertAlmostEqual(target,.55)
        model,_=econ.calculate(c,[100,300],.4,.1,.3,capital_hurdle=0,pricing_mode='NYMEX + adder',transport_pass_through=True,fixed_quote=target)
        self.assertAlmostEqual(model.gross_profit.sum()/400,.3)
        manual,_=econ.calculate(c,[100,300],.4,.1,.3,capital_hurdle=0,pricing_mode='NYMEX + adder',transport_pass_through=True,fixed_quote=.9)
        s=econ.deal_summary(manual)
        self.assertAlmostEqual(s['nymex_h_reference'],3.75)
        self.assertAlmostEqual(s['modeled_nymex_adder'],.55)
        self.assertAlmostEqual(s['offer_nymex_adder'],.9)
        self.assertAlmostEqual(s['contract_price'],5.05)
        np.testing.assert_allclose(manual.supply_rate,[2.7,4.3])
        np.testing.assert_allclose(manual.customer_rate,[4.3,5.3])

    def test_reference_equivalents_do_not_change_fixed_or_regional_quotes(self):
        p=prices(2);c=econ.curve(p,p.trade_date.max(),SELECTIONS,pd.date_range('2026-12-01',periods=2,freq='MS'))
        fixed,_=econ.calculate(c,[100,300],.4,.1,.3,capital_hurdle=0,fixed_quote=4.8,transport_pass_through=True)
        self.assertAlmostEqual(econ.deal_summary(fixed)['offer_nymex_adder'],1.05)
        regional,_=econ.calculate(c,[100,300],.4,.1,.3,capital_hurdle=0,fixed_quote=.8,pricing_mode='Index + adder',transport_pass_through=True)
        self.assertAlmostEqual(econ.deal_summary(regional)['offer_nymex_adder'],.95)
        np.testing.assert_allclose(regional.customer_rate,regional.market_rate+1.2)

    def test_missing_duplicate_old_or_invalid_h_never_falls_back(self):
        p=prices(2);months=pd.date_range('2026-12-01',periods=2,freq='MS')
        old=prices(2,'2026-11-13')
        missing=pd.concat([p.loc[~p.contract.eq('H')],old],ignore_index=True)
        self.assertTrue(henry_hub(missing,pd.Timestamp('2026-11-16'),months).nymex_h_settlement.isna().all())
        for bad in [missing,pd.concat([p,p.loc[p.contract.eq('H')]],ignore_index=True)]:
            h=henry_hub(bad,pd.Timestamp('2026-11-16'),months)
            c=pd.DataFrame({'component_1':[2.7,4.3]},index=months).join(h)
            with self.assertRaisesRegex(ValueError,'requires one valid product H'):econ.calculate(c,[100,100],.4,.1,.3,pricing_mode='NYMEX + adder')
            # A reference gap alone cannot block a previously supported fixed quote.
            r,_=econ.calculate(c,[100,100],.4,.1,.3)
            self.assertEqual(econ.deal_summary(r)['nymex_reference_status'].split(':')[0],'Unavailable')

    def test_standard_term_manual_h_offer_and_client_alternatives_preserve_reviewed_rates(self):
        from client_term_proposal import schedules
        p=prices();usage=pd.DataFrame({'month':pd.date_range('2026-12-01',periods=12,freq='MS'),'usage_mmbtu':[100.]*12})
        common=(p,p.trade_date.max(),SELECTIONS,pd.Timestamp('2026-12-01'),usage,{i:100. for i in range(1,13)},.4,.1,.3,0.,0.,30.,.2,0.,0.)
        options=dict(pricing_mode='NYMEX + adder',customer_adder=.38,transport_pass_through=True)
        offers=econ.standard_offers(*common,overrides={12:.9,24:1.1},**options)
        self.assertTrue(offers.status.eq('Available').all())
        self.assertEqual(offers.set_index('term_months').loc[12,'offer_price'],.9)
        tables=schedules(econ.standard_offers,common,'NYMEX + adder',offers,options)
        self.assertEqual(tuple(tables),MODES)
        self.assertEqual(tables['NYMEX + adder']['rows'][0]['offer_price'],.9)
        self.assertNotIn('contract_profit',tables['NYMEX + adder']['rows'][0])
        self.assertNotIn('npv',tables['NYMEX + adder']['rows'][0])

    def test_newest_dated_source_beats_generic_alias_and_preserves_manual_inputs(self):
        files=[Path('icecleared_latest.xlsx'),Path('icecleared_2026_11_13.xlsx'),Path('icecleared_2026_11_16.xlsx')]
        self.assertEqual(ordered(files)[0].name,'icecleared_2026_11_16.xlsx')
        state={'wk_custom_price':.9,'offer_override_12':.8,'wk_customer':'Synthetic customer','wk_asof':'old','_saved_trade_date':'old','_current_snapshot':{'usage':[1,2]}}
        select_newest(state,'icecleared_2026_11_16.xlsx')
        self.assertEqual(state['wk_custom_price'],.9);self.assertEqual(state['offer_override_12'],.8)
        self.assertEqual(state['_pricing_carry'],{'usage':[1,2]});self.assertTrue(state['_restore_pricing'])
        self.assertNotIn('wk_asof',state);self.assertIsNone(state['_saved_trade_date'])

    def test_full_app_can_load_newest_and_switch_to_nymex(self):
        from streamlit.testing.v1 import AppTest
        with tempfile.TemporaryDirectory() as folder:
            directory=Path(folder)/'input';directory.mkdir()
            prices().to_csv(directory/'icecleared_2026_11_16.csv',index=False)
            prices(asof='2026-11-13').to_csv(directory/'icecleared_2026_11_13.csv',index=False)
            script=f'''
from pathlib import Path
from unittest.mock import patch
import deal_economics as econ
import streamlit as st
def saved(repo,token,name,snapshot,*args):
    st.session_state['_test_saved']=snapshot
    return 'workstreams/test.json','synthetic-sha',dict(snapshot,name=name,saved_at='test')
internal_pdf=econ.build_summary_pdf
client_pdf=econ.build_client_proposal_pdf
def capture_internal(*args):
    data=internal_pdf(*args);st.session_state['_test_internal_pdf']=data;return data
def capture_client(*args):
    data=client_pdf(*args);st.session_state['_test_client_pdf']=data;return data
with patch.object(econ,'ROOT',Path({folder!r})),patch('app_workflow_help.begin'):
    with patch.object(econ,'save_workstream',side_effect=saved),patch.object(econ,'build_summary_pdf',side_effect=capture_internal),patch.object(econ,'build_client_proposal_pdf',side_effect=capture_client):econ.run_economics_app()
'''
            app=AppTest.from_string(script,default_timeout=30)
            app.secrets['WORKSTREAM_REPO']='synthetic/private'
            app.secrets['WORKSTREAM_GITHUB_TOKEN']='synthetic-not-a-token'
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(app.selectbox(key='wk_pricing_file').value,'icecleared_2026_11_16.csv')
            app.button(key='wk_use_newest_pricing').click().run()
            self.assertFalse(app.exception)
            app.radio(key='wk_pricing_mode').set_value('NYMEX + adder').run()
            self.assertFalse(app.exception)
            self.assertTrue(any(x.label=='Offered NYMEX adder ($/MMBtu)' for x in app.metric))
            self.assertFalse(app.error)
            app.number_input(key='wk_client_current_rate').set_value(5.).run()
            from pypdf import PdfReader
            for key in ('_test_internal_pdf','_test_client_pdf'):
                pdf=PdfReader(BytesIO(app.session_state[key]))
                self.assertEqual(len(pdf.pages),4)
                self.assertIn('NYMEX',pdf.pages[3].extract_text())
            # Both explicit choices survive Save, Open, reruns and pricing-mode changes.
            for choice in (True,False):
                app.checkbox(key='wk_transport_pass_through').set_value(choice).run()
                app.text_input(key='ws_name').set_value('Synthetic saved deal').run()
                next(x for x in app.button if x.label=='Save current workstream').click().run()
                self.assertFalse(app.exception)
                saved=app.session_state['_test_saved']
                self.assertIs(saved['widgets']['wk_transport_pass_through'],choice)
                self.assertIs(saved['pricing_structure']['transport_pass_through'],choice)
                app.session_state['_pending_workstream']=saved
                app.run()
                self.assertIs(app.checkbox(key='wk_transport_pass_through').value,choice)
                app.radio(key='wk_pricing_mode').set_value('Fixed price').run()
                self.assertIs(app.checkbox(key='wk_transport_pass_through').value,choice)
                app.radio(key='wk_pricing_mode').set_value('NYMEX + adder').run()
                self.assertIs(app.checkbox(key='wk_transport_pass_through').value,choice)
            legacy=dict(saved,widgets={k:v for k,v in saved['widgets'].items() if k!='wk_transport_pass_through'})
            legacy['pricing_structure']=dict(saved['pricing_structure'],transport_pass_through=True)
            app.session_state['_pending_workstream']=legacy
            app.run()
            self.assertTrue(app.checkbox(key='wk_transport_pass_through').value)


if __name__=='__main__':unittest.main()
