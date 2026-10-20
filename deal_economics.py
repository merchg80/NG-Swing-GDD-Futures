import os
from pathlib import Path
import pandas as pd
import streamlit as st
from pathlib import Path
import pandas as pd
import numpy as np

REQUIRED = ['trade_date','hub','product','strip','contract','settlement_price']

def normalize_prices(df):
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(' ','_') for c in df.columns]
    missing = set(REQUIRED)-set(df.columns)
    if missing: raise ValueError('Missing pricing columns: '+', '.join(sorted(missing)))
    df = df[REQUIRED].copy()
    for col in ['trade_date','strip']:
        df[col] = pd.to_datetime(df[col].astype(str), errors='coerce', format='mixed').dt.normalize()
    df['settlement_price'] = pd.to_numeric(df['settlement_price'], errors='coerce')
    df = df.dropna(subset=REQUIRED)
    for col in ['contract','hub','product']: df[col] = df[col].astype(str).str.strip()
    if df.empty: raise ValueError('No valid pricing records found.')
    return df

def read_prices(source):
    name = str(getattr(source,'name',source)).lower().split('?')[0]
    if hasattr(source,'seek'): source.seek(0)
    if name.endswith('.parquet'): df = pd.read_parquet(source)
    elif name.endswith('.csv'): df = pd.read_csv(source)
    else: df = pd.read_excel(source, sheet_name=0, usecols='A:K')
    return normalize_prices(df)

def workbook_defaults(source):
    import openpyxl
    if hasattr(source,'seek'): source.seek(0)
    s = openpyxl.load_workbook(source,data_only=True,read_only=True)['Pricer']
    usage = {int(s.cell(r,32).value):float(s.cell(r,43).value) for r in range(17,29)}
    return dict(customer=s['AN1'].value,start=str(s['N3'].value)[:10],transport=float(s['AH1'].value),broker=float(s['AJ1'].value),margin=float(s['AG40'].value),usage=usage)

def curve(prices, asof, selections, months):
    snapshot = prices[prices.trade_date.eq(pd.Timestamp(asof))]
    out = pd.DataFrame(index=pd.DatetimeIndex(months,name='month'))
    for i,(hub,product,contract) in enumerate(selections):
        x = snapshot[(snapshot.hub==hub)&(snapshot['product']==product)&(snapshot.contract==contract)]
        if x.strip.duplicated().any(): raise ValueError('Duplicate settlements for selected product/date; clean the source before pricing.')
        out[f'component_{i+1}'] = x.set_index('strip').settlement_price.reindex(out.index)
    return out

def calculate(pricing, usage, transport, broker, margin, fuel=0., fixed_quote=None, capital_hurdle=0.08, financing_days=30.):
    x=pricing.copy()
    components=[c for c in x if c.startswith('component_')]
    if x[components].isna().any().any(): raise ValueError('Missing settlement months; incomplete terms cannot be quoted.')
    x['usage_mmbtu']=np.asarray(usage,dtype=float)
    if not np.isfinite(x.usage_mmbtu).all() or (x.usage_mmbtu<0).any() or x.usage_mmbtu.sum()<=0: raise ValueError('Usage must be finite, nonnegative and total above zero.')
    if not 0<=fuel<1: raise ValueError('Fuel must be between zero and 100%.')
    x['market_rate']=x[components].sum(axis=1)
    x['supply_rate']=x.market_rate/(1-fuel)
    x['transport_rate']=transport
    x['broker_rate']=broker
    x['landed_rate']=x.supply_rate+transport+broker
    if capital_hurdle<0 or financing_days<0: raise ValueError('Capital hurdle and financing days must be nonnegative.')
    x['capital_base_rate']=x.landed_rate.clip(lower=0.)
    x['capital_rate']=x.capital_base_rate*capital_hurdle*financing_days/365.
    x['all_in_cost_rate']=x.landed_rate+x.capital_rate
    target=float(np.average(x.all_in_cost_rate,weights=x.usage_mmbtu)+margin)
    quote=target if fixed_quote is None else float(fixed_quote)
    x['customer_rate']=quote
    x['revenue']=quote*x.usage_mmbtu
    for name,col in [('supply_cost','supply_rate'),('transport_cost','transport_rate'),('broker_cost','broker_rate')]: x[name]=x[col]*x.usage_mmbtu
    x['capital_cost']=x.capital_rate*x.usage_mmbtu
    x['gross_profit_before_capital']=x.revenue-x.supply_cost-x.transport_cost-x.broker_cost
    x['gross_profit']=x.gross_profit_before_capital-x.capital_cost
    x['margin_per_mmbtu']=x.customer_rate-x.all_in_cost_rate
    return x,target

def deal_summary(result, annual_discount=0.10, initial_investment=0., terminal_recovery=0.):
    """Monthly cash proxy: GP at month-end; investment at delivery start."""
    if annual_discount<=-1 or initial_investment<0 or terminal_recovery<0:
        raise ValueError('Invalid discount rate or investment inputs.')
    n=len(result)
    flows=result.gross_profit.to_numpy(dtype=float).copy()
    flows[-1]+=terminal_recovery
    monthly_discount=(1+annual_discount)**(1/12)-1
    npv=-initial_investment+sum(cf/(1+monthly_discount)**(i+1) for i,cf in enumerate(flows))
    cash=np.r_[-initial_investment,flows]
    irr=None
    status='Requires an initial investment and a cash-flow sign change'
    if initial_investment>0 and (cash>0).any():
        roots=np.polynomial.polynomial.polyroots(cash)
        candidates=[float(r.real) for r in roots if abs(r.imag)<1e-7 and r.real>0]
        if len(candidates)==1:
            annual=(1/candidates[0])**12-1
            if np.isfinite(annual): irr=annual;status='Annual effective IRR'
        elif len(candidates)>1: status='Multiple IRRs; no unique return'
        else: status='No valid IRR'
    revenue=float(result.revenue.sum());volume=float(result.usage_mmbtu.sum());gp=float(result.gross_profit.sum())
    return dict(annual_revenue=revenue*12/n, volume=volume, annual_volume=volume*12/n,
                contract_price=float(result.customer_rate.iloc[0]),npv=float(npv),irr=irr,irr_status=status,
                spread=gp/volume,gross_margin=gp/revenue if revenue>0 else None)


ROOT=Path(__file__).resolve().parent
st.set_page_config(page_title='REM | Deal Economics',layout='wide',initial_sidebar_state='expanded')
st.title('Riverbend Energy Management')
st.caption('Natural gas deal economics • Monthly volume-weighted pricing')
st.info('Build REM-INPUT-1020 • Pricing comes directly from the GitHub input folder. No pricing upload needed.')
summary_panel=st.container()

@st.cache_data(ttl=300)
def load_location(path, modified_ns):
    return read_prices(path)

# Streamlit checks out this GitHub repository before running the app.
# Read its input folder directly: no upload or separate GitHub credentials needed.
input_candidates=[ROOT/'input',ROOT.parent/'input']
INPUT_DIR=next((d for d in input_candidates if d.is_dir()),ROOT/'input')
defaults=dict(customer='New customer', start=str(pd.Timestamp.today().to_period('M').to_timestamp().date()), transport=.35, broker=.30, margin=.18, usage={m:1700. for m in range(1,13)})
prices=None
with st.expander('1. Select ICE pricing file and contract dates', expanded=True):
    st.caption('Pricing source: merchg80/NG-Swing-GDD-Futures • main branch • input folder')
    if st.button('Refresh pricing files'):
        load_location.clear();st.rerun()
    if not INPUT_DIR.is_dir():
        st.error('The repository input folder was not found. Deploy this app from merchg80/NG-Swing-GDD-Futures, branch main, with streamlit_app.py at the repository root.')
        st.code(f'App location: {ROOT}\nExpected data folder: {ROOT / "input"}')
        st.stop()
    available=sorted([p for p in INPUT_DIR.rglob('*') if p.is_file() and p.suffix.lower() in {'.xlsx','.csv','.parquet'}],key=lambda p: ('ice' not in p.name.lower(),p.name.lower()))
    if not available:
        st.warning('The input folder contains no supported pricing files (.xlsx, .csv or .parquet).')
        st.stop()
    selected_path=st.selectbox('ICE Cleared Gas pricing file',available,format_func=lambda p:str(p.relative_to(INPUT_DIR)))
    st.caption(f'Selected repository file: input/{selected_path.relative_to(INPUT_DIR)}')
    stamp=f'{selected_path}:{selected_path.stat().st_mtime_ns}'
    if st.button('Load selected pricing file',type='primary'):
        st.session_state['loaded_pricing_stamp']=stamp
    if st.session_state.get('loaded_pricing_stamp')!=stamp:
        st.info('Choose an existing ICE file above, then click Load selected pricing file. No upload is needed.')
        st.stop()
    try:
        with st.spinner(f'Reading {selected_path.name} — large ICE files may take time…'):
            prices=load_location(str(selected_path),selected_path.stat().st_mtime_ns)
        st.success(f'Loaded {len(prices):,} settlement rows • latest trade date {prices.trade_date.max():%Y-%m-%d}')
    except Exception as e:
        st.error(f'Could not read input/{selected_path.relative_to(INPUT_DIR)}: {e}')
        st.caption('Select an ICE settlement file with trade date, hub, product, strip, contract and settlement price columns. A daily Date-by-hub master is not a monthly forward curve.')
        st.stop()
    dates=sorted(prices.trade_date.unique(),reverse=True)
    asof=st.selectbox('Settlement trade date',dates,format_func=lambda d:pd.Timestamp(d).strftime('%b %d, %Y'))
    st.caption('All components use this exact trade date. Missing months block a quote.')
    customer=st.text_input('Customer',value=defaults['customer'])
    start_entered=st.date_input('Contract start date',value=(pd.Timestamp(asof).to_period('M')+1).to_timestamp().date())
    start=start_entered.replace(day=1)
    if start_entered.day!=1: st.caption('Monthly pricing uses the start month. Enter first-month usage for the remaining delivery days; usage is not automatically prorated.')
    term=st.slider('Term (months)',min_value=1,max_value=48,value=12,step=1)

st.subheader('Market components')
st.caption('Choose products by name. Additive mode follows the workbook: Henry Hub + HSC basis + HSC index adjustment. Outright mode uses one full-price product.')
kind=st.radio('Curve construction',['Additive components','Outright fixed price'],horizontal=True)
options=list(prices[['hub','product','contract']].drop_duplicates().itertuples(index=False,name=None))
selections=[]
cols=st.columns(3 if kind=='Additive components' else 1)
for i,c in enumerate(cols):
    preferred=['H','HXS','SHS'][i] if kind=='Additive components' else 'H'
    idx=next((j for j,x in enumerate(options) if x[2]==preferred),0)
    selections.append(c.selectbox(f'Component {i+1}',options,index=idx,format_func=lambda x:f'{x[0]} | {x[1]} ({x[2]})'))
if len(set(selections))!=len(selections): st.error('Select distinct components.');st.stop()
st.subheader('Deal assumptions')
a,b,c,d=st.columns(4)
transport=a.number_input('Upstream transport ($/delivered MMBtu)',min_value=0.,value=defaults['transport'],step=.01,format='%.4f')
broker=b.number_input('Brokerage ($/delivered MMBtu)',min_value=0.,value=defaults['broker'],step=.01,format='%.4f')
margin=c.number_input('Target gross margin ($/MMBtu)',value=defaults['margin'],step=.01,format='%.4f')
fuel=d.number_input('Supply fuel retention (%)',min_value=0.,max_value=99.,value=0.,step=.1)/100
st.caption('Margin is after supply, transport and brokerage, after the capital charge, before overhead, other credit fees and taxes. Transport and brokerage apply to delivered usage. Fuel gross-up applies to gas supply only.')
st.subheader('Cost of capital')
cap1,cap2=st.columns(2)
capital_hurdle=cap1.number_input('Annual capital hurdle (%)',min_value=0.,value=8.,step=.5)/100
financing_days=cap2.number_input('Days of costs financed',min_value=0.,value=30.,step=1.)
st.caption('Capital charge per delivered MMBtu = positive supply + transport + brokerage cost per MMBtu × annual hurdle × financing days ÷ 365. The charge is included in the target offer and break-even price, and deducted from modeled profit. Financing days represent the cash collection gap. Negative total costs have a zero capital base.')
with st.expander('NPV and IRR assumptions'):
    a,b,c=st.columns(3)
    discount=a.number_input('Annual discount rate (%)',min_value=0.,value=10.,step=.5)/100
    investment=b.number_input('Initial cash investment ($)',min_value=0.,value=0.,step=1000.)
    recovery=c.number_input('Cash recovered at contract end ($)',min_value=0.,value=0.,step=1000.)
    st.caption('NPV is measured at contract start. Cash flows use month-end profit after the modeled capital charge, before overhead and taxes, less initial investment, plus terminal recovery. Enter posted collateral as initial cash and its expected return as terminal recovery. IRR is annualized and requires an initial cash outflow.')
months=pd.date_range(start,periods=term,freq='MS')
base=pd.DataFrame({'month':months,'usage_mmbtu':[defaults['usage'].get(m.month,0.) for m in months]})
st.subheader('Customer monthly usage')
st.caption('Enter total delivered MMBtu for each month. The attached workbook starts at 1,700 MMBtu/month. Values are estimates, not daily quantities.')
usage_file=st.file_uploader('Load a saved usage schedule (optional)',type=['csv'])
if usage_file is not None:
    try:
        loaded=pd.read_csv(usage_file)
        loaded['month']=pd.to_datetime(loaded['month']).dt.to_period('M').dt.to_timestamp()
        if loaded.month.duplicated().any(): raise ValueError('Duplicate usage months')
        base['usage_mmbtu']=base.month.map(loaded.set_index('month')['usage_mmbtu'])
        if base.usage_mmbtu.isna().any(): raise ValueError('Schedule must cover every selected month')
    except Exception as e: st.error(f'Invalid usage schedule: {e}');st.stop()
usage=st.data_editor(base,hide_index=True,disabled=['month'],column_config={'usage_mmbtu':st.column_config.NumberColumn('Usage (MMBtu/month)',min_value=0.,required=True)},key=f'usage_{start}_{term}',use_container_width=True)
manual=st.checkbox('Evaluate a proposed customer fixed price')
quote=st.number_input('Proposed customer price ($/MMBtu)',value=4.,step=.01) if manual else None
try:
    p=curve(prices,asof,selections,months)
    if p.isna().any().any():
        st.error('Pricing is missing for these delivery months/components. Select another trade date, product or term.')
        st.dataframe(p[p.isna().any(axis=1)]);st.stop()
    result,target=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,quote,capital_hurdle,financing_days)
except ValueError as e: st.error(str(e));st.stop()
summary=deal_summary(result,discount,investment,recovery)
with summary_panel:
    st.subheader('Deal summary')
    st.caption(f'{customer} • Start {start_entered:%b %d, %Y} • {term} month term • Settlements {pd.Timestamp(asof):%b %d, %Y}')
    top=st.columns(4)
    top[0].metric('Annual revenue (annualized)',f"${summary['annual_revenue']:,.0f}")
    top[1].metric('Volume sold (contract MMBtu)',f"{summary['volume']:,.0f}")
    top[2].metric('Contract price ($/MMBtu)',f"${summary['contract_price']:,.4f}")
    top[3].metric('NPV at contract start',f"${summary['npv']:,.0f}")
    top[3].metric('Break-even price ($/MMBtu)',f'${target-margin:,.4f}')
    bottom=st.columns(3)
    bottom[0].metric('IRR (annualized)',f"{summary['irr']:.2%}" if summary['irr'] is not None else 'N/A')
    bottom[1].metric('MMBtu spread ($/MMBtu)',f"${summary['spread']:,.4f}")
    bottom[2].metric('Gross margin (%)',f"{summary['gross_margin']:.2%}" if summary['gross_margin'] is not None else 'N/A')
    st.caption(f"Capital hurdle: {capital_hurdle:.1%} • Financing days: {financing_days:g} • Weighted capital charge: ${(result.capital_cost.sum()/result.usage_mmbtu.sum()):.4f}/MMBtu • Contract capital charge: ${result.capital_cost.sum():,.0f}")
    st.caption('Annual revenue = total contract revenue × 12 ÷ term months. Spread = gross profit ÷ delivered volume. Gross margin = gross profit ÷ revenue. '+summary['irr_status']+'.')
    st.download_button('Download deal summary CSV',pd.DataFrame([dict(customer=customer,start_date=str(start_entered),term_months=term,trade_date=str(pd.Timestamp(asof).date()),capital_hurdle=capital_hurdle,financing_days=financing_days,capital_charge_per_mmbtu=result.capital_cost.sum()/result.usage_mmbtu.sum(),annual_discount_rate=discount,initial_investment=investment,terminal_recovery=recovery,**summary)]).to_csv(index=False).encode(),'deal_summary.csv','text/csv')
vol=result.usage_mmbtu.sum();gp=result.gross_profit.sum()
a,b,c,d=st.columns(4)
a.metric('Target customer fixed price',f'${target:,.4f}/MMBtu')
b.metric('Total contract volume',f'{vol:,.0f} MMBtu')
c.metric('Contract gross profit',f'${gp:,.0f}')
d.metric('Realized gross margin',f'${gp/vol:,.4f}/MMBtu')
st.caption(f'Customer: {customer} • Settlements: {pd.Timestamp(asof):%Y-%m-%d} • Annualized gross profit: ${gp*12/term:,.0f}')
st.subheader('Offer comparison')
st.caption('Offer uses the proposed customer price, or target price when no proposed price is entered. Costs and volume remain constant. Break-even includes the capital charge and excludes overhead and taxes.')
scenario_rows=[]
scenario_monthly=pd.DataFrame(index=result.index)
for label,reduction in [('Offer',0.),('Offer less $0.10',.10),('Offer less $0.20',.20)]:
    offer_price=float(result.customer_rate.iloc[0])-reduction
    sr,_=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,offer_price,capital_hurdle,financing_days)
    ss=deal_summary(sr,discount,investment,recovery)
    scenario_rows.append({'Scenario':label,'Offer $/MMBtu':offer_price,'Break-even $/MMBtu':target-margin,'Spread $/MMBtu':ss['spread'],'Gross margin %':ss['gross_margin']*100 if ss['gross_margin'] is not None else None,'Average monthly GP $':sr.gross_profit.sum()/term,'Annualized GP $':sr.gross_profit.sum()*12/term,'Contract GP $':sr.gross_profit.sum(),'NPV $':ss['npv'],'Annualized IRR %':ss['irr']*100 if ss['irr'] is not None else None})
    scenario_monthly[label]=sr.gross_profit
scenario_table=pd.DataFrame(scenario_rows)
st.dataframe(scenario_table,hide_index=True,use_container_width=True)
st.caption('Monthly and annualized averages use the selected contract term. The schedule below shows actual monthly profits. NPV includes initial investment and recovery; break-even price is the volume-weighted cost including the capital charge.')
with st.expander('Monthly offer economics',expanded=True):
    st.line_chart(scenario_monthly)
    st.dataframe(scenario_monthly.reset_index(),hide_index=True,use_container_width=True)
st.download_button('Download offer comparison CSV',scenario_table.to_csv(index=False).encode(),'offer_comparison.csv','text/csv')
st.download_button('Download monthly offer economics CSV',scenario_monthly.reset_index().to_csv(index=False).encode(),'monthly_offer_economics.csv','text/csv')
t1,t2,t3=st.tabs(['Monthly economics','Term comparisons','Method & source'])
with t1:
    st.line_chart(result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']])
    st.bar_chart(result[['gross_profit']])
    st.dataframe(result.reset_index(),hide_index=True,use_container_width=True)
    totals=result[['usage_mmbtu','revenue','supply_cost','transport_cost','broker_cost','capital_cost','gross_profit']].sum()
    st.write('Contract totals',totals.to_frame('Total').T)
    st.write('Selected-strip averages',pd.DataFrame({'Simple average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mean(),'Volume-weighted average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mul(result.usage_mmbtu,axis=0).sum()/vol}))
with t2:
    st.caption('Compares standard terms and your selected term within the entered usage schedule. Longer terms require extending the schedule above. Each term is priced separately to the same target margin.')
    rows=[]
    for n in sorted(set([1,3,6,12,18,24,36,48,term])):
        if n>term or usage.usage_mmbtu.iloc[:n].sum()<=0: continue
        r,q=calculate(p.iloc[:n],usage.usage_mmbtu.iloc[:n],transport,broker,margin,fuel,capital_hurdle=capital_hurdle,financing_days=financing_days)
        rows.append({'Months':n,'Volume MMBtu':r.usage_mmbtu.sum(),'Target fixed price':q,'Gross profit':r.gross_profit.sum(),'Annualized GP':r.gross_profit.sum()*12/n})
    st.dataframe(pd.DataFrame(rows),hide_index=True)
with t3:
    st.write('Supply = sum of selected components ÷ (1 − fuel retention). Landed cost = supply + upstream transport + brokerage. Target fixed price = volume-weighted landed cost + target margin. Capital charge = max(landed cost, 0) × annual hurdle × financing days / 365. Target fixed price = weighted landed cost + weighted capital charge + target margin. Monthly modeled profit = (customer fixed price − landed cost − capital charge) × usage.')
    st.write('Uses actual entered monthly volumes, calendar delivery months and a single settlement snapshot. No interpolation, silent zero filling or fallback to older settlements. Daily swing observations are not converted to forward prices.')
    st.write('The capital charge is a modeled financing expense, not a full working-capital schedule. NPV separately discounts profit after that charge; initial investment and terminal recovery are additional cash flows and should not duplicate the financed operating costs. The workbook’s NPV/carry formulas and daily-volume approximation are not reproduced. This model quotes a level fixed price over the selected term. It excludes utility charges, imbalance exposure, credit/collateral financing and overhead unless included in your rates.')
    st.dataframe(pd.DataFrame(selections,columns=['Hub','Product','Contract']),hide_index=True)
export=result.reset_index(); export.insert(0,'customer',customer);export.insert(1,'trade_date',pd.Timestamp(asof));export.insert(2,'contract_start_date',str(start_entered));export.insert(3,'term_months',term);export.insert(2,'components',' + '.join(x[2] for x in selections))
st.download_button('Download monthly economics CSV',export.to_csv(index=False).encode(),'rem_deal_economics.csv','text/csv')
st.download_button('Download usage schedule CSV',usage.to_csv(index=False).encode(),'customer_usage.csv','text/csv')
