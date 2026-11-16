"""Customer-facing Texas fixed/index term worksheets. Internal returns stay private."""
import math
import pandas as pd

TERMS=(12,18,24,36,48)


def schedules(offer_fn,common,active_mode,active_offers,deal_options,internal=False,**credit_options):
    """Retain reviewed quotes for active structure; other structure uses model targets."""
    out={}
    for mode in ('Fixed price','Index + adder'):
        if mode==active_mode:
            frame=active_offers.copy()
            basis='Uses the reviewed standard-term offers in the app.'
        else:
            options=dict(deal_options,pricing_mode=mode)
            frame=offer_fn(*common,**credit_options,**options)
            targets={int(r.term_months):float(r.target_price) for r in frame.itertuples() if r.status=='Available'}
            frame=offer_fn(*common,overrides=targets,**credit_options,**options)
            basis='Alternative pricing structure calculated at the shared target margin.'
        if frame.term_months.duplicated().any():raise ValueError('Duplicate proposal terms.')
        safe=[]
        for term in TERMS:
            hits=frame.loc[frame.term_months.eq(term)]
            if hits.empty:
                safe.append(dict(term_months=term,status='Unavailable',offer_price=None,projected_all_in_price=None,contract_volume=None));continue
            row=hits.iloc[0]
            fields={k:row.get(k) for k in ('term_months','offer_price','projected_all_in_price','contract_volume')}
            valid=row.status=='Available' and all(pd.notna(fields[k]) and math.isfinite(float(fields[k])) for k in ('offer_price','projected_all_in_price','contract_volume'))
            if internal:
                fields.update({k:(float(row[k]) if pd.notna(row.get(k)) else None) for k in ['break_even','spread','gross_margin','annual_revenue','contract_profit','npv','upfront_credit','ft_credit','supply_credit','credit_recovery']})
            safe.append(dict(fields,status='Available' if valid else 'Unavailable - incomplete pricing'))
        out[mode]=dict(rows=safe,basis=basis)
    return out


def pages(meta,styles,logo_bytes):
    """Return two explicitly separated portrait worksheets using summary typography."""
    from io import BytesIO
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph,Table,TableStyle,Spacer,PageBreak,Image
    green=colors.HexColor('#005448');pale=colors.HexColor('#EDF5F1');line=colors.HexColor('#D8E3DE')
    def p(v,style='centerbody'):return Paragraph(escape(str(v)).replace('\n','<br/>'),styles[style])
    def table(rows,widths,header=False):
        t=Table(rows,colWidths=widths,hAlign='CENTER')
        cmds=[('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,0),(-1,-1),'CENTER'),('LEFTPADDING',(0,0),(-1,-1),9),('RIGHTPADDING',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),13),('BOTTOMPADDING',(0,0),(-1,-1),13)]
        if header:
            cmds += [('BACKGROUND',(0,0),(-1,0),green),('ROWBACKGROUNDS',(0,1),(-1,-1),[pale,colors.white]),('LINEBELOW',(0,0),(-1,-1),.4,line)]
        else:cmds += [('BACKGROUND',(0,0),(-1,-1),pale)]
        t.setStyle(TableStyle(cmds));return t
    out=[]
    for mode,title in [('Fixed price','Fixed-price options'),('Index + adder','Index-adder options')]:
        data=meta['term_worksheets'][mode]
        out.append(PageBreak())
        logo=Image(BytesIO(logo_bytes),width=250,height=250*402/2096);logo.hAlign='CENTER'
        out.extend([logo,Spacer(1,20),p(title,'title'),Spacer(1,5),p('12, 18, 24, 36 and 48 months','label'),Spacer(1,20)])
        customer=p(meta['customer'],'centerbold')
        info=p((meta.get('utility') or 'Utility not specified')+'\n'+(meta.get('client_location') or 'Location not specified'))
        out.extend([table([[customer,info]],[270,270]),Spacer(1,15)])
        indexed=mode=='Index + adder'
        if indexed:
            out.extend([p('Monthly index: '+meta.get('client_index_name','Selected monthly index'),'centerbold'),Spacer(1,7)])
        headers=['TERM','DELIVERY PERIOD','INDEX ADDER\n$/MMBtu' if indexed else 'FIXED RATE\n$/MMBtu','PROJECTED ALL-IN\n$/MMBtu' if indexed else 'DELIVERED RATE\n$/MMBtu','CONTRACT USAGE\nMMBtu']
        rows=[[p(h,'white') for h in headers]]
        for r in data['rows']:
            months=int(r['term_months']);start=pd.Timestamp(meta['start']);end=(start.to_period('M')+months-1).end_time
            period=f'{start:%b %d, %Y}\nto {end:%b %d, %Y}'
            if r['status']!='Available':values=[f'{months} months',period,'Unavailable','Unavailable','Unavailable']
            else:
                offer=float(r['offer_price']);rate=(('+' if offer>=0 else '-')+f'${abs(offer):,.4f}') if indexed else f'${offer:,.4f}'
                values=[f'{months} months',period,rate,f"${r['projected_all_in_price']:,.4f}",f"{r['contract_volume']:,.0f}"]
            rows.append([p(v,'centerbold' if j==2 else 'centerbody') for j,v in enumerate(values)])
        out.extend([table(rows,[58,135,107,120,120],True),Spacer(1,22)])
        out.extend([p('Transportation','heading'),Spacer(1,7),p(meta['proposed_transport_text']),Spacer(1,18)])
        note=('The quoted adder is applied to the published monthly index. Projected all-in rates use the current forward curve and volume-weighted usage for each term; they are not fixed prices.' if indexed else 'The fixed rate applies over the indicated delivery term. The delivered rate includes the transportation treatment shown above.')
        note+=' Exact entered usage is used where available; longer terms repeat calendar-month usage assumptions. Utility distribution charges, taxes and other fees are excluded unless expressly included. Unavailable terms require complete pricing before quoting.'
        out.extend([p('Offer basis','bold'),Spacer(1,5),p(note,'small'),Spacer(1,12),p(f"Pricing as of {meta['asof']} | Offer due: {meta.get('due_date') or 'Not specified'}",'small'),p('Subject to final agreement and credit approval.','small')])
    return out
