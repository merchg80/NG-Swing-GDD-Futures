"""Confidential three-page Texas opportunity package; never used by client renderer."""
from io import BytesIO
from xml.sax.saxutils import escape
import pandas as pd


def build(meta,summary,result,scenarios,logo_bytes):
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import BaseDocTemplate,Frame,PageTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,Image
    import reportlab
    from pathlib import Path
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    fonts=Path(reportlab.__file__).parent/'fonts'
    pdfmetrics.registerFont(TTFont('Opportunity',str(fonts/'Vera.ttf')))
    pdfmetrics.registerFont(TTFont('OpportunityBold',str(fonts/'VeraBd.ttf')))
    green=colors.HexColor('#005448');red=colors.HexColor('#862131');pale=colors.HexColor('#EDF5F1');line=colors.HexColor('#D8E3DE')
    styles={
        'body':ParagraphStyle('InternalBody',fontName='Opportunity',fontSize=8.3,leading=11,textColor=green),
        'small':ParagraphStyle('InternalSmall',fontName='Opportunity',fontSize=7,leading=9,textColor=green),
        'bold':ParagraphStyle('InternalBold',fontName='OpportunityBold',fontSize=8.3,leading=11,textColor=green),
        'title':ParagraphStyle('InternalTitle',fontName='OpportunityBold',fontSize=19,leading=23,textColor=green,alignment=1),
        'white':ParagraphStyle('InternalWhite',fontName='OpportunityBold',fontSize=8,leading=11,textColor=colors.white,alignment=1),
        'money':ParagraphStyle('InternalMoney',fontName='OpportunityBold',fontSize=14,leading=20,textColor=green,alignment=1),
    }
    def p(v,style='body'):return Paragraph(escape(str(v)).replace('\n','<br/>'),styles[style])
    def money(v):return f'${float(v):,.0f}' if pd.notna(v) else 'Unavailable'
    def rate(v):return f'${float(v):,.4f}' if pd.notna(v) else 'Unavailable'
    def percent(v):return f'{float(v):.2%}' if pd.notna(v) else 'Unavailable'
    def table(rows,widths,header=False,pad=7):
        t=Table(rows,colWidths=widths,hAlign='CENTER')
        cmds=[('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),pad),('BOTTOMPADDING',(0,0),(-1,-1),pad),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('LINEBELOW',(0,0),(-1,-1),.3,line)]
        if header:cmds += [('BACKGROUND',(0,0),(-1,0),green),('ROWBACKGROUNDS',(0,1),(-1,-1),[pale,colors.white])]
        else:cmds += [('BACKGROUND',(0,0),(-1,-1),pale)]
        t.setStyle(TableStyle(cmds));return t
    def brand(title):
        logo=Image(BytesIO(logo_bytes),width=255,height=255*402/2096);logo.hAlign='CENTER'
        return [logo,Spacer(1,12),p(title,'title'),Spacer(1,5),p('CONFIDENTIAL - INTERNAL OPPORTUNITY REVIEW','bold'),Spacer(1,10)]
    def identity():return p(f"{meta['customer']} | {meta.get('utility') or 'Utility not specified'} | {meta.get('client_location') or 'Location not specified'}",'bold')
    profit=float(result.gross_profit.sum());annual=profit*12/meta['term']
    story=brand('Opportunity summary')+[identity(),Spacer(1,5),p(f"Delivery start: {meta['start']} | Selected term: {meta['term']} months | Pricing: {meta.get('pricing_mode','Fixed price')}"),p(f"Broker: {meta.get('broker_name') or 'Not specified'} | Offer due: {meta.get('due_date') or 'Not specified'} | Settlements: {meta['asof']}"),Spacer(1,12)]
    cards=[]
    for label,value in [('ANNUALIZED PROFIT',money(annual)),('CONTRACT PROFIT',money(profit)),('NET PRESENT VALUE',money(summary['npv']))]:cards.append([p(label,'bold'),Spacer(1,4),p(value,'money')])
    story += [table([cards],[180]*3),Spacer(1,12)]
    indexed=meta.get('pricing_mode')=='Index + adder'
    rows=[['Selected offer / adder',rate(meta['customer_offer_rate']),'Projected all-in price',rate(summary['contract_price'])],['Contract usage',f"{summary['volume']:,.0f} MMBtu",'Spread after capital',rate(summary['spread'])],['Gross margin after capital',percent(summary['gross_margin']),'Break-even adder' if indexed else 'Break-even offer',rate(meta['break_even'])],['Upfront credit',money(meta['investment']),'Terminal recovery',money(meta['recovery'])]]
    story += [table([[p(v,'bold' if j%2==0 else 'body') for j,v in enumerate(r)] for r in rows],[150,120,150,120]),Spacer(1,12),p('Selected-term offer scenarios','bold'),Spacer(1,5)]
    scenario_table=[[p(h,'white') for h in ['Scenario','Quote / adder','Spread','Annual profit','Contract profit','NPV']]]
    for r in scenarios:
        val=next((v for k,v in r.items() if k.startswith('NPV @')),None)
        scenario_table.append([p(v,'small') for v in [r['Scenario'],rate(r['Offer $/MMBtu']),rate(r['Spread $/MMBtu']),money(r['Annualized GP $']),money(r['Contract GP $']),money(val)]])
    story += [table(scenario_table,[116,82,76,88,90,88],True,6),Spacer(1,12),p('Shared pricing and credit assumptions','bold'),Spacer(1,5)]
    assumptions=[['Transport',rate(meta['transport']),'Brokerage',rate(meta['broker'])],['Target spread',rate(meta['margin']),'Supply fuel',percent(meta['fuel'])],['Capital hurdle',percent(meta['capital_hurdle']),'Financing days',f"{meta['days']:g}"],['Annual NPV discount',percent(meta['discount']),'Credit sizing',meta.get('credit_mode','Manual')],['Transportation recovery','Full pass-through' if meta.get('transport_pass_through') else 'Included in quote','Supplier premium',rate(meta.get('supplier_premium',0))]]
    story += [table([[p(v,'small') for v in r] for r in assumptions],[150,120,150,120],pad=5),Spacer(1,10),p('Sources: '+meta['source'],'small'),p('Market components: '+meta['components'],'small'),Spacer(1,5),p('Profit is after modeled supply, transportation, brokerage and capital charges; overhead and taxes are excluded. NPV uses monthly cash flows, upfront credit and final recovery. Index revenue and profit remain projections.','small')]
    metrics=[('Quoted fixed rate','offer_price',rate),('Projected all-in rate','projected_all_in_price',rate),('Break-even rate / adder','break_even',rate),('Spread after capital','spread',rate),('Gross margin after capital','gross_margin',percent),('Contract usage (MMBtu)','contract_volume',lambda v:f'{v:,.0f}' if pd.notna(v) else 'Unavailable'),('Annualized revenue','annual_revenue',money),('Annualized profit','annualized_profit',money),('Contract profit','contract_profit',money),(f"NPV @ {meta['discount']:.2%}",'npv',money),('Upfront credit posted','upfront_credit',money),('FT credit component','ft_credit',money),('Supply credit component','supply_credit',money),('Terminal credit recovery','credit_recovery',money)]
    for mode,title in [('Fixed price','Fixed-price opportunity terms'),('Index + adder','Index-adder opportunity terms')]:
        data=meta['internal_term_worksheets'][mode];rows=data['rows']
        story += [PageBreak()]+brand(title)+[identity(),Spacer(1,8),p('Pricing as of '+meta['asof']+' | Delivery starts '+meta['start']),Spacer(1,4),p(data['basis'],'small'),Spacer(1,12)]
        if mode=='Index + adder':story += [p('Monthly index: '+meta.get('client_index_name','Selected monthly index'),'bold'),Spacer(1,8)]
        matrix=[[p('ECONOMICS','white')]+[p(f"{r['term_months']} months",'white') for r in rows]]
        for label,key,fmt in metrics:
            if key=='offer_price' and mode=='Index + adder':label='Quoted index adder'
            cells=[p(label,'bold')]
            for r in rows:
                value=r.get(key)
                if key=='annualized_profit':value=r.get('contract_profit')*12/r['term_months'] if pd.notna(r.get('contract_profit')) else None
                if r['status']!='Available':value=None
                cells.append(p(fmt(value),'small'))
            matrix.append(cells)
        ends=[p('Delivery ends','bold')]+[p((pd.Timestamp(meta['start']).to_period('M')+int(r['term_months'])-1).end_time.strftime('%b %d, %Y'),'small') for r in rows]
        matrix.append(ends)
        story += [table(matrix,[150]+[78]*5,True,8),Spacer(1,14),p('Calculation and decision basis','bold'),Spacer(1,5),p('All five terms use the same start date, selected pricing curves and operating assumptions. Exact entered usage is used first; longer terms repeat calendar-month usage estimates. Annualized profit equals contract profit x 12 / term months. NPV and credit are calculated separately for each term. FT/supply components show model sizing; posted credit follows the manual/automatic election. Missing terms are withheld.','small'),Spacer(1,8),p('Transportation: '+('additional full-cost pass-through' if meta.get('transport_pass_through') else 'included in quote')+'. '+('Index adders apply above the monthly index; all-in rates are forward estimates, not fixed quotes.' if mode=='Index + adder' else 'Fixed rates shown are the quoted rate under the indicated transportation treatment.'),'small')]
    out=BytesIO()
    def footer(c,doc):
        c.setFillColor(green);c.rect(0,782,612,10,fill=1,stroke=0);c.setStrokeColor(line);c.line(36,44,576,44)
        c.setFillColor(red);c.setFont('OpportunityBold',7);c.drawString(36,28,'CONFIDENTIAL - INTERNAL DISTRIBUTION ONLY')
        c.setFillColor(green);c.drawRightString(576,28,f'Opportunity review | {doc.page}')
    doc=BaseDocTemplate(out,pagesize=(612,792),title='Internal opportunity review',author='Riverbend Gas Marketing')
    doc.addPageTemplates(PageTemplate(id='InternalOpportunity',frames=[Frame(36,52,540,710,leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0)],onPage=footer))
    doc.build(story)
    return out.getvalue()
