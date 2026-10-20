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

def investment_cashflows(result, initial_investment=0., terminal_recovery=0.):
    """One initial outlay and at most one recovery, in the actual final month."""
    n=len(result)
    if n<1: raise ValueError('At least one operating month is required.')
    operating=np.r_[0.,result.gross_profit.to_numpy(dtype=float)]
    collateral=np.zeros(n+1,dtype=float)
    collateral[0]=-float(initial_investment)
    collateral[n]=float(terminal_recovery)
    return pd.DataFrame({'month_number':np.arange(n+1),'operating_profit':operating,'collateral_cashflow':collateral,'net_cashflow':operating+collateral})

def deal_summary(result, annual_discount=0.10, initial_investment=0., terminal_recovery=0.):
    """Monthly cash proxy: GP at month-end; investment at delivery start."""
    if annual_discount<=-1 or initial_investment<0 or terminal_recovery<0:
        raise ValueError('Invalid discount rate or investment inputs.')
    n=len(result)
    schedule=investment_cashflows(result,initial_investment,terminal_recovery)
    flows=schedule.net_cashflow.to_numpy()[1:]
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
                contract_price=float(result.customer_rate.iloc[0]),collateral_recovery_month=n if terminal_recovery>0 else None,npv=float(npv),irr=irr,irr_status=status,
                spread=gp/volume,gross_margin=gp/revenue if revenue>0 else None)

def standard_offers(prices, asof, selections, start, usage, fallback_profile, transport, broker, margin, fuel, capital_hurdle, financing_days, discount=0.10, investment=0., recovery=0., overrides=None):
    """Exact entered months first, then repeat first-year calendar-month estimates."""
    schedule=usage.copy()
    schedule['month']=pd.to_datetime(schedule['month'])
    exact=dict(zip(schedule['month'],schedule['usage_mmbtu']))
    profile=dict(fallback_profile)
    for row in schedule.iloc[:12].itertuples(index=False): profile[row.month.month]=row.usage_mmbtu
    rows=[]
    for n in [12,18,24,36,48]:
        months=pd.date_range(start,periods=n,freq='MS')
        volumes=[exact[m] if m in exact else profile[m.month] for m in months]
        projected=sum(m not in exact for m in months)
        pricing=curve(prices,asof,selections,months)
        missing=int(pricing.isna().any(axis=1).sum())
        if missing:
            rows.append(dict(term_months=n,offer_price=None,break_even=None,contract_volume=None,contract_profit=None,npv=None,irr=None,discount_rate=discount,collateral_recovery_month=None,irr_status="Unavailable",projected_usage_months=projected,status=f'Unavailable: {missing} months missing settlements'))
            continue
        try:
            result,offer=calculate(pricing,volumes,transport,broker,margin,fuel,fixed_quote=(overrides or {}).get(n),capital_hurdle=capital_hurdle,financing_days=financing_days)
            returns=deal_summary(result,discount,investment,recovery)
            rows.append(dict(term_months=n,target_price=offer,offer_price=float(result.customer_rate.iloc[0]),break_even=offer-margin,spread=result.gross_profit.sum()/result.usage_mmbtu.sum(),gross_margin=returns['gross_margin'],annual_revenue=returns['annual_revenue'],contract_volume=result.usage_mmbtu.sum(),contract_profit=result.gross_profit.sum(),npv=returns['npv'],irr=returns['irr'],discount_rate=discount,collateral_recovery_month=n if recovery>0 else None,irr_status=returns['irr_status'],projected_usage_months=projected,status='Available'))
        except ValueError as e:
            rows.append(dict(term_months=n,offer_price=None,break_even=None,contract_volume=None,contract_profit=None,npv=None,irr=None,discount_rate=discount,collateral_recovery_month=None,irr_status="Unavailable",projected_usage_months=projected,status=str(e)))
    return pd.DataFrame(rows)

def build_summary_pdf(meta, summary, result, scenarios):
    from io import BytesIO
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.enums import TA_LEFT
    buffer=BytesIO()
    navy=colors.HexColor('#153449');teal=colors.HexColor('#197D87');pale=colors.HexColor('#EDF4F6')
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Brand',fontName='Helvetica-Bold',fontSize=17,leading=21,textColor=navy))
    styles.add(ParagraphStyle(name='SmallREM',fontName='Helvetica',fontSize=8,leading=11,textColor=navy))
    styles.add(ParagraphStyle(name='SectionREM',fontName='Helvetica-Bold',fontSize=10,leading=14,textColor=teal,spaceBefore=12,spaceAfter=7))
    def para(text):return Paragraph(escape(str(text)),styles['SmallREM'])
    def usd(v):return f'${v:,.0f}'
    def pct(v):return f'{v:.2%}' if v is not None else 'N/A'
    story=[Paragraph('RIVERBEND GAS MARKETING',styles['Brand']),Paragraph('DEAL ECONOMICS | INTERNAL REVIEW',styles['SectionREM']),para(meta['customer']),para(f"Broker: {meta.get('broker_name','Not specified')}   |   Offer due: {meta.get('due_date','Not specified')}"),para(f"Delivery start: {meta['start']}   |   Term: {meta['term']} months   |   Settlements: {meta['asof']}"),Spacer(1,14)]
    metrics=[('Annualized revenue',usd(summary['annual_revenue'])),('Contract volume',f"{summary['volume']:,.0f} MMBtu"),('All-in contract price',f"${summary['contract_price']:.4f}/MMBtu"),(f"NPV @ {meta['discount']:.2%}",usd(summary['npv'])),('Break-even price',f"${meta['break_even']:.4f}/MMBtu"),('Annualized IRR',pct(summary['irr'])),('Spread after capital',f"${summary['spread']:.4f}/MMBtu"),('Margin after capital',pct(summary['gross_margin']))]
    cards=[]
    for i in [0,4]:
        cards.append([Paragraph(f'<font size="8">{escape(label)}</font><br/><br/><b>{escape(value)}</b>',styles['SmallREM']) for label,value in metrics[i:i+4]])
    table=Table(cards,colWidths=[132]*4,rowHeights=[58,58])
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),pale),('VALIGN',(0,0),(-1,-1),'TOP'),('BOX',(0,0),(-1,-1),.4,colors.white),('INNERGRID',(0,0),(-1,-1),4,colors.white),('LEFTPADDING',(0,0),(-1,-1),10),('TOPPADDING',(0,0),(-1,-1),10)]));story.append(table)
    story.append(Paragraph('Offer comparison',styles['SectionREM']))
    rows=[['Scenario','Price\n$/MMBtu','Spread\n$/MMBtu','Annualized\nprofit','Contract\nprofit','NPV']]
    for row in scenarios:
        npv=next(v for k,v in row.items() if k.startswith('NPV @'))
        rows.append([row['Scenario'],f"${row['Offer $/MMBtu']:.4f}",f"${row['Spread $/MMBtu']:.4f}",usd(row['Annualized GP $']),usd(row['Contract GP $']),usd(npv)])
    def styled_table(rows,widths):
        t=Table([[para(v) for v in row] for row in rows],colWidths=widths,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),pale),('LINEBELOW',(0,0),(-1,0),1,teal),('LINEBELOW',(0,1),(-1,-1),.3,colors.HexColor('#DBE5E9')),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8)]));return t
    story.append(styled_table(rows,[116,78,78,88,88,80]))
    story.append(Paragraph('Pricing and credit assumptions',styles['SectionREM']))
    assumptions=[['Upstream transport',f"${meta['transport']:.4f}/MMBtu",'Brokerage',f"${meta['broker']:.4f}/MMBtu"],['Target spread',f"${meta['margin']:.4f}/MMBtu",'Supply fuel',pct(meta['fuel'])],['Capital hurdle',pct(meta['capital_hurdle']),'Financing period',f"{meta['days']:g} days"],['Capital charge',f"${result.capital_cost.sum()/summary['volume']:.4f}/MMBtu",'Discount rate',pct(meta['discount'])],['Upfront credit cash',usd(meta['investment']),'Terminal credit recovery',usd(meta['recovery'])]]
    story.append(styled_table(assumptions,[132,132,132,132]))
    story.append(Paragraph('Source and valuation basis',styles['SectionREM']))
    story.append(para('Source: '+meta['source']))
    story.append(para('Market components: '+meta['components']))
    story.append(Spacer(1,6))
    story.append(para('NPV is measured at contract start using the stated effective annual rate and month-end profit after the modeled capital charge. IRR is annualized on upfront credit cash and includes terminal credit recovery. These are cash-flow estimates, not a full working-capital forecast.'))
    story.append(Spacer(1,5))
    story.append(para('Annualized revenue and profit equal contract totals multiplied by 12 / term months. The quoted all-in price includes transportation and brokerage; these are deducted once, not billed again. Break-even includes supply, transport, brokerage and capital charge; overhead and taxes are excluded. The offer reductions hold volumes and costs constant.'))
    def footer(c,doc):
        c.setStrokeColor(teal);c.line(42,39,570,39);c.setFont('Helvetica',7);c.setFillColor(navy);c.drawString(42,27,'CONFIDENTIAL | Riverbend Gas Marketing');c.drawRightString(570,27,f'Page {doc.page}')
    SimpleDocTemplate(buffer,pagesize=(612,792),leftMargin=42,rightMargin=42,topMargin=36,bottomMargin=50,title='Riverbend Deal Economics',author='Riverbend Gas Marketing').build(story,onFirstPage=footer,onLaterPages=footer)
    return buffer.getvalue()

"""Durable named workstreams in a private GitHub repository."""
import base64
import json
import re
import uuid
from datetime import datetime, timezone
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from urllib.parse import quote

class WorkstreamError(ValueError): pass

def github_request(repo,token,path='',method='GET',body=None):
    if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+',repo): raise WorkstreamError('Set WORKSTREAM_REPO to owner/private-repository.')
    url=f'https://api.github.com/repos/{repo}'+path
    headers={'Accept':'application/vnd.github+json','Authorization':f'Bearer {token}','X-GitHub-Api-Version':'2022-11-28','User-Agent':'REM-Deal-Economics'}
    data=None if body is None else json.dumps(body,allow_nan=False).encode()
    try:
        with urlopen(Request(url,data=data,headers=headers,method=method),timeout=20) as response:return json.load(response)
    except HTTPError as e:
        if e.code==404: raise WorkstreamError('Not found. Check the private repository, branch and token access.') from None
        if e.code in [409,422]: raise WorkstreamError('Save conflict: reopen the current saved version or save a new workstream.') from None
        raise WorkstreamError(f'GitHub request failed (HTTP {e.code}). Check token permissions or expiry.') from None
    except (URLError,TimeoutError):raise WorkstreamError('Unable to reach saved-workstream storage. Your current inputs remain on screen.') from None

def verify_private(repo,token):
    info=github_request(repo,token)
    if not info.get('private'): raise WorkstreamError('Use a private repository for saved customer economics. WORKSTREAM_REPO currently points to a public repository.')

def list_workstreams(repo,token,branch='main'):
    verify_private(repo,token)
    try:items=github_request(repo,token,'/contents/workstreams?ref='+quote(branch,safe=''))
    except WorkstreamError as e:
        if str(e).startswith('Not found'):return []
        raise
    return sorted([x for x in items if x.get('type')=='file' and x['name'].endswith('.json')],key=lambda x:x['name'].lower())

def load_workstream(repo,token,path,branch='main'):
    if not re.fullmatch(r'workstreams/[A-Za-z0-9_.-]+\.json',path):raise WorkstreamError('Invalid saved workstream path.')
    item=github_request(repo,token,'/contents/'+path+'?ref='+quote(branch,safe=''))
    payload=json.loads(base64.b64decode(item['content']))
    if payload.get('schema_version')!=1: raise WorkstreamError('Unsupported workstream version.')
    return payload,item['sha']

def save_workstream(repo,token,name,snapshot,branch='main',path=None,sha=None):
    verify_private(repo,token)
    if not name.strip():raise WorkstreamError('Enter a workstream name.')
    if path is None:
        slug=re.sub('[^a-z0-9]+','-',name.lower()).strip('-')[:60] or 'deal'
        path=f'workstreams/{slug}-{uuid.uuid4().hex[:10]}.json'
    if not re.fullmatch(r'workstreams/[A-Za-z0-9_.-]+\.json',path):raise WorkstreamError('Invalid saved workstream path.')
    if sha is None and snapshot.get('_existing'):raise WorkstreamError('Reopen the saved workstream before updating it.')
    payload=dict(snapshot,schema_version=1,name=name.strip(),saved_at=datetime.now(timezone.utc).isoformat())
    body=dict(message='Save deal workstream: '+name.strip(),branch=branch,content=base64.b64encode(json.dumps(payload,indent=2,allow_nan=False).encode()).decode())
    if sha:body['sha']=sha
    response=github_request(repo,token,'/contents/'+path,'PUT',body)
    return path,response['content']['sha'],payload


ROOT=Path(__file__).resolve().parent
st.set_page_config(page_title='REM | Deal Economics',layout='wide',initial_sidebar_state='expanded')
st.markdown("""<style>
.block-container {max-width:1550px;padding-top:2rem;}
h1,h2,h3 {font-family:Arial,sans-serif;text-align:center;color:#153449;}
[data-testid="stVerticalBlockBorderWrapper"] {border-color:#a6bdc9 !important;border-radius:12px;}
[data-testid="stMetric"] {background:#f0f5f8;border:1px solid #d0dee5;border-radius:8px;padding:12px 6px;text-align:center;min-height:95px;}
[data-testid="stMetricLabel"] {justify-content:center;font-size:.82rem;}
[data-testid="stMetricValue"] {font-size:1.35rem;font-weight:600;color:#153449;}
[data-testid="stMetricLabel"] p {font-size:.82rem;}
[data-testid="stNumberInput"] input {text-align:center;}
</style>""",unsafe_allow_html=True)
st.title('Riverbend Gas Marketing')
st.caption('Natural gas deal economics • Monthly volume-weighted pricing')
st.caption('Pricing: GitHub input folder | Build REM-TERM-RECOVERY-1020')
def workstream_config():
    def setting(name,default=''):
        try:return str(st.secrets.get(name,os.getenv(name,default)))
        except Exception:return os.getenv(name,default)
    return setting('WORKSTREAM_REPO'),setting('WORKSTREAM_GITHUB_TOKEN'),setting('WORKSTREAM_BRANCH','main')

ws_repo,ws_token,ws_branch=workstream_config()
if '_pending_workstream' in st.session_state:
    pending=st.session_state.pop('_pending_workstream')
    st.session_state.pop('_current_snapshot',None)
    for k in list(st.session_state):
        if k.startswith(('wk_','offer_mode_','offer_override_','usage_')):del st.session_state[k]
    for k,v in pending.get('widgets',{}).items():
        if k.startswith(('wk_','offer_mode_','offer_override_')):
            if k in ['wk_start','wk_due']:v=pd.Timestamp(v).date() if v else None
            if k.startswith('wk_component_'):v=tuple(v)
            st.session_state[k]=v
    st.session_state['_restored_usage']=pending.get('usage',[])
    st.session_state['_saved_pricing_file']=pending.get('pricing_file')
    st.session_state['_saved_trade_date']=pending.get('trade_date')
    st.session_state['_restore_pricing']=True
    st.session_state['ws_latest']=bool(pending.get('latest_pricing',False))
    st.session_state['ws_name']=pending['name']

def save_current_workstream(copy=False):
    try:
        snapshot=st.session_state.get('_current_snapshot')
        if not snapshot:raise WorkstreamError('Load pricing and enter deal assumptions before saving.')
        path,sha,payload=save_workstream(ws_repo,ws_token,st.session_state.get('ws_name',''),snapshot,ws_branch,None if copy else st.session_state.get('_ws_path'),None if copy else st.session_state.get('_ws_sha'))
        st.session_state['_ws_path']=path;st.session_state['_ws_sha']=sha
        st.session_state['_saved_message']='Saved '+payload['name']+' at '+payload['saved_at']
        st.session_state['_last_saved_snapshot']=snapshot
    except Exception as e:st.session_state['_save_error']=str(e)

def open_workstream(path):
    try:
        payload,sha=load_workstream(ws_repo,ws_token,path,ws_branch)
        st.session_state['_pending_workstream']=payload
        st.session_state['_ws_path']=path;st.session_state['_ws_sha']=sha
        st.session_state['_saved_message']='Opened '+payload['name']+' | saved '+payload.get('saved_at','')
    except Exception as e:st.session_state['_save_error']=str(e)

with st.container(border=True):
    st.subheader('Saved workstreams')
    st.caption('Save a named working deal across browser refreshes and days. Saving is explicit; pricing updates do not change the saved assumptions until you save again.')
    st.text_input('Workstream name',key='ws_name',placeholder='Customer - December 2026 evaluation')
    if not ws_repo or not ws_token:
        st.info('One-time setup: configure a private workstream repository and GitHub token in Streamlit Secrets to enable durable Save / Open. See the included setup guide.')
    else:
        if st.checkbox('Show saved workstreams',key='ws_show'):
            try:
                saved=list_workstreams(ws_repo,ws_token,ws_branch)
                if saved:
                    selected=st.selectbox('Saved workstream',saved,format_func=lambda x:x['name'])
                    st.button('Open saved workstream',on_click=open_workstream,args=(selected['path'],))
                else:st.caption('No saved workstreams yet. Save the current deal below.')
            except Exception as e:st.error(str(e))
        one,two=st.columns(2)
        one.button('Save current workstream',on_click=save_current_workstream,disabled='_current_snapshot' not in st.session_state)
        two.button('Save as new workstream',on_click=save_current_workstream,args=(True,),disabled='_current_snapshot' not in st.session_state)
    if '_saved_message' in st.session_state:st.success(st.session_state['_saved_message'])
    if '_save_error' in st.session_state:st.error(st.session_state.pop('_save_error'))
    backup=st.file_uploader('Restore a workstream backup (optional; not pricing)',type=['json'],key='ws_backup')
    if backup is not None and st.button('Open backup workstream'):
        try:
            import json
            payload=json.loads(backup.getvalue())
            if payload.get('schema_version')!=1:raise ValueError('Unsupported backup version')
            st.session_state['_pending_workstream']=payload
            st.session_state.pop('_ws_path',None);st.session_state.pop('_ws_sha',None)
            st.rerun()
        except Exception as e:st.error(str(e))
    st.checkbox('Use latest available pricing file',key='ws_latest',help='Uses the newest date in an ICE filename, falling back to file modified time. Deal assumptions and manual offers remain unchanged.')
    def refresh_latest_pricing():
        st.session_state['ws_latest']=True
        st.session_state['_restore_pricing']=True
        st.session_state['_saved_trade_date']=None
        st.session_state.pop('loaded_pricing_stamp',None)
        st.session_state.pop('wk_asof',None)
        st.session_state.pop('wk_pricing_file',None)
    st.button('Refresh latest pricing',on_click=refresh_latest_pricing)

summary_panel=st.container()

def format_table(df):
    formats={}
    for name in df.columns:
        if not pd.api.types.is_numeric_dtype(df[name]): continue
        key=str(name).lower()
        if '%' in key: formats[name]='{:,.2f}%'
        elif key in {'irr','gross_margin','discount_rate'}: formats[name]='{:.2%}'
        elif any(t in key for t in ['price','profit','revenue','cost','rate','npv','spread','margin','capital','supply','broker','transport','offer','break_even']):
            formats[name]='${:,.4f}' if any(t in key for t in ['price','rate','spread','break_even']) or '/mmbtu' in key else '${:,.0f}'
        else: formats[name]='{:,.0f}'
    return df.style.format(formats,na_rep='N/A')

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
    import re
    def pricing_recency(path):
        match=re.search(r'(20\d{2})[-_](\d{2})[-_](\d{2})',path.name)
        try:return (pd.Timestamp('-'.join(match.groups())).timestamp() if match else path.stat().st_mtime,path.name)
        except ValueError:return (path.stat().st_mtime,path.name)
    names=[str(p.relative_to(INPUT_DIR)) for p in available]
    if st.session_state.get('ws_latest'):
        ice_files=[p for p in available if 'ice' in p.name.lower()]
        desired=str(max(ice_files or available,key=pricing_recency).relative_to(INPUT_DIR))
    else:desired=st.session_state.get('_saved_pricing_file')
    if desired and desired in names and (st.session_state.get('_restore_pricing') or st.session_state.get('ws_latest')):
        st.session_state['wk_pricing_file']=desired
    if st.session_state.get('wk_pricing_file') not in names:st.session_state.pop('wk_pricing_file',None)
    selected_name=st.selectbox('ICE Cleared Gas pricing file',names,key='wk_pricing_file')
    selected_path=INPUT_DIR/selected_name
    st.caption(f'Selected repository file: input/{selected_path.relative_to(INPUT_DIR)}')
    stamp=f'{selected_path}:{selected_path.stat().st_mtime_ns}'
    if st.session_state.pop('_restore_pricing',False) or st.session_state.get('ws_latest'):
        st.session_state['loaded_pricing_stamp']=stamp
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
    saved_date=st.session_state.get('_saved_trade_date')
    if saved_date and not st.session_state.get('ws_latest') and pd.Timestamp(saved_date) in dates:st.session_state.setdefault('wk_asof',pd.Timestamp(saved_date))
    if st.session_state.get('ws_latest'):st.session_state['wk_asof']=dates[0]
    if st.session_state.get('wk_asof') not in dates:st.session_state.pop('wk_asof',None)
    asof=st.selectbox('Settlement trade date',dates,format_func=lambda d:pd.Timestamp(d).strftime('%b %d, %Y'),key='wk_asof')
    st.caption('All components use this exact trade date. Missing months block a quote.')
    customer=st.text_input('Customer',value=defaults['customer'],key='wk_customer')
    broker_name=st.text_input('Broker name / firm',value='',key='wk_broker_name')
    offer_due=st.date_input('Offer due date',value=None,key='wk_due')
    start_entered=st.date_input('Contract start date',value=(pd.Timestamp(asof).to_period('M')+1).to_timestamp().date(),key='wk_start')
    start=start_entered.replace(day=1)
    if start_entered.day!=1: st.caption('Monthly pricing uses the start month. Enter first-month usage for the remaining delivery days; usage is not automatically prorated.')
    term=st.slider('Term (months)',min_value=1,max_value=48,value=12,step=1,key='wk_term')

st.subheader('Market components')
st.caption('Choose products by name. Additive mode follows the workbook: Henry Hub + HSC basis + HSC index adjustment. Outright mode uses one full-price product.')
kind=st.radio('Curve construction',['Additive components','Outright fixed price'],horizontal=True,key='wk_kind')
options=list(prices[['hub','product','contract']].drop_duplicates().itertuples(index=False,name=None))
selections=[]
cols=st.columns(3 if kind=='Additive components' else 1)
for i,c in enumerate(cols):
    preferred=['H','HXS','SHS'][i] if kind=='Additive components' else 'H'
    idx=next((j for j,x in enumerate(options) if x[2]==preferred),0)
    if st.session_state.get(f'wk_component_{i}') not in options:
        if st.session_state.get(f'wk_component_{i}') is not None:st.warning(f'Saved component {i+1} is not in this pricing file. Review the replacement selection before quoting.')
        st.session_state.pop(f'wk_component_{i}',None)
    selections.append(c.selectbox(f'Component {i+1}',options,index=idx,format_func=lambda x:f'{x[0]} | {x[1]} ({x[2]})',key=f'wk_component_{i}'))
if len(set(selections))!=len(selections): st.error('Select distinct components.');st.stop()
st.subheader('Deal assumptions')
a,b,c,d=st.columns(4)
transport=a.number_input('Upstream transport ($/delivered MMBtu)',min_value=0.,value=defaults['transport'],step=.01,format='%.4f',key='wk_transport')
broker=b.number_input('Brokerage ($/delivered MMBtu)',min_value=0.,value=defaults['broker'],step=.01,format='%.4f',key='wk_broker')
margin=c.number_input('Target gross margin ($/MMBtu)',value=defaults['margin'],step=.01,format='%.4f',key='wk_margin')
fuel=d.number_input('Supply fuel retention (%)',min_value=0.,max_value=99.,value=0.,step=.1,key='wk_fuel')/100
st.caption('Margin is after supply, transport and brokerage, after the capital charge, before overhead, other credit fees and taxes. Transport and brokerage apply to delivered usage. Fuel gross-up applies to gas supply only.')
st.subheader('Cost of capital')
cap1,cap2=st.columns(2)
capital_hurdle=cap1.number_input('Annual capital hurdle (%)',min_value=0.,value=8.,step=.5,key='wk_capital')/100
financing_days=cap2.number_input('Days of costs financed',min_value=0.,value=30.,step=1.,key='wk_days')
st.caption('Capital charge per delivered MMBtu = positive supply + transport + brokerage cost per MMBtu × annual hurdle × financing days ÷ 365. The charge is included in the target offer and break-even price, and deducted from modeled profit. Financing days represent the cash collection gap. Negative total costs have a zero capital base.')
with st.expander('Upfront credit and investment returns',expanded=True):
    a,b,c=st.columns(3)
    discount=a.number_input('Required annual discount rate (%)',key='wk_discount',min_value=0.,value=10.,step=.5,format='%.2f',help='Effective annual rate used for NPV. This is separate from the offer-price capital hurdle and does not change IRR.')/100
    investment=b.number_input('Upfront cash posted for credit ($)',key='wk_investment',min_value=0.,value=0.,step=1000.,help='Cash collateral posted at contract start. Do not enter the face amount of an unfunded letter of credit.')
    full_recovery=c.checkbox('Return all posted credit cash at contract end',value=True,key='wk_full_recovery')
    recovery=investment if full_recovery else c.number_input('Credit cash returned at contract end ($)',key='wk_recovery',min_value=0.,value=0.,step=1000.)
    c.caption(f'Credit cash recovered: ${recovery:,.0f}')
    st.caption('NPV is measured at contract start. Cash flows use month-end profit after the modeled capital charge, before overhead and taxes, less initial investment, plus terminal recovery. The upfront credit amount is a cash outflow at contract start. Credit cash is returned in the final month under the selected recovery assumption. Monthly profit is received at month-end. The posted credit cash is separate from the financing-days capital charge; no additional interest charge on collateral is added. IRR is annualized and requires an initial cash outflow.')
months=pd.date_range(start,periods=term,freq='MS')
base=pd.DataFrame({'month':months,'usage_mmbtu':[defaults['usage'].get(m.month,0.) for m in months]})
if st.session_state.get('_restored_usage'):
    saved_usage={pd.Timestamp(row['month']):float(row['usage_mmbtu']) for row in st.session_state['_restored_usage']}
    base['usage_mmbtu']=[saved_usage.get(m,v) for m,v in zip(base.month,base.usage_mmbtu)]
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
usage=st.data_editor(base,hide_index=True,disabled=['month'],column_config={'usage_mmbtu':st.column_config.NumberColumn('Usage (MMBtu/month)',min_value=0.,required=True,format='localized')},key=f'usage_{start}_{term}',use_container_width=True)
offers=standard_offers(prices,asof,selections,start,usage,defaults['usage'],transport,broker,margin,fuel,capital_hurdle,financing_days,discount,investment,recovery)
with summary_panel:
    with st.container(border=True):
        st.subheader('Projected offers | standard contract terms')
        st.caption(f'Start: {start_entered:%b %d, %Y} • Target margin after capital: ${margin:.4f}/MMBtu • Settlements: {pd.Timestamp(asof):%Y-%m-%d}')
        st.caption(f'Shared assumptions: transport ${transport:.4f}/MMBtu | brokerage ${broker:.4f}/MMBtu | capital hurdle {capital_hurdle:.2%} | financing {financing_days:g} days | target margin ${margin:.4f}/MMBtu')
        offer_cols=st.columns(5)
        overrides={}
        output_panels=[]
        for col,row in zip(offer_cols,offers.itertuples(index=False)):
            with col:
                with st.container(border=True):
                    st.markdown(f'### {row.term_months} months')
                    if row.status!='Available':
                        st.warning(row.status)
                        output_panels.append(None)
                        continue
                    st.caption(f'All-in target: ${row.offer_price:,.4f}/MMBtu')
                    setting=st.selectbox('Offer setting',['Model target','Target less $0.10','Target less $0.20','Manual price'],key=f'offer_mode_{row.term_months}')
                    if setting=='Manual price':
                        overrides[row.term_months]=st.number_input('All-in offer ($/MMBtu)',value=float(row.offer_price),step=.01,format='%.4f',key=f'offer_override_{row.term_months}')
                    elif setting!='Model target':
                        overrides[row.term_months]=row.offer_price-(.1 if setting=='Target less $0.10' else .2)
                    output_panels.append(st.container())
        offers=standard_offers(prices,asof,selections,start,usage,defaults['usage'],transport,broker,margin,fuel,capital_hurdle,financing_days,discount,investment,recovery,overrides)
        for panel,row in zip(output_panels,offers.itertuples(index=False)):
            if panel is None: continue
            with panel:
                st.metric('All-in offer ($/MMBtu)',f'${row.offer_price:,.4f}')
                st.metric(f'NPV @ {discount:.2%}',f'${row.npv:,.0f}' if pd.notna(row.npv) else 'N/A')
                st.metric('Annualized IRR',f'{row.irr:.2%}' if pd.notna(row.irr) else 'N/A')
                st.caption(f'Collateral returned: ${recovery:,.0f} in month {row.term_months}' if recovery>0 else 'No collateral recovery included')
                st.metric('Contract profit',f'${row.contract_profit:,.0f}')
                st.metric('Spread ($/MMBtu)',f'${row.spread:,.4f}')
                st.metric('Gross margin',f'{row.gross_margin:.2%}' if pd.notna(row.gross_margin) else 'N/A')
        st.caption('Change each term independently. Model target follows the shared assumptions; manual prices remain as entered. Offer reductions are relative to the current model target.')
        st.caption(f'Returns use ${investment:,.0f} upfront credit cash, ${recovery:,.0f} recovered at the end of each respective term, and a {discount:.2%} effective annual discount rate. NPV is valued at contract start; IRR is annualized.')
        st.caption('Each offer is a separately calculated fixed price including supply, fuel, upstream transport, brokerage, capital charge and target margin. Entered months use your exact volumes. Beyond the entered schedule, the first 12 months repeat by calendar month; any calendar months not entered use the displayed 1,700 MMBtu/month default. Each term reflects its own offer setting above. For the selected contract term, that price also feeds the detailed summary, comparisons and PDF.')
        with st.expander('Standard offer detail and volume assumptions'):
            st.dataframe(format_table(offers),hide_index=True,use_container_width=True)
        st.download_button('Download standard offers CSV',offers.to_csv(index=False).encode(),'standard_contract_offers.csv','text/csv')
with summary_panel:
    with st.container(border=True):
        st.subheader('Detailed summary selection')
        st.caption(f'The summary, monthly economics, offer comparisons and PDF below use the selected {term}-month contract. Change the contract term in the pricing setup to review another term.')
        if term in [12,18,24,36,48]:
            quote=overrides.get(term)
        else:
            custom_override=st.checkbox(f'Manual offer for selected {term}-month contract',key='wk_custom_override')
            quote=st.number_input('Selected contract offer ($/MMBtu)',value=4.,step=.01,format='%.4f',key='wk_custom_price') if custom_override else None
import json
widget_snapshot={}
for k,v in st.session_state.items():
    if k.startswith(('wk_','offer_mode_','offer_override_')) and k not in ['wk_pricing_file','wk_asof']:
        if hasattr(v,'isoformat'):v=v.isoformat()
        elif isinstance(v,tuple):v=list(v)
        widget_snapshot[k]=v
current_usage=[dict(month=str(pd.Timestamp(row.month).date()),usage_mmbtu=float(row.usage_mmbtu)) for row in usage.itertuples(index=False)]
snapshot=dict(schema_version=1,widgets=widget_snapshot,usage=current_usage,pricing_file=str(selected_path.relative_to(INPUT_DIR)),trade_date=str(pd.Timestamp(asof).date()),latest_pricing=bool(st.session_state.get('ws_latest')),name=st.session_state.get('ws_name',''))
st.session_state['_current_snapshot']=snapshot
with summary_panel:
    st.download_button('Download workstream backup (JSON)',json.dumps(snapshot,indent=2,allow_nan=False).encode(),'REM_Workstream.json','application/json')
try:
    p=curve(prices,asof,selections,months)
    if p.isna().any().any():
        st.error('Pricing is missing for these delivery months/components. Select another trade date, product or term.')
        st.dataframe(p[p.isna().any(axis=1)].style.format('${:,.4f}',na_rep='Missing'));st.stop()
    result,target=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,quote,capital_hurdle,financing_days)
except ValueError as e: st.error(str(e));st.stop()
summary=deal_summary(result,discount,investment,recovery)
with summary_panel:
    with st.expander('Verify collateral timing and investment cash flows'):
        st.caption(f'For this {term}-month contract, upfront credit is posted in month 0. '+(f'Collateral is returned once, in month {term}, alongside the final operating profit.' if recovery>0 else 'No collateral repayment is included.'))
        cash_audit=investment_cashflows(result,investment,recovery)
        monthly_discount=(1+discount)**(1/12)-1
        cash_audit['discount_factor']=1/(1+monthly_discount)**cash_audit.month_number
        cash_audit['present_value']=cash_audit.net_cashflow*cash_audit.discount_factor
        st.dataframe(cash_audit.style.format({'month_number':'{:,.0f}','operating_profit':'${:,.2f}','collateral_cashflow':'${:,.2f}','net_cashflow':'${:,.2f}','discount_factor':'{:.6f}','present_value':'${:,.2f}'}),hide_index=True,use_container_width=True)
        st.caption(f'Sum of present values = NPV ${cash_audit.present_value.sum():,.2f}. IRR uses the same net cash-flow column, including month 0.')
        st.download_button('Download investment cash-flow audit CSV',cash_audit.to_csv(index=False).encode(),'investment_cashflow_audit.csv','text/csv')

with summary_panel:
    pdf_panel=st.container()
with summary_panel.container(border=True):
    st.subheader('Deal summary')
    st.caption(f'Broker: {broker_name or "Not specified"} • Offer due: {offer_due or "Not specified"}')
    st.caption(f'{customer} • Start {start_entered:%b %d, %Y} • {term} month term • Settlements {pd.Timestamp(asof):%b %d, %Y}')
    top=st.columns(4)
    top[0].metric('Annual revenue (annualized)',f"${summary['annual_revenue']:,.0f}")
    top[1].metric('Volume sold (contract MMBtu)',f"{summary['volume']:,.0f}")
    top[2].metric('All-in contract price ($/MMBtu)',f"${summary['contract_price']:,.4f}")
    top[3].metric(f'NPV @ {discount:.2%} annual',f"${summary['npv']:,.0f}")
    bottom_break_even=target-margin
    bottom=st.columns(4)
    bottom[3].metric('Break-even ($/MMBtu)',f'${bottom_break_even:,.4f}')
    bottom[0].metric('IRR (annualized)',f"{summary['irr']:.2%}" if summary['irr'] is not None else 'N/A')
    bottom[1].metric('MMBtu spread ($/MMBtu)',f"${summary['spread']:,.4f}")
    bottom[2].metric('Gross margin (%)',f"{summary['gross_margin']:.2%}" if summary['gross_margin'] is not None else 'N/A')
    st.markdown('**All-in price breakdown ($/delivered MMBtu)**')
    weighted_supply=float(result.supply_cost.sum()/result.usage_mmbtu.sum())
    weighted_capital=float(result.capital_cost.sum()/result.usage_mmbtu.sum())
    bridge=pd.DataFrame({'Component':['Gas supply (volume-weighted)','Transportation included','Brokerage included','Capital charge included','Marketer margin after capital','Total all-in customer price'],'Rate ($/MMBtu)':[weighted_supply,transport,broker,weighted_capital,summary['spread'],summary['contract_price']]})
    st.dataframe(bridge.style.format({'Rate ($/MMBtu)':'${:,.4f}'}),hide_index=True,use_container_width=True)
    st.caption('Transportation and brokerage are included inside the quoted all-in price. They are not added to the customer bill again. Customer revenue is the all-in price multiplied by usage; supply, transport, brokerage and the modeled capital charge are each deducted once.')
    st.caption(f'PV basis: {discount:.2%} effective annual discount rate; equivalent monthly rate {(1+discount)**(1/12)-1:.4%}. Valuation date: contract start. Monthly cash flows occur at month-end. IRR is solved independently of this rate.')
    st.caption(f"Capital hurdle: {capital_hurdle:.1%} • Financing days: {financing_days:g} • Weighted capital charge: ${(result.capital_cost.sum()/result.usage_mmbtu.sum()):.4f}/MMBtu • Contract capital charge: ${result.capital_cost.sum():,.0f}")
    st.caption('Annual revenue = total contract revenue × 12 ÷ term months. Spread = gross profit ÷ delivered volume. Gross margin = gross profit ÷ revenue. '+summary['irr_status']+'.')
    st.download_button('Download deal summary CSV',pd.DataFrame([dict(customer=customer,broker_name=broker_name,offer_due_date=str(offer_due) if offer_due else '',start_date=str(start_entered),term_months=term,trade_date=str(pd.Timestamp(asof).date()),capital_hurdle=capital_hurdle,financing_days=financing_days,capital_charge_per_mmbtu=result.capital_cost.sum()/result.usage_mmbtu.sum(),annual_discount_rate=discount,upfront_credit_cash=investment,credit_cash_recovery=recovery,**summary)]).to_csv(index=False).encode(),'deal_summary.csv','text/csv')
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
    scenario_rows.append({'Scenario':label,'Offer $/MMBtu':offer_price,'Break-even $/MMBtu':target-margin,'Spread $/MMBtu':ss['spread'],'Gross margin %':ss['gross_margin']*100 if ss['gross_margin'] is not None else None,'Average monthly GP $':sr.gross_profit.sum()/term,'Annualized GP $':sr.gross_profit.sum()*12/term,'Contract GP $':sr.gross_profit.sum(),f'NPV @ {discount:.2%} $':ss['npv'],'Annual discount rate %':discount*100,'Annualized IRR %':ss['irr']*100 if ss['irr'] is not None else None})
    scenario_monthly[label]=sr.gross_profit
scenario_table=pd.DataFrame(scenario_rows)
with pdf_panel:
    try:
        pdf_bytes=build_summary_pdf(dict(customer=customer,broker_name=broker_name or 'Not specified',due_date=str(offer_due) if offer_due else 'Not specified',start=str(start_entered),term=term,asof=str(pd.Timestamp(asof).date()),discount=discount,break_even=target-margin,transport=transport,broker=broker,margin=margin,fuel=fuel,capital_hurdle=capital_hurdle,days=financing_days,investment=investment,recovery=recovery,source='input/'+str(selected_path.relative_to(INPUT_DIR)),components=' + '.join(' | '.join(x) for x in selections)),summary,result,scenario_rows)
        st.download_button('Download professional deal summary (PDF)',pdf_bytes,'REM_Deal_Summary.pdf','application/pdf')
    except ImportError:
        st.warning('PDF export requires reportlab in requirements.txt. CSV downloads remain available.')

st.dataframe(format_table(scenario_table),hide_index=True,use_container_width=True)
st.caption('Monthly and annualized averages use the selected contract term. The schedule below shows actual monthly profits. NPV includes initial investment and recovery; break-even price is the volume-weighted cost including the capital charge.')
with st.expander('Monthly offer economics',expanded=True):
    st.line_chart(scenario_monthly)
    st.dataframe(format_table(scenario_monthly.reset_index().rename(columns={c:c+" profit ($)" for c in scenario_monthly.columns})), hide_index=True,use_container_width=True)
st.download_button('Download offer comparison CSV',scenario_table.to_csv(index=False).encode(),'offer_comparison.csv','text/csv')
st.download_button('Download monthly offer economics CSV',scenario_monthly.reset_index().to_csv(index=False).encode(),'monthly_offer_economics.csv','text/csv')
t1,t2,t3=st.tabs(['Monthly economics','Term comparisons','Method & source'])
with t1:
    st.line_chart(result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']])
    st.bar_chart(result[['gross_profit']])
    st.dataframe(format_table(result.reset_index()),hide_index=True,use_container_width=True)
    totals=result[['usage_mmbtu','revenue','supply_cost','transport_cost','broker_cost','capital_cost','gross_profit']].sum()
    st.write('Contract totals')
    st.dataframe(format_table(totals.to_frame('Total').T),hide_index=True,use_container_width=True)
    st.write('Selected-strip averages',pd.DataFrame({'Simple average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mean(),'Volume-weighted average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mul(result.usage_mmbtu,axis=0).sum()/vol}))
with t2:
    st.caption('Compares standard terms and your selected term within the entered usage schedule. Longer terms require extending the schedule above. Each term is priced separately to the same target margin.')
    rows=[]
    for n in sorted(set([1,3,6,12,18,24,36,48,term])):
        if n>term or usage.usage_mmbtu.iloc[:n].sum()<=0: continue
        r,q=calculate(p.iloc[:n],usage.usage_mmbtu.iloc[:n],transport,broker,margin,fuel,capital_hurdle=capital_hurdle,financing_days=financing_days)
        rows.append({'Months':n,'Volume MMBtu':r.usage_mmbtu.sum(),'Target fixed price':q,'Gross profit':r.gross_profit.sum(),'Annualized GP':r.gross_profit.sum()*12/n})
    st.dataframe(format_table(pd.DataFrame(rows)),hide_index=True)
with t3:
    st.write('Supply = sum of selected components ÷ (1 − fuel retention). Landed cost = supply + upstream transport + brokerage. Target fixed price = volume-weighted landed cost + target margin. Capital charge = max(landed cost, 0) × annual hurdle × financing days / 365. Target fixed price = weighted landed cost + weighted capital charge + target margin. Monthly modeled profit = (customer fixed price − landed cost − capital charge) × usage.')
    st.write('Uses actual entered monthly volumes, calendar delivery months and a single settlement snapshot. No interpolation, silent zero filling or fallback to older settlements. Daily swing observations are not converted to forward prices.')
    st.write('The capital charge is a modeled financing expense, not a full working-capital schedule. NPV separately discounts profit after that charge; initial investment and terminal recovery are additional cash flows and should not duplicate the financed operating costs. The workbook’s NPV/carry formulas and daily-volume approximation are not reproduced. This model quotes a level fixed price over the selected term. It excludes utility charges, imbalance exposure, credit/collateral financing and overhead unless included in your rates.')
    st.dataframe(pd.DataFrame(selections,columns=['Hub','Product','Contract']),hide_index=True)
export=result.reset_index(); export.insert(0,'customer',customer);export.insert(1,'trade_date',pd.Timestamp(asof));export.insert(2,'contract_start_date',str(start_entered));export.insert(3,'term_months',term);export.insert(2,'components',' + '.join(x[2] for x in selections))
st.download_button('Download monthly economics CSV',export.to_csv(index=False).encode(),'rem_deal_economics.csv','text/csv')
st.download_button('Download usage schedule CSV',usage.to_csv(index=False).encode(),'customer_usage.csv','text/csv')
