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

def calculate(pricing, usage, transport, broker, margin, fuel=0., fixed_quote=None):
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
    target=float(np.average(x.landed_rate,weights=x.usage_mmbtu)+margin)
    quote=target if fixed_quote is None else float(fixed_quote)
    x['customer_rate']=quote
    x['revenue']=quote*x.usage_mmbtu
    for name,col in [('supply_cost','supply_rate'),('transport_cost','transport_rate'),('broker_cost','broker_rate')]: x[name]=x[col]*x.usage_mmbtu
    x['gross_profit']=x.revenue-x.supply_cost-x.transport_cost-x.broker_cost
    x['margin_per_mmbtu']=x.customer_rate-x.landed_rate
    return x,target


ROOT=Path(__file__).resolve().parent
st.set_page_config(page_title='REM | Deal Economics',layout='wide')
st.title('Riverbend Energy Management')
st.caption('Natural gas deal economics • Monthly volume-weighted pricing')

@st.cache_data
def bundled(): return read_prices(ROOT/'data/rempr.xlsx'),workbook_defaults(ROOT/'data/rempr.xlsx')
@st.cache_data(ttl=300)
def load_location(path): return read_prices(path)

defaults=dict(customer='New customer', start=str(pd.Timestamp.today().to_period('M').to_timestamp().date()), transport=.35, broker=.30, margin=.18, usage={m:1700. for m in range(1,13)})
prices=None
if (ROOT/'data/rempr.xlsx').exists():
    prices,defaults=bundled()
with st.sidebar:
    st.header('Pricing source')
    mode=st.selectbox('Load pricing from',['Attached workbook','Upload pricing file','Shared file / URL','Existing app session'],index=0 if prices is not None else 1)
    try:
        if mode=='Upload pricing file':
            upload=st.file_uploader('ICE or normalized pricing',type=['xlsx','csv','parquet'])
            if upload is None: st.info('Upload a file to continue.');st.stop()
            prices=read_prices(upload)
            try: defaults=workbook_defaults(upload)
            except (KeyError,ValueError,TypeError): pass
        elif mode=='Shared file / URL':
            loc=st.text_input('Pricing file path or raw HTTPS URL',value=os.getenv('REM_PRICING_PATH',''))
            if not loc: st.info('Enter the same master file used by the pricing app.');st.stop()
            prices=load_location(loc)
        elif mode=='Existing app session':
            if 'pricing_data' not in st.session_state: st.info('The existing app must set st.session_state["pricing_data"] to its pricing DataFrame.');st.stop()
            prices=normalize_prices(st.session_state['pricing_data'])
    except Exception as e: st.error(f'Unable to load pricing: {e}');st.stop()
    if st.button('Refresh file pricing'): load_location.clear();st.rerun()
    if prices is None:
        st.info('Upload rempr.xlsx or a pricing master, or select Shared file / URL.');st.stop()
    dates=sorted(prices.trade_date.unique(),reverse=True)
    asof=st.selectbox('Settlement trade date',dates,format_func=lambda d:pd.Timestamp(d).strftime('%b %d, %Y'))
    st.caption('All components use this exact trade date. Missing months block a quote.')
    customer=st.text_input('Customer',value=defaults['customer'])
    start_entered=st.date_input('Contract start date',value=pd.Timestamp(defaults['start']).date())
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
st.caption('Margin is after supply, transport and brokerage, before overhead, credit fees and taxes. Transport and brokerage apply to delivered usage. Fuel gross-up applies to gas supply only.')
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
    result,target=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,quote)
except ValueError as e: st.error(str(e));st.stop()
vol=result.usage_mmbtu.sum();gp=result.gross_profit.sum()
a,b,c,d=st.columns(4)
a.metric('Target customer fixed price',f'${target:,.4f}/MMBtu')
b.metric('Total contract volume',f'{vol:,.0f} MMBtu')
c.metric('Contract gross profit',f'${gp:,.0f}')
d.metric('Realized gross margin',f'${gp/vol:,.4f}/MMBtu')
st.caption(f'Customer: {customer} • Settlements: {pd.Timestamp(asof):%Y-%m-%d} • Annualized gross profit: ${gp*12/term:,.0f}')
t1,t2,t3=st.tabs(['Monthly economics','Term comparisons','Method & source'])
with t1:
    st.line_chart(result[['market_rate','landed_rate','customer_rate']])
    st.bar_chart(result[['gross_profit']])
    st.dataframe(result.reset_index(),hide_index=True,use_container_width=True)
    totals=result[['usage_mmbtu','revenue','supply_cost','transport_cost','broker_cost','gross_profit']].sum()
    st.write('Contract totals',totals.to_frame('Total').T)
    st.write('Selected-strip averages',pd.DataFrame({'Simple average':result[['market_rate','landed_rate','customer_rate']].mean(),'Volume-weighted average':result[['market_rate','landed_rate','customer_rate']].mul(result.usage_mmbtu,axis=0).sum()/vol}))
with t2:
    st.caption('Compares standard terms and your selected term within the entered usage schedule. Longer terms require extending the schedule above. Each term is priced separately to the same target margin.')
    rows=[]
    for n in sorted(set([1,3,6,12,18,24,36,48,term])):
        if n>term or usage.usage_mmbtu.iloc[:n].sum()<=0: continue
        r,q=calculate(p.iloc[:n],usage.usage_mmbtu.iloc[:n],transport,broker,margin,fuel)
        rows.append({'Months':n,'Volume MMBtu':r.usage_mmbtu.sum(),'Target fixed price':q,'Gross profit':r.gross_profit.sum(),'Annualized GP':r.gross_profit.sum()*12/n})
    st.dataframe(pd.DataFrame(rows),hide_index=True)
with t3:
    st.write('Supply = sum of selected components ÷ (1 − fuel retention). Landed cost = supply + upstream transport + brokerage. Target fixed price = volume-weighted landed cost + target margin. Monthly gross profit = (customer fixed price − landed cost) × monthly usage.')
    st.write('Uses actual entered monthly volumes, calendar delivery months and a single settlement snapshot. No interpolation, silent zero filling or fallback to older settlements. Daily swing observations are not converted to forward prices.')
    st.write('The workbook’s NPV/carry formulas and daily-volume approximation are not reproduced. This model quotes a level fixed price over the selected term. It excludes utility charges, imbalance exposure, credit/collateral financing and overhead unless included in your rates.')
    st.dataframe(pd.DataFrame(selections,columns=['Hub','Product','Contract']),hide_index=True)
export=result.reset_index(); export.insert(0,'customer',customer);export.insert(1,'trade_date',pd.Timestamp(asof));export.insert(2,'contract_start_date',str(start_entered));export.insert(3,'term_months',term);export.insert(2,'components',' + '.join(x[2] for x in selections))
st.download_button('Download monthly economics CSV',export.to_csv(index=False).encode(),'rem_deal_economics.csv','text/csv')
st.download_button('Download usage schedule CSV',usage.to_csv(index=False).encode(),'customer_usage.csv','text/csv')
