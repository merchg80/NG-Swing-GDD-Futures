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
st.set_page_config(page_title='REM | Deal Economics',layout='wide')
st.title('Riverbend Energy Management')
st.caption('Natural gas deal economics • Monthly volume-weighted pricing')
summary_panel=st.container()

@st.cache_data
def bundled(): return read_prices(ROOT/'data/rempr.xlsx'),workbook_defaults(ROOT/'data/rempr.xlsx')
@st.cache_data(ttl=300)
def load_location(path): return read_prices(path)

defaults=dict(customer='New customer', start=str(pd.Timestamp.today().to_period('M').to_timestamp().date()), transport=.35, broker=.30, margin=.18, usage={m:1700. for m in range(1,13)})
prices=None
if (ROOT/'data/rempr.xlsx').exists():
    prices,defaults=bundled()
@st.cache_data(ttl=60)
def repository_files():
    # Restrict discovery to the app repository and omit hidden/dependency folders.
    excluded={'venv','env','node_modules','__pycache__','site-packages'}
    files=[]
    for path in ROOT.rglob('*'):
        relative=path.relative_to(ROOT)
        if any(part.startswith('.') or part in excluded for part in relative.parts): continue
        if path.is_file() and path.suffix.lower() in {'.xlsx','.csv','.parquet'}:
            files.append(str(relative))
    return sorted(files,key=lambda name: ('ice' not in name.lower(),name.lower()))
