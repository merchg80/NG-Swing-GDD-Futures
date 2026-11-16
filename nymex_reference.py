"""Exact-date Henry Hub H reference; independent from regional supply components."""
import numpy as np
import pandas as pd

MODES=('Fixed price','Index + adder','NYMEX + adder')
INDEXED_MODES=MODES[1:]


def henry_hub(prices,asof,months):
    index=pd.DatetimeIndex(months,name='month')
    selected=prices.loc[prices.trade_date.eq(pd.Timestamp(asof)) & prices.contract.astype(str).str.strip().str.upper().eq('H')]
    records=[]
    for month in index:
        hits=selected.loc[selected.strip.eq(month),'settlement_price']
        status='Available' if len(hits)==1 else 'Missing H settlement' if hits.empty else 'Duplicate H settlements'
        value=float(hits.iloc[0]) if len(hits)==1 else np.nan
        if not np.isfinite(value):
            value=np.nan
            if status=='Available':status='Invalid H settlement'
        records.append(dict(nymex_h_settlement=value,nymex_h_status=status))
    return pd.DataFrame(records,index=index,columns=['nymex_h_settlement','nymex_h_status'])


def summary(result):
    required=['nymex_h_settlement','offer_nymex_adder','modeled_nymex_adder']
    if any(c not in result or not np.isfinite(pd.to_numeric(result[c],errors='coerce')).all() for c in required):
        return dict(nymex_h_reference=None,offer_nymex_adder=None,modeled_nymex_adder=None,nymex_reference_status='Unavailable: complete exact-date H settlements required')
    weights=result.usage_mmbtu
    return dict(nymex_h_reference=float(np.average(result.nymex_h_settlement,weights=weights)),
        offer_nymex_adder=float(np.average(result.offer_nymex_adder,weights=weights)),
        modeled_nymex_adder=float(np.average(result.modeled_nymex_adder,weights=weights)),nymex_reference_status='Available')
