from io import BytesIO, StringIO
from datetime import date
import hashlib
import math
import re
import json
import zipfile
import sqlite3
from pathlib import Path
from urllib.request import Request, build_opener, HTTPRedirectHandler
from urllib.parse import quote
from urllib.error import HTTPError, URLError

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

VERSION = 'Stage 3.73 — Draft cleanup and bulk delivery strips'
TRANSITION = pd.Timestamp('2026-06-01')
OCT_START, OCT_END = pd.Timestamp('2026-10-01'), pd.Timestamp('2026-10-31')
DIRECT = {'DOM':'Eastern Gas-South','EGTS':'Eastern Gas-South','TCO':'TCO',
          'REXZ3D':'REX-Z3','TENN219':'TGP-Z4 200L','ANRSE':'ANR-SE',
          'TETCOM2':'TETCO-M2 (receipts)','TENN500':'TGP-500L','HSC':'HSC'}
LEGACY = {'DOM':8,'EGTS':9,'TCO':21,'REXZ3D':22,'TENN219':20,
          'ANRSE':19,'TETCOM2':15,'TENN500':23,'HSC':12}
PROXIES = {'TENN219','ANRSE','TENN500'}
ALIASES = {'DOM':'DOM','DOMSP':'DOM','EGTS':'DOM','TCO':'TCO',
           'REXZ3':'REXZ3D','REXZ3D':'REXZ3D','NYM':'NYMEX','NYMEX':'NYMEX'}
RULES = {'FIXED':'Contract fixed price','GDD':'Daily index plus premium',
         'INDEX':'Monthly index plus premium',
         'PHYSB':'Monthly index plus contractual amount; interpretation requires review'}
INDEX_DATA = '''egts,dom,tco,rexz3d,nymex
4.47,4.47,4.63,,5.34
6.387,6.387,6.58,,7.27
-0.995,-0.995,-0.89,,8.908
-0.85,-0.85,-0.75,,6.551
-0.925,-0.925,-0.775,,8.686
7.9,-2.3,-2.15,,9.353
-2.325,-2.325,-1.9825,,6.868
-1.37,-1.37,-1.1675,,5.186
5.9,5.9,6.1,,6.7
3.605,3.605,4.07,,4.709
2.46,2.46,2.61,,3.31
2.02,2.02,2.08,,2.45
1.6285,1.6285,1.7185,,1.991
1.207,1.207,1.297,,2.117
1.211,1.211,1.526,,2.181
1.233,1.233,1.59,,2.603
1.13,1.13,1.37,,2.492
0.96,0.96,1.07,,2.54
1.249,1.249,1.38,,2.764
1.7755,1.7755,1.938,,3.163
2.005,2.005,2.025,,2.705
1.9815,1.9815,2.1115,,2.619
1.7125,1.7125,2.0375,,2.49
1.2,1.2,1.3,,1.615
1.2625,1.2625,1.4275,,1.575
1.259,1.259,1.334,,1.614
1.483,1.483,1.713,,2.493
1.838,1.838,1.943,,2.628
1.236,1.236,1.341,,2.036
1.27,1.27,1.38,,1.93
1.575,1.575,1.725,,2.585
1.73,1.73,1.77,2.1,2.346
3.026,3.026,3.11,3.54,3.431
3.034,3.034,3.1765,3.77,3.514
3.065,3.065,3.28,,3.535
3.47,3.47,3.52,3.58,3.906
3.15,3.15,3.23,3.37,3.949
2.33,2.33,2.54,2.67,3.17
2.09,2.09,2.43,2.74,3.204
2.04,2.04,2.2,,3.261
2.29,2.29,2.53,,3.081
1.77,1.77,2.04,,2.867
1.61,1.61,1.71,,2.835
2.63,2.63,2.741,,3.376
3.8,3.8,3.95,4.28,4.424
4.07,4.07,4.33,4.49,4.687
7.15,7.15,7.33,,7.46
2.36,2.36,2.447,,2.97
2.16,2.16,2.3225,,3.05
1.75,1.75,2.06,,2.559
1.9,1.9,2.1,2.46,2.894
2.33,2.33,2.49,2.74,3.231
1.865,1.865,1.975,,2.725
1.672,1.672,1.786,,2.906
1.42,1.42,1.55,,3
'''


def txt(v):
    return '' if v is None or pd.isna(v) else str(v).strip()


def canonical_counterparty(value):
    name=txt(value)
    aliases={'snyder':'Snyder','mercuria':'Mercuria'}
    return aliases.get(name.casefold(),name[:1].upper()+name[1:])


def num(v):
    try:
        n = float(str(v).replace(',', '').strip())
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None


def day(v):
    if v is None or isinstance(v, (int, float)):
        return pd.NaT
    d = pd.to_datetime(v, errors='coerce')
    return d.normalize() if pd.notna(d) else pd.NaT


def market(v):
    label = txt(v).upper()
    return ALIASES.get(label, label)


def same(a, b):
    a, b = num(a), num(b)
    return a is not None and b is not None and math.isclose(a, b, abs_tol=1e-6, rel_tol=0)


def show(df, filename):
    # Rebuilt on every rerun: never mix exports from different inputs.
    if '_run_reports' in st.session_state:
        st.session_state['_run_reports'][filename] = df.copy()
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button('Download CSV', df.to_csv(index=False).encode(), filename,
                       'text/csv', key=filename)


def monthly_history():
    raw = pd.read_csv(StringIO(INDEX_DATA))
    dates = pd.date_range('2022-04-01','2026-10-01',freq='MS')
    if len(raw) != len(dates):
        raise ValueError('Monthly settlement count does not match dates.')
    rows = []
    for d, (_, r) in zip(dates, raw.iterrows()):
        for c in ['egts','dom','tco','rexz3d','nymex']:
            original = r[c]
            converted = pd.notna(original) and c != 'nymex' and original < 0
            calculated = round(r.nymex + original, 6) if converted else original
            approved = d == pd.Timestamp('2022-09-01') and c in ['egts','dom']
            rows.append({'delivery_month':d,'index_code':c.upper(),
                         'original_value':original,'nymex_reference':r.nymex,
                         'calculated_before_correction':calculated,
                         'settlement_price':7.053 if approved else calculated,
                         'method':'Owner-confirmed settlement' if approved else
                         'Not supplied' if pd.isna(original) else
                         'Converted from basis' if converted else 'Supplied settlement',
                         'correction_note':'DOM/EGTS approved at 7.053; original retained.' if approved else ''})
    return pd.DataFrame(rows)


def normalize_cash(raw, date_column, source):
    frame = raw.copy()
    frame.columns = [txt(c) for c in frame.columns]
    if date_column not in frame:
        raise ValueError(f'{source}: missing {date_column} column.')
    original = frame[date_column]
    parsed = pd.to_datetime(original, errors='coerce')
    invalid = int((original.notna() & parsed.isna()).sum())
    frame[date_column] = parsed
    frame = frame[frame[date_column].notna()].rename(columns={date_column:'Date'}).copy()
    frame['Date'] = frame['Date'].dt.normalize()
    frame.attrs.update(source=source, invalid_dates=invalid)
    return frame


def cash_sheet(wb, name, position):
    ws = wb[name]
    ws.reset_dimensions()
    it = ws.iter_rows(values_only=True)
    header = next(it, None)
    if header is None or len(header) <= position:
        raise ValueError(f'{name}: no usable header.')
    headers = [txt(v) or f'Unnamed_{i+1}' for i,v in enumerate(header)]
    rows = []
    for r in it:
        if len(r) > position and r[position] is not None:
            values = list(r[:len(headers)])
            rows.append(values + [None] * (len(headers)-len(values)))
    return normalize_cash(pd.DataFrame(rows,columns=headers), headers[position], f'Workbook {name}')


def duplicate_trade_review(ledger):
    keys=['source_sheet','trade_date','start','end','daily_volume','counterparty',
          'index_code','secondary_label','contract_price','premium','fixed_price']
    prior=ledger.duplicated(keys,keep=False)
    current=ledger.duplicated(keys+['location'],keep=False)
    distinguished=ledger.loc[prior & ~current].copy()
    distinguished['review_resolution']='Distinct delivery locations; no longer flagged by the duplicate-field check. Trade retained.'
    return current,distinguished


def apply_confirmed_trade_corrections(ledger):
    """Owner-confirmed duplicate; require row identity AND original trade terms.

    Preserve source values for audit. Never zero an unrelated replacement row.
    """
    out=ledger.copy()
    out['source_daily_volume']=out.daily_volume
    out['trade_correction']=''
    for i,r in out.iterrows():
        if (r.source_id=='cashblotter:13475' and r.source_sheet=='cashblotter'
            and txt(r.counterparty).casefold()=='shell'
            and day(r.trade_date)==pd.Timestamp('2026-10-01')
            and day(r.start)==pd.Timestamp('2026-10-02')
            and day(r.end)==pd.Timestamp('2026-10-02')
            and ''.join(c for c in txt(r.location).lower() if c.isalnum())=='deodlvd'
            and txt(r.index_code).upper()=='EGTS'
            and same(r.daily_volume,3400) and same(r.premium,.11)):
            out.loc[i,'daily_volume']=0.
            out.loc[i,'trade_correction']='Owner confirmed duplicate: volume zeroed; cashblotter:13487 remains valid.'
    return out


@st.cache_data(show_spinner=False)
def import_book(content):
    wb = load_workbook(BytesIO(content),read_only=True,data_only=True)
    fw = load_workbook(BytesIO(content),read_only=True,data_only=False)
    trades, flags, support, lookups = [], [], [], []
    def flag(sid, issue, detail='', severity='REVIEW'):
        flags.append(dict(source_id=sid,severity=severity,issue=issue,detail=detail))
    try:
        missing = {'cashblotter','fwdblotter','cash','cash2'}-set(wb.sheetnames)
        if missing:
            raise ValueError('Missing sheets: '+', '.join(sorted(missing)))
        for sheet, first, left, right in [('cashblotter',5,4,17),('fwdblotter',3,3,12)]:
            id_columns=[]
            for header in wb[sheet].iter_rows(min_row=1,max_row=first-1,max_col=64,values_only=True):
                id_columns.extend(i+1 for i,value in enumerate(header) if txt(value).lower().replace(' ','_')=='risk_deal_id')
            if len(id_columns)>1:raise ValueError('Multiple Risk Deal ID headers in '+sheet+'; keep one unique ID column.')
            id_column=id_columns[0] if len(id_columns)==1 else None
            id_rows=wb[sheet].iter_rows(min_row=first,min_col=id_column,max_col=id_column,values_only=True) if id_column else None
            vr=wb[sheet].iter_rows(min_row=first,min_col=left,max_col=right,values_only=True)
            fr=fw[sheet].iter_rows(min_row=first,min_col=left,max_col=right,values_only=True)
            for rn,(v,f) in enumerate(zip(vr,fr),first):
                risk_id=txt(next(id_rows,('',))[0]) if id_rows else ''
                if not any(txt(f[i]) for i in range(7)):
                    continue
                sid=f'{sheet}:{rn}'
                r=dict(source_id=sid,source_sheet=sheet,source_row=rn,risk_deal_id=risk_id,
                       trade_date=day(v[0]),premium=None,fixed_price=None,
                       saved_cash_index=None,saved_transaction_price=None,
                       saved_invoice_amount=None,contract_price=None,
                       secondary_label='',classification_code='',location='',pricing_label='',
                       premium_issue=False)
                if sheet=='cashblotter':
                    r.update(start=day(v[1]),end=day(v[2]),daily_volume=num(v[3]),
                             location=txt(v[4]),counterparty=txt(v[5]),pricing_label=txt(v[6]),
                             index_code=txt(v[7]).upper(),premium=num(v[8]),
                             fixed_price=num(v[10]),saved_cash_index=num(v[9]),
                             saved_transaction_price=num(v[11]),saved_invoice_amount=num(v[12]))
                    r['premium_issue'] = (bool(txt(f[8])) and num(v[8]) is None)
                    important=[1,2,3]
                else:
                    r.update(daily_volume=num(v[1]),index_code=txt(v[2]).upper(),
                             contract_price=num(v[3]),start=day(v[4]),end=day(v[5]),
                             counterparty=txt(v[6]),secondary_label=txt(v[7]),
                             classification_code=txt(v[9]))
                    important=[1,3,4,5]
                for i in important:
                    if isinstance(f[i],str) and f[i].startswith('=') and v[i] is None:
                        flag(sid,'Missing cached formula result','Recalculate and save in Excel.','ERROR')
                if pd.isna(r['start']) or pd.isna(r['end']):
                    flag(sid,'Invalid delivery dates',severity='ERROR')
                elif r['end']<r['start'] or (r['end']-r['start']).days>7305:
                    flag(sid,'Invalid or excessive delivery period',severity='ERROR')
                if r['daily_volume'] is None:
                    flag(sid,'Invalid volume',severity='ERROR')
                if not r['counterparty']:
                    flag(sid,'Missing counterparty',severity='ERROR')
                if not r['index_code']:
                    flag(sid,'Missing index code')
                expected={'domgdd':'DOMSP','tcogdd':'TCO'}.get(r['secondary_label'].lower())
                if expected and expected!=r['index_code']:
                    flag(sid,'Index/secondary label disagreement')
                r['counterparty_raw']=r['counterparty']
                r['counterparty']=canonical_counterparty(r['counterparty'])
                trades.append(r)
        tables={n:cash_sheet(wb,n,p) for n,p in [('cash',2),('cash2',0)]}
        headers=list(tables['cash'].columns)
        for rn,r in enumerate(wb['cashblotter'].iter_rows(min_row=4,max_row=161,min_col=43,max_col=44,values_only=True),4):
            if not txt(r[0]) or not txt(r[1]): continue
            m=re.search(r'cash!\$?([a-z]+):\$?([a-z]+)',txt(r[1]),re.I)
            col=column_index_from_string(m.group(1).upper()) if m else 0
            lookups.append(dict(index_code=txt(r[0]).upper(),reference=txt(r[1]),
                                historical_market=headers[col-1] if 0<col<=len(headers) else '',source_row=rn))
        if 'wapp' in wb.sheetnames:
            for rn,v in enumerate(wb['wapp'].iter_rows(min_row=5,min_col=9,max_col=17,values_only=True),5):
                if num(v[0]) is None and not txt(v[5]): continue
                support.append(dict(source_id=f'wapp:{rn}',daily_volume=num(v[0]),product_label=txt(v[1]),
                                    contract_price_or_premium=num(v[2]),start=day(v[3]),end=day(v[4]),
                                    counterparty=canonical_counterparty(v[5]),counterparty_raw=txt(v[5]),location=txt(v[6]),settlement_index=txt(v[7]).upper(),
                                    pricing_classification=txt(v[8]).upper()))
        ledger=pd.DataFrame(trades)
        if not ledger.empty:
            ledger=apply_confirmed_trade_corrections(ledger)
            duplicate_mask,_=duplicate_trade_review(ledger)
            for sid in ledger.loc[duplicate_mask,'source_id']:
                flag(sid,'Possible duplicate trade','All records remain included pending review.')
        return ledger,pd.DataFrame(flags,columns=['source_id','severity','issue','detail']),tables,pd.DataFrame(lookups),pd.DataFrame(support)
    finally:
        wb.close(); fw.close()


@st.cache_data(show_spinner=False)
def read_master(content):
    return normalize_cash(pd.read_csv(BytesIO(content)),'Date','Uploaded cash master')


@st.cache_data(show_spinner=False)
def read_ice(content):
    f=pd.read_excel(BytesIO(content))
    f.columns=[txt(c).upper() for c in f.columns]
    required={'TRADE DATE','STRIP','CONTRACT','CONTRACT TYPE','SETTLEMENT PRICE','PRODUCT','HUB'}
    if not required.issubset(f): raise ValueError('Missing ICE columns: '+', '.join(sorted(required-set(f.columns))))
    for c in ['TRADE DATE','STRIP']: f[c]=pd.to_datetime(f[c],errors='coerce')
    f['SETTLEMENT PRICE']=pd.to_numeric(f['SETTLEMENT PRICE'],errors='coerce')
    return f


class NoGitHubRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def github_request(endpoint, token, raw=False):
    request = Request('https://api.github.com/' + endpoint, headers={
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/vnd.github.raw+json' if raw else 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'Riverbend-Risk',
    })
    try:
        with build_opener(NoGitHubRedirect()).open(request, timeout=60) as response:
            data = response.read()
    except HTTPError as exc:
        messages = {401:'GitHub token is invalid or expired.',
                    403:'GitHub denied access or its API limit was reached.',
                    404:'Repository, branch, or file is inaccessible. Check token repository access and spelling.'}
        raise ValueError(messages.get(exc.code, f'GitHub request failed (HTTP {exc.code}).')) from None
    except (URLError, TimeoutError):
        raise ValueError('Unable to reach GitHub. Retry Refresh sources.') from None
    if raw:
        if data.startswith(b'version https://git-lfs.github.com/spec'):
            raise ValueError('This file uses Git LFS. Its pricing bytes require a different download connection.')
        return data
    return json.loads(data)


@st.cache_data(ttl=300, show_spinner=False)
def github_inventory(repo, branch, token):
    if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repo):
        raise ValueError('Repository must have the format owner/repository.')
    root = 'repos/' + repo
    if not branch:
        branch = github_request(root, token)['default_branch']
    commit = github_request(root + '/commits/' + quote(branch, safe=''), token)
    sha = commit['sha']
    tree_sha = commit['commit']['tree']['sha']
    tree = github_request(root + '/git/trees/' + tree_sha + '?recursive=1', token)
    if tree.get('truncated'):
        raise ValueError('Repository file listing was truncated. Explicit file paths need to be configured.')
    return dict(commit=sha, branch=branch,
                paths=sorted(x['path'] for x in tree['tree'] if x['type']=='blob'),
                checked_at=pd.Timestamp.now(tz='UTC').isoformat())


@st.cache_data(ttl=300, show_spinner=False)
def github_bytes(repo, commit, path, token):
    return github_request('repos/' + repo + '/contents/' + quote(path, safe='/')
                          + '?ref=' + quote(commit, safe=''), token, raw=True)


def ice_auto_paths(paths):
    dated = []
    aliases = []
    for path in paths:
        name = path.rsplit('/',1)[-1].lower()
        if not name.endswith('.xlsx') or not name.startswith('icecleared'):
            continue
        match = re.search(r'(20\d{2})[_-](\d{2})[_-](\d{2})', name)
        if match:
            try:
                stamp = pd.Timestamp('-'.join(match.groups()))
                dated.append((stamp,path))
            except ValueError:
                pass
        elif 'latest' in name:
            aliases.append(path)
    # Inspect recent dated snapshots and the latest alias. Select valuation
    # date from the file contents, not the filename. Old history is not loaded.
    return list(dict.fromkeys(aliases + [p for _,p in sorted(dated,reverse=True)[:2]]))


def github_pricing_controls():
    try:
        config = dict(st.secrets.get('github', {}))
    except (FileNotFoundError, KeyError):
        config = {}
    token = txt(config.get('token'))
    if not token or token == 'PASTE_YOUR_TOKEN_HERE':
        st.info('Add [github] token in Streamlit Settings → Secrets. Upload validation mode remains available.')
        return None
    repo = txt(config.get('repository')) or 'merchg80/NG-Swing-GDD-Futures'
    branch = txt(config.get('branch'))
    try:
        inventory = github_inventory(repo, branch, token)
        paths = inventory['paths']
        cash_candidates = [p for p in paths if p.lower().endswith('.csv') and 'master_ng' in p.rsplit('/',1)[-1].lower()]
        csv_paths = [p for p in paths if p.lower().endswith('.csv')]
        preferred = txt(config.get('cash_path'))
        cash_options = ['Select cash master'] + sorted(set(cash_candidates + csv_paths))
        default = preferred if preferred in cash_options else cash_candidates[0] if len(cash_candidates)==1 else cash_options[0]
        cash_path = st.selectbox('GitHub cash-master path', cash_options, index=cash_options.index(default))
        excel_paths = [p for p in paths if p.lower().endswith('.xlsx')]
        auto_label = 'Automatic: compare recent ICE snapshots'
        ice_options = [auto_label] + excel_paths
        preferred_ice = txt(config.get('ice_path'))
        selected = st.selectbox('GitHub full ICE path', ice_options,
                                index=ice_options.index(preferred_ice) if preferred_ice in ice_options else 0)
        chosen = ice_auto_paths(paths) if selected == auto_label else [selected]
        st.caption(f"Repository: {repo}; branch: {inventory['branch']}; commit: {inventory['commit'][:8]}")
        if cash_path == cash_options[0]:
            st.warning('Select the cash master file from the dropdown. No daily upload is required.')
            return None
        if not chosen:
            st.warning('Select the full ICE spreadsheet. No matching icecleared files were discovered automatically.')
            return None
        with st.spinner('Retrieving pricing from GitHub...'):
            master = read_master(github_bytes(repo,inventory['commit'],cash_path,token))
            frames = []
            for path in chosen:
                f = read_ice(github_bytes(repo,inventory['commit'],path,token)).copy()
                f['source_file'] = path
                frames.append(f)
        master.attrs['source'] = f"GitHub {repo}/{cash_path} @ {inventory['commit'][:8]}"
        return dict(master=master, frames=frames, cash_path=cash_path, ice_paths=chosen,
                    repository=repo, inventory=inventory)
    except Exception as exc:
        # Only our sanitized connection errors and parser errors are displayed.
        st.error(f'GitHub pricing unavailable: {exc}')
        st.caption('No silent fallback to workbook pricing occurs in GitHub mode.')
        return None


def reconcile(ledger,wapp):
    forward=ledger[ledger.source_sheet.eq('fwdblotter')]
    s=wapp.copy()
    s['_cp']=s.counterparty.map(lambda x:txt(x).upper())
    s['_market']=s.apply(lambda r:market(r.settlement_index or r.product_label),axis=1)
    candidates={}; reverse={}
    for _,r in forward.iterrows():
        ids=[]
        if txt(r.counterparty) and txt(r.index_code) and pd.notna(r.start) and pd.notna(r.end):
            pool=s[s['_cp'].eq(txt(r.counterparty).upper()) & s.start.eq(r.start) &
                   s.end.eq(r.end) & s['_market'].eq(market(r.index_code))]
            for _,v in pool.iterrows():
                if same(r.daily_volume,v.daily_volume) and same(r.contract_price,v.contract_price_or_premium):
                    ids.append(v.source_id)
                    reverse.setdefault(v.source_id,[]).append(r.source_id)
        candidates[r.source_id]=ids
    rows=[]; used=set()
    for _,r in forward.iterrows():
        ids=candidates[r.source_id]; cls=''; idx=''; cardinality=''; note=''
        if not ids:
            status='UNMATCHED'
        else:
            p=s[s.source_id.isin(ids)]
            signatures={(txt(v.pricing_classification).upper(),market(v.settlement_index)) for _,v in p.iterrows()}
            unique=len(ids)==1 and len(reverse[ids[0]])==1
            cardinality='Unique row match' if unique else 'Shared supporting records'
            if len(signatures)!=1:
                status='CONFLICTING CLASSIFICATIONS'
            else:
                cls,idx=next(iter(signatures))
                if not cls or not idx:
                    status='MATCHED — INCOMPLETE CLASSIFICATION'
                elif cls not in RULES:
                    status='MATCHED — REVIEW CLASSIFICATION'
                else:
                    status='MATCHED — REVIEW' if unique else 'SHARED — CONSISTENT CLASSIFICATION'
                    used.update(ids)
                    note='Each blotter trade remains separate. Classification agreement does not prove duplicate status.'
        rows.append(dict(source_id=r.source_id,trade_date=r.trade_date,counterparty=r.counterparty,
                         start=r.start,end=r.end,daily_volume=r.daily_volume,blotter_index=r.index_code,
                         secondary_label=r.secondary_label,contract_price_or_premium=r.contract_price,
                         match_status=status,candidate_count=len(ids),candidate_wapp_ids=', '.join(ids),
                         match_cardinality=cardinality,candidate_classification=cls,
                         candidate_settlement_index=idx,pricing_rule=RULES.get(cls,''),review_note=note))
    return pd.DataFrame(rows),wapp[~wapp.source_id.isin(used)]


def acknowledge_expired_exceptions(ledger, flags):
    """Apply the user's specific historical acknowledgments, never a blanket waiver."""
    specs = {
        'cashblotter:24': ('snyder', '2022-04-03', None, 'Invalid volume'),
        'cashblotter:11538': ('snyder', '2026-02-01', '2026-01-28', 'Invalid or excessive delivery period'),
        'cashblotter:11539': ('carbonbetter', '2026-02-01', '2026-01-28', 'Invalid or excessive delivery period'),
        'fwdblotter:379': ('snyder', '2023-04-01', '2023-03-31', 'Invalid or excessive delivery period'),
        'fwdblotter:380': ('snyder', '2023-04-01', '2023-03-31', 'Invalid or excessive delivery period'),
    }
    ledger = ledger.copy()
    ledger['expired_exception_acknowledged'] = False
    acknowledged = pd.Series(False, index=flags.index)
    for sid, (party, start, end, issue) in specs.items():
        hits = ledger[ledger.source_id.eq(sid)]
        if len(hits) != 1:
            continue
        r = hits.iloc[0]
        if str(r.counterparty).strip().casefold() != party or r.start != pd.Timestamp(start):
            continue
        if pd.isna(r.end) or r.end >= pd.Timestamp('2026-10-24'):
            continue
        if end is not None and r.end != pd.Timestamp(end):
            continue
        if issue == 'Invalid volume' and (pd.notna(r.daily_volume) or r.end < r.start):
            continue
        match = flags.source_id.eq(sid) & flags.issue.eq(issue)
        if not match.any():
            continue
        acknowledged |= match
        ledger.loc[hits.index, 'expired_exception_acknowledged'] = True
    audit = flags.loc[acknowledged].merge(
        ledger[['source_id','counterparty','start','end','daily_volume']], on='source_id', how='left')
    audit['acknowledgment'] = 'User acknowledged expired record; original data retained'
    audit['acknowledged_on'] = '2026-10-24'
    return ledger, flags.loc[~acknowledged].copy(), audit


def scope_import_flags(ledger, flags, start, end):
    """Shared app/runner policy: unknown identity or delivery scope stays reviewable."""
    columns=['source_id','source_sheet','counterparty','start','end','daily_volume','index_code','location']
    if ledger.source_id.duplicated().any():
        raise ValueError('Duplicate ledger source IDs prevent reliable flag scoping.')
    out=flags.merge(ledger[columns],on='source_id',how='left',validate='many_to_one')
    out['start']=pd.to_datetime(out.start,errors='coerce')
    out['end']=pd.to_datetime(out.end,errors='coerce')
    out['reporting_scope']='UNRESOLVED — review required'
    for sheet,a,b in [('cashblotter',OCT_START,OCT_END),('fwdblotter',start,end)]:
        a,b=pd.Timestamp(a),pd.Timestamp(b)
        if pd.isna(a) or pd.isna(b) or b<a or (b-a).days>3660:
            continue
        known=out.source_sheet.eq(sheet) & out.start.notna() & out.end.notna() & out.end.ge(out.start) & (out.end-out.start).dt.days.le(7305)
        out.loc[known & out.end.lt(a),'reporting_scope']='HISTORICAL — outside reporting period'
        out.loc[known & out.start.gt(b),'reporting_scope']='FUTURE — outside reporting period'
        out.loc[known & out.start.le(b) & out.end.ge(a),'reporting_scope']='IN REPORTING PERIOD'
    out['requires_scheduled_review']=out.reporting_scope.isin(['IN REPORTING PERIOD','UNRESOLVED — review required'])
    out['volume_note']=out.daily_volume.map(lambda v:'ZERO VOLUME — record retained' if pd.notna(v) and v==0 else '')
    return out


def active_unallocatable_records(ledger, active_flags):
    invalid = ledger.start.isna() | ledger.end.isna() | ledger.daily_volume.isna() | (ledger.end < ledger.start) | ((ledger.end-ledger.start).dt.days > 7305)
    # A separate unresolved issue on an acknowledged row must still block it.
    unresolved = ledger.source_id.isin(active_flags.source_id)
    return invalid & (~ledger.expired_exception_acknowledged | unresolved)


def monthly_volumes(ledger,start,end,positive_buy,all_physical=False):
    rows=[]
    selected = ledger if all_physical else ledger[ledger.source_sheet.eq('fwdblotter')]
    for _,r in selected.iterrows():
        if pd.isna(r.start) or pd.isna(r.end) or num(r.daily_volume) is None or r.end<r.start or (r.end-r.start).days>7305: continue
        a,b=max(start,r.start),min(end,r.end)
        if a>b: continue
        for m in pd.date_range(a.to_period('M').start_time,b.to_period('M').start_time,freq='MS'):
            days=(min(b,m+pd.offsets.MonthEnd(0))-max(a,m)).days+1
            v=r.daily_volume*days; d=v*(1 if positive_buy else -1)
            rows.append(dict(month=m,source_id=r.source_id,counterparty=r.counterparty,index_code=r.index_code,
                             location=r.get("location", ""),secondary_label=r.get("secondary_label", ""),source_sheet=r.get("source_sheet", ""),
                             delivery_days=days,signed_volume=v,purchase_volume=max(d,0),sale_volume=max(-d,0),net_volume=d))
    return pd.DataFrame(rows)


def daily_prices(code,start,end,legacy,newer,master):
    rows=[]; headers=list(legacy.columns)
    for d in pd.date_range(start,end):
        note=''; method='Direct market'
        if d<TRANSITION:
            f=legacy; col=LEGACY.get(code); c=headers[col-1] if col and col<=len(headers) else None
            method='Historical proxy' if code in PROXIES else 'Historical direct'
            if code in ['DOM','EGTS']:
                hit=f[f.Date.eq(d)]; available=[]
                if len(hit)==1:
                    for i in ([8,9] if code=='DOM' else [9,8]):
                        if i<=len(headers) and num(hit.iloc[0][headers[i-1]]) is not None:
                            available.append((headers[i-1],num(hit.iloc[0][headers[i-1]])))
                c=available[0][0] if available else None
                if len(available)==2 and not same(available[0][1],available[1][1]): note='Historical DOM/EGTS conflict'
        else:
            f=master if master is not None else newer; c=DIRECT.get(code)
        hit=f[f.Date.eq(d)]; p=None
        status='UNMAPPED INDEX' if c is None else 'MISSING'
        if len(hit)>1: status='DUPLICATE DATE'
        elif len(hit)==1 and c in f.columns:
            p=num(hit.iloc[0][c]); status='AVAILABLE' if p is not None else 'MISSING'
        if note and status=='AVAILABLE': status='REVIEW'
        rows.append(dict(delivery_date=d,index_code=code,market=c,price=p,status=status,
                         source=f.attrs.get('source',''),method=method,review_note=note))
    return pd.DataFrame(rows)


def settlement_index_price(code,day_,legacy,newer,master,price_days,ice,asof):
    """One dated source policy for cash and counterparty settlement views."""
    code=market(code);day_=pd.Timestamp(day_).normalize();asof=pd.Timestamp(asof).normalize()
    price=None;kind='Unavailable';source='No supported price source';check='No supported dated source/forward mapping'
    if code in ['DOM','TCO'] and price_days is not None and not price_days.empty:
        source='Storage daily price report';check='Missing, duplicate or blocked daily price row'
        hits=price_days.loc[price_days.index_code.map(market).eq(code)&pd.to_datetime(price_days.delivery_date).eq(day_)]
        if len(hits)==1 and hits.iloc[0].status in ['Historical cash','Forward estimate']:
            price=num(hits.iloc[0].price);kind=txt(hits.iloc[0].price_kind)
    elif day_<asof or code in ['TENN219','ANRSE']:
        source='cash' if day_<TRANSITION else 'cash2'
        h=daily_prices(code,day_,day_,legacy,newer,None).iloc[0]
        if h.status=='MISSING' and day_>=TRANSITION and master is not None:
            source='Master cash file (cash2 missing)';h=daily_prices(code,day_,day_,legacy,newer,master).iloc[0]
        check=str(h.status)
        if h.status=='AVAILABLE':price=num(h.price);kind='Historical' if day_<asof else 'Dated cash cache'
    elif code=='REXZ3D' and day_.to_period('M')==asof.to_period('M'):
        source='ICE exact-date daily forward';price,_,check=storage_daily_reference(ice,code,day_,asof)
        if price is not None:kind='Forward'
    return dict(delivery_date=day_.date().isoformat(),index_code=code,expected_cash_column=DIRECT.get(code,''),price=price,price_kind=kind,source=source,status='AVAILABLE' if price is not None else 'BLOCKED',issue='' if price is not None else check)


def october_cash(ledger,legacy,newer,master,positive_buy,price_days=None,ice=None,asof=None):
    reports=[]; details=[];cache={}
    asof=pd.Timestamp(asof) if asof is not None else pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    cash=ledger[ledger.source_sheet.eq('cashblotter')]
    for _,r in cash.iterrows():
        # Invalid dates are retained in import flags; cannot assign a month.
        if pd.isna(r.start) or pd.isna(r.end) or r.end<r.start: continue
        a,b=max(r.start,OCT_START),min(r.end,OCT_END)
        if a>b: continue
        resolved=[]
        for delivery_date in pd.date_range(a,b):
            key=(market(r.index_code),delivery_date)
            if key not in cache:cache[key]=settlement_index_price(r.index_code,delivery_date,legacy,newer,master,price_days,ice,asof)
            result=cache[key];resolved.append(dict(delivery_date=delivery_date,price=result['price'],status=result['status'],source=result['source'],market=result['expected_cash_column'],price_kind=result['price_kind'],issue=result['issue']))
        prices=pd.DataFrame(resolved)
        premium=0.0 if pd.isna(r.premium) else r.premium
        issues=[]
        if num(r.daily_volume) is None: issues.append('Invalid volume')
        if r.premium_issue: issues.append('Invalid or uncached premium')
        if not txt(r.counterparty): issues.append('Missing counterparty')
        unpriced=int(prices.status.ne('AVAILABLE').sum())
        if unpriced: issues.append(f'{unpriced} days missing/duplicate/unmapped')
        complete=not issues
        volume=r.daily_volume*len(prices) if num(r.daily_volume) is not None else None
        average=float(prices.price.mean()) if complete else None
        unit_price=average+premium if complete else None
        amount=volume*unit_price if complete else None
        # Compare only the October portion against the saved unit price.
        saved=num(r.saved_transaction_price)
        legacy_comparison=volume*saved if saved is not None and volume is not None else None
        signed_amount=amount*(1 if positive_buy else -1) if amount is not None else None
        reports.append(dict(source_id=r.source_id,counterparty=r.counterparty,location=r.location,
                            index_code=r.index_code,original_pricing_label=r.pricing_label,
                            delivery_start=a,delivery_end=b,delivery_days=len(prices),signed_volume=volume,
                            system_index_average=average,premium_applied=premium,
                            premium_note='Blank premium treated as zero' if pd.isna(r.premium) and not r.premium_issue else '',
                            system_unit_price=unit_price,system_signed_amount=signed_amount,
                            worksheet_fixed_price_ignored=r.fixed_price,worksheet_saved_unit_price=saved,
                            worksheet_october_comparison=legacy_comparison*(1 if positive_buy else -1) if legacy_comparison is not None else None,
                            system_minus_saved_amount=(amount-legacy_comparison)*(1 if positive_buy else -1) if amount is not None and legacy_comparison is not None else None,
                            status='SYSTEM PRICED — SETTLEMENT STATUS UNVERIFIED' if complete else 'BLOCKED',
                            issue='; '.join(issues),source='; '.join(sorted(set(prices.source))),
                            pricing_rule='System daily index plus premium; no fixed-price fallback'))
        for _,p in prices.iterrows():
            valid=p.status=='AVAILABLE' and num(r.daily_volume) is not None and not r.premium_issue
            details.append(dict(source_id=r.source_id,counterparty=r.counterparty,
                                delivery_date=p.delivery_date,index_code=r.index_code,market=p.market,
                                system_index=p.price,premium=premium,daily_volume=r.daily_volume,
                                daily_amount=r.daily_volume*(p.price+premium)*(1 if positive_buy else -1) if valid else None,
                                price_status=p.status,source=p.source,price_kind=p.price_kind,price_issue=p.issue))
    return pd.DataFrame(reports),pd.DataFrame(details)


def settlement_reports(report, detail, confirmed_through=None, confirmed_markets=()):
    """Attribute amounts by delivery day; confirmation is explicit, never inferred."""
    report, detail = report.copy(), detail.copy()
    cutoff = day(confirmed_through) if confirmed_through is not None else pd.NaT
    markets = set(confirmed_markets)
    if detail.empty:
        return report, detail, pd.DataFrame()
    delivery = pd.to_datetime(detail.delivery_date)
    confirmed = (delivery.le(cutoff) & detail.index_code.isin(markets)
                 if pd.notna(cutoff) else pd.Series(False, index=detail.index))
    valid = detail.price_status.eq('AVAILABLE') & detail.daily_amount.notna()
    detail['settlement_status'] = 'ESTIMATE — UNCONFIRMED'
    detail.loc[confirmed & valid, 'settlement_status'] = 'USER-CONFIRMED PRICE'
    detail.loc[~valid, 'settlement_status'] = 'BLOCKED'
    detail['confirmation_cutoff'] = cutoff
    detail['confirmed_signed_amount'] = detail.daily_amount.where(confirmed & valid)
    detail['estimated_signed_amount'] = detail.daily_amount.where(~confirmed & valid)
    report['confirmed_signed_amount'] = float('nan')
    report['estimated_signed_amount'] = float('nan')
    for i, r in report.iterrows():
        days = detail[detail.source_id.eq(r.source_id)]
        n = int(days.settlement_status.eq('USER-CONFIRMED PRICE').sum())
        report.loc[i, 'confirmed_price_days'] = n
        report.loc[i, 'estimated_price_days'] = int(days.settlement_status.eq('ESTIMATE — UNCONFIRMED').sum())
        report.loc[i, 'confirmation_cutoff'] = cutoff
        if r.status != 'BLOCKED':
            report.loc[i, 'confirmed_signed_amount'] = days.confirmed_signed_amount.sum()
            report.loc[i, 'estimated_signed_amount'] = days.estimated_signed_amount.sum()
            report.loc[i, 'status'] = ('SYSTEM PRICED — USER-CONFIRMED PRICES' if n == len(days)
                                     else 'SYSTEM PRICED — MIXED CONFIRMED/ESTIMATE' if n
                                     else 'SYSTEM PRICED — ESTIMATE')
    summary = []
    for cp, g in report.groupby('counterparty', dropna=False):
        blocked = g.status.eq('BLOCKED')
        confirmed_amount = g.loc[~blocked, 'confirmed_signed_amount'].sum()
        estimated_amount = g.loc[~blocked, 'estimated_signed_amount'].sum()
        summary.append(dict(counterparty=cp, trades=len(g), blocked_trades=int(blocked.sum()),
                            confirmed_signed_amount=confirmed_amount,
                            estimated_signed_amount=estimated_amount,
                            priced_subset_signed_amount=g.system_signed_amount.sum(min_count=1),
                            complete_signed_amount=None if blocked.any() else g.system_signed_amount.sum(),
                            estimated_price_days=int(g.estimated_price_days.sum()),
                            status='INCOMPLETE — total withheld' if blocked.any() else
                                   'ESTIMATE INCLUDED' if g.estimated_price_days.sum() else
                                   'USER-CONFIRMED PRICES'))
    return report, detail, pd.DataFrame(summary)


def invoice_reconciliation(summary, entries):
    """One signed commodity subtotal per counterparty, preserving missing inputs."""
    out = summary.copy()
    if entries.counterparty.duplicated().any():
        raise ValueError('Use one invoice subtotal per counterparty; combine multiple invoices first.')
    unknown = set(entries.counterparty) - set(summary.counterparty)
    if unknown:
        raise ValueError('Invoice counterparties absent from this cash report: ' + ', '.join(map(str, sorted(unknown))))
    out = out.merge(entries[['counterparty', 'invoice_reference', 'invoice_signed_amount']],
                    on='counterparty', how='left', validate='one_to_one')
    original = out.invoice_signed_amount
    values = pd.to_numeric(original, errors='coerce')
    invalid = original.notna() & (values.isna() | ~values.map(lambda x: math.isfinite(x) if pd.notna(x) else True))
    if invalid.any():
        raise ValueError('Invoice amounts must be finite numbers or blank.')
    out['invoice_signed_amount'] = values
    out['system_minus_invoice'] = out.complete_signed_amount - values
    out['reconciliation_status'] = 'NO INVOICE ENTERED'
    present = values.notna()
    out.loc[present & out.blocked_trades.gt(0), 'reconciliation_status'] = 'BLOCKED — pricing incomplete'
    ready = present & out.blocked_trades.eq(0)
    out.loc[ready, 'reconciliation_status'] = 'PROVISIONAL — estimates included'
    final = ready & out.estimated_price_days.eq(0)
    out.loc[final, 'reconciliation_status'] = 'DIFFERENCE — review'
    out.loc[final & out.system_minus_invoice.abs().le(0.01), 'reconciliation_status'] = 'MATCHED — commodity subtotal'
    return out


def forward_price_coverage(allocations, classifications, ice, mappings):
    """Check explicit monthly product mappings at one ICE snapshot; never value trades."""
    latest = ice['TRADE DATE'].max()
    if pd.isna(latest):
        raise ValueError('No valid ICE trade date.')
    snapshot = ice[ice['TRADE DATE'].eq(latest)].copy()
    snapshot['_month'] = snapshot['STRIP'].dt.to_period('M').dt.to_timestamp()
    rows = []
    by_id = classifications.set_index('source_id')
    for _, a in allocations.iterrows():
        c = by_id.loc[a.source_id]
        cls = c.candidate_classification
        mapping = mappings.get(a.index_code)
        def resolve(component):
            if not component:
                return None, 'UNMAPPED — select a monthly ICE product', pd.DataFrame()
            if component['CONTRACT TYPE'] != 'F':
                return None, 'UNSUPPORTED — select a monthly futures component', pd.DataFrame()
            hits = snapshot[snapshot.CONTRACT.eq(component['CONTRACT'])
                            & snapshot.PRODUCT.eq(component['PRODUCT'])
                            & snapshot.HUB.eq(component['HUB'])
                            & snapshot['CONTRACT TYPE'].eq(component['CONTRACT TYPE'])
                            & snapshot['STRIP'].eq(a.month)]
            prices = hits['SETTLEMENT PRICE'].dropna().unique()
            if hits.empty:
                return None, 'MISSING — no price in newest snapshot', hits
            if hits['SETTLEMENT PRICE'].isna().any() or any(not math.isfinite(float(x)) for x in prices):
                return None, 'INVALID — non-finite or missing settlement', hits
            if len(prices) != 1:
                return None, 'CONFLICT — multiple settlements', hits
            return float(prices[0]), '', hits
        base, issue, matches = resolve(mapping)
        adjustment = None
        adjustment_mapping = mapping.get('_adjustment') if mapping else None
        reference_type = mapping.get('_reference_type','Monthly IF reference') if mapping else ''
        if reference_type == 'Daily cash proxy: IF + GDD adjustment':
            adjustment, second_issue, extra = resolve(adjustment_mapping)
            matches = pd.concat([matches, extra], ignore_index=True)
            if second_issue:
                issue = issue or 'ADJUSTMENT '+second_issue
        price = (base + adjustment if base is not None and adjustment is not None
                 else base if reference_type != 'Daily cash proxy: IF + GDD adjustment' else None)
        classification_issue = ('Unresolved pricing classification' if not cls or cls not in RULES
                                else 'PHYSB interpretation requires confirmation' if cls == 'PHYSB'
                                else 'Classification candidate; approval pending')
        rows.append(dict(month=a.month, source_id=a.source_id, counterparty=a.counterparty,
                         index_code=a.index_code, signed_volume=a.signed_volume,
                         candidate_classification=cls, classification_status=c.match_status,
                         classification_note=classification_issue,
                         ice_trade_date=latest, ice_contract=mapping['CONTRACT'] if mapping else '',
                         ice_product=mapping['PRODUCT'] if mapping else '',
                         ice_hub=mapping['HUB'] if mapping else '',
                         ice_contract_type=mapping['CONTRACT TYPE'] if mapping else '',
                         reference_type=reference_type, base_component_price=base,
                         adjustment_contract=adjustment_mapping['CONTRACT'] if adjustment_mapping else '',
                         adjustment_component_price=adjustment,
                         ice_settlement_price=price, price_status=issue or 'AVAILABLE — mapped product',
                         source_files=', '.join(sorted(matches.source_file.dropna().astype(str).unique()))
                                      if 'source_file' in matches else '',
                         valuation_status='NOT VALUED — product interpretation and classification approval required'))
    return pd.DataFrame(rows)


def value_forward_trades(coverage, ledger, positive_buy=True):
    """Undiscounted commodity value relative to the selected same-location reference."""
    if coverage.empty:
        return pd.DataFrame()
    if coverage.duplicated(['source_id','month']).any():
        raise ValueError('Duplicate trade-month allocations; valuation withheld.')
    terms = ledger[ledger.source_sheet.eq('fwdblotter')][
        ['source_id','contract_price','start','end','daily_volume']]
    out = coverage.merge(terms,on='source_id',how='left',validate='many_to_one')
    results=[]
    accepted = {'MATCHED — REVIEW','SHARED — CONSISTENT CLASSIFICATION'}
    for _, r in out.iterrows():
        row = r.to_dict()
        price, term, volume = num(r.ice_settlement_price), num(r.contract_price), num(r.signed_volume)
        cls = txt(r.candidate_classification)
        issues=[]
        if r.classification_status not in accepted:
            issues.append('Unresolved classification match')
        if cls not in {'FIXED','GDD','INDEX'}:
            issues.append('Unsupported classification or PHYSB interpretation unresolved')
        if r.price_status != 'AVAILABLE — mapped product' or price is None:
            issues.append('Missing or conflicting reference price')
        if term is None:
            issues.append('Missing or invalid contract price/premium')
        if volume is None:
            issues.append('Invalid volume')
        if cls == 'GDD' and r.reference_type != 'Daily cash proxy: IF + GDD adjustment':
            issues.append('GDD requires a daily cash reference')
        if cls == 'INDEX' and r.reference_type != 'Monthly IF reference':
            issues.append('INDEX requires a monthly IF reference')
        row.update(valuation_status='BLOCKED' if issues else 'INDICATIVE — undiscounted commodity only',
                   valuation_issue='; '.join(issues), contract_unit_price=None,
                   directional_volume=None, projected_signed_contract_amount=None,
                   projected_net_cash_receipt=None, signed_market_reference_amount=None,
                   indicative_mtm=None, outright_price_exposure_volume=None,
                   valuation_rule='Fixed: V × (reference − fixed); floating: −V × premium. V is purchases positive.',
                   assumptions='Selected reference applies at the same location with no additional delivery adjustment; monthly reference is flat over allocated delivery days. Excludes fees, fuel, tax, credit and discounting.')
        if not issues:
            v = volume * (1 if positive_buy else -1)
            unit = term if cls == 'FIXED' else price + term
            contract_amount = v * unit
            reference_amount = v * price
            row.update(contract_unit_price=unit, directional_volume=v,
                       projected_signed_contract_amount=contract_amount,
                       projected_net_cash_receipt=-contract_amount,
                       signed_market_reference_amount=reference_amount,
                       indicative_mtm=reference_amount-contract_amount,
                       outright_price_exposure_volume=v if cls == 'FIXED' else 0.0)
        results.append(row)
    return pd.DataFrame(results)


def filter_forward_valuation(values,selections):
    """OR within each selection, AND across fields; empty selection means all."""
    result=values.copy()
    allowed={'counterparty','month','index_code','candidate_classification','valuation_status'}
    for field,selected in selections.items():
        if field not in allowed:raise ValueError('Unsupported forward valuation filter: '+field)
        if not selected:continue
        labels=pd.to_datetime(result[field]).dt.strftime('%Y-%m') if field=='month' else result[field].map(lambda x:txt(x) or '(blank)')
        result=result.loc[labels.isin(selected)]
    return result


def summarize_forward_values(values, groups):
    rows=[]
    for key,g in values.groupby(groups,dropna=False):
        key = key if isinstance(key,tuple) else (key,)
        row=dict(zip(groups,key))
        incomplete=g.valuation_status.eq('BLOCKED').any()
        row.update(trades=g.source_id.nunique(), trade_months=len(g),
                   blocked_allocations=int(g.valuation_status.eq('BLOCKED').sum()),
                   priced_subset_mtm=g.indicative_mtm.sum(min_count=1),
                   complete_mtm=None if incomplete else g.indicative_mtm.sum(),
                   projected_signed_contract_amount=None if incomplete else g.projected_signed_contract_amount.sum(),
                   projected_net_cash_receipt=None if incomplete else g.projected_net_cash_receipt.sum(),
                   outright_price_exposure_volume=None if incomplete else g.outright_price_exposure_volume.sum(),
                   status='INCOMPLETE — full totals withheld' if incomplete else 'INDICATIVE — commodity only')
        rows.append(row)
    return pd.DataFrame(rows)


def forward_scenario(values, shocks):
    """Parallel monthly reference changes by blotter index, holding contract terms fixed."""
    out = values.copy()
    for shock in shocks.values():
        if num(shock) is None:
            raise ValueError('Scenario price changes must be finite numbers.')
    out['reference_price_change'] = out.index_code.map(shocks).fillna(0.0).astype(float)
    valid = out.valuation_status.ne('BLOCKED') & out.indicative_mtm.notna()
    fixed = out.candidate_classification.eq('FIXED')
    out['scenario_reference_price'] = (out.ice_settlement_price + out.reference_price_change).where(valid)
    out['scenario_mtm_change'] = (out.outright_price_exposure_volume * out.reference_price_change).where(valid)
    out['scenario_mtm'] = (out.indicative_mtm + out.scenario_mtm_change).where(valid)
    out['scenario_contract_amount_change'] = (out.directional_volume * out.reference_price_change).where(~fixed,0.0).where(valid)
    out['scenario_projected_signed_contract_amount'] = (out.projected_signed_contract_amount + out.scenario_contract_amount_change).where(valid)
    out['scenario_projected_net_cash_receipt'] = -out.scenario_projected_signed_contract_amount
    return out


def summarize_scenario(scenario, groups):
    rows=[]
    for key,g in scenario.groupby(groups,dropna=False):
        key=key if isinstance(key,tuple) else (key,)
        blocked=g.valuation_status.eq('BLOCKED') | g.scenario_mtm.isna()
        row=dict(zip(groups,key))
        row.update(blocked_allocations=int(blocked.sum()),
                   base_mtm=None if blocked.any() else g.indicative_mtm.sum(),
                   scenario_mtm=None if blocked.any() else g.scenario_mtm.sum(),
                   mtm_change=None if blocked.any() else g.scenario_mtm_change.sum(),
                   scenario_projected_signed_contract_amount=None if blocked.any() else g.scenario_projected_signed_contract_amount.sum(),
                   status='INCOMPLETE — full totals withheld' if blocked.any() else 'INDICATIVE — commodity only')
        rows.append(row)
    return pd.DataFrame(rows)


def storage_physical_bridge(reports):
    """Inventory-boundary Dth, not delivered supply or automatic fuel gross-up."""
    out=[]
    accounts=reports.get('storage_account_inputs.csv',pd.DataFrame())
    forecast=reports.get('storage_withdrawal_schedule.csv',pd.DataFrame())
    if not accounts.empty and not forecast.empty:
        if accounts.account.duplicated().any():raise ValueError('Duplicate Columbia storage account.')
        factors=accounts.set_index('account').dth_per_mcf
        for month,rows in forecast.groupby('month',sort=True):
            if rows.account.duplicated().any() or set(rows.account)!=set(accounts.account):raise ValueError('Storage bridge requires every Columbia account for each month.')
            inj=wd=0.;missing=False
            for _,r in rows.iterrows():
                factor=storage_number(factors.get(r.account),'Columbia Dth/Mcf',0,True)
                if factor is None or factor<=0:missing=True;continue
                inj+=storage_number(r.injections_mcf,'Columbia injections',0)*factor
                wd+=storage_number(r.withdrawals_mcf,'Columbia withdrawals',0)*factor
            out.append(dict(month=str(month),index_code='TCO',storage_injections=None if missing else inj,storage_withdrawals=None if missing else wd,
                source_report='storage_withdrawal_schedule.csv',basis='Inventory-boundary Dth; delivery losses and transport not added',
                status='BLOCKED — conversion missing' if missing else 'PROVISIONAL — confirm delivery basis and entitlements'))
    deo=reports.get('storage_deo_inventory_forecast.csv',pd.DataFrame())
    if not deo.empty:
        if deo.month.duplicated().any():raise ValueError('Duplicate DEO storage month.')
        for _,r in deo.iterrows():
            out.append(dict(month=str(r.month),index_code='DOM',storage_injections=storage_number(r.injections_dth,'DEO injections',0),
                storage_withdrawals=storage_number(r.withdrawals_dth,'DEO withdrawals',0),source_report='storage_deo_inventory_forecast.csv',
                basis='Inventory-boundary Dth; adjustments excluded from gas flows',
                status='PROVISIONAL — '+str(r.basis)+'; CIV, inventory and delivery basis require review'))
    return pd.DataFrame(out,columns=['month','index_code','storage_injections','storage_withdrawals','source_report','basis','status'])


def link_storage_balance_inputs(inputs,bridge):
    result=inputs.copy();result['month']=pd.to_datetime(result.month).dt.strftime('%Y-%m-01');result['index_code']=result.index_code.map(market)
    if bridge.duplicated(['month','index_code']).any():raise ValueError('Duplicate linked storage month/market.')
    source=bridge.set_index(['month','index_code'])
    for i,r in result.iterrows():
        if r.index_code not in ['DOM','TCO']:continue
        key=(r.month,r.index_code)
        if key not in source.index or str(source.loc[key,'status']).startswith('BLOCKED'):
            result.loc[i,['storage_injections','storage_withdrawals']]=float('nan')
        else:
            result.loc[i,['storage_injections','storage_withdrawals']]=source.loc[key,['storage_injections','storage_withdrawals']].to_numpy()
    return result


def winter_customer_demand_reference():
    return pd.DataFrame([dict(month=m,east_customer_demand_dth=e,west_customer_demand_dth=w,columbia_customer_demand_dth=c,
        source='User-supplied Dth customer demand; Usage row excluded because it is net of storage') for m,e,w,c in [
        ('2026-11-01',164128.,21557.,59641.),('2026-12-01',274359.,43567.,123306.),
        ('2027-01-01',310069.,52602.,134349.),('2027-02-01',272835.,45865.,98006.),('2027-03-01',243320.,39328.,101281.)]])


def retail_demand_defaults(inputs):
    """Fill missing winter retail demand; explicit edits, including zero, take precedence."""
    result=inputs.copy()
    reference=winter_customer_demand_reference().set_index('month')
    for i,r in result.iterrows():
        month=pd.Timestamp(r.month).strftime('%Y-%m-01');code=market(r.index_code)
        if pd.notna(r.additional_customer_demand) or month not in reference.index:continue
        source=reference.loc[month]
        if code=='DOM':result.loc[i,'additional_customer_demand']=source.east_customer_demand_dth+source.west_customer_demand_dth
        elif code=='TCO':result.loc[i,'additional_customer_demand']=source.columbia_customer_demand_dth
    return result


def delivered_pricing_defaults(inputs):
    """Apply user-approved delivery differentials without overwriting explicit prices or edits."""
    result=inputs.copy()
    for col in ['delivered_price_usd_dth','delivery_differential_usd_dth']:
        if col not in result:result[col]=float('nan')
    if 'delivered_price_source' not in result:result['delivered_price_source']=''
    defaults={'DOM':(.20,'Gas Daily index — Dominion delivered'),
              'TCO':(.30,'Gas Daily index — TCO/Columbia delivered')}
    for i,r in result.iterrows():
        code=market(r.index_code)
        if code not in defaults or pd.notna(r.delivered_price_usd_dth):continue
        differential,source=defaults[code]
        if pd.isna(r.delivery_differential_usd_dth):result.loc[i,'delivery_differential_usd_dth']=differential
        if not txt(r.delivered_price_source):result.loc[i,'delivered_price_source']=source
    return result


# Exact delivery labels observed in the supplied ledger; no numeric range guessing.
DELIVERY_LOCATIONS = {'deodlvd':'DOM','deodelivered':'DOM','anrconvoy':'DOM',
    **{k:'TCO' for k in ['coh22','coh231','coh23n2','coh233','coh234','coh235',
        'coh236','coh23n7','coh238','coh239','coh2345','coh2435','coh2439']}}
POOL_LABELS = {'domspool','tcopool','domsp','domgdd','tcogdd','domif','domindex',
    'domphysb','domspphysb','tcophysb','gddpool','pool','balmopool','ixpool','physbpool','fixed'}


def physical_delivery_classification(allocations):
    out=allocations.copy()
    kinds=[];books=[];points=[];reasons=[]
    def clean(value):return ''.join(c for c in txt(value).lower() if c.isalnum())
    for _,r in out.iterrows():
        loc=clean(r.get('location',''));label=clean(r.get('secondary_label',''));idx=clean(r.index_code)
        # Location is the delivery evidence; index_code is generally a pricing basis.
        evidence=loc or (label if label in DELIVERY_LOCATIONS else '') or (idx if idx in DELIVERY_LOCATIONS else '')
        if evidence in DELIVERY_LOCATIONS:
            kind='DELIVERED';book=DELIVERY_LOCATIONS[evidence];point=evidence;reason='Exact delivered location/code'
        elif loc in {'domspool','tcopool'} or (not loc and (label in POOL_LABELS or not label) and market(r.index_code) in ['DOM','TCO']):
            kind='POOL';book=market(r.index_code);point=loc or ('tcopool' if book=='TCO' else 'domspool');reason='Pool trade; excluded from delivered netting'
        else:
            kind='REVIEW';book=market(r.index_code);point=loc or label or idx;reason='Unmapped delivery point; excluded pending review'
        kinds.append(kind);books.append(book);points.append(point);reasons.append(reason)
    out['supply_scope']=kinds;out['delivery_book']=books;out['delivery_point']=points;out['classification_reason']=reasons
    return out


def delivered_balance_reports(allocations,inputs):
    audit=physical_delivery_classification(allocations)
    delivered=audit.loc[audit.supply_scope.eq('DELIVERED')].copy()
    if not delivered.empty:delivered['index_code']=delivered.delivery_book
    balance=physical_balance(delivered,inputs)
    balance['balance_scope']='DELIVERED'
    # Unknown points cannot silently become zero delivery purchases.
    uncertain=set(pd.to_datetime(audit.loc[audit.supply_scope.eq('REVIEW'),'month'])) if not audit.empty else set()
    mask=balance.month.isin(uncertain)
    balance.loc[mask,['net_surplus_shortfall','additional_purchase_required','surplus_available']]=float('nan')
    balance.loc[mask,'status']='REVIEW DELIVERY CLASSIFICATION'
    pool=audit.loc[audit.supply_scope.eq('POOL')]
    keys=['month','delivery_point','index_code']
    cols=['purchase_volume','sale_volume','net_volume']
    pools=pool.groupby(keys,as_index=False)[cols].sum() if not pool.empty else pd.DataFrame(columns=keys+cols)
    return balance,audit,pools


POOL_NAMES={'domspool':'DEO / Dominion pool','tcopool':'Columbia / TCO pool'}
POOL_ROUTE_COLUMNS=['route_id','source_pool','destination_pool','destination_description','start_date','end_date','departure_dth_day','receipt_dth_day','stage','source_reference']


def pool_allocation_defaults(profile):
    """One-time, editable owner-confirmed allocation; never books storage inventory."""
    result=json.loads(json.dumps(profile))
    if result.get('storage_allocation_defaults_version',0)>=1:return result
    routes=result.setdefault('routes',[])
    # Existing external TCO assignments require reconciliation, not automatic stacking.
    existing=any(r.get('source_pool')=='tcopool' and r.get('destination_pool')=='External'
        and str(r.get('start_date',''))[:10]<='2026-10-31' and str(r.get('end_date',''))[:10]>='2026-10-27' for r in routes)
    if not existing:
        routes.append(dict(route_id='TCO-STORAGE-20261027-31',source_pool='tcopool',destination_pool='External',
            destination_description='Columbia storage — pool allocation only; reconcile existing injections to monthly statements',
            start_date='2026-10-27',end_date='2026-10-31',departure_dth_day=1160.,receipt_dth_day=None,
            stage='Confirmed',source_reference='Owner confirmed October 27–31 TCO surplus is being injected; net inventory and account split await statements'))
    result['storage_allocation_defaults_version']=1
    return result


def validate_pool_profile(profile):
    if profile.get('schema')!='pool-molecule-balance-1' or type(profile.get('positive_volume_means_purchase')) is not bool:raise ValueError('Invalid pool molecule profile.')
    start=day(profile.get('start'));end=day(profile.get('end'))
    if pd.isna(start) or pd.isna(end) or end<start or (end-start).days>3660:raise ValueError('Choose a valid pool balance range of 10 years or less.')
    routes=pd.DataFrame(profile.get('routes',[]),columns=POOL_ROUTE_COLUMNS)
    for col in ['route_id','source_pool','destination_pool','destination_description','stage','source_reference']:routes[col]=routes[col].map(txt)
    if routes.route_id.eq('').any() or routes.route_id.str.casefold().duplicated().any():raise ValueError('Transport route IDs must be nonblank and unique.')
    if not routes.source_pool.isin(POOL_NAMES).all() or not routes.destination_pool.isin([*POOL_NAMES,'External']).all():raise ValueError('Choose a recognized source pool and a destination pool or External.')
    if routes.source_pool.eq(routes.destination_pool).any():raise ValueError('Source and destination pools must differ.')
    if not routes.stage.isin(['Planned','Confirmed']).all():raise ValueError('Transport stage must be Planned or Confirmed.')
    for col in ['start_date','end_date']:
        routes[col]=pd.to_datetime(routes[col],errors='coerce')
        if routes[col].isna().any() or routes[col].dt.tz is not None:raise ValueError('Transport dates are required and must be valid.')
        routes[col]=routes[col].dt.normalize()
    if (routes.end_date<routes.start_date).any():raise ValueError('Transport end date precedes start date.')
    for col in ['departure_dth_day','receipt_dth_day']:
        raw=routes[col].copy();routes[col]=routes[col].map(num).astype(float)
        if (raw.map(lambda x:bool(txt(x))) & routes[col].isna()).any():raise ValueError('Transport quantities must be finite numeric values or blank.')
    if routes.departure_dth_day.isna().any() or routes.departure_dth_day.le(0).any():raise ValueError('Transport departure must be a positive finite Dth/day quantity.')
    internal=routes.destination_pool.ne('External')
    if (internal & (routes.receipt_dth_day.isna()|routes.receipt_dth_day.lt(0)|routes.receipt_dth_day.gt(routes.departure_dth_day))).any():raise ValueError('Pool-to-pool receipt Dth/day must be supplied, nonnegative and no greater than departure; account for fuel once.')
    if (routes.receipt_dth_day.notna() & (routes.receipt_dth_day.lt(0)|routes.receipt_dth_day.gt(routes.departure_dth_day))).any():raise ValueError('Receipt volume cannot be negative or exceed departure volume.')
    if (routes.stage.eq('Confirmed') & routes.source_reference.eq('')).any():raise ValueError('Confirmed transport requires a source reference.')
    if (routes.destination_pool.eq('External') & routes.destination_description.eq('')).any():raise ValueError('Describe the external transport destination.')
    return start,end,routes


def pool_molecule_reports(ledger,flags,profile,asof):
    profile=pool_allocation_defaults(profile)
    start,end,routes=validate_pool_profile(profile);today=pd.Timestamp(asof).normalize();start=max(start,today)
    columns=['date','month','pool','forward_purchases_dth','forward_sales_dth','cash_purchases_dth','cash_sales_dth','confirmed_transport_out_dth','confirmed_transport_in_dth','planned_transport_out_dth','planned_transport_in_dth']
    dates=pd.date_range(start,end) if start<=end else pd.DatetimeIndex([])
    daily=pd.DataFrame([(d,d.to_period('M').start_time,p,*([0.]*8)) for d in dates for p in POOL_NAMES],columns=columns)
    allocations=monthly_volumes(ledger,start,end,profile['positive_volume_means_purchase'],all_physical=True) if len(dates) else pd.DataFrame()
    audit=physical_delivery_classification(allocations)
    if ledger.source_id.duplicated().any():raise ValueError('Duplicate ledger source IDs prevent pool balancing.')
    invalid=int(active_unallocatable_records(ledger,flags).sum())
    unknown=pd.DataFrame();review_dates=set();trade_detail=[]
    if not audit.empty:
        unknown=audit.loc[audit.supply_scope.eq('REVIEW')].copy()
        ids=audit.source_id.unique();terms=ledger.set_index('source_id')
        for sid in ids:
            t=terms.loc[sid];a=audit.loc[audit.source_id.eq(sid)].iloc[0]
            first=max(start,t.start);last=min(end,t.end)
            if a.supply_scope=='REVIEW':review_dates.update(pd.date_range(first,last));continue
            if a.supply_scope!='POOL':continue
            if a.delivery_point not in POOL_NAMES:review_dates.update(pd.date_range(first,last));continue
            if t.source_sheet not in ['fwdblotter','cashblotter']:review_dates.update(pd.date_range(first,last));continue
            q=float(t.daily_volume)*(1 if profile['positive_volume_means_purchase'] else -1)
            mask=daily.pool.eq(a.delivery_point)&daily.date.between(first,last)
            prefix='forward' if t.source_sheet=='fwdblotter' else 'cash'
            daily.loc[mask,prefix+'_purchases_dth']+=max(q,0.)
            daily.loc[mask,prefix+'_sales_dth']+=max(-q,0.)
            trade_detail.append(dict(source_id=sid,counterparty=canonical_counterparty(t.counterparty),source_sheet=t.source_sheet,pool=a.delivery_point,index_code=t.index_code,start_date=first.date().isoformat(),end_date=last.date().isoformat(),signed_dth_day=q,purchase_dth=max(q,0.)*((last-first).days+1),sale_dth=max(-q,0.)*((last-first).days+1),classification_basis=a.classification_reason))
    for _,r in routes.iterrows():
        prefix='confirmed' if r.stage=='Confirmed' else 'planned'
        mask=daily.date.between(r.start_date,r.end_date)
        daily.loc[mask & daily.pool.eq(r.source_pool),prefix+'_transport_out_dth']+=r.departure_dth_day
        if r.destination_pool!='External':daily.loc[mask & daily.pool.eq(r.destination_pool),prefix+'_transport_in_dth']+=r.receipt_dth_day
    daily['forward_net_dth']=daily.forward_purchases_dth-daily.forward_sales_dth
    daily['cash_net_dth']=daily.cash_purchases_dth-daily.cash_sales_dth
    daily['booked_net_dth']=daily.forward_net_dth+daily.cash_net_dth
    daily['after_confirmed_transport_dth']=daily.booked_net_dth+daily.confirmed_transport_in_dth-daily.confirmed_transport_out_dth
    daily['after_all_transport_plans_dth']=daily.after_confirmed_transport_dth+daily.planned_transport_in_dth-daily.planned_transport_out_dth
    daily['classification_review']=daily.date.isin(review_dates)|bool(invalid)
    net=daily.after_all_transport_plans_dth
    daily['status']=['REVIEW — incomplete trade/location data' if bad else 'PLANNING BALANCED' if abs(q)<=1e-6 else 'SURPLUS — sell or assign transport' if q>0 else 'DEFICIT — buy or reduce commitments' for q,bad in zip(net,daily.classification_review)]
    daily['unassigned_surplus_dth']=net.clip(lower=0).where(~daily.classification_review)
    daily['uncovered_deficit_dth']=(-net).clip(lower=0).where(~daily.classification_review)
    monthly=[]
    sumcols=columns[3:]+['forward_net_dth','cash_net_dth','booked_net_dth','after_confirmed_transport_dth','after_all_transport_plans_dth']
    for (month,pool),g in daily.groupby(['month','pool'],sort=True):
        bad=g.classification_review.any();pos=int(g.after_all_transport_plans_dth.gt(1e-6).sum());neg=int(g.after_all_transport_plans_dth.lt(-1e-6).sum())
        row=dict(month=month.date().isoformat(),pool=pool,pool_name=POOL_NAMES[pool],period_start=g.date.min().date().isoformat(),period_end=g.date.max().date().isoformat(),days_in_scope=len(g))
        row.update({c:float(g[c].sum()) for c in sumcols})
        row.update(average_planning_net_dth_day=None if bad else float(g.after_all_transport_plans_dth.mean()),surplus_days=None if bad else pos,deficit_days=None if bad else neg,
            unassigned_surplus_dth=None if bad else float(g.unassigned_surplus_dth.sum()),uncovered_deficit_dth=None if bad else float(g.uncovered_deficit_dth.sum()),
            peak_daily_surplus_dth=None if bad else float(g.unassigned_surplus_dth.max()),peak_daily_deficit_dth=None if bad else float(g.uncovered_deficit_dth.max()),
            status='REVIEW — complete balance withheld' if bad else 'MIXED DAILY SURPLUS / DEFICIT' if pos and neg else 'SURPLUS — sell or assign transport' if pos else 'DEFICIT — buy or reduce commitments' if neg else 'PLANNING BALANCED',
            scope='Contractual daily molecules from as-of date inclusive; forward plus cash trades, less transport departures plus explicit pool receipts. Known-trade subtotal columns remain visible during review. Planned transport is not execution. No daily inventory carry, nomination approval, capacity check, delivered-demand offset, or inter-pool netting. Prices and MTM do not determine molecule balance.')
        monthly.append(row)
    status=pd.DataFrame([dict(as_of=today.date().isoformat(),period_start=start.date().isoformat(),period_end=end.date().isoformat(),status='NO FUTURE DAYS IN SAVED RANGE' if not len(dates) else 'REVIEW' if invalid or review_dates else 'CONTRACTUAL PLANNING VIEW',unallocatable_records=invalid,unmapped_trade_months=len(unknown),note='Current day uses full contractual quantity, not intraday actual remaining gas. Pool receipts from transport are explicit net receipts; external destinations are not credited to delivered demand automatically.')])
    reports={'pool_molecule_balance_monthly.csv':pd.DataFrame(monthly),'pool_molecule_balance_daily.csv':daily,'pool_molecule_trade_detail.csv':pd.DataFrame(trade_detail),'pool_molecule_location_review.csv':unknown,'pool_molecule_transport_inputs.csv':routes,'pool_molecule_status.csv':status}
    return reports,int(daily.classification_review.sum())


def render_pool_molecules(ledger,flags):
    st.subheader('Forward molecule balance at physical pools')
    st.caption('Balance the DEO/Dominion and Columbia/TCO pools independently. Forward purchases minus sales, plus in-month cash deals, determine booked molecules. Transport commitments consume gas at the source pool; explicit net receipts add gas at another pool. All pool trades count as molecules regardless of pricing type. Delivered customer demand remains in the separate delivered balance.')
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    ends=ledger.loc[ledger.source_sheet.eq('fwdblotter')&ledger.end.ge(today),'end'];last=ends.max() if not ends.empty else today+pd.offsets.MonthEnd(0)
    saved=st.session_state.get('_pool_molecule_saved',dict(schema='pool-molecule-balance-1',start=today.to_period('M').start_time.date().isoformat(),end=last.date().isoformat(),positive_volume_means_purchase=True,routes=[]))
    upload=st.file_uploader('Restore pool balance settings JSON',type=['json'],key='pool_settings_restore')
    if upload is not None and st.button('Load pool balance settings',key='pool_settings_load'):
        try:
            loaded=json.loads(upload.getvalue());validate_pool_profile(loaded)
            st.session_state['_pool_molecule_saved']=loaded;st.session_state['_pool_molecule_revision']=st.session_state.get('_pool_molecule_revision',0)+1
            saved=loaded
        except (ValueError,TypeError,KeyError) as exc:st.error('Pool settings not loaded: '+str(exc))
    saved=pool_allocation_defaults(saved)
    rev=str(st.session_state.get('_pool_molecule_revision',0))
    period=st.date_input('Pool balance delivery range',value=(pd.Timestamp(saved['start']).date(),pd.Timestamp(saved['end']).date()),key='pool_range_'+rev)
    positive=st.checkbox('Pool balance: positive blotter volume means purchase',value=saved['positive_volume_means_purchase'],key='pool_sign_'+rev)
    st.caption('Current-month output includes today through the selected end date; earlier days are excluded. Today uses full contractual quantities. This is not a measured inventory or actual nomination report.')
    st.markdown('**Transport assignments — Dth per day**')
    st.caption('Leave this table empty until gas is assigned to transport. Enter each movement once, only if it is not already booked as a pool sale/purchase. Planned and confirmed movements are separated. For a pool-to-pool transfer, enter net receipt Dth/day after fuel; no automatic fuel rate is assumed. External destinations do not automatically reduce delivered customer demand in the other tab.')
    seed=pd.DataFrame(saved['routes'],columns=POOL_ROUTE_COLUMNS)
    for col in ['start_date','end_date']:seed[col]=pd.to_datetime(seed[col])
    for col in ['departure_dth_day','receipt_dth_day']:seed[col]=pd.to_numeric(seed[col],errors='coerce').astype(float)
    for col in ['route_id','source_pool','destination_pool','destination_description','stage','source_reference']:seed[col]=seed[col].astype('string')
    routes=st.data_editor(seed,num_rows='dynamic',hide_index=True,key='pool_routes_'+rev,column_config={
        'source_pool':st.column_config.SelectboxColumn('Source pool',options=list(POOL_NAMES),required=True),
        'destination_pool':st.column_config.SelectboxColumn('Destination pool',options=[*POOL_NAMES,'External'],required=True),
        'stage':st.column_config.SelectboxColumn('Stage',options=['Planned','Confirmed'],default='Planned',required=True),
        'start_date':st.column_config.DateColumn('Start date',required=True),'end_date':st.column_config.DateColumn('End date',required=True),
        'departure_dth_day':st.column_config.NumberColumn('Gross departure Dth/day',min_value=0.,required=True),
        'receipt_dth_day':st.column_config.NumberColumn('Net pool receipt Dth/day',min_value=0.)})
    if len(period)!=2:return
    profile=dict(storage_allocation_defaults_version=1,schema='pool-molecule-balance-1',start=str(period[0]),end=str(period[1]),positive_volume_means_purchase=positive,routes=json.loads(routes.to_json(orient='records',date_format='iso')))
    try:
        reports,review=pool_molecule_reports(ledger,flags,profile,today);st.session_state['_run_reports'].update(reports)
        show(reports['pool_molecule_balance_monthly.csv'],'pool_molecule_balance_monthly.csv')
        st.caption('Read both surplus days and deficit days. A zero monthly net can still require purchases on some days and sales or transport on others. After all transport plans is a planning residual; after confirmed transport excludes unconfirmed plans. Forward and cash volumes remain separate columns.')
        with st.expander('Daily molecules and trade detail',expanded=False):
            show(reports['pool_molecule_balance_daily.csv'],'pool_molecule_balance_daily.csv')
            show(reports['pool_molecule_trade_detail.csv'],'pool_molecule_trade_detail.csv')
        if review:st.warning('Unmapped or unallocatable trades prevent a complete balance. Resolve source data before using surplus/deficit quantities.');show(reports['pool_molecule_location_review.csv'],'pool_molecule_location_review.csv')
        show(reports['pool_molecule_status.csv'],'pool_molecule_status.csv')
        st.session_state['_scheduled_storage_inputs']['pool_molecule_balance']=profile
        st.download_button('Download pool balance settings JSON',json.dumps(profile,indent=2),file_name='pool_molecule_balance_settings.json',mime='application/json',key='pool_settings_download')
        st.caption('The combined profile includes this range, sign convention and transport assignments. Re-export it after changes for scheduled runs. Transport availability, tariffs, fuel and nominations must be confirmed separately; this screen does not authorize transport.')
    except (ValueError,KeyError,TypeError) as exc:
        st.session_state['_scheduled_storage_inputs'].pop('pool_molecule_balance',None);st.error('Pool balance withheld: '+str(exc))


def physical_balance(allocations, inputs):
    fields=['additional_customer_demand','other_supply','storage_withdrawals','storage_injections']
    required={'month','index_code',*fields}
    if not required.issubset(inputs.columns):
        raise ValueError('Balance inputs require: '+', '.join(sorted(required)))
    x=inputs.copy()
    x['month']=pd.to_datetime(x.month,errors='coerce')
    if x.month.isna().any() or x.index_code.isna().any():
        raise ValueError('Month and market are required.')
    x['month']=x.month.dt.to_period('M').dt.to_timestamp()
    x['index_code']=x.index_code.map(market)
    if x.duplicated(['month','index_code']).any():
        raise ValueError('Use one input row per month and market.')
    for col in fields:
        original=x[col]
        x[col]=pd.to_numeric(original,errors='coerce')
        invalid=original.notna() & (x[col].isna() | x[col].lt(0) | x[col].isin([float('inf'),float('-inf')]))
        if invalid.any(): raise ValueError(col+' must be nonnegative finite volumes or blank.')
    if 'delivered_price_usd_dth' not in x:x['delivered_price_usd_dth']=float('nan')
    if 'delivered_price_source' not in x:x['delivered_price_source']=''
    for col in ['delivered_price_usd_dth','delivery_differential_usd_dth']:
        if col not in x:x[col]=float('nan')
        original=x[col]
        x[col]=pd.to_numeric(original,errors='coerce')
        if (original.notna() & (x[col].isna() | x[col].isin([float('inf'),float('-inf')]))).any():
            raise ValueError(col+' must be finite or blank; zero and negative values are allowed.')
    if allocations.empty:
        totals=pd.DataFrame(columns=['month','index_code','purchase_volume','sale_volume'])
    else:
        a=allocations.copy();a['month']=pd.to_datetime(a.month);a['index_code']=a.index_code.map(market)
        totals=a.groupby(['month','index_code'],as_index=False)[['purchase_volume','sale_volume']].sum()
    out=x.merge(totals,on=['month','index_code'],how='outer',validate='one_to_one')
    out[['purchase_volume','sale_volume']]=out[['purchase_volume','sale_volume']].fillna(0)
    complete=out[fields].notna().all(axis=1)
    out['total_supply']=out.purchase_volume+out.other_supply+out.storage_withdrawals
    out['total_requirement']=out.sale_volume+out.additional_customer_demand+out.storage_injections
    out['net_surplus_shortfall']=(out.total_supply-out.total_requirement).where(complete)
    out['additional_purchase_required']=(-out.net_surplus_shortfall).clip(lower=0)
    out['surplus_available']=out.net_surplus_shortfall.clip(lower=0)
    out['status']='INPUTS REQUIRED'
    out.loc[complete,'status']='BALANCED'
    out.loc[complete & out.net_surplus_shortfall.gt(0),'status']='SURPLUS'
    out.loc[complete & out.net_surplus_shortfall.lt(0),'status']='SHORTFALL'
    return out.sort_values(['month','index_code'])


SCENARIO_COLUMNS=['month','index_code','delivered_price_change_usd_dth','pool_price_change_usd_dth']


def validate_monthly_scenario(profile):
    if profile.get('schema')!='monthly-market-scenario-1':raise ValueError('Unsupported monthly scenario schema.')
    frame=pd.DataFrame(profile.get('rows',[]),columns=SCENARIO_COLUMNS)
    dates=pd.to_datetime(frame.month,errors='coerce')
    if dates.isna().any() or dates.dt.tz is not None:raise ValueError('Scenario delivery months must be valid dates.')
    frame['month']=dates.dt.to_period('M').dt.to_timestamp();frame['index_code']=frame.index_code.map(market)
    if not frame.index_code.isin(['DOM','TCO']).all() or frame.duplicated(['month','index_code']).any():raise ValueError('Use unique month/book rows for DOM or TCO.')
    for col in SCENARIO_COLUMNS[2:]:
        frame[col]=frame[col].map(num)
        if frame[col].isna().any():raise ValueError('Enter finite price changes; use zero for unchanged prices.')
    return frame


def monthly_scenario_defaults(reports):
    keys=set()
    for name in ['monthly_physical_balance.csv','physical_open_exposure_monthly.csv']:
        f=reports.get(name,pd.DataFrame())
        if not f.empty:keys.update((pd.Timestamp(r.month).to_period('M').start_time,market(r.index_code)) for _,r in f.iterrows() if market(r.index_code) in ['DOM','TCO'])
    return dict(schema='monthly-market-scenario-1',rows=[dict(month=m.date().isoformat(),index_code=c,delivered_price_change_usd_dth=0.,pool_price_change_usd_dth=0.) for m,c in sorted(keys)])


def monthly_market_scenario(reports,profile):
    inputs=validate_monthly_scenario(profile)
    expected=validate_monthly_scenario(monthly_scenario_defaults(reports))
    required=set(zip(expected.month,expected.index_code));provided=set(zip(inputs.month,inputs.index_code))
    if not required.issubset(provided):raise ValueError('Scenario lacks one or more current physical month/book rows; enter zero for unchanged months.')
    source={}
    for name in ['monthly_physical_balance.csv','physical_open_exposure_monthly.csv','physical_trade_allocations.csv','forward_trade_valuation.csv']:
        f=reports.get(name,pd.DataFrame()).copy()
        if not f.empty:f['month']=pd.to_datetime(f.month).dt.to_period('M').dt.to_timestamp()
        source[name]=f
    balance=source['monthly_physical_balance.csv'];exposure=source['physical_open_exposure_monthly.csv']
    audit=source['physical_trade_allocations.csv'];values=source['forward_trade_valuation.csv']
    for f in [balance,exposure]:
        if not f.empty and f.duplicated(['month','index_code']).any():raise ValueError('Duplicate physical month/book source rows.')
    if not audit.empty and audit.duplicated(['month','source_id']).any():raise ValueError('Duplicate physical allocation IDs in scenario.')
    rows=[];detail=[]
    for _,s in inputs.iterrows():
        b=balance.loc[balance.month.eq(s.month)&balance.index_code.map(market).eq(s.index_code)] if not balance.empty else balance
        e=exposure.loc[exposure.month.eq(s.month)&exposure.index_code.map(market).eq(s.index_code)] if not exposure.empty else exposure
        volume=price=cost=None;physical_status='NOT GENERATED — delivered balance and price required'
        if len(b)==1 and len(e)==1:
            br=b.iloc[0];er=e.iloc[0];volume=num(er.net_open_purchase_dth);price=num(er.forward_usd_dth);cost=num(er.net_open_cost_usd);net=num(br.net_surplus_shortfall)
            physical_status='INDICATIVE — fixed volumes; '+str(br.status)
            if br.get('balance_scope')!='DELIVERED' or net is None or volume is None or abs(volume+net)>1e-6 or not (br.status in ['BALANCED','SURPLUS','SHORTFALL'] or str(br.status).startswith('PROVISIONAL STORAGE — ')):
                volume=None;physical_status='BLOCKED — incomplete or non-delivered balance'
            if volume is None or price is None or cost is None or not str(er.status).startswith('INDICATIVE') or (volume is not None and price is not None and cost is not None and abs(cost-volume*price)>.01):
                cost=None;physical_status='BLOCKED — incomplete or inconsistent delivered forecast'
        pool=audit.loc[audit.month.eq(s.month)&audit.delivery_book.eq(s.index_code)&audit.supply_scope.eq('POOL')&audit.source_sheet.eq('fwdblotter')] if not audit.empty else audit
        base_mtm=change=0.;blocked=0
        for _,trade in pool.iterrows():
            candidates=values.loc[values.month.eq(s.month)&values.source_id.eq(trade.source_id)] if not values.empty else values
            mtm=sensitivity=None;status='BLOCKED — matching forward valuation required'
            if len(candidates)==1:
                v=candidates.iloc[0];mtm=num(v.indicative_mtm);sensitivity=num(v.get('outright_price_exposure_volume'));direction=num(v.get('directional_volume'))
                valid=(str(v.valuation_status).startswith('INDICATIVE') and mtm is not None and sensitivity is not None and direction is not None and abs(direction-(trade.purchase_volume-trade.sale_volume))<=1e-6 and market(v.index_code)==market(trade.index_code) and v.candidate_classification in ['FIXED','GDD'] and abs(sensitivity-(direction if v.candidate_classification=='FIXED' else 0.))<=1e-6)
                if valid:status='INDICATIVE — pool forward only'
                else:mtm=sensitivity=None
            delta=sensitivity*s.pool_price_change_usd_dth if sensitivity is not None else None
            if mtm is None:blocked+=1
            else:base_mtm+=mtm;change+=delta
            detail.append(dict(month=s.month.date().isoformat(),index_code=s.index_code,source_id=trade.source_id,candidate_classification=candidates.iloc[0].get('candidate_classification') if len(candidates)==1 else None,outright_price_exposure_dth=sensitivity,base_mtm_usd=mtm,pool_price_change_usd_dth=s.pool_price_change_usd_dth,mtm_change_usd=delta,scenario_mtm_usd=mtm+delta if mtm is not None else None,status=status))
        pool_status='NOT GENERATED — physical allocations required' if 'physical_trade_allocations.csv' not in reports else 'NO POOL FORWARD ALLOCATIONS' if pool.empty else 'BLOCKED — matching pool valuations incomplete' if blocked else 'INDICATIVE — pool forwards only'
        if pool.empty or blocked:base_mtm=change=None
        cost_change=volume*s.delivered_price_change_usd_dth if cost is not None else None
        rows.append(dict(month=s.month.date().isoformat(),index_code=s.index_code,
            delivered_price_change_usd_dth=s.delivered_price_change_usd_dth,pool_price_change_usd_dth=s.pool_price_change_usd_dth,
            net_open_purchase_dth=volume,base_delivered_price_usd_dth=price,scenario_delivered_price_usd_dth=price+s.delivered_price_change_usd_dth if cost is not None else None,
            base_procurement_cost_usd=cost,procurement_cost_change_usd=cost_change,scenario_procurement_cost_usd=cost+cost_change if cost is not None else None,
            delivered_status=physical_status,pool_forward_allocations=len(pool),blocked_pool_allocations=blocked,
            base_pool_forward_mtm_usd=base_mtm,pool_forward_mtm_change_usd=change,scenario_pool_forward_mtm_usd=base_mtm+change if base_mtm is not None else None,pool_status=pool_status))
    monthly=pd.DataFrame(rows);summary=[]
    if not monthly.empty:
        for code,f in list(monthly.groupby('index_code'))+[('TOTAL',monthly)]:
            cost_ok=f.scenario_procurement_cost_usd.notna().all()
            pool_ok=not f.pool_status.str.startswith(('BLOCKED','NOT GENERATED')).any()
            has_pool=f.pool_forward_allocations.sum()>0
            row=dict(index_code=code,months=len(f),delivered_status='INDICATIVE — selected months only' if cost_ok else 'INCOMPLETE — cost totals withheld',pool_status='NO POOL FORWARD ALLOCATIONS' if pool_ok and not has_pool else 'INDICATIVE — selected pool forwards only' if pool_ok else 'INCOMPLETE — pool MTM totals withheld')
            for col in ['base_procurement_cost_usd','procurement_cost_change_usd','scenario_procurement_cost_usd']:row[col]=float(f[col].sum()) if cost_ok else None
            for col in ['base_pool_forward_mtm_usd','pool_forward_mtm_change_usd','scenario_pool_forward_mtm_usd']:row[col]=float(f[col].sum()) if pool_ok and has_pool else None
            row['scope']='Independent delivered and pool $/Dth shifts; volumes, contract premiums, storage movements and basis differentials held fixed. Selected physical reporting months and matching pool forwards only. Procurement cost and hedge MTM are separate, not combined or realized P&L.'
            summary.append(row)
    return {'monthly_market_scenario_inputs.csv':inputs,'monthly_market_scenario.csv':monthly,'monthly_market_scenario_summary.csv':pd.DataFrame(summary),'monthly_market_scenario_pool_detail.csv':pd.DataFrame(detail)}


def render_monthly_market_scenario():
    st.subheader('Monthly delivered and pool price scenario')
    st.caption('Enter independent $/Dth changes for each month and utility. All changes default to zero. Delivered changes affect open procurement costs; pool changes affect matching forward hedge MTM. Enter both columns when both prices move. Volumes, delivery differentials, contract terms and storage movements stay fixed; this is price stress, not a weather-driven demand model.')
    reports=st.session_state['_run_reports'];seed=validate_monthly_scenario(monthly_scenario_defaults(reports))
    saved=st.session_state.get('_monthly_scenario_saved')
    if saved:
        prior=validate_monthly_scenario(saved).set_index(['month','index_code']);seed=seed.set_index(['month','index_code'])
        overlap=seed.index.intersection(prior.index);seed.loc[overlap,SCENARIO_COLUMNS[2:]]=prior.loc[overlap,SCENARIO_COLUMNS[2:]];seed=seed.reset_index()
    entries=st.data_editor(seed,hide_index=True,disabled=['month','index_code'],key='monthly_market_scenario_'+str(st.session_state.get('_monthly_scenario_revision',0)),column_config={
        'month':st.column_config.DateColumn('Delivery month'),'delivered_price_change_usd_dth':st.column_config.NumberColumn('Delivered price change ($/Dth)',step=.05,format='%.3f'),
        'pool_price_change_usd_dth':st.column_config.NumberColumn('Pool price change ($/Dth)',step=.05,format='%.3f')})
    profile=dict(schema='monthly-market-scenario-1',rows=json.loads(entries.to_json(orient='records',date_format='iso')))
    try:
        result=monthly_market_scenario(reports,profile);reports.update(result)
        show(result['monthly_market_scenario_summary.csv'],'monthly_market_scenario_summary.csv')
        show(result['monthly_market_scenario.csv'],'monthly_market_scenario.csv')
        with st.expander('Matching pool forward scenario detail'):show(result['monthly_market_scenario_pool_detail.csv'],'monthly_market_scenario_pool_detail.csv')
        st.session_state['_scheduled_storage_inputs']['monthly_market_scenario']=profile
        st.caption('Positive procurement change means higher cost; positive pool MTM change means higher hedge value. These are separate outputs, not combined P&L. GDD floating-price hedges have zero outright-price MTM sensitivity under this parallel reference assumption; basis and contract premiums are held fixed. Missing valuations withhold pool totals. Provisional storage remains provisional. Re-export the combined JSON to retain these edits and use them in scheduled runs.')
    except (ValueError,KeyError,TypeError) as exc:
        st.session_state['_scheduled_storage_inputs'].pop('monthly_market_scenario',None)
        st.error('Monthly scenario withheld: '+str(exc))


def physical_open_exposure(balance,ice,mappings,asof):
    """Indicative cash required for open physical volumes; not trade MTM or full P&L."""
    today=pd.Timestamp(asof).normalize();rows=[]
    for _,r in balance.iterrows():
        month=pd.Timestamp(r.month).to_period('M').start_time;code=market(r.index_code)
        volume=storage_number(r.net_surplus_shortfall,'Net physical balance',optional=True)
        valid_status=str(r.status) in ['BALANCED','SURPLUS','SHORTFALL'] or str(r.status).startswith('PROVISIONAL STORAGE — ')
        price=snapshot=pool_price=differential=None;method='Unavailable';source='';pool_note='';note='Unsupported market; no automatic proxy'
        if month<today.to_period('M').start_time:note='Historical delivery month; forward valuation withheld'
        elif r.get('balance_scope','')=='DELIVERED':
            direct=storage_number(r.get('delivered_price_usd_dth'),'Delivered price',optional=True)
            differential=storage_number(r.get('delivery_differential_usd_dth'),'Delivery differential',optional=True)
            source=txt(r.get('delivered_price_source',''))
            if direct is not None and differential is not None:
                note='Choose either a full delivered price or a delivery differential; clear the other field'
            elif direct is not None:
                method='Direct delivered forecast'
                note='User-entered delivered Gas Daily forecast: '+source if source else 'Delivered forecast source required'
                if source:price=direct
            else:
                method='Pool reference plus delivery differential'
                if code in ['DOM','TCO']:
                    pool_price,snapshot,pool_note=storage_forward_reference(ice,code,month,mappings,today)
                if differential is None:note='Delivery differential required; blank is not zero'
                elif not source:note='Delivery differential source required'
                elif pool_price is None:note='Pool reference unavailable: '+(pool_note or 'Unsupported market')
                else:
                    price=pool_price+differential
                    note='Delivered Gas Daily forecast = pool reference + delivery differential; '+source
        elif code in ['DOM','TCO']:price,snapshot,note=storage_forward_reference(ice,code,month,mappings,today)
        if volume is None or not valid_status:status='BLOCKED — incomplete physical balance'
        elif price is None:status='BLOCKED — '+note
        else:status='INDICATIVE — provisional storage' if str(r.status).startswith('PROVISIONAL') else 'INDICATIVE'
        ready=not status.startswith('BLOCKED');open_volume=-volume if volume is not None and valid_status else None
        purchases=max(open_volume,0.) if open_volume is not None else None
        surplus=max(-open_volume,0.) if open_volume is not None else None
        rows.append(dict(month=month.date().isoformat(),index_code=code,net_open_purchase_dth=open_volume,
            purchase_required_dth=purchases,surplus_available_dth=surplus,forward_usd_dth=price,
            pool_reference_usd_dth=pool_price,delivery_differential_usd_dth=differential,price_method=method,
            delivered_price_source=source,pool_reference=pool_note,
            purchase_cost_usd=purchases*price if ready else None,surplus_value_usd=surplus*price if ready else None,
            net_open_cost_usd=open_volume*price if ready else None,
            ice_trade_date=None if snapshot is None or pd.isna(snapshot) else pd.Timestamp(snapshot).date().isoformat(),
            reference=note,physical_balance_status=r.status,status=status))
    monthly=pd.DataFrame(rows)
    if monthly.empty:raise ValueError('Physical exposure requires balance rows.')
    if monthly.duplicated(['month','index_code']).any():raise ValueError('Duplicate month/market in physical exposure.')
    totals=[];scenarios=[]
    groups=[(code,frame) for code,frame in monthly.groupby('index_code',sort=True)]+[('TOTAL',monthly)]
    for code,frame in groups:
        blocked=int(frame.status.str.startswith('BLOCKED').sum());complete=blocked==0
        provisional=bool(frame.physical_balance_status.astype(str).str.startswith('PROVISIONAL').any())
        status='INCOMPLETE — cost totals withheld' if blocked else 'INDICATIVE — provisional storage' if provisional else 'INDICATIVE'
        def total(col):
            volume_column=col in ['net_open_purchase_dth','purchase_required_dth','surplus_available_dth']
            usable=frame[col].notna().all() if volume_column else complete
            return float(frame[col].sum()) if usable else None
        totals.append(dict(index_code=code,months=len(frame),blocked_months=blocked,net_open_purchase_dth=total('net_open_purchase_dth'),
            purchase_required_dth=total('purchase_required_dth'),surplus_available_dth=total('surplus_available_dth'),
            purchase_cost_usd=total('purchase_cost_usd'),surplus_value_usd=total('surplus_value_usd'),net_open_cost_usd=total('net_open_cost_usd'),status=status))
        for shock in [-1.,-.5,-.25,0.,.25,.5,1.]:
            change=total('net_open_purchase_dth')*shock if complete else None
            base=total('net_open_cost_usd')
            scenarios.append(dict(index_code=code,parallel_price_change_usd_dth=shock,base_net_open_cost_usd=base,
                net_open_cost_change_usd=change,scenario_net_open_cost_usd=base+change if complete else None,
                status=status,scope='Fixed open volumes; parallel DOM/TCO price shift; positive change increases net procurement cost; excludes revenue, existing trade MTM, transport/fuel and other costs'))
    return {'delivered_price_forecast.csv':monthly[['month','index_code','price_method','pool_reference_usd_dth','delivery_differential_usd_dth','forward_usd_dth','delivered_price_source','ice_trade_date','pool_reference','reference','status']].rename(columns={'forward_usd_dth':'delivered_forecast_usd_dth'}),
        'physical_open_exposure_monthly.csv':monthly,'physical_open_exposure_summary.csv':pd.DataFrame(totals),
        'physical_open_price_sensitivity.csv':pd.DataFrame(scenarios)},int(monthly.status.str.startswith('BLOCKED').sum())


INVOICE_COLUMNS=['invoice_id','counterparty','direction','delivery_month','invoice_date','due_date','amount_usd']
PAYMENT_COLUMNS=['payment_id','invoice_id','payment_date','amount_usd']


def invoice_aging_reports(profile,asof):
    """USD invoice ledger, payments and aging; amounts calculated in integer cents."""
    if profile.get('schema')!='counterparty-ledger-1':raise ValueError('Unsupported invoice ledger schema.')
    today=pd.Timestamp(asof).normalize()
    invoices=pd.DataFrame(profile.get('invoices',[]));payments=pd.DataFrame(profile.get('payments',[]))
    if 'delivery_month' not in invoices:invoices['delivery_month']=None
    for frame,columns in [(invoices,INVOICE_COLUMNS),(payments,PAYMENT_COLUMNS)]:
        if frame.empty:
            for col in columns:frame[col]=pd.Series(dtype='object')
        if not set(columns).issubset(frame.columns):raise ValueError('Missing invoice/payment columns: '+', '.join(columns))
        for col in [c for c in columns if c.endswith('_id')]:
            frame[col]=frame[col].map(txt)
            if frame[col].eq('').any():raise ValueError(col+' is required.')
        for col in [c for c in columns if c.endswith('_date')]:
            frame[col]=pd.to_datetime(frame[col],errors='coerce')
            if frame[col].isna().any():raise ValueError(col+' is required and must be a valid date.')
            frame[col]=frame[col].dt.normalize()
        cents=[]
        for value in frame.amount_usd:
            value=num(value)
            if value is None or value<=0 or abs(value*100-round(value*100))>1e-6:
                raise ValueError('Invoice and payment amounts must be positive USD amounts with at most two decimal places.')
            cents.append(int(round(value*100)))
        frame['_cents']=pd.Series(cents,index=frame.index,dtype='int64')
    raw_month=invoices.delivery_month
    months=pd.to_datetime(raw_month,errors='coerce')
    if (raw_month.map(lambda x:bool(txt(x))) & months.isna()).any():raise ValueError('Delivery month must be a valid date or blank.')
    invoices['delivery_month']=months.dt.to_period('M').dt.to_timestamp()
    if invoices.invoice_id.duplicated().any():raise ValueError('Invoice IDs must be unique across this ledger.')
    if payments.payment_id.duplicated().any():raise ValueError('Payment IDs must be unique across this ledger.')
    invoices['counterparty']=invoices.counterparty.map(canonical_counterparty)
    if invoices.counterparty.eq('').any():raise ValueError('Invoice counterparty is required.')
    if not invoices.direction.isin(['Payable','Receivable']).all():raise ValueError('Invoice direction must be Payable or Receivable.')
    if (invoices.due_date<invoices.invoice_date).any():raise ValueError('Due date cannot precede invoice date.')
    if not set(payments.invoice_id).issubset(set(invoices.invoice_id)):raise ValueError('Every payment must reference an existing invoice ID.')
    if not payments.empty:
        issue_dates=invoices.set_index('invoice_id').invoice_date
        if (payments.payment_date<payments.invoice_id.map(issue_dates)).any():raise ValueError('Payments before invoice date require separate prepayment handling; do not attach them to this ledger.')
    rows=[]
    for _,r in invoices.iterrows():
        linked=payments.loc[payments.invoice_id.eq(r.invoice_id)]
        if int(linked._cents.sum())>r._cents:raise ValueError('Payments exceed invoice amount: '+r.invoice_id+'; resolve overpayments separately.')
        paid=int(linked.loc[linked.payment_date.le(today),'_cents'].sum())
        issued=r.invoice_date<=today;balance=r._cents-paid if issued else None
        overdue=max((today-r.due_date).days,0) if issued and balance else 0
        bucket='Future invoice' if not issued else 'Paid' if balance==0 else 'Not overdue' if overdue==0 else '1–30 days' if overdue<=30 else '31–60 days' if overdue<=60 else '61–90 days' if overdue<=90 else 'Over 90 days'
        rows.append(dict(invoice_id=r.invoice_id,counterparty=r.counterparty,direction=r.direction,
            delivery_month=r.delivery_month.date().isoformat() if pd.notna(r.delivery_month) else None,
            invoice_date=r.invoice_date.date().isoformat(),due_date=r.due_date.date().isoformat(),as_of=today.date().isoformat(),
            invoice_amount_usd=r._cents/100,applied_payments_usd=paid/100,
            future_dated_payments_usd=int(linked.loc[linked.payment_date.gt(today),'_cents'].sum())/100,
            outstanding_usd=balance/100 if balance is not None else None,days_overdue=overdue,aging_bucket=bucket,
            status='FUTURE — excluded from outstanding' if not issued else 'PAID' if balance==0 else 'OVERDUE' if overdue else 'OPEN'))
    detail=pd.DataFrame(rows);summary=[]
    if not detail.empty:
        for (cp,direction),group in detail.groupby(['counterparty','direction']):
            issued=group.loc[~group.status.str.startswith('FUTURE')]
            row=dict(counterparty=cp,direction=direction,as_of=today.date().isoformat(),issued_invoices=len(issued),
                future_invoices=int(group.status.str.startswith('FUTURE').sum()),
                outstanding_usd=round(issued.outstanding_usd.sum(),2),overdue_invoices=int(issued.status.eq('OVERDUE').sum()))
            for bucket,col in [('Not overdue','not_overdue_usd'),('1–30 days','days_1_30_usd'),('31–60 days','days_31_60_usd'),('61–90 days','days_61_90_usd'),('Over 90 days','over_90_days_usd')]:
                row[col]=round(issued.loc[issued.aging_bucket.eq(bucket),'outstanding_usd'].sum(),2)
            row['scope']='Entered invoices only; payables and receivables remain separate. No forecast/MTM netting, collateral, credit limits or unentered-invoice completeness assertion.'
            summary.append(row)
    overdue=int(detail.status.eq('OVERDUE').sum()) if not detail.empty else 0
    status=pd.DataFrame([dict(as_of=today.date().isoformat(),status='NOT STARTED' if invoices.empty else 'OVERDUE ITEMS' if overdue else 'RECORDED LEDGER',invoice_count=len(invoices),payment_count=len(payments),overdue_invoices=overdue,scope='User-entered USD invoices and payments only; an empty ledger is not a confirmed zero balance.')])
    return {'counterparty_invoice_inputs.csv':invoices[INVOICE_COLUMNS],
        'counterparty_payment_inputs.csv':payments[PAYMENT_COLUMNS],
        'counterparty_invoice_aging.csv':detail,'counterparty_invoice_summary.csv':pd.DataFrame(summary),
        'counterparty_invoice_status.csv':status},overdue


def expected_invoice_summary(settlements):
    columns=['month','counterparty','valuation_date','cash_projected_signed_amount_usd','forward_projected_signed_amount_usd',
        'projected_payables_usd','projected_receivables_usd','projected_net_payable_usd','blocked_trade_months','status']
    if settlements.empty:return pd.DataFrame(columns=columns+['scope'])
    if not set(columns).issubset(settlements.columns):raise ValueError('Expected-invoice summary requires complete settlement report columns.')
    out=settlements[columns].copy()
    if out.duplicated(['month','counterparty']).any():raise ValueError('Duplicate counterparty/month in expected invoices.')
    blocked=out.status.astype(str).str.startswith('INCOMPLETE')|out.blocked_trade_months.gt(0)
    amounts=['cash_projected_signed_amount_usd','forward_projected_signed_amount_usd','projected_payables_usd','projected_receivables_usd','projected_net_payable_usd']
    out.loc[blocked,amounts]=float('nan')
    out['status']=out.status.where(blocked,'FORECAST — compare with invoices when received')
    out['scope']='Commodity forecast only, not a received invoice or unpaid balance. Positive net = expected payment; negative = expected receipt. Fees, fuel, taxes and other invoice charges excluded. No contractual netting assumption.'
    return out.sort_values(['month','counterparty'])


def invoice_reconciliation_reports(reports,asof):
    """Compare recorded invoice face amounts to commodity forecasts, never to unpaid balances."""
    expected=reports.get('counterparty_invoice_expectations.csv',pd.DataFrame()).copy()
    aging=reports.get('counterparty_invoice_aging.csv',pd.DataFrame()).copy()
    today=pd.Timestamp(asof).normalize()
    if not expected.empty:
        expected['month']=pd.to_datetime(expected.month).dt.to_period('M').dt.to_timestamp()
        expected['counterparty']=expected.counterparty.map(canonical_counterparty)
        if expected.duplicated(['month','counterparty']).any():raise ValueError('Duplicate counterparty/month forecast for reconciliation.')
    if not aging.empty:
        aging=aging.loc[pd.to_datetime(aging.invoice_date).le(today)].copy()
        if 'delivery_month' not in aging:aging['delivery_month']=None
        aging['delivery_month']=pd.to_datetime(aging.delivery_month,errors='coerce').dt.to_period('M').dt.to_timestamp()
    missing=aging.loc[aging.delivery_month.isna()].copy() if not aging.empty else pd.DataFrame()
    if not missing.empty:missing['reconciliation_issue']='Delivery month required; invoice date is not the delivery month.'
    assigned=aging.loc[aging.delivery_month.notna()] if not aging.empty else pd.DataFrame()
    keys=set(zip(expected.month,expected.counterparty)) if not expected.empty else set()
    if not assigned.empty:keys.update(zip(assigned.delivery_month,assigned.counterparty))
    rows=[]
    for month,cp in sorted(keys):
        forecasts=expected.loc[expected.month.eq(month)&expected.counterparty.eq(cp)] if not expected.empty else pd.DataFrame()
        invoices=assigned.loc[assigned.delivery_month.eq(month)&assigned.counterparty.eq(cp)] if not assigned.empty else pd.DataFrame()
        unknown=int(missing.counterparty.eq(cp).sum()) if not missing.empty else 0
        payable=sum(int(round(v*100)) for v in invoices.loc[invoices.direction.eq('Payable'),'invoice_amount_usd'])/100 if not invoices.empty else None
        receivable=sum(int(round(v*100)) for v in invoices.loc[invoices.direction.eq('Receivable'),'invoice_amount_usd'])/100 if not invoices.empty else None
        actual=round(payable-receivable,2) if payable is not None else None
        forecast=num(forecasts.iloc[0].projected_net_payable_usd) if not forecasts.empty else None
        if not forecasts.empty and (str(forecasts.iloc[0].status).startswith('INCOMPLETE') or num(forecasts.iloc[0].blocked_trade_months)!=0):forecast=None
        difference=round(actual-round(forecast,2),2) if actual is not None and forecast is not None and not unknown else None
        status='REVIEW — assign delivery months' if unknown else 'AWAITING INVOICE' if actual is None else 'REVIEW — forecast unavailable or incomplete' if forecast is None else 'VARIANCE — review invoice coverage and charges' if difference!=0 else 'AMOUNTS AGREE — coverage and charges unconfirmed'
        rows.append(dict(month=month.date().isoformat(),counterparty=cp,as_of=today.date().isoformat(),
            recorded_invoice_count=len(invoices),invoice_ids='; '.join(invoices.invoice_id) if not invoices.empty else '',
            invoices_missing_month=unknown,recorded_invoice_payables_usd=payable,recorded_invoice_receivables_usd=receivable,
            recorded_net_invoice_amount_usd=actual,forecast_net_payable_usd=forecast,invoice_less_forecast_usd=difference,status=status,
            scope='Recorded invoice face amounts versus commodity forecast. Includes paid invoices; excludes future-issued invoices. Positive difference = more payable or less receivable. Invoice fees, fuel, taxes and missing invoices can explain differences. Amount agreement is not completed reconciliation, legal netting or approval to pay.'))
    out=pd.DataFrame(rows,columns=['month','counterparty','as_of','recorded_invoice_count','invoice_ids','invoices_missing_month','recorded_invoice_payables_usd','recorded_invoice_receivables_usd','recorded_net_invoice_amount_usd','forecast_net_payable_usd','invoice_less_forecast_usd','status','scope'])
    attention=int(out.status.str.startswith(('REVIEW','VARIANCE')).sum())+len(missing)
    return {'counterparty_invoice_reconciliation.csv':out,'counterparty_invoice_unassigned.csv':missing},attention


def render_invoice_reconciliation(invoice_reports,asof):
    st.subheader('Invoice reconciliation by delivery month')
    st.caption('Compare recorded invoice totals with expected commodity amounts. Paid invoices remain in this comparison. Fees, fuel, taxes and partial invoice coverage can cause differences; equal amounts do not establish completed reconciliation. Enter each invoice delivery month above. A statement covering multiple months requires separate allocation before it can be compared here.')
    st.caption('The Counterparty delivery range above must include the invoice delivery month to supply its forecast. Scheduled comparisons use the configured reporting range; out-of-range invoices remain flagged without an invented forecast.')
    reports,_=invoice_reconciliation_reports(dict(st.session_state['_run_reports'],**invoice_reports),asof)
    st.session_state['_run_reports'].update(reports)
    full=reports['counterparty_invoice_reconciliation.csv']
    if full.empty:st.info('Generate expected invoices or enter an invoice with a delivery month to populate this report.')
    else:
        months=sorted(full.month.unique())
        target=st.session_state.get('expected_invoice_month')
        index=months.index(target) if target in months else len(months)-1
        month=st.selectbox('Reconciliation delivery month',months,index=index,key='invoice_reconciliation_month',format_func=lambda m:pd.Timestamp(m).strftime('%B %Y'))
        show(full.loc[full.month.eq(month)],'counterparty_invoice_reconciliation_selected.csv')
    missing=reports['counterparty_invoice_unassigned.csv']
    if not missing.empty:
        st.warning(str(len(missing))+' issued invoice(s) need a delivery month. Differences for those counterparties are withheld until assigned.')
        show(missing,'counterparty_invoice_unassigned.csv')


def render_expected_invoices():
    st.subheader('Expected invoices by counterparty')
    st.caption('Use this commodity forecast as the reference while invoices are pending. October invoices are expected in early November. Leave actual invoice entries blank until received; this forecast is not an overdue balance. Compare fees, fuel, taxes and other invoice charges separately.')
    source=st.session_state['_run_reports'].get('counterparty_settlement_monthly.csv',pd.DataFrame())
    expected=expected_invoice_summary(source)
    st.session_state['_run_reports']['counterparty_invoice_expectations.csv']=expected
    if expected.empty:
        st.info('Generate the combined counterparty settlement report above to populate expected invoices.');return
    months=sorted(expected.month.astype(str).unique())
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None)
    target=(today.to_period('M').start_time-pd.DateOffset(months=1) if today.day<=10 else today.to_period('M').start_time).strftime('%Y-%m-%d')
    available=[m for m in months if m<=target];default=available[-1] if available else months[0]
    if st.session_state.get('expected_invoice_month') not in months:st.session_state['expected_invoice_month']=default
    month=st.selectbox('Expected invoice delivery month',months,key='expected_invoice_month',format_func=lambda m:pd.Timestamp(m).strftime('%B %Y'))
    view=expected.loc[expected.month.astype(str).eq(month)]
    complete=not view.status.str.startswith('INCOMPLETE').any() and view.projected_net_payable_usd.notna().all()
    a,b,c=st.columns(3)
    for widget,label,col in [(a,'Forecast payables','projected_payables_usd'),(b,'Forecast receivables','projected_receivables_usd'),(c,'Net forecast payable','projected_net_payable_usd')]:
        widget.metric(label,f"${view[col].sum():,.2f}" if complete else 'Withheld — incomplete')
    show(view,'counterparty_invoice_expectations_selected.csv')
    st.caption('Positive net amounts indicate expected payments; negative amounts indicate expected receipts. Full expectations are retained in the run archive. No invoice is automatically created from this table.')


def invoice_pdf_text(content,reader):
    """OCR image-only pages locally; never send invoices to an external API."""
    import shutil
    import subprocess
    import tempfile
    pages=[page.extract_text() or '' for page in reader.pages]
    missing=[i for i,text in enumerate(pages) if len(text.strip())<30]
    if not missing:return pages,False
    if not shutil.which('pdftoppm') or not shutil.which('tesseract'):
        raise ValueError('Scanned PDF detected. Add poppler-utils and tesseract-ocr to packages.txt in the Streamlit repository, then redeploy.')
    try:
        with tempfile.TemporaryDirectory(prefix='invoice_ocr_') as folder:
            source=Path(folder)/'invoice.pdf';source.write_bytes(content)
            for i in missing:
                prefix=Path(folder)/('page_'+str(i+1))
                subprocess.run(['pdftoppm','-f',str(i+1),'-l',str(i+1),'-singlefile','-scale-to','2200','-png',str(source),str(prefix)],check=True,capture_output=True,timeout=40)
                ocr=subprocess.run(['tesseract',str(prefix)+'.png','stdout','-l','eng'],check=True,capture_output=True,text=True,timeout=40)
                pages[i]=ocr.stdout
                if i==0:
                    alternate=subprocess.run(['tesseract',str(prefix)+'.png','stdout','-l','eng','--psm','6'],check=True,capture_output=True,text=True,timeout=40)
                    pages[i]+='\n'+alternate.stdout
                Path(str(prefix)+'.png').unlink(missing_ok=True)
    except (subprocess.SubprocessError,OSError):raise ValueError('OCR could not finish. Try a clearer or smaller PDF, or enter the invoice manually.')
    return pages,True


def invoice_pdf_draft(content,known_counterparties):
    """Extract labeled candidates only; missing/conflicting fields require review."""
    if len(content)>20*1024*1024:raise ValueError('Use a PDF under 20 MB.')
    try:
        from pypdf import PdfReader
    except ImportError:raise ValueError('Install pypdf>=5,<7 in the app requirements.txt to enable PDF upload.')
    try:
        reader=PdfReader(BytesIO(content))
        if reader.is_encrypted:raise ValueError('Upload an unlocked PDF.')
        if len(reader.pages)>30:raise ValueError('Upload one invoice with at most 30 pages.')
        pages,ocr_used=invoice_pdf_text(content,reader)
        text='\n'.join(pages)
    except ValueError:raise
    except Exception:raise ValueError('Could not read the PDF. Use an unlocked, text-based invoice.')
    # Normalize actual text separators without interpreting document instructions.
    text=text.replace('\\n','\n')
    if len(text.strip())<30:raise ValueError('This PDF has no usable text. Use a searchable/OCR PDF or enter the invoice manually.')
    warnings=['OCR was used. Check invoice digits, dates and total against the PDF before adding.'] if ocr_used else []
    evidence=[]
    def unique(field,items):
        pairs={str(value): (value,line.strip()) for value,line in items}
        if len(pairs)==1:
            value,line=next(iter(pairs.values()));evidence.append(dict(field=field,matched_text=line));return value
        warnings.append(field+(': multiple candidates; review required' if pairs else ': not found; enter manually'))
        return None
    ids=[]
    for m in re.finditer(r'(?im)\b(?:invoice\s*(?:number|no\.?|\#|id)|invoice(?=\s*:)|inv\.?\s*(?:no\.?|\#))\s*[:#-]?\s*([A-Z0-9][A-Z0-9._/\-]{0,69})',text):
        value=m.group(1)
        if value.lower() not in ['date','number','no','amount','total']:ids.append((value,m.group(0)))
    invoice_id=unique('Invoice number',ids)
    date_pattern=r'(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})'
    dates={}
    for field,label in [('invoice_date',r'invoice\s+date'),('due_date',r'(?:payment\s+due(?:\s+date)?|due\s+date)')]:
        found=[]
        for m in re.finditer(r'(?im)\b'+label+r'\s*:?\s*'+date_pattern,text):
            date_=pd.to_datetime(m.group(1),errors='coerce')
            if pd.notna(date_):found.append((date_.date().isoformat(),m.group(0)))
        dates[field]=unique(field,found)
    amounts=[]
    # Never interpret a subtotal, bank account or prior balance as the invoice total.
    for label in [r'(?:invoice\s+total|total\s+invoice(?:\s+amount)?|invoice\s+amount|grand\s+total|total\s+amount\s+due)',r'(?:amount\s+due|total\s+due)']:
        for m in re.finditer(r'(?im)\b'+label+r'\s*:?\s*(?:USD\s*)?\$?\s*(-?\d[\d,]*\.\d{2})\b',text):
            amounts.append((float(m.group(1).replace(',','')),m.group(0)))
        if amounts:break
    # Snyder's cover page nets its purchases and sales; use only its explicit net-due line.
    snyder_net=[]
    if pages and re.search(r'(?i)\bSnyder\s+Brothers\s+Inc',pages[0]):
        for m in re.finditer(r'(?im)Total\s+Due\s+to\s+Snyder\s+Brothers\s+Inc\.?\s*:\s*\$?\s*(\(?-?\d[\d,]*\.\d{2}\)?)',pages[0]):
            token=m.group(1);value=float(token.replace(',','').replace('(','').replace(')',''))
            if token.startswith('('):value=-abs(value)
            snyder_net.append((value,m.group(0)))
    if snyder_net:amounts=snyder_net
    amount=unique('Invoice amount',amounts)
    parties=[]
    for cp in sorted(set(canonical_counterparty(v) for v in known_counterparties if txt(v))):
        if re.search(r'(?i)(?<!\w)'+re.escape(cp)+r'(?!\w)',text):parties.append((cp,cp))
    if snyder_net:parties=[('Snyder','Snyder Brothers Inc. — explicit net amount due on cover page')]
    cp=unique('Counterparty',parties)
    direction='Payable'
    if snyder_net and amount is not None and amount<0:
        direction='Receivable';amount=abs(amount)
        warnings.append('Negative net due to Snyder interpreted as Receivable; confirm the direction.')
    unsupported=bool(re.search(r'(?i)\b(?:credit\s+memo|credit\s+note|EUR|GBP|CAD)\b|C\$|€|£',text))
    if unsupported:warnings.append('Credit note or non-USD currency detected; this ledger supports positive USD invoices only.')
    return dict(row=dict(invoice_id=invoice_id or '',counterparty=cp or '',direction=direction,
        invoice_date=dates['invoice_date'],due_date=dates['due_date'],amount_usd=amount),
        text=text,ocr_used=ocr_used,warnings=warnings,evidence=evidence,unsupported=unsupported,sha256=hashlib.sha256(content).hexdigest())


def add_uploaded_invoice(profile,row,document,asof):
    if document.get('unsupported'):raise ValueError('Credit notes and non-USD documents cannot be imported into this invoice ledger.')
    result=json.loads(json.dumps(profile))
    existing=result.get('source_documents',[])
    if any(x.get('sha256')==document['sha256'] for x in existing):raise ValueError('This PDF has already been added to the ledger.')
    if any(txt(x['invoice_id']).casefold()==txt(row['invoice_id']).casefold() for x in result['invoices']):
        raise ValueError('This invoice ID already exists; review the existing entry instead of importing it twice.')
    result['invoices'].append(row)
    invoice_aging_reports(result,asof)
    result['source_documents']=existing+[dict(invoice_id=row['invoice_id'],file_name=document['file_name'],sha256=document['sha256'])]
    return result


def render_invoice_pdf_upload(current,today):
    with st.expander('Upload a received invoice PDF',expanded=True):
        st.caption('PDF, one invoice per file; scanned pages are read using local OCR. Extracted fields are drafts: review the counterparty, total, dates and Payable/Receivable direction before adding. Received invoices default to Payable. Review OCR results carefully. The original PDF is not retained by this ledger; keep it with your accounting records.')
        upload=st.file_uploader('Actual invoice PDF',type=['pdf'],key='actual_invoice_pdf')
        if upload is None:return
        content=upload.getvalue();digest=hashlib.sha256(content).hexdigest()
        if st.button('Extract invoice details',key='invoice_extract'):
            try:
                parties=st.session_state['_run_reports'].get('counterparty_settlement_monthly.csv',pd.DataFrame()).get('counterparty',pd.Series(dtype=str)).tolist()
                parties += [r['counterparty'] for r in current['invoices']]
                with st.spinner('Reading invoice and scanned pages...'):
                    draft=invoice_pdf_draft(content,parties)
                draft['file_name']=upload.name
                st.session_state['_invoice_pdf_draft']=draft
            except ValueError as exc:
                st.session_state.pop('_invoice_pdf_draft',None);st.error(str(exc))
        draft=st.session_state.get('_invoice_pdf_draft')
        if not draft or draft['sha256']!=digest:return
        for message in draft['warnings']:st.warning(message)
        with st.expander('Extracted invoice text and matching labels'):
            st.text(draft['text']);st.dataframe(pd.DataFrame(draft['evidence']),hide_index=True)
        row=pd.DataFrame([dict(draft['row'],delivery_month=None)])
        st.caption('Enter the delivery month shown on the invoice. It is not inferred from the issue or due date. Leave blank if unknown; reconciliation will flag it.')
        for col in ['invoice_date','due_date','delivery_month']:row[col]=pd.to_datetime(row[col],errors='coerce')
        row['amount_usd']=pd.to_numeric(row.amount_usd,errors='coerce').astype(float)
        reviewed=st.data_editor(row,hide_index=True,key='invoice_pdf_review_'+digest,
            column_config={'delivery_month':st.column_config.DateColumn('Delivery month (any day)'), 'direction':st.column_config.SelectboxColumn('Direction',options=['Payable','Receivable'],required=True),
                'invoice_date':st.column_config.DateColumn('Invoice date',required=True),'due_date':st.column_config.DateColumn('Due date',required=True),
                'amount_usd':st.column_config.NumberColumn('Invoice amount (USD)',min_value=.01,step=.01,format='%.2f',required=True)})
        if st.button('Add reviewed invoice to ledger',key='invoice_pdf_add',disabled=draft['unsupported']):
            try:
                added=json.loads(reviewed.to_json(orient='records',date_format='iso'))[0]
                updated=add_uploaded_invoice(current,added,draft,today)
                st.session_state['_invoice_saved']=updated
                st.session_state['_invoice_revision']=st.session_state.get('_invoice_revision',0)+1
                st.session_state.pop('_invoice_pdf_draft',None)
                st.session_state['_invoice_upload_success']=True
                st.rerun()
            except (ValueError,TypeError,KeyError) as exc:st.error('Invoice not added: '+str(exc))


CREDIT_DEFAULT_LIMITS={'Bce':100000.,'Bpec':350000.,'Carbonbetter':0.,'Eco':150000.,'Egts':10000.,'Mercuria':1000000.,'Shell':500000.,'Snyder':1000000.,'Uet':100000.}


def default_credit_row(counterparty):
    cp=canonical_counterparty(counterparty)
    key=next((k for k in CREDIT_DEFAULT_LIMITS if k.casefold()==cp.casefold()),None)
    return dict(counterparty=cp,receivables_limit_usd=CREDIT_DEFAULT_LIMITS.get(key),eligible_collateral_usd=0.,receivables_complete_through=None,source_reference='',payment_terms='Prepay only' if key=='Carbonbetter' else 'Credit')


def default_credit_profile():
    return dict(schema='counterparty-credit-1',defaults_version='approved-limits-1',rows=[default_credit_row(cp) for cp in CREDIT_DEFAULT_LIMITS])


CREDIT_COLUMNS=['counterparty','receivables_limit_usd','eligible_collateral_usd','receivables_complete_through','source_reference','payment_terms']


def counterparty_credit_reports(profile,reports,asof):
    if profile.get('schema')!='counterparty-credit-1':raise ValueError('Unsupported counterparty credit settings schema.')
    today=pd.Timestamp(asof).normalize();settings=pd.DataFrame(profile.get('rows',[]))
    if settings.empty:settings=pd.DataFrame(columns=CREDIT_COLUMNS)
    if 'payment_terms' not in settings:settings['payment_terms']=settings.counterparty.map(lambda cp:default_credit_row(cp)['payment_terms']) if 'counterparty' in settings else 'Credit'
    if not set(CREDIT_COLUMNS).issubset(settings.columns):raise ValueError('Credit settings require: '+', '.join(CREDIT_COLUMNS))
    settings=settings[CREDIT_COLUMNS].copy();settings['counterparty']=settings.counterparty.map(canonical_counterparty)
    if settings.counterparty.eq('').any() or settings.counterparty.duplicated().any():raise ValueError('Credit settings require unique, nonblank counterparties.')
    for col in ['receivables_limit_usd','eligible_collateral_usd']:
        raw=settings[col];settings[col]=pd.to_numeric(raw,errors='coerce')
        if (raw.notna() & (settings[col].isna()|settings[col].lt(0)|settings[col].isin([float('inf'),float('-inf')]))).any():
            raise ValueError(col+' must be a nonnegative finite amount or blank.')
    if not settings.payment_terms.isin(['Credit','Prepay only']).all():raise ValueError('Payment terms must be Credit or Prepay only.')
    if (settings.payment_terms.eq('Prepay only') & settings.receivables_limit_usd.ne(0)).any():raise ValueError('Prepay-only counterparties must have a zero receivables limit.')
    dates=pd.to_datetime(settings.receivables_complete_through,errors='coerce')
    present=settings.receivables_complete_through.map(lambda x:bool(txt(x)))
    if (present & dates.isna()).any() or dates.gt(today).any():raise ValueError('Receivables coverage date must be valid and cannot be in the future.')
    settings['receivables_complete_through']=dates
    aging=reports.get('counterparty_invoice_aging.csv',pd.DataFrame())
    ledger_status=reports.get('counterparty_invoice_status.csv',pd.DataFrame())
    ledger_available=(not ledger_status.empty and 'as_of' in ledger_status and
        pd.to_datetime(ledger_status.as_of,errors='coerce').eq(today).all() and
        not ledger_status.status.isin(['NOT CONFIGURED']).any())
    detail=reports.get('counterparty_settlement_detail.csv',pd.DataFrame()).copy()
    if not detail.empty:detail['month']=pd.to_datetime(detail.month)
    rows=[]
    for _,r in settings.iterrows():
        current=aging.loc[(aging.counterparty==r.counterparty)&(aging.direction=='Receivable')&(~aging.status.str.startswith('FUTURE'))] if not aging.empty else aging
        recorded=float(current.outstanding_usd.sum()) if not current.empty else 0.
        complete=ledger_available and pd.notna(r.receivables_complete_through) and r.receivables_complete_through>=today and bool(txt(r.source_reference))
        limit=num(r.receivables_limit_usd);collateral=num(r.eligible_collateral_usd)
        exposure=max(recorded-collateral,0.) if complete and collateral is not None else None
        headroom=limit-exposure if limit is not None and exposure is not None else None
        status='NOT CONFIGURED' if limit is None else 'REVIEW — confirm current receivables coverage and collateral' if headroom is None else 'BREACH' if headroom<0 else 'AT LIMIT' if headroom==0 else 'WITHIN RECEIVABLES LIMIT'
        if r.payment_terms=='Prepay only':
            exposure=recorded if complete else None
            headroom=-exposure if exposure is not None else None
            status='BREACH' if ledger_available and recorded>0 else 'PREPAY ONLY — no unpaid receivables recorded' if complete else 'REVIEW — confirm prepay-only receivables coverage'
        future=detail.loc[(detail.counterparty==r.counterparty)&detail.source_sheet.eq('fwdblotter')&(detail.month>today.to_period('M').start_time)] if not detail.empty else detail
        mtm_complete=not future.empty and future.forward_mtm_usd.notna().all() and not future.status.eq('BLOCKED').any()
        rows.append(dict(counterparty=r.counterparty,payment_terms=r.payment_terms,as_of=today.date().isoformat(),
            recorded_receivables_usd=recorded if ledger_available else None,
            receivables_coverage_status='USER-CONFIRMED THROUGH AS-OF DATE' if complete else 'NOT CONFIRMED FOR AS-OF DATE',
            receivables_limit_usd=limit,eligible_collateral_usd=collateral,
            receivables_after_collateral_usd=exposure,remaining_receivables_limit_usd=headroom,
            excess_over_limit_usd=max(-headroom,0.) if headroom is not None else None,status=status,
            future_positive_trade_mtm_usd=float(future.forward_mtm_usd.clip(lower=0).sum()) if mtm_complete else None,
            future_net_mtm_usd=float(future.forward_mtm_usd.sum()) if mtm_complete else None,
            future_mtm_status='INDICATIVE — selected reporting range only' if mtm_complete else 'NOT AVAILABLE OR INCOMPLETE',
            future_mtm_start=future.month.min().date().isoformat() if not future.empty else None,
            future_mtm_end=future.month.max().date().isoformat() if not future.empty else None,
            source_reference=txt(r.source_reference),scope='Limit check covers confirmed unpaid receivables less entered eligible collateral only. Future MTM is separate and excluded from this limit. No payable offset, contractual netting, collateral valuation or total credit-exposure certification.'))
    result=pd.DataFrame(rows)
    attention=int(result.status.str.startswith('REVIEW').sum()+result.status.eq('BREACH').sum()) if not result.empty else 0
    return {'counterparty_credit_inputs.csv':settings,'counterparty_credit_monitor.csv':result},attention


def render_counterparty_credit():
    st.subheader('Counterparty limits and collateral — receivables monitor')
    st.caption('Enter the approved receivables limit and eligible collateral held for each counterparty. This limit check covers unpaid receivables only; future hedge MTM is displayed separately. Blank limits are unconfigured, never unlimited. No payables are offset against receivables.')
    st.caption('Approved limits are prefilled and editable. Carbonbetter defaults to Prepay only with a zero limit. Any recorded unpaid receivable flags a prepay breach even if collateral is entered. This monitor does not verify advance payment before gas delivery or block trades.')
    st.caption('Receivables complete through means you have reconciled all unpaid receivables for this counterparty through that date. Confirm the current as-of date and a source reference before relying on headroom. An empty invoice ledger alone does not establish a zero balance. Collateral amounts are entered by you; eligibility, valuation and expiration are not independently checked.')
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    upload=st.file_uploader('Restore counterparty credit settings JSON',type=['json'],key='credit_restore')
    if upload is not None and st.button('Load credit settings',key='credit_load'):
        try:
            data=json.loads(upload.getvalue());counterparty_credit_reports(data,st.session_state['_run_reports'],today)
            st.session_state['_credit_saved']=data;st.session_state['_credit_revision']=st.session_state.get('_credit_revision',0)+1
        except (ValueError,TypeError,KeyError) as exc:st.error('Credit settings not loaded: '+str(exc))
    saved=st.session_state.get('_credit_saved',dict(schema='counterparty-credit-1',rows=[]))
    parties=set(st.session_state['_run_reports'].get('counterparty_settlement_monthly.csv',pd.DataFrame()).get('counterparty',[]))
    parties.update(CREDIT_DEFAULT_LIMITS)
    parties.update(r['counterparty'] for r in saved['rows'])
    parties={canonical_counterparty(cp) for cp in parties if canonical_counterparty(cp)}
    by_name={canonical_counterparty(r['counterparty']):r for r in saved['rows']}
    seeded=[]
    for cp in sorted(parties):
        row=default_credit_row(cp)
        prior=by_name.get(cp,{})
        row.update(prior)
        if saved.get('defaults_version')!='approved-limits-1' and num(row.get('receivables_limit_usd')) is None:
            row['receivables_limit_usd']=default_credit_row(cp)['receivables_limit_usd']
        seeded.append(row)
    seed=pd.DataFrame(seeded,columns=CREDIT_COLUMNS)
    seed['receivables_complete_through']=pd.to_datetime(seed.receivables_complete_through,errors='coerce')
    for col in ['receivables_limit_usd','eligible_collateral_usd']:seed[col]=pd.to_numeric(seed[col],errors='coerce').astype(float)
    entries=st.data_editor(seed,num_rows='dynamic',hide_index=True,key='credit_rows_'+str(st.session_state.get('_credit_revision',0)),
        column_config={'payment_terms':st.column_config.SelectboxColumn('Payment terms',options=['Credit','Prepay only'],required=True),'receivables_complete_through':st.column_config.DateColumn('Receivables complete through'),
            'receivables_limit_usd':st.column_config.NumberColumn('Receivables limit (USD)',min_value=0.,format='%.2f'),
            'eligible_collateral_usd':st.column_config.NumberColumn('Eligible collateral held (USD)',min_value=0.,format='%.2f')})
    profile=dict(schema='counterparty-credit-1',defaults_version='approved-limits-1',rows=json.loads(entries.to_json(orient='records',date_format='iso')))
    try:
        reports,_=counterparty_credit_reports(profile,st.session_state['_run_reports'],today)
        for name,report in reports.items():show(report,name)
        if entries.receivables_limit_usd.notna().any() or '_credit_saved' in st.session_state:
            st.session_state['_scheduled_storage_inputs']['counterparty_credit']=profile
        st.download_button('Download credit settings JSON',json.dumps(profile,indent=2),file_name='counterparty_credit_settings.json',mime='application/json',key='credit_download')
        st.caption('Retain settings with the JSON download. To use them in scheduled runs, re-export and upload the combined profile from Storage pricing & summary. Coverage dates older than the run date withhold headroom until reviewed.')
    except (ValueError,TypeError,KeyError) as exc:
        st.session_state['_scheduled_storage_inputs'].pop('counterparty_credit',None)
        st.error('Credit report withheld: '+str(exc))


def render_invoice_ledger():
    render_expected_invoices()
    st.subheader('Invoices, payments and aging')
    st.caption('Enter actual USD invoices and payments. Payable = we owe the counterparty; Receivable = the counterparty owes us. Payments are positive amounts applied to one invoice ID. Use unique IDs and split a payment across invoices with distinct allocation IDs. Forecast trades are not copied into this ledger. Credit notes, prepayments and overpayments need separate handling.')
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    if st.session_state.pop('_invoice_upload_success',False):st.success('Reviewed invoice added. Download the updated ledger JSON to retain it.')
    uploaded=st.file_uploader('Restore invoice and payment ledger JSON',type=['json'],key='invoice_ledger_restore')
    if uploaded is not None and st.button('Load invoice and payment ledger',key='invoice_ledger_load'):
        try:
            data=json.loads(uploaded.getvalue());invoice_aging_reports(data,today)
            st.session_state['_invoice_saved']=data
            st.session_state['_invoice_revision']=st.session_state.get('_invoice_revision',0)+1
        except (ValueError,TypeError,KeyError) as exc:st.error('Ledger not loaded: '+str(exc))
    saved=st.session_state.get('_invoice_saved',dict(schema='counterparty-ledger-1',invoices=[],payments=[]))
    rev=str(st.session_state.get('_invoice_revision',0))
    def frame(rows,columns):
        out=pd.DataFrame(rows,columns=columns)
        for col in columns:
            if col.endswith('_date') or col=='delivery_month':out[col]=pd.to_datetime(out[col],errors='coerce')
            elif col=='amount_usd':out[col]=pd.to_numeric(out[col],errors='coerce').astype(float)
            else:out[col]=out[col].astype('string')
        return out
    st.markdown('**Invoices**')
    inv=st.data_editor(frame(saved['invoices'],INVOICE_COLUMNS),num_rows='dynamic',hide_index=True,key='invoice_rows_'+rev,
        column_config={'delivery_month':st.column_config.DateColumn('Delivery month (any day)'), 'direction':st.column_config.SelectboxColumn('Direction',options=['Payable','Receivable'],required=True),
            'invoice_date':st.column_config.DateColumn('Invoice date',required=True),'due_date':st.column_config.DateColumn('Due date',required=True),
            'amount_usd':st.column_config.NumberColumn('Invoice amount (USD)',min_value=.01,step=.01,format='%.2f',required=True)})
    st.markdown('**Payments made or received**')
    pay=st.data_editor(frame(saved['payments'],PAYMENT_COLUMNS),num_rows='dynamic',hide_index=True,key='payment_rows_'+rev,
        column_config={'payment_date':st.column_config.DateColumn('Payment date',required=True),
            'amount_usd':st.column_config.NumberColumn('Applied payment (USD)',min_value=.01,step=.01,format='%.2f',required=True)})
    selected=st.date_input('Invoice aging as of',value=today.date(),key='invoice_aging_asof')
    data=dict(schema='counterparty-ledger-1',invoices=json.loads(inv.to_json(orient='records',date_format='iso')),payments=json.loads(pay.to_json(orient='records',date_format='iso')),source_documents=saved.get('source_documents',[]))
    render_invoice_pdf_upload(data,today)
    try:
        reports,_=invoice_aging_reports(data,selected)
        for name,report in reports.items():show(report,name)
        render_invoice_reconciliation(reports,selected)
        if data['invoices'] or '_invoice_saved' in st.session_state:
            st.session_state['_scheduled_storage_inputs']['counterparty_ledger']=data
        st.download_button('Download invoice and payment ledger JSON',json.dumps(data,indent=2),file_name='counterparty_invoice_ledger.json',mime='application/json',key='invoice_ledger_download')
        st.caption('Download JSON to retain these entries across app restarts. To include them in scheduled runs, export the combined profile from Storage pricing & summary and replace the configured profile JSON in GitHub. Scheduled aging uses the run date; the as-of control here changes only this view.')
    except (ValueError,TypeError,KeyError) as exc:
        st.session_state['_scheduled_storage_inputs'].pop('counterparty_ledger',None)
        st.error('Invoice reports and scheduled ledger export withheld: '+str(exc))


def counterparty_settlement_reports(ledger,flags,values,price_days,legacy,newer,master,start,end,asof,positive_buy=True,ice=None,classifications=None):
    """Commodity settlement forecasts, not unpaid balances or legal credit netting."""
    start=pd.Timestamp(start).normalize();end=pd.Timestamp(end).normalize();asof=pd.Timestamp(asof).normalize()
    if end<start or (end-start).days>3660:raise ValueError('Invalid counterparty reporting period.')
    if ledger.source_id.duplicated().any():raise ValueError('Duplicate ledger source IDs.')
    allocations=monthly_volumes(ledger,start,end,positive_buy,all_physical=True)
    columns=['month','counterparty','source_id','source_sheet','index_code','purchase_volume','sale_volume','net_volume',
        'projected_signed_amount','projected_payable_usd','projected_receivable_usd','forward_mtm_usd','status','issue']
    if allocations.empty:return {'counterparty_settlement_detail.csv':pd.DataFrame(columns=columns),
        'counterparty_settlement_monthly.csv':pd.DataFrame()},0
    terms=ledger.set_index('source_id');v=values.copy();prices=price_days.copy()
    if not v.empty:v['month']=pd.to_datetime(v.month).dt.to_period('M').dt.to_timestamp()
    if not prices.empty:
        prices['delivery_date']=pd.to_datetime(prices.delivery_date)
        prices['index_code']=prices.index_code.map(market)
    price_cache={};price_checks=[];rows=[]
    def index_day(code,day_):
        key=(market(code),day_)
        if key not in price_cache:
            result=settlement_index_price(code,day_,legacy,newer,master,prices,ice,asof)
            price_checks.append(result);price_cache[key]=(result['price'],result['price_kind'])
        return price_cache[key]
    for _,a in allocations.iterrows():
        r=terms.loc[a.source_id];amount=mtm=None;issues=[];historical=forecast=cached=0;missing_dates=[]
        if not txt(r.counterparty):issues.append('Missing counterparty')
        if a.net_volume==0:
            amount=0.;mtm=0. if r.source_sheet=='fwdblotter' else None
            basis='Zero-volume commodity allocation; no index price required; fees excluded'
        elif r.source_sheet=='fwdblotter':
            hits=v.loc[(v.source_id==a.source_id)&(v.month==a.month)] if not v.empty else v
            current_fallback=len(hits)<=1 and a.month.to_period('M')==asof.to_period('M') and classifications is not None
            basis='Existing matched forward valuation; selected market reference and contract terms'
            if current_fallback:
                cls=classifications.loc[classifications.source_id.eq(a.source_id)] if not classifications.empty else classifications
                accepted={'MATCHED — REVIEW','SHARED — CONSISTENT CLASSIFICATION'}
                term=num(r.get('contract_price'))
                if len(cls)!=1 or cls.iloc[0].match_status not in accepted:
                    issues.append('Current-month forward classification missing or unresolved')
                elif term is None:issues.append('Current-month forward contract price/premium missing')
                elif cls.iloc[0].candidate_classification=='FIXED':
                    amount=a.net_volume*term
                elif cls.iloc[0].candidate_classification=='GDD':
                    daily=[];missing=0
                    first=max(start,r.start,a.month);last=min(end,r.end,a.month+pd.offsets.MonthEnd(0))
                    for day_ in pd.date_range(first,last):
                        price,kind=index_day(a.index_code,day_)
                        if price is None:missing+=1;missing_dates.append(day_.date().isoformat())
                        else:
                            daily.append(price);historical+=int(kind=='Historical');forecast+=int(kind=='Forward');cached+=int(kind=='Dated cash cache')
                    if missing:issues.append(str(missing)+' current-month forward days missing usable index prices')
                    else:amount=a.net_volume*(sum(daily)/len(daily)+term)
                else:issues.append('Current-month settlement fallback supports reconciled FIXED/GDD only')
                basis='Current-month settlement: reconciled FIXED contract amount or GDD daily index plus booked premium; historical/forecast blend; no current-month MTM inferred'
            elif len(hits)!=1:issues.append('Forward valuation missing or duplicated for selected month')
            else:
                f=hits.iloc[0];direction=num(f.get('directional_volume'))
                if (not txt(f.get('valuation_status')).startswith('INDICATIVE') or direction is None
                        or abs(direction-a.net_volume)>1e-6 or market(f.index_code)!=market(a.index_code)
                        or canonical_counterparty(f.counterparty)!=canonical_counterparty(r.counterparty)):
                    issues.append('Forward valuation blocked or scope/sign mismatch')
                else:
                    amount=num(f.get('projected_signed_contract_amount'));mtm=num(f.get('indicative_mtm'))
                    if amount is None or mtm is None:issues.append('Forward amount or MTM missing')
        else:
            first=max(start,r.start,a.month);last=min(end,r.end,a.month+pd.offsets.MonthEnd(0))
            premium=num(r.get('premium'));premium=0. if premium is None and pd.isna(r.get('premium')) else premium
            if bool(r.get('premium_issue',False)) or premium is None:issues.append('Invalid cash premium')
            daily=[];missing=0
            for day_ in pd.date_range(first,last):
                price,kind=index_day(a.index_code,day_)
                if price is None:missing+=1;missing_dates.append(day_.date().isoformat())
                else:
                    daily.append(price);historical+=int(kind=='Historical');forecast+=int(kind=='Forward');cached+=int(kind=='Dated cash cache')
            if missing:issues.append(str(missing)+' cash days missing usable index prices: '+a.index_code+' on '+', '.join(missing_dates))
            if not issues:amount=a.net_volume*(sum(daily)/len(daily)+premium)
            basis='Exact-date system index plus booked cash premium; historical, dated cash-cache or forward source identified in price checks; settlement unconfirmed; no default delivery differential added'
        if issues:amount=mtm=None
        rows.append(dict(month=a.month.date().isoformat(),counterparty=canonical_counterparty(r.counterparty) or '(missing)',
            source_id=a.source_id,source_sheet=r.source_sheet,index_code=a.index_code,location=txt(r.get('location')),delivery_days=a.delivery_days,
            delivery_start=max(start,r.start,a.month).date().isoformat(),delivery_end=min(end,r.end,a.month+pd.offsets.MonthEnd(0)).date().isoformat(),missing_index_dates='; '.join(missing_dates),
            purchase_volume=a.purchase_volume,sale_volume=a.sale_volume,net_volume=a.net_volume,
            projected_signed_amount=amount,projected_payable_usd=max(amount,0.) if amount is not None else None,
            projected_receivable_usd=max(-amount,0.) if amount is not None else None,forward_mtm_usd=mtm,
            historical_cash_days=historical if r.source_sheet=='cashblotter' else 0,forecast_cash_days=forecast if r.source_sheet=='cashblotter' else 0,
            historical_index_days=historical,forecast_index_days=forecast,dated_cash_cache_days=cached,
            forward_mtm_status='NOT APPLICABLE' if r.source_sheet!='fwdblotter' else 'INDICATIVE' if mtm is not None else 'NOT AVAILABLE — settlement calculation is separate',
            status='BLOCKED' if issues else 'ESTIMATE — settlement unconfirmed',issue='; '.join(issues),pricing_basis=basis))
    detail=pd.DataFrame(rows);summary=[]
    unresolved=int(active_unallocatable_records(ledger,flags).sum())
    for (month,cp),g in detail.groupby(['month','counterparty'],dropna=False):
        row=dict(month=month,counterparty=cp,report_start=start.date().isoformat(),report_end=end.date().isoformat(),valuation_date=asof.date().isoformat(),trade_months=len(g),blocked_trade_months=int(g.status.eq('BLOCKED').sum()),unallocatable_ledger_records=unresolved)
        for book,label in [('cashblotter','cash'),('fwdblotter','forward')]:
            group=g.loc[g.source_sheet.eq(book)];complete=not group.status.eq('BLOCKED').any() and not unresolved
            row[label+'_purchases_dth']=group.purchase_volume.sum();row[label+'_sales_dth']=group.sale_volume.sum()
            row[label+'_net_volume_dth']=group.net_volume.sum()
            row[label+'_projected_signed_amount_usd']=group.projected_signed_amount.sum() if complete else None
        complete=not row['blocked_trade_months'] and not unresolved
        row.update(net_volume_dth=g.net_volume.sum(),projected_payables_usd=g.projected_payable_usd.sum() if complete else None,
            projected_receivables_usd=g.projected_receivable_usd.sum() if complete else None,
            projected_net_payable_usd=g.projected_signed_amount.sum() if complete else None,
            pool_and_delivered_scope='All booked locations; counterparty settlement view only, not delivered supply netting',
            status='INCOMPLETE — totals withheld' if not complete else 'ESTIMATE — settlement unconfirmed',
            scope='Positive net payable means forecast payment; negative means forecast receipt. MTM is separate. Excludes fees, fuel, taxes, invoices, payments, collateral and credit limits; no legal netting assumption.')
        forwards=g.loc[g.source_sheet.eq('fwdblotter')]
        row['forward_mtm_usd']=forwards.forward_mtm_usd.sum() if not forwards.empty and not forwards.status.eq('BLOCKED').any() and forwards.forward_mtm_usd.notna().all() and not unresolved else None
        row['forward_mtm_status']='NO FORWARD ALLOCATIONS' if forwards.empty else 'WITHHELD — MTM unavailable for one or more forward allocations' if pd.isna(row['forward_mtm_usd']) else 'INDICATIVE'
        summary.append(row)
    result=pd.DataFrame(summary)
    return {'counterparty_settlement_detail.csv':detail,'counterparty_settlement_monthly.csv':result,'counterparty_index_price_checks.csv':pd.DataFrame(price_checks,columns=['delivery_date','index_code','expected_cash_column','price','price_kind','source','status','issue'])},int(result.status.str.startswith('INCOMPLETE').sum())


def monthly_delivery_hedge_summary(reports):
    """Display procurement and pool forward MTM side by side; never add them into P&L."""
    balance=reports.get('monthly_physical_balance.csv',pd.DataFrame()).copy()
    if balance.empty:return pd.DataFrame()
    audit=reports.get('physical_trade_allocations.csv',pd.DataFrame()).copy()
    exposure=reports.get('physical_open_exposure_monthly.csv',pd.DataFrame()).copy()
    values=reports.get('forward_trade_valuation.csv',pd.DataFrame()).copy()
    for frame in [balance,audit,exposure,values]:
        if not frame.empty:
            frame['month']=pd.to_datetime(frame.month).dt.to_period('M').dt.to_timestamp()
            if 'index_code' in frame:frame['index_code']=frame.index_code.map(market)
    for frame in [balance,exposure]:
        if not frame.empty and frame.duplicated(['month','index_code']).any():
            raise ValueError('Duplicate month/book in risk summary source.')
    if not audit.empty and audit.duplicated(['month','source_id']).any():
        raise ValueError('Duplicate physical trade-month in risk summary source.')
    rows=[]
    for _,b in balance.iterrows():
        a=audit.loc[(audit.month==b.month)&(audit.delivery_book==b.index_code)] if not audit.empty else audit
        pool=a.loc[a.supply_scope.eq('POOL')] if not a.empty else a
        forwards=pool.loc[pool.source_sheet.eq('fwdblotter')] if not pool.empty else pool
        mtm=0.;missing=0
        for _,trade in forwards.iterrows():
            candidates=values.loc[(values.month==b.month)&(values.source_id==trade.source_id)] if not values.empty else values
            if len(candidates)!=1:
                missing+=1;continue
            v=candidates.iloc[0];value=num(v.get('indicative_mtm'));direction=num(v.get('directional_volume'))
            if (not txt(v.get('valuation_status')).startswith('INDICATIVE') or value is None or direction is None
                    or abs(direction-(trade.purchase_volume-trade.sale_volume))>1e-6 or v.index_code!=market(trade.index_code)):
                missing+=1;continue
            mtm+=value
        if forwards.empty:mtm=None;mtm_status='NO POOL FORWARD ALLOCATIONS'
        elif missing:mtm=None;mtm_status='INCOMPLETE — matching pool forward MTM withheld'
        else:mtm_status='INDICATIVE — pool forwards only'
        e=exposure.loc[(exposure.month==b.month)&(exposure.index_code==b.index_code)] if not exposure.empty else exposure
        er=e.iloc[0] if len(e)==1 else {}
        def quantity(col):return float(pool[col].sum()) if not pool.empty else 0.
        rows.append(dict(month=b.month.date().isoformat(),index_code=b.index_code,
            retail_demand_dth=b.additional_customer_demand,booked_delivered_purchases_dth=b.purchase_volume,
            booked_delivered_sales_dth=b.sale_volume,storage_injections_dth=b.storage_injections,
            storage_withdrawals_dth=b.storage_withdrawals,other_delivered_supply_dth=b.other_supply,
            delivered_purchase_required_dth=b.additional_purchase_required,
            delivered_surplus_dth=b.surplus_available,delivered_balance_status=b.status,
            delivered_forecast_usd_dth=er.get('forward_usd_dth'),
            delivered_purchase_cost_usd=er.get('purchase_cost_usd'),
            delivered_net_procurement_cost_usd=er.get('net_open_cost_usd'),
            delivered_cost_status=er.get('status','NOT GENERATED'),
            pool_purchases_dth=quantity('purchase_volume'),pool_sales_dth=quantity('sale_volume'),
            pool_net_position_dth=quantity('purchase_volume')-quantity('sale_volume'),
            pool_forward_allocations=len(forwards),pool_forward_missing_or_blocked_mtm=missing,
            pool_forward_mtm_usd=mtm,pool_forward_mtm_status=mtm_status,
            scope='Procurement cash forecast and pool forward MTM are separate measures; no combined P&L. Pool quantities include cash and forward trades; MTM covers matching forward allocations only. Storage delivery adjustments remain provisional.'))
    return pd.DataFrame(rows)


def validate_physical_profile(profile):
    if profile.get('schema')!='physical-balance-1' or type(profile.get('link_storage')) is not bool or type(profile.get('positive_volume_means_purchase')) is not bool:
        raise ValueError('Invalid saved physical-balance profile or volume convention.')
    inputs=pd.DataFrame(profile['inputs'])
    if inputs.empty:raise ValueError('Saved physical balance requires at least one input row.')
    checked=delivered_pricing_defaults(physical_balance(pd.DataFrame(),inputs))
    if (checked.month.max()-checked.month.min()).days>3660:raise ValueError('Physical input period exceeds ten years.')
    columns=['month','index_code','additional_customer_demand','other_supply','storage_withdrawals','storage_injections','delivered_price_usd_dth','delivery_differential_usd_dth','delivered_price_source']
    checked['month']=checked.month.dt.strftime('%Y-%m-01')
    return checked[columns]


def physical_profile_reports(ledger,flags,profile,reports):
    inputs=validate_physical_profile(profile)
    if profile['link_storage']:inputs=link_storage_balance_inputs(inputs,storage_physical_bridge(reports))
    start=pd.Timestamp(inputs.month.min());end=pd.Timestamp(inputs.month.max())+pd.offsets.MonthEnd(0)
    allocations=monthly_volumes(ledger,start,end,profile['positive_volume_means_purchase'],all_physical=True)
    balance,audit,pools=delivered_balance_reports(allocations,inputs)
    if profile['link_storage']:
        linked=balance.index_code.isin(['DOM','TCO']) & balance.status.isin(['BALANCED','SURPLUS','SHORTFALL'])
        balance.loc[linked,'status']='PROVISIONAL STORAGE — '+balance.loc[linked,'status']
    if active_unallocatable_records(ledger,flags).any():
        balance[['net_surplus_shortfall','additional_purchase_required','surplus_available']]=float('nan')
        balance['status']='REVIEW UNALLOCATABLE TRADES'
    review=int((~balance.status.isin(['BALANCED','SURPLUS','SHORTFALL'])).sum())
    return {'physical_balance_inputs.csv':inputs,'physical_trade_allocations.csv':audit,'pool_hedge_positions.csv':pools,'monthly_physical_balance.csv':balance,
        'physical_balance_scope.csv':pd.DataFrame([dict(link_storage=profile['link_storage'],positive_volume_means_purchase=profile['positive_volume_means_purchase'],
            scope='Delivered balance only: exact delivery locations net against retail demand; pool hedges are separate. Inventory-boundary storage is provisional. Utility-level totals do not establish point-level transport capability. Costs require delivered Gas Daily forecasts; not full P&L.')])},review


def build_run_archive(reports, provenance):
    """Create a portable evidence bundle; no credentials or source workbook bytes."""
    manifest = dict(schema_version=1, app_version=VERSION, provenance=provenance,
                    reports=[], missing_core_reports=[])
    core = ['complete_trade_ledger.csv', 'forward_pricing_coverage.csv', 'forward_trade_valuation.csv']
    manifest['missing_core_reports'] = [name for name in core if name not in reports]
    output = BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, frame in sorted(reports.items()):
            if not re.fullmatch(r'[A-Za-z0-9_-]+\.csv', name):
                raise ValueError('Invalid report filename.')
            content = frame.to_csv(index=False).encode('utf-8')
            status_counts = {}
            for column in frame.columns:
                if column == 'status' or column.endswith('_status'):
                    status_counts[column] = {str(k): int(v) for k, v in frame[column].fillna('MISSING').value_counts().items()}
            manifest['reports'].append(dict(file='reports/'+name, rows=len(frame),
                                             sha256=hashlib.sha256(content).hexdigest(),
                                             status_counts=status_counts))
            archive.writestr('reports/'+name, content)
        archive.writestr('manifest.json', json.dumps(manifest, indent=2, allow_nan=False))
        archive.writestr('README.txt',
            'Riverbend Risk run archive\n\n'
            'Reports reflect the current run and selected delivery ranges. Missing reports are not zero.\n'
            'Inspect manifest.json for provenance, report hashes, row counts and status counts.\n'
            'An archive is not approval of valuation, mapping, settlement or manual inputs.\n'
            'Scenario and physical balance reports may contain default or incomplete inputs.\n'
            'All trades are physical. Valuations remain indicative commodity estimates.\n'
            'Trade identifiers follow worksheet rows and may change when a workbook is reordered.\n'
            'Source files and credentials are excluded. Retain original source files separately.\n'
            'This portable download is not a server database or an automatic scheduled backup.\n')
    return output.getvalue(), manifest


def snapshot_id(manifest):
    identity = dict(manifest)
    identity['provenance'] = {k:v for k,v in manifest['provenance'].items() if k != 'created_at_utc'}
    return hashlib.sha256(json.dumps(identity,sort_keys=True,allow_nan=False).encode()).hexdigest()


SHADOW_FIELDS=['source_sheet','trade_date','start','end','buy_sell','quantity_dth_day','counterparty','index_code','physical_location','pricing_type','price_or_premium']


def shadow_terms(raw):
    t={k:raw.get(k) for k in SHADOW_FIELDS}
    for k in ['trade_date','start','end']:
        d=day(t[k])
        if pd.isna(d):raise ValueError(k+' is required.')
        t[k]=d.date().isoformat()
    if t['end']<t['start'] or (pd.Timestamp(t['end'])-pd.Timestamp(t['start'])).days>7305:raise ValueError('Invalid delivery period.')
    if t['source_sheet'] not in ['fwdblotter','cashblotter'] or t['buy_sell'] not in ['Buy','Sell']:raise ValueError('Choose a valid book and Buy/Sell direction.')
    t['counterparty']=canonical_counterparty(t['counterparty']);t['index_code']=market(t['index_code'])
    t['physical_location']=''.join(c for c in txt(t['physical_location']).lower() if c.isalnum())
    if t['physical_location']=='deodlvd':t['physical_location']='deodelivered'
    if not all(t[k] for k in ['counterparty','index_code','physical_location']):raise ValueError('Counterparty, price index and physical location are required.')
    if t['pricing_type'] not in ['FIXED','GDD','INDEX','PHYSB']:raise ValueError('Pricing classification is unresolved; review the Excel/WAPP mapping.')
    for k in ['quantity_dth_day','price_or_premium']:
        t[k]=num(t[k])
        if t[k] is None:raise ValueError(k+' must be a finite number.')
    if t['quantity_dth_day']<0:raise ValueError('Use nonnegative Dth/day and a separate Buy/Sell direction.')
    return t


def shadow_fingerprint(terms):
    t=shadow_terms(terms);t['counterparty']=t['counterparty'].casefold()
    return hashlib.sha256(json.dumps(t,sort_keys=True,allow_nan=False).encode()).hexdigest()


@st.cache_data(show_spinner=False,max_entries=8)
def shadow_excel_rows(ledger,classifications,cutoff,positive_buy=True):
    acknowledged=ledger.get('expired_exception_acknowledged',pd.Series(False,index=ledger.index)).fillna(False).astype(bool)
    scope=ledger.loc[ledger.end.ge(pd.Timestamp(cutoff)) | (ledger.end.isna() & ~acknowledged)].copy()
    classified=physical_delivery_classification(scope);rows=[]
    id_counts=ledger.get('risk_deal_id',pd.Series(dtype=str)).map(txt).value_counts()
    for _,r in classified.iterrows():
        cls=txt(r.get('pricing_label')).upper();term=r.get('fixed_price') if cls=='FIXED' else r.get('premium')
        if r.source_sheet=='fwdblotter':
            hits=classifications.loc[classifications.source_id.eq(r.source_id)] if not classifications.empty else classifications
            cls=txt(hits.iloc[0].get('candidate_classification')) if len(hits)==1 else ''
            term=r.get('contract_price')
        elif cls in ['GDD','INDEX'] and pd.isna(term) and not r.get('premium_issue',False):term=0.
        signed=num(r.daily_volume);signed=signed*(1 if positive_buy else -1) if signed is not None else None
        raw=dict(source_sheet=r.source_sheet,trade_date=r.trade_date,start=r.start,end=r.end,buy_sell='Sell' if signed is not None and signed<0 else 'Buy',quantity_dth_day=abs(signed) if signed is not None else None,counterparty=r.counterparty,index_code=r.index_code,physical_location=r.delivery_point,pricing_type=cls,price_or_premium=term)
        try:terms=shadow_terms(raw);fingerprint=shadow_fingerprint(terms);issue=''
        except ValueError as exc:terms=raw;fingerprint='';issue=str(exc)
        excel_id=txt(r.get('risk_deal_id'))
        if excel_id and (not re.fullmatch(r'RB-[0-9a-f]{32}',excel_id) or id_counts.get(excel_id,0)!=1):issue='Invalid or duplicate Risk Deal ID in Excel.'
        rows.append(dict(source_id=r.source_id,excel_deal_id=excel_id,terms=terms,fingerprint=fingerprint,issue=issue))
    return rows


def validate_shadow_register(data):
    if data.get('schema')!='shadow-deal-register-1' or not isinstance(data.get('records'),list):raise ValueError('Invalid shadow register format.')
    ids=set()
    for r in data['records']:
        if not re.fullmatch(r'RB-[0-9a-f]{32}',r.get('deal_id','')) or r['deal_id'] in ids:raise ValueError('Invalid or duplicate permanent deal ID.')
        ids.add(r['deal_id']);shadow_terms(r['terms'])
        if r.get('state') not in ['Draft','Excel linked','Cancelled draft']:raise ValueError('Invalid shadow deal state.')
    return data


def shadow_store(target,data=None,expected_revision=None,event='Read'):
    """Append-only revisions; a unique revision prevents concurrent overwrite."""
    if not target:raise ValueError('Configure durable history storage before saving a shadow register.')
    if data is not None:validate_shadow_register(data)
    pg=isinstance(target,dict);namespace=target['namespace'] if pg else 'riverbend-risk';connection=None
    try:
        if pg:
            import psycopg
            from psycopg.conninfo import conninfo_to_dict
            options=conninfo_to_dict(target['database_url'])
            if options.get('sslmode') not in ['require','verify-ca','verify-full']:options['sslmode']='require'
            options.update(connect_timeout=10,options='-c statement_timeout=15000 -c lock_timeout=10000')
            connection=psycopg.connect(**options)
        else:connection=history_connection(target)
        connection.execute('CREATE TABLE IF NOT EXISTS riverbend_shadow_register (namespace TEXT NOT NULL, revision INTEGER NOT NULL, created_at TEXT NOT NULL, event TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(namespace,revision))')
        connection.commit()
        if not pg and data is not None:connection.execute('BEGIN IMMEDIATE')
        mark='%s' if pg else '?'
        found=connection.execute('SELECT revision,payload FROM riverbend_shadow_register WHERE namespace='+mark+' ORDER BY revision DESC LIMIT 1',(namespace,)).fetchone()
        revision=int(found[0]) if found else 0
        current=validate_shadow_register(json.loads(found[1])) if found else dict(schema='shadow-deal-register-1',records=[])
        if data is None:return revision,current
        if expected_revision!=revision:raise ValueError('The register changed in another session. Refresh and review before saving again.')
        if data==current:return revision,current
        payload=json.dumps(data,allow_nan=False)
        if len(payload.encode())>20*1024*1024:raise ValueError('Register exceeds the supported size.')
        connection.execute('INSERT INTO riverbend_shadow_register (namespace,revision,created_at,event,payload) VALUES ('+','.join([mark]*5)+')',(namespace,revision+1,pd.Timestamp.now(tz='UTC').isoformat(),event,payload))
        connection.commit();return revision+1,data
    except ValueError:raise
    except Exception:
        raise ValueError('Shadow register could not be read or saved. Check database access, or refresh after a concurrent edit. No overwrite was applied.') from None
    finally:
        if connection is not None:connection.close()


def shadow_reconcile(data,excel,workbook_hash,cutoff):
    validate_shadow_register(data);valid=[x for x in excel if not x['issue']]
    records=[r for r in data['records'] if r['state']!='Cancelled draft' and r['terms']['end']>=str(cutoff)]
    rows=[];used=set()
    for r in records:
        fp=shadow_fingerprint(r['terms']);id_hits=[x for x in valid if x.get('excel_deal_id')==r['deal_id']]
        hits=[x for x in valid if x['fingerprint']==fp and (not x.get('excel_deal_id') or x['excel_deal_id']==r['deal_id'])]
        peers=[x for x in records if shadow_fingerprint(x['terms'])==fp]
        bound=[x for x in hits if x['source_id']==r.get('excel_source_id') and workbook_hash==r.get('workbook_hash')]
        match=id_hits[0] if len(id_hits)==1 and id_hits[0]['fingerprint']==fp and r['state']=='Excel linked' else None
        if not id_hits:match=bound[0] if len(bound)==1 else hits[0] if len(hits)==1 and len(peers)==1 and r['state']=='Excel linked' else None
        if match is not None and match['source_id'] not in used:
            used.add(match['source_id']);status='MATCHED';sid=match['source_id']
        else:
            sid='';status='REVIEW — identified Excel deal amended' if id_hits and id_hits[0]['fingerprint']!=fp else 'DRAFT MATCH CANDIDATE — link explicitly' if r['state']=='Draft' and len(hits)==1 and len(peers)==1 else 'AMBIGUOUS — confirm individual IDs' if hits else 'DRAFT — not found in Excel' if r['state']=='Draft' else 'REVIEW — missing or amended in Excel'
        rows.append(dict(deal_id=r['deal_id'],state=r['state'],counterparty=r['terms']['counterparty'],source_id=sid,candidate_source_ids='; '.join(x['source_id'] for x in hits),status=status))
    for x in excel:
        if x['source_id'] not in used:rows.append(dict(deal_id='',state='Excel',counterparty=x['terms'].get('counterparty',''),source_id=x['source_id'],candidate_source_ids='',status='INVALID — '+x['issue'] if x['issue'] else 'EXCEL ROW — not linked in register'))
    return pd.DataFrame(rows,columns=['deal_id','state','counterparty','source_id','candidate_source_ids','status'])


def shadow_apply(data,excel,workbook_hash,cutoff,operation,source_ids=None,deal_id=None,terms=None):
    import uuid
    result=json.loads(json.dumps(validate_shadow_register(data)));stamp=pd.Timestamp.now(tz='UTC').isoformat()
    if operation=='draft':
        result['records'].append(dict(deal_id='RB-'+uuid.uuid4().hex,state='Draft',terms=shadow_terms(terms),created_at=stamp,updated_at=stamp,excel_source_id='',workbook_hash=''))
        return result
    report=shadow_reconcile(data,excel,workbook_hash,cutoff)
    occupied=set(report.loc[report.status.eq('MATCHED'),'source_id'])
    sources={x['source_id']:x for x in excel}
    if operation=='register':
        for sid in dict.fromkeys(source_ids or []):
            if sid in occupied:continue
            x=sources[sid]
            if x['issue']:raise ValueError('Resolve invalid Excel row '+sid+' before registering it.')
            # A matching draft must be linked rather than registered a second time.
            if any(r['state']=='Draft' and shadow_fingerprint(r['terms'])==x['fingerprint'] for r in result['records']):raise ValueError('A matching draft exists. Link the draft to '+sid+' instead of creating another ID.')
            new_id=x.get('excel_deal_id') or 'RB-'+uuid.uuid4().hex
            if any(r['deal_id']==new_id for r in result['records']):raise ValueError('Excel references an existing deal ID. Use link/amendment instead of registering it again.')
            result['records'].append(dict(deal_id=new_id,state='Excel linked',terms=x['terms'],created_at=stamp,updated_at=stamp,excel_source_id=sid,workbook_hash=workbook_hash))
    elif operation in ['link','amend']:
        targets=[r for r in result['records'] if r['deal_id']==deal_id]
        if len(targets)!=1 or len(source_ids or [])!=1:raise ValueError('Select one existing deal and one Excel row.')
        r=targets[0];sid=source_ids[0];x=sources[sid]
        if x['issue']:raise ValueError('The selected Excel row has unresolved fields.')
        if x.get('excel_deal_id') and x['excel_deal_id']!=deal_id:raise ValueError('Excel specifies a different permanent deal ID. Correct the ID or choose that deal.')
        occupied_other=report.loc[report.status.eq('MATCHED') & report.deal_id.ne(deal_id),'source_id']
        if sid in set(occupied_other):raise ValueError('This Excel row is already linked to another deal ID.')
        if operation=='link' and shadow_fingerprint(r['terms'])!=x['fingerprint']:raise ValueError('Terms differ. Review them and use Accept Excel amendment explicitly.')
        if r['state']=='Cancelled draft':raise ValueError('Cancelled drafts cannot be linked.')
        r.update(state='Excel linked',terms=x['terms'],updated_at=stamp,excel_source_id=sid,workbook_hash=workbook_hash)
    else:raise ValueError('Unsupported shadow register action.')
    return validate_shadow_register(result)


def shadow_bulk_validate(frame,data,excel):
    """Validate the entire batch before a single database revision is written."""
    seen={shadow_fingerprint(r['terms']) for r in data['records'] if r['state']!='Cancelled draft'}
    seen.update(x['fingerprint'] for x in excel if not x['issue'])
    terms=[];checks=[]
    for i,r in enumerate(frame.to_dict('records'),1):
        if not any(txt(r.get(k)) for k in SHADOW_FIELDS):continue
        issue='';t=None
        try:
            t=shadow_terms(r)
            if t['quantity_dth_day']<=0:raise ValueError('Enter a positive quantity for a new draft.')
            fp=shadow_fingerprint(t)
            duplicate=fp in seen
            allow=r.get('allow_identical')
            confirmed=allow is True or str(allow).lower()=='true'
            if duplicate and not confirmed:raise ValueError('Matching terms exist in Excel, register or this batch. Link an existing deal, or confirm this is a separate deal with identical terms.')
            seen.add(fp)
        except (ValueError,TypeError,KeyError) as exc:issue=str(exc)
        if t is not None and not issue:terms.append(t)
        checks.append(dict(row=i,status='BLOCKED' if issue else 'READY',issue=issue))
    return terms,pd.DataFrame(checks,columns=['row','status','issue'])


def shadow_bulk_apply(frame,data,excel,workbook_hash,cutoff):
    terms,checks=shadow_bulk_validate(frame,data,excel)
    if checks.empty or checks.status.ne('READY').any():raise ValueError('Resolve every row before saving; no drafts were saved.')
    updated=data
    for t in terms:updated=shadow_apply(updated,excel,workbook_hash,cutoff,'draft',terms=t)
    return updated


def shadow_shift_rows(rows,days):
    result=[]
    for raw in rows:
        r=dict(raw)
        for k in ['start','end']:
            d=day(r.get(k))
            if pd.isna(d):raise ValueError('Valid start and end dates are required before shifting.')
            r[k]=(d+pd.Timedelta(days=int(days))).date().isoformat()
        r['selected']=False;r['allow_identical']=False
        result.append(r)
    return result


def shadow_deal_group(excel,mode,selected_date,filters=None):
    chosen=day(selected_date)
    if pd.isna(chosen):return []
    filters=filters or {};result=[]
    for x in excel:
        t=x['terms']
        if mode=='Trade date':matches=day(t.get('trade_date'))==chosen
        elif mode=='Delivery date':matches=day(t.get('start'))<=chosen<=day(t.get('end'))
        else:raise ValueError('Choose Trade date or Delivery date.')
        if matches and all(not values or t.get(k) in values for k,values in filters.items()):result.append(x)
    return result


def shadow_copy_group(group,shift,new_trade_date):
    if not group:raise ValueError('No deals match this group.')
    if any(x['issue'] for x in group):raise ValueError('Resolve the flagged worksheet deals or narrow the filters before copying this group.')
    date_=day(new_trade_date)
    if pd.isna(date_):raise ValueError('Choose a valid new trade date.')
    rows=shadow_shift_rows([x['terms'] for x in group],shift)
    for r in rows:r['trade_date']=date_.date().isoformat()
    return rows


def render_shadow_group(excel,rows,load):
    st.markdown('**Copy a group of worksheet deals**')
    st.caption('All matching rows are included automatically. Source deals follow the reconciliation cutoff above; move it earlier to include older deals. Delivery-date selection includes contracts active on that date and copies their full duration.')
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    dates=[day(x['terms'].get('trade_date')) for x in excel]
    prior=[d for d in dates if pd.notna(d) and d<today]
    preset=st.selectbox('Group selection',['Previous trade day','Specific trade date','Delivery date'],key='bulk_group_mode')
    if preset=='Previous trade day':
        selected_date=max(prior) if prior else None
        if selected_date is None:st.info('No earlier trade date is available within the current reconciliation cutoff.');return
        st.caption('Previous available trade day: '+selected_date.date().isoformat())
    else:
        selected_date=st.date_input('Source group date',value=today.date(),key='bulk_group_date_'+preset)
    filters={}
    for field,label in [('source_sheet','Book'),('counterparty','Counterparty'),('physical_location','Physical location')]:
        options=sorted({txt(x['terms'].get(field)) for x in excel if txt(x['terms'].get(field))})
        filters[field]=st.multiselect(label+' filter (empty means all)',options,key='bulk_group_filter_'+field)
    group=shadow_deal_group(excel,'Delivery date' if preset=='Delivery date' else 'Trade date',selected_date,filters)
    st.write('Matching worksheet deals:',len(group))
    if group:
        st.dataframe(pd.DataFrame([dict(source_id=x['source_id'],**x['terms'],issue=x['issue']) for x in group]),hide_index=True,use_container_width=True)
    new_trade_date=st.date_input('New trade date for copied group',value=today.date(),key='bulk_group_new_trade_date')
    default_shift=(pd.Timestamp(new_trade_date)-pd.Timestamp(selected_date)).days
    shift=st.number_input('Group delivery date shift (days)',value=int(default_shift),step=1,key='bulk_group_shift_'+str(selected_date)+'_'+str(new_trade_date))
    try:
        copies=shadow_copy_group(group,shift,new_trade_date)
    except ValueError as exc:
        copies=[]
        if group:st.warning(str(exc))
    if copies:
        st.caption('New dates below: delivery durations and commercial terms are preserved. Review the draft grid before saving.')
        st.dataframe(pd.DataFrame(copies)[['counterparty','trade_date','start','end','buy_sell','quantity_dth_day','physical_location','price_or_premium']],hide_index=True,use_container_width=True)
    if st.button('Copy entire group to draft grid',disabled=not copies,key='bulk_copy_group'):
        load(rows+copies)


def shadow_recent_dates(excel,asof):
    today=day(asof)
    return sorted({day(x['terms'].get('trade_date')) for x in excel if pd.notna(day(x['terms'].get('trade_date'))) and day(x['terms'].get('trade_date'))<today},reverse=True)[:2]


@st.cache_data(show_spinner=False,max_entries=4)
def shadow_template_bytes(excel,start,end):
    """Runtime Excel export: reference rows never become new trades on import."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb=Workbook();guide=wb.active;guide.title='Instructions'
    for line in ['Mercury — deal entry workbook',
        'Reference trade dates: '+str(start)+' through '+str(end)+' (inclusive).',
        'Copy columns A:K from Reference deals to New deals. Edit trade and delivery dates, volumes and prices.',
        'Upload this workbook in the app. Only New deals is read; reference rows are never imported.',
        'Review the draft grid and save. Excel remains the official position source.',
        'Rows with a source_issue need correction before use. Negative differentials are allowed; quantity is positive and Buy/Sell determines direction.']:
        guide.append([line])
    guide.column_dimensions['A'].width=110
    for row in guide:
        row[0].alignment=Alignment(wrap_text=True,vertical='top');guide.row_dimensions[row[0].row].height=32
    ref=wb.create_sheet('Reference deals');ref.append(SHADOW_FIELDS+['source_id','source_issue'])
    def safe(value):
        if value is None or (not isinstance(value,str) and pd.isna(value)):return None
        return value.to_pydatetime() if isinstance(value,pd.Timestamp) else value
    for x in excel:
        t=x['terms'];ref.append([safe(t.get(k)) for k in SHADOW_FIELDS]+[x['source_id'],x['issue']])
    new=wb.create_sheet('New deals');new.append(SHADOW_FIELDS)
    for ws in [ref,new]:
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.fill=PatternFill('solid',fgColor='17365D');cell.font=Font(color='FFFFFF',bold=True)
        for col in range(1,ws.max_column+1):ws.column_dimensions[ws.cell(1,col).column_letter].width=23 if col!=13 else 65
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                # Treat source strings as text, never executable Excel formulas.
                if isinstance(cell.value,str):cell.data_type='s'
                if cell.column in [7,11]:cell.number_format='#,##0.00000;[Red](#,##0.00000);–'
    output=BytesIO();wb.save(output);wb.close();return output.getvalue()


def shadow_template_import(content):
    wb=load_workbook(BytesIO(content),read_only=True,data_only=False)
    try:
        if 'New deals' not in wb.sheetnames:raise ValueError('The workbook must contain a New deals sheet.')
        ws=wb['New deals'];it=ws.iter_rows(values_only=True);header=list(next(it,()))
        if header!=SHADOW_FIELDS:raise ValueError('Keep the New deals header unchanged; use the downloaded template.')
        rows=[]
        for i,values in enumerate(it,2):
            if not any(txt(v) for v in values):continue
            if len(rows)>=1000:raise ValueError('Import at most 1,000 new deals per batch.')
            if any(isinstance(v,str) and v.startswith('=') for v in values):raise ValueError('Paste formulas as values in New deals before importing (row '+str(i)+').')
            rows.append(dict(zip(SHADOW_FIELDS,values)))
        if not rows:raise ValueError('New deals is empty. Copy and edit reference rows into New deals first.')
        return rows
    finally:wb.close()


def render_shadow_checkboxes(excel,rows,load,digest):
    st.markdown('**Check worksheet deals to copy**')
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    dates=shadow_recent_dates(excel,today)
    available=sorted({day(x['terms'].get('trade_date')).date() for x in excel if pd.notna(day(x['terms'].get('trade_date')))})
    with st.form('source_filter_form'):
        chosen=st.multiselect('Source trade dates',available,default=[d.date() for d in dates],key='check_dates_'+digest[:12])
        filters={}
        for field,label in [('source_sheet','Book'),('counterparty','Counterparty'),('physical_location','Location')]:
            filters[field]=st.multiselect(label+' (empty means all)',sorted({txt(x['terms'].get(field)) for x in excel if txt(x['terms'].get(field))}),key='check_filter_'+field)
        st.form_submit_button('Apply source filters')
    group=[x for x in excel if pd.notna(day(x['terms'].get('trade_date'))) and day(x['terms'].get('trade_date')).date() in chosen and all(not v or x['terms'].get(k) in v for k,v in filters.items())]
    signature=hashlib.sha256(json.dumps([digest,chosen,filters,[x['source_id'] for x in group]],default=str,sort_keys=True).encode()).hexdigest()[:16]
    base='source_checks_'+signature;revision=st.session_state.get(base+'_revision',0)
    selected_ids=set(st.session_state.get(base+'_ids',[]))
    if st.button('Select all displayed deals',key=base+'_all'):
        selected_ids={x['source_id'] for x in group};revision+=1
        st.session_state[base+'_ids']=list(selected_ids);st.session_state[base+'_revision']=revision
    if st.button('Clear selection',key=base+'_clear'):
        selected_ids=set();revision+=1
        st.session_state[base+'_ids']=[];st.session_state[base+'_revision']=revision
    cols=['copy','source_id',*SHADOW_FIELDS,'issue']
    frame=pd.DataFrame([dict(copy=x['source_id'] in selected_ids,source_id=x['source_id'],**x['terms'],issue=x['issue']) for x in group],columns=cols)
    frame['copy']=frame['copy'].astype(bool)
    with st.form('source_copy_form_'+signature):
        checked=st.data_editor(frame,disabled=cols[1:],hide_index=True,use_container_width=True,key=base+'_'+str(revision),column_config={'copy':st.column_config.CheckboxColumn('Copy')})
        st.caption(str(len(group))+' displayed deals. Check all desired rows, then click Copy; individual ticks do not refresh the app.')
        newdate=st.date_input('Trade date for new copies',value=today.date(),key='checked_newdate')
        shift=st.number_input('Shift copied delivery dates by days (0 keeps original dates)',value=0,step=1,key='checked_shift')
        copy_clicked=st.form_submit_button('Copy checked deals to draft grid',disabled=not group)
    if copy_clicked:
        ids=set(checked.loc[checked['copy'].eq(True),'source_id'])
        selected=[x for x in group if x['source_id'] in ids]
        try:
            copies=shadow_copy_group(selected,shift,newdate)
            st.session_state[base+'_ids']=[];st.session_state[base+'_revision']=revision+1
            load(rows+copies)
        except ValueError as exc:st.error(str(exc))


def shadow_grid_action(rows,action,scope='All draft rows',trade_date=None,start=None,end=None,shift=0):
    result=[dict(r) for r in rows]
    selected=lambda r:r.get('selected') is True or str(r.get('selected')).lower()=='true'
    if action=='clear':return []
    if action=='delete':
        if not any(selected(r) for r in result):raise ValueError('Check the Select box beside the rows to delete.')
        return [r for r in result if not selected(r)]
    targets=[i for i,r in enumerate(result) if scope=='All draft rows' or selected(r)]
    if not targets:raise ValueError('No draft rows selected for this action.')
    td=day(trade_date)
    if pd.isna(td):raise ValueError('Choose a valid bulk trade date.')
    if action=='duplicate':
        picked=[r for r in result if selected(r)]
        if not picked:raise ValueError('Select draft rows to duplicate.')
        copies=shadow_shift_rows(picked,shift)
        for r in copies:r['trade_date']=td.date().isoformat()
        return [dict(r,selected=False) for r in result]+copies
    if action=='set_dates':
        first,last=day(start),day(end)
        if pd.isna(first) or pd.isna(last) or last<first:raise ValueError('Delivery end must be on or after delivery start.')
        for i in targets:result[i].update(trade_date=td.date().isoformat(),start=first.date().isoformat(),end=last.date().isoformat(),allow_identical=False)
    elif action=='shift':
        for i in targets:
            updated=shadow_shift_rows([result[i]],shift)[0]
            updated.update(trade_date=td.date().isoformat(),selected=selected(result[i]));result[i]=updated
    else:raise ValueError('Unsupported draft-grid action.')
    return result


def shadow_cancel_drafts(data,ids):
    result=json.loads(json.dumps(validate_shadow_register(data)))
    ids=set(ids)
    if not ids:raise ValueError('Select one or more saved drafts.')
    matches=[r for r in result['records'] if r['deal_id'] in ids]
    if len(matches)!=len(ids) or any(r['state']!='Draft' for r in matches):raise ValueError('Only unlinked saved drafts can be removed here; Excel-linked trades are unchanged.')
    stamp=pd.Timestamp.now(tz='UTC').isoformat()
    for r in matches:r.update(state='Cancelled draft',updated_at=stamp)
    return validate_shadow_register(result)


def render_shadow_bulk(data,excel,digest,cutoff,save,ledger=None,classifications=None,positive_buy=True):
    with st.expander('Excel-style bulk draft entry',expanded=True):
        st.caption('Paste a table, copy existing worksheet deals, or edit the grid. Delivery dates shift together; trade dates remain editable. Saved drafts stay outside the official Excel risk book.')
        today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
        with st.form('template_period_form'):
            period=st.date_input('Template and source trade-date range (last two months)',value=((today-pd.DateOffset(months=2)).date(),today.date()),key='template_trade_range')
            st.form_submit_button('Apply source date range')
        sources=[]
        if len(period)==2 and ledger is not None:
            source_ledger=ledger.loc[pd.to_datetime(ledger.trade_date).between(pd.Timestamp(period[0]),pd.Timestamp(period[1]))].copy()
            sources=shadow_excel_rows(source_ledger,classifications,pd.Timestamp('1900-01-01'),positive_buy)
        validation_excel=list({x['source_id']:x for x in [*excel,*sources]}.values())
        columns=['selected',*SHADOW_FIELDS,'allow_identical']
        rev=st.session_state.get('_bulk_rev',0)
        def load(rows):
            st.session_state['_bulk_rows']=rows;st.session_state['_bulk_rev']=rev+1;st.rerun()
        source_container=st.container()
        st.markdown('**New draft rows — edit copies here**')
        seed=pd.DataFrame(st.session_state.get('_bulk_rows',[]),columns=columns)
        for k in SHADOW_FIELDS:
            if k in ['quantity_dth_day','price_or_premium']:seed[k]=seed[k].map(num).astype(float)
            elif k in ['trade_date','start','end']:seed[k]=pd.to_datetime(seed[k],errors='coerce')
            else:seed[k]=seed[k].astype('string')
        for k in ['selected','allow_identical']:seed[k]=seed[k].fillna(False).astype(bool)
        with st.form('draft_edit_form_'+str(rev)):
            grid=st.data_editor(seed,num_rows='dynamic',hide_index=True,use_container_width=True,key='bulk_grid_'+str(rev),column_config={
                'selected':st.column_config.CheckboxColumn('Select'),
                'source_sheet':st.column_config.SelectboxColumn('Book',options=['cashblotter','fwdblotter']),
                'buy_sell':st.column_config.SelectboxColumn('Buy / Sell',options=['Buy','Sell']),
                'pricing_type':st.column_config.SelectboxColumn('Pricing',options=['FIXED','GDD','INDEX','PHYSB']),
                'trade_date':st.column_config.DateColumn('Trade date'),'start':st.column_config.DateColumn('Delivery start'),'end':st.column_config.DateColumn('Delivery end'),
                'allow_identical':st.column_config.CheckboxColumn('Separate identical deal'),
                'quantity_dth_day':st.column_config.NumberColumn('Dth/day',min_value=0.),
                'price_or_premium':st.column_config.NumberColumn('Price / differential',format='%.5f')})
            st.markdown('**Batch dates and cleanup**')
            scope=st.selectbox('Apply dates to',['All draft rows','Selected draft rows'],key='draft_date_scope')
            batch_trade=st.date_input('Bulk trade date',value=today.date(),key='draft_batch_trade')
            batch_start=st.date_input('Bulk delivery start',value=(today+pd.Timedelta(days=1)).date(),key='draft_batch_start')
            batch_end=st.date_input('Bulk delivery end',value=(today+pd.Timedelta(days=1)).date(),key='draft_batch_end')
            set_dates=st.form_submit_button('Apply trade date and delivery strip')
            batch_shift=st.number_input('Shift existing delivery dates by days',value=1,step=1,key='draft_batch_shift')
            shift_dates=st.form_submit_button('Apply trade date and shift delivery dates')
            duplicate_rows=st.form_submit_button('Duplicate selected rows with trade date and delivery shift')
            st.caption('Set delivery strip gives every targeted row the same start/end dates. Shift moves both existing delivery dates and preserves each duration. Both actions apply Bulk trade date. Quantity remains Dth/day.')
            delete_rows=st.form_submit_button('Delete selected draft rows')
            clear_rows=st.form_submit_button('Clear entire draft grid')
            undo_grid=st.form_submit_button('Undo last grid date/delete change')
            apply_edits=st.form_submit_button('Apply draft edits and validate')
            save_batch=st.form_submit_button('Validate and save batch as shadow drafts')
        rows=grid.to_dict('records')
        if set_dates or shift_dates or duplicate_rows or delete_rows or clear_rows:
            try:
                action='set_dates' if set_dates else 'shift' if shift_dates else 'duplicate' if duplicate_rows else 'delete' if delete_rows else 'clear'
                updated=shadow_grid_action(rows,action,scope,batch_trade,batch_start,batch_end,batch_shift)
                st.session_state['_bulk_undo']=rows
                load(updated)
            except ValueError as exc:st.error(str(exc))
        if undo_grid:
            previous=st.session_state.pop('_bulk_undo',None)
            if previous is None:st.info('No grid date/delete change to undo.')
            else:load(previous)
        if apply_edits:st.success('Draft edits applied; validation results are below.')
        st.caption('Apply draft edits before using other controls. Save validates the current edits before writing any drafts.')
        selected=[r for r in rows if r.get('selected') is True or str(r.get('selected')).lower()=='true']
        with source_container:
            render_shadow_checkboxes(sources,rows,load,digest)
        if len(period)==2:
            st.download_button('Download two-month Excel template',shadow_template_bytes(sources,*period),file_name='mercury_deal_entry_two_months.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='two_month_template')
        upload=st.file_uploader('Upload completed template (New deals sheet only)',type=['xlsx'],key='completed_entry_template')
        if upload is not None and st.button('Append New deals to draft grid',key='append_template'):
            try:load(rows+shadow_template_import(upload.getvalue()))
            except (ValueError,KeyError,zipfile.BadZipFile) as exc:st.error(str(exc))
        paste=st.text_area('Paste Excel cells here (include column headers from the template)',key='bulk_paste_text')
        if st.button('Append pasted Excel rows',disabled=not paste.strip(),key='bulk_append_paste'):
            try:
                pasted=pd.read_csv(StringIO(paste),sep='\t',dtype=str,keep_default_na=False)
                if set(pasted.columns)!=set(SHADOW_FIELDS):raise ValueError('Paste all template columns including the header row; do not include selection or confirmation columns.')
                if len(pasted)>1000:raise ValueError('Paste at most 1,000 rows per batch.')
                load(rows+pasted.to_dict('records'))
            except (ValueError,pd.errors.ParserError) as exc:st.error(str(exc))
        terms,checks=shadow_bulk_validate(grid,data,validation_excel)
        st.dataframe(checks,hide_index=True,use_container_width=True)
        if terms:
            summary=pd.DataFrame(terms)
            summary['delivery_days']=(pd.to_datetime(summary.end)-pd.to_datetime(summary.start)).dt.days+1
            summary['total_dth']=summary.quantity_dth_day*summary.delivery_days
            st.dataframe(summary.groupby('buy_sell',as_index=False).agg(deals=('buy_sell','size'),total_dth=('total_dth','sum')),hide_index=True)
            st.caption('Preview totals include valid rows only. All rows must pass before the batch can be saved.')
        if save_batch:
            try:save(shadow_bulk_apply(grid,data,validation_excel,digest,cutoff),'Bulk shadow drafts saved; Excel risk book unchanged')
            except ValueError as exc:st.error(str(exc))


def render_shadow_register(ledger,classifications,book,positive_buy):
    st.subheader('Parallel-run deal register — Excel remains authoritative')
    st.caption('This separate shadow register never adds drafts or register rows to the risk book. Valuations and molecule balances continue to read Excel only. Permanent IDs and register revisions are retained in the configured history database. Namespaces are not user permissions; this release does not implement approvals or automatic trade promotion.')
    try:
        target=history_target(dict(st.secrets.get('history',{})))
        revision,data=shadow_store(target)
    except (ValueError,FileNotFoundError,KeyError) as exc:
        st.info(str(exc) if isinstance(exc,ValueError) else 'Configure durable history storage to enable the parallel-run register.');return
    cutoff=st.date_input('Reconcile deals delivering on or after',value=pd.Timestamp.now(tz='America/Chicago').date().replace(day=1),key='shadow_cutoff')
    digest=hashlib.sha256(book.getvalue()).hexdigest();excel=shadow_excel_rows(ledger,classifications,cutoff,positive_buy)
    report=shadow_reconcile(data,excel,digest,cutoff)
    st.write('Register revision:',revision)
    st.session_state['_run_reports']['parallel_register_status.csv']=pd.DataFrame([dict(revision=revision,status='NOT STARTED' if not revision else 'REVIEW' if report.status.ne('MATCHED').any() else 'MATCHED',cutoff=str(cutoff),scope='Excel remains authoritative; shadow register never supplies risk positions.')])
    show(report,'parallel_deal_reconciliation.csv')
    flat=[dict(deal_id=r['deal_id'],state=r['state'],**r['terms'],excel_source_id=r.get('excel_source_id',''),workbook_hash=r.get('workbook_hash',''),updated_at=r['updated_at']) for r in data['records']]
    show(pd.DataFrame(flat),'parallel_deal_register.csv')
    st.download_button('Download register backup JSON',json.dumps(data,indent=2),file_name='shadow_deal_register_backup.json',mime='application/json',key='shadow_backup')
    def save(updated,event):
        shadow_store(target,updated,revision,event)
        if event.startswith('Bulk'):
            st.session_state['_bulk_rows']=[];st.session_state.pop('_bulk_undo',None);st.session_state['_bulk_rev']=st.session_state.get('_bulk_rev',0)+1
        st.session_state['_shadow_saved_message']=event;st.rerun()
    if st.session_state.get('_shadow_saved_message'):st.success(st.session_state.pop('_shadow_saved_message'))
    drafts=[r for r in data['records'] if r['state']=='Draft']
    if drafts:
        with st.expander('Remove saved shadow drafts'):
            st.caption('Removes drafts from the active register and retains their history. Excel-linked trades and the unsaved draft grid are unchanged.')
            with st.form('cancel_saved_drafts'):
                frame=pd.DataFrame([dict(remove=False,deal_id=r['deal_id'],**r['terms']) for r in drafts])
                checked=st.data_editor(frame,disabled=list(frame.columns[1:]),hide_index=True,use_container_width=True,key='cancel_saved_'+str(revision))
                remove_selected=st.form_submit_button('Remove checked saved drafts')
                remove_all=st.form_submit_button('Clear all saved unlinked drafts')
            if remove_selected or remove_all:
                ids=[r['deal_id'] for r in drafts] if remove_all else checked.loc[checked.remove.eq(True),'deal_id'].tolist()
                try:save(shadow_cancel_drafts(data,ids),'Saved drafts cancelled; Excel book unchanged')
                except ValueError as exc:st.error(str(exc))
    preview=pd.DataFrame([dict(source_id=x['source_id'],risk_deal_id=x.get('excel_deal_id',''),**x['terms'],issue=x['issue']) for x in excel])
    with st.expander('Excel import preview and registration',expanded=False):
        st.caption('Review the loaded workbook selection. Initial registration assigns one ID per valid worksheet row, preserving identical-term deals as separate records. Reimporting matched rows does not create new IDs. Missing rows are flagged, never automatically cancelled.')
        st.dataframe(preview,hide_index=True,use_container_width=True)
        occupied=set(report.loc[report.status.eq('MATCHED'),'source_id'])
        eligible=[x['source_id'] for x in excel if not x['issue'] and x['source_id'] not in occupied]
        if not data['records']:
            if st.button('Register valid Excel rows as initial baseline',disabled=not eligible,key='shadow_initialize'):
                try:save(shadow_apply(data,excel,digest,cutoff,'register',source_ids=eligible),'Initial Excel baseline registered')
                except (ValueError,KeyError) as exc:st.error(str(exc))
        else:
            selected=st.multiselect('Excel rows explicitly confirmed as new deals',eligible,key='shadow_new_rows')
            if st.button('Register selected rows as new deals',disabled=not selected,key='shadow_register_new'):
                try:save(shadow_apply(data,excel,digest,cutoff,'register',source_ids=selected),'Selected Excel deals registered')
                except (ValueError,KeyError) as exc:st.error(str(exc))
    render_shadow_bulk(data,excel,digest,cutoff,save,ledger,classifications,positive_buy)
    with st.expander('Enter a shadow deal draft'):
        st.caption('Price value = outright $/Dth for FIXED; index premium $/Dth for GDD/INDEX. Quantity is Dth/day. Save here, enter the same deal in Excel, then reload Excel and link the match. A draft stays outside the official book.')
        with st.form('shadow_draft_form',clear_on_submit=False):
            sheet=st.selectbox('Draft book',['fwdblotter','cashblotter']);cp=st.text_input('Draft counterparty')
            today=pd.Timestamp.now(tz='America/Chicago').date()
            td=st.date_input('Draft trade date',value=today);start=st.date_input('Draft delivery start',value=today);end=st.date_input('Draft delivery end',value=today)
            direction=st.selectbox('Draft Buy/Sell',['Buy','Sell']);qty=st.number_input('Draft quantity Dth/day',min_value=0.,value=0.)
            idx=st.text_input('Draft pricing index code');loc=st.text_input('Draft physical location',help='Use the Excel physical location, e.g. domspool, tcopool, deodelivered or coh22.')
            cls=st.selectbox('Draft pricing type',['FIXED','GDD','INDEX','PHYSB']);price=st.number_input('Draft price or premium ($/Dth)',value=0.,format='%.5f')
            allow_identical=st.checkbox('This is a separate physical deal with identical terms',value=False)
            submitted=st.form_submit_button('Save shadow draft')
        if submitted:
            try:
                terms=dict(source_sheet=sheet,trade_date=str(td),start=str(start),end=str(end),buy_sell=direction,quantity_dth_day=qty,counterparty=cp,index_code=idx,physical_location=loc,pricing_type=cls,price_or_premium=price)
                validated=shadow_terms(terms)
                if not allow_identical and any(r['state']=='Draft' and shadow_fingerprint(r['terms'])==shadow_fingerprint(validated) for r in data['records']):raise ValueError('An identical draft already exists. Review it, or explicitly mark this as a separate physical deal with identical terms.')
                save(shadow_apply(data,excel,digest,cutoff,'draft',terms=validated),'Shadow draft saved; Excel risk book unchanged')
            except ValueError as exc:st.error(str(exc))
    if any(r['state']!='Cancelled draft' for r in data['records']) and excel:
        with st.expander('Link a draft, confirm an ID, or accept an Excel amendment'):
            ids=[r['deal_id'] for r in data['records'] if r['state']!='Cancelled draft']
            selected_id=st.selectbox('Permanent deal ID',ids,key='shadow_link_id')
            selected_source=st.selectbox('Excel source row',[x['source_id'] for x in excel],key='shadow_link_source')
            old=next(r for r in data['records'] if r['deal_id']==selected_id);new=next(x for x in excel if x['source_id']==selected_source)
            st.dataframe(pd.DataFrame([dict(field=k,register=old['terms'].get(k),excel=new['terms'].get(k)) for k in SHADOW_FIELDS]).astype(str),hide_index=True)
            action=st.selectbox('Action',['Link matching terms / confirm ID','Accept Excel amendment'],key='shadow_link_action')
            if st.button('Save selected link or amendment',key='shadow_link_save'):
                try:save(shadow_apply(data,excel,digest,cutoff,'link' if action.startswith('Link') else 'amend',source_ids=[selected_source],deal_id=selected_id),'Shadow register link/amendment saved')
                except (ValueError,KeyError) as exc:st.error(str(exc))
    st.caption('Export parallel_deal_register.csv to retain the permanent-ID mapping for Excel. Excel can carry the ID in a Risk Deal ID column: use the existing header area above the first trade row, in a spare column within A:BL without moving existing columns. IDs must be unique across both blotters. Identical-term groups without IDs require explicit confirmation after a workbook change. Generic CSV trade import, cancellation workflows, authenticated editor attribution and official promotion remain subsequent work.')


def history_target(config):
    backend = str(config.get('backend', '')).strip().lower()
    path = str(config.get('database_path', '')).strip()
    url = str(config.get('database_url', '')).strip()
    if not backend:
        if url:
            raise ValueError('Set history.backend to postgresql to activate the configured database URL.')
        return path
    if backend == 'session':
        return ''
    if backend == 'sqlite':
        if not path:
            raise ValueError('SQLite history requires history.database_path on a persistent mounted disk.')
        return path
    if backend != 'postgresql':
        raise ValueError('History backend must be session, sqlite, or postgresql.')
    if not url.startswith(('postgresql://', 'postgres://')):
        raise ValueError('PostgreSQL history requires a PostgreSQL connection URL in history.database_url.')
    namespace = str(config.get('namespace', 'riverbend-risk')).strip()
    if not re.fullmatch(r'[A-Za-z0-9_-]{1,80}', namespace):
        raise ValueError('History namespace must contain 1–80 letters, digits, underscores, or hyphens.')
    return dict(backend=backend, database_url=url, namespace=namespace)


def postgres_history(target, operation, bundle=None, manifest=None, run_id=None, heartbeat=None):
    """Bound parameters, encrypted connections, and redacted backend errors."""
    try:
        import psycopg
        from psycopg.conninfo import conninfo_to_dict
    except ImportError:
        raise ValueError('Hosted history requires psycopg[binary]>=3.2,<4 in requirements.txt. Redeploy after adding it.') from None
    try:
        options = conninfo_to_dict(target['database_url'])
        if options.get('sslmode') not in ('require', 'verify-ca', 'verify-full'):
            options['sslmode'] = 'require'
        options.update(connect_timeout=10, options='-c statement_timeout=15000 -c lock_timeout=10000')
        with psycopg.connect(**options) as connection:
            connection.execute('CREATE TABLE IF NOT EXISTS riverbend_risk_runs ('
                'namespace TEXT NOT NULL, run_id TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, '
                'manifest TEXT NOT NULL, archive BYTEA NOT NULL, archive_hash TEXT NOT NULL, '
                'PRIMARY KEY (namespace, run_id))')
            namespace = target['namespace']
            if operation in {'heartbeat_write','heartbeat_read'}:
                connection.execute('CREATE TABLE IF NOT EXISTS riverbend_background_health ('
                    'namespace TEXT NOT NULL, trigger TEXT NOT NULL, completed_at TIMESTAMPTZ NOT NULL, '
                    'payload TEXT NOT NULL, PRIMARY KEY (namespace,trigger))')
                if operation == 'heartbeat_write':
                    connection.execute('INSERT INTO riverbend_background_health '
                        '(namespace,trigger,completed_at,payload) VALUES (%s,%s,%s,%s) '
                        'ON CONFLICT (namespace,trigger) DO UPDATE SET completed_at=EXCLUDED.completed_at,payload=EXCLUDED.payload '
                        'WHERE riverbend_background_health.completed_at <= EXCLUDED.completed_at',
                        (namespace,heartbeat['trigger'],heartbeat['completed_at_utc'],json.dumps(heartbeat,allow_nan=False)))
                    return None
                rows=connection.execute('SELECT payload FROM riverbend_background_health WHERE namespace=%s ORDER BY completed_at DESC',(namespace,)).fetchall()
                return [json.loads(row[0]) for row in rows]
            if operation == 'save':
                run_id = snapshot_id(manifest)
                connection.execute('INSERT INTO riverbend_risk_runs '
                    '(namespace,run_id,created_at,manifest,archive,archive_hash) VALUES (%s,%s,%s,%s,%s,%s) '
                    'ON CONFLICT (namespace,run_id) DO NOTHING',
                    (namespace,run_id,manifest['provenance']['created_at_utc'],
                     json.dumps(manifest,allow_nan=False),bundle,hashlib.sha256(bundle).hexdigest()))
                return run_id
            if operation == 'list':
                rows = connection.execute('SELECT run_id,manifest FROM riverbend_risk_runs '
                    'WHERE namespace=%s ORDER BY created_at DESC,run_id DESC LIMIT 100', (namespace,)).fetchall()
                return {key:json.loads(value) for key,value in rows}
            if operation == 'read':
                row = connection.execute('SELECT archive,archive_hash FROM riverbend_risk_runs '
                    'WHERE namespace=%s AND run_id=%s', (namespace,run_id)).fetchone()
                if row is None:
                    raise ValueError('Saved run no longer exists.')
                data = bytes(row[0])
                if hashlib.sha256(data).hexdigest() != row[1]:
                    raise ValueError('Saved archive checksum failed.')
                if snapshot_id(validate_run_archive(data)) != run_id:
                    raise ValueError('Saved archive identity check failed.')
                return data
            raise ValueError('Unsupported history operation.')
    except psycopg.Error:
        # Driver errors may contain a hostname, username or connection string.
        raise ValueError('Hosted history unavailable. Check the database URL, network access, TLS, and table read/write permissions. No successful save is reported; the current archive download remains available.') from None


def history_connection(path):
    target = Path(path).expanduser()
    if not target.is_absolute():
        raise ValueError('History database path must be absolute.')
    target.parent.mkdir(parents=True,exist_ok=True)
    connection = sqlite3.connect(str(target),timeout=15)
    try:
        connection.execute('CREATE TABLE IF NOT EXISTS risk_runs ('
                           'run_id TEXT PRIMARY KEY, created_at TEXT NOT NULL, '
                           'manifest TEXT NOT NULL, archive BLOB NOT NULL, archive_hash TEXT NOT NULL)')
        connection.commit()
    except Exception:
        connection.close()
        raise
    return connection


def save_history(path, bundle, manifest):
    checked = validate_run_archive(bundle)
    if checked != manifest:
        raise ValueError('Archive manifest does not match the supplied run metadata.')
    if isinstance(path,dict):
        return postgres_history(path,'save',bundle=bundle,manifest=manifest)
    run_id = snapshot_id(manifest)
    connection = history_connection(path)
    try:
        with connection:
            connection.execute('INSERT OR IGNORE INTO risk_runs VALUES (?,?,?,?,?)',
                               (run_id,manifest['provenance']['created_at_utc'],
                                json.dumps(manifest,allow_nan=False),bundle,hashlib.sha256(bundle).hexdigest()))
    finally:
        connection.close()
    return run_id


def list_history(path):
    if isinstance(path,dict):
        return postgres_history(path,'list')
    connection = history_connection(path)
    try:
        rows = connection.execute('SELECT run_id,manifest FROM risk_runs ORDER BY created_at DESC,run_id DESC LIMIT 100').fetchall()
        return {key:json.loads(manifest) for key,manifest in rows}
    finally:
        connection.close()


def read_history(path, run_id):
    if isinstance(path,dict):
        return postgres_history(path,'read',run_id=run_id)
    connection = history_connection(path)
    try:
        row = connection.execute('SELECT archive,archive_hash FROM risk_runs WHERE run_id=?',(run_id,)).fetchone()
        if row is None:
            raise ValueError('Saved run no longer exists.')
        if hashlib.sha256(row[0]).hexdigest() != row[1]:
            raise ValueError('Saved archive checksum failed.')
        if snapshot_id(validate_run_archive(row[0])) != run_id:
            raise ValueError('Saved archive identity check failed.')
        return row[0]
    finally:
        connection.close()


def record_background_run(path, run_id, status, trigger, completed_at=None):
    if not re.fullmatch(r'[a-f0-9]{64}',run_id) or status not in {'COMPLETE','REVIEW'}:
        raise ValueError('Invalid background run health record.')
    if trigger not in {'schedule','workflow_dispatch','local'}:
        raise ValueError('Unrecognized background run trigger.')
    stamp=pd.Timestamp(completed_at) if completed_at is not None else pd.Timestamp.now(tz='UTC')
    if pd.isna(stamp) or stamp.tzinfo is None:
        raise ValueError('Background completion time must include a timezone.')
    item=dict(run_id=run_id,status=status,trigger=trigger,completed_at_utc=stamp.tz_convert('UTC').isoformat())
    if isinstance(path,dict):
        return postgres_history(path,'heartbeat_write',heartbeat=item)
    if not path:
        raise ValueError('Background run health requires durable history storage.')
    connection=history_connection(path)
    try:
        with connection:
            connection.execute('CREATE TABLE IF NOT EXISTS background_health (trigger TEXT PRIMARY KEY, completed_at TEXT NOT NULL, payload TEXT NOT NULL)')
            connection.execute('INSERT INTO background_health VALUES (?,?,?) ON CONFLICT(trigger) DO UPDATE SET '
                'completed_at=excluded.completed_at,payload=excluded.payload WHERE background_health.completed_at <= excluded.completed_at',
                (trigger,item['completed_at_utc'],json.dumps(item,allow_nan=False)))
    finally:
        connection.close()


def read_background_runs(path):
    if isinstance(path,dict):
        return postgres_history(path,'heartbeat_read')
    if not path:
        return []
    connection=history_connection(path)
    try:
        with connection:
            connection.execute('CREATE TABLE IF NOT EXISTS background_health (trigger TEXT PRIMARY KEY, completed_at TEXT NOT NULL, payload TEXT NOT NULL)')
        return [json.loads(row[0]) for row in connection.execute('SELECT payload FROM background_health ORDER BY completed_at DESC').fetchall()]
    finally:
        connection.close()


def background_health_rows(items, now=None, stale_hours=36):
    now=pd.Timestamp(now) if now is not None else pd.Timestamp.now(tz='UTC')
    if now.tzinfo is None or pd.isna(now) or stale_hours<=0:
        raise ValueError('Invalid background health time or threshold.')
    rows=[]
    for item in items:
        stamp=pd.Timestamp(item['completed_at_utc'])
        if stamp.tzinfo is None or pd.isna(stamp) or item.get('status') not in {'COMPLETE','REVIEW'}:
            raise ValueError('Invalid saved background health record.')
        age=(now-stamp).total_seconds()/3600
        health='CLOCK REVIEW' if age<0 else 'STALE' if age>stale_hours else 'REVIEW' if item['status']=='REVIEW' else 'RECENT'
        rows.append(dict(trigger=item['trigger'],last_verified_run_central=stamp.tz_convert('America/Chicago').strftime('%Y-%m-%d %I:%M %p %Z'),
            hours_since_completion=round(age,1),health=health,run_result=item['status'],run_id=item['run_id']))
    return pd.DataFrame(rows,columns=['trigger','last_verified_run_central','hours_since_completion','health','run_result','run_id'])


def render_background_health(path):
    st.subheader('Background run health')
    if not path:
        st.info('Connect durable history storage to view background run health.')
        return
    try:
        rows=background_health_rows(read_background_runs(path))
    except (ValueError,OSError,sqlite3.Error):
        st.warning('Background run health could not be read. Check database access; snapshot downloads remain available.')
        return
    if rows.empty:
        st.info('No Stage 3.2 runner completion recorded yet. Start a new GitHub workflow run after updating both app files.')
    else:
        st.dataframe(rows,hide_index=True,use_container_width=True)
        scheduled=rows[rows.trigger.eq('schedule')]
        if scheduled.empty:
            st.info('No timer-triggered run recorded yet. A manual workflow run does not verify that the daily schedule is enabled.')
        elif scheduled.health.isin(['STALE','CLOCK REVIEW']).any():
            st.warning('The last timer-triggered completion is stale or has a clock discrepancy. Check GitHub Actions and RISK_SCHEDULE_ENABLED.')
        elif scheduled.run_result.eq('REVIEW').any():
            st.warning('The latest timer-triggered snapshot was saved with review flags. Inspect its archive.')
    st.caption('Updated only after the runner saves and reads back a verified archive, including unchanged snapshots. More than 36 hours since completion is stale. Failed attempts do not advance this timestamp; open the GitHub workflow run summary for its diagnostic step and next action. This panel does not read the current schedule setting. Manual app refreshes and archive restores do not update runner health.')


def archived_monthly_values(bundle):
    manifest=validate_run_archive(bundle)
    name='reports/forward_valuation_monthly.csv'
    if name not in {r['file'] for r in manifest['reports']}:
        raise ValueError('A selected run has no monthly forward valuation report.')
    with zipfile.ZipFile(BytesIO(bundle)) as archive:
        frame=pd.read_csv(BytesIO(archive.read(name)),keep_default_na=False)
    required={'month','index_code','complete_mtm','status','blocked_allocations','trade_months'}
    if not required.issubset(frame.columns) or frame.empty:
        raise ValueError('Monthly forward valuation report is empty or has an unsupported format.')
    dates=pd.to_datetime(frame.month,errors='coerce')
    if dates.isna().any() or dates.dt.tz is not None:
        raise ValueError('Monthly forward valuation report has invalid dates.')
    frame['month']=dates.dt.to_period('M').astype(str)
    frame['index_code']=frame.index_code.astype(str).str.strip()
    if frame.index_code.eq('').any() or frame.duplicated(['month','index_code']).any():
        raise ValueError('Monthly forward valuation report has missing or duplicate market keys.')
    for col in ['complete_mtm','blocked_allocations','trade_months']:
        frame[col]=pd.to_numeric(frame[col],errors='coerce')
    valid_amount=frame.complete_mtm.map(lambda v:pd.notna(v) and math.isfinite(float(v)))
    valid_counts=frame.trade_months.gt(0) & frame.trade_months.mod(1).eq(0)
    frame['usable']=frame.status.astype(str).str.startswith('INDICATIVE') & frame.blocked_allocations.eq(0) & valid_amount & valid_counts
    return frame,manifest


def comparison_delivery_scope(manifest):
    p=manifest.get('provenance',{})
    try:
        start=pd.Timestamp(p.get('delivery_start'))
        end=pd.Timestamp(p.get('delivery_end'))
        if pd.isna(start) or pd.isna(end) or start.tzinfo is not None or end.tzinfo is not None or end<start:
            return None
        return start.normalize(),end.normalize()
    except (ValueError,TypeError):
        return None


def compare_archived_mtm(baseline_bundle,current_bundle):
    before,before_manifest=archived_monthly_values(baseline_bundle)
    after,after_manifest=archived_monthly_values(current_bundle)
    baseline_scope=comparison_delivery_scope(before_manifest)
    current_scope=comparison_delivery_scope(after_manifest)
    scope_issue=('DELIVERY RANGE NOT RECORDED' if baseline_scope is None or current_scope is None
                 else 'DIFFERENT DELIVERY RANGES' if baseline_scope!=current_scope else '')
    joined=before.merge(after,on=['month','index_code'],how='outer',suffixes=('_baseline','_current'),indicator=True,validate='one_to_one')
    rows=[]
    for _,r in joined.iterrows():
        old_ok=r['_merge']!='right_only' and r['usable_baseline']==True
        new_ok=r['_merge']!='left_only' and r['usable_current']==True
        status=('ADDED GROUP — difference withheld' if r['_merge']=='right_only'
                else 'REMOVED GROUP — difference withheld' if r['_merge']=='left_only'
                else scope_issue+' — difference withheld' if scope_issue
                else 'INCOMPLETE — difference withheld' if not old_ok or not new_ok
                else 'COMPARABLE — reported MTM difference')
        comparable=status.startswith('COMPARABLE')
        rows.append(dict(month=r['month'],index_code=r.index_code,
            baseline_mtm=r.complete_mtm_baseline if old_ok else None,
            current_mtm=r.complete_mtm_current if new_ok else None,
            reported_mtm_difference=r.complete_mtm_current-r.complete_mtm_baseline if comparable else None,
            baseline_allocations=r.trade_months_baseline,current_allocations=r.trade_months_current,
            baseline_blocked_allocations=r.blocked_allocations_baseline,current_blocked_allocations=r.blocked_allocations_current,
            baseline_status=r.status_baseline,current_status=r.status_current,comparison_status=status))
    table=pd.DataFrame(rows).sort_values(['month','index_code'])
    complete=not table.empty and table.comparison_status.str.startswith('COMPARABLE').all()
    metadata=dict(baseline_delivery_scope='Not recorded' if baseline_scope is None else ' to '.join(str(v.date()) for v in baseline_scope),
        current_delivery_scope='Not recorded' if current_scope is None else ' to '.join(str(v.date()) for v in current_scope),
        baseline_price_date=before_manifest['provenance'].get('latest_ice_trade_date','Not recorded'),
        current_price_date=after_manifest['provenance'].get('latest_ice_trade_date','Not recorded'),
        full_difference=table.reported_mtm_difference.sum() if complete else None)
    return table,metadata


def archived_delivered_exposure(bundle):
    manifest=validate_run_archive(bundle)
    required_reports=['physical_open_exposure_monthly.csv','monthly_physical_balance.csv']
    inventory={r['file'] for r in manifest['reports']}
    if any('reports/'+n not in inventory for n in required_reports):
        raise ValueError('Both selected runs need delivered physical balance and exposure reports.')
    frames=[]
    with zipfile.ZipFile(BytesIO(bundle)) as archive:
        for name in required_reports:
            frame=pd.read_csv(BytesIO(archive.read('reports/'+name)),keep_default_na=False)
            required={'month','index_code','status'} | ({'net_open_purchase_dth','forward_usd_dth','net_open_cost_usd'} if name==required_reports[0] else {'balance_scope','net_surplus_shortfall'})
            if frame.empty or not required.issubset(frame.columns):raise ValueError('Selected archive lacks a supported delivered exposure format.')
            dates=pd.to_datetime(frame.month,errors='coerce')
            if dates.isna().any() or dates.dt.tz is not None:raise ValueError('Invalid delivery months in archived exposure.')
            frame['month']=dates.dt.to_period('M').astype(str)
            frame['index_code']=frame.index_code.map(market)
            if frame.index_code.eq('').any() or frame.duplicated(['month','index_code']).any():raise ValueError('Missing or duplicate book/month in archived exposure.')
            frames.append(frame)
    exposure,balance=frames
    joined=exposure.merge(balance[['month','index_code','balance_scope','net_surplus_shortfall','status']],on=['month','index_code'],how='outer',suffixes=('','_balance'),validate='one_to_one',indicator=True)
    for col in ['net_open_purchase_dth','forward_usd_dth','net_open_cost_usd','net_surplus_shortfall']:
        joined[col]=pd.to_numeric(joined[col],errors='coerce')
        joined[col]=joined[col].where(joined[col].map(lambda v:pd.notna(v) and math.isfinite(float(v))))
    valid_balance=joined.status_balance.fillna('').map(lambda v:v in ['BALANCED','SURPLUS','SHORTFALL'] or v.startswith('PROVISIONAL STORAGE — '))
    joined['volume_usable']=(joined['_merge'].eq('both') & joined.balance_scope.eq('DELIVERED') & valid_balance & joined.net_open_purchase_dth.notna() & joined.net_surplus_shortfall.notna() & (joined.net_open_purchase_dth+joined.net_surplus_shortfall).abs().le(1e-6))
    product=joined.net_open_purchase_dth*joined.forward_usd_dth
    joined['cost_usable']=(joined.volume_usable & joined.status.fillna('').str.startswith('INDICATIVE') & joined.forward_usd_dth.notna() & joined.net_open_cost_usd.notna() & (joined.net_open_cost_usd-product).abs().le(.01))
    return joined.drop(columns='_merge'),manifest


def compare_archived_delivered(baseline_bundle,current_bundle):
    before,bm=archived_delivered_exposure(baseline_bundle);after,cm=archived_delivered_exposure(current_bundle)
    bs=comparison_delivery_scope(bm);cs=comparison_delivery_scope(cm)
    scope_issue='DELIVERY RANGE NOT RECORDED' if bs is None or cs is None else 'DIFFERENT DELIVERY RANGES' if bs!=cs else ''
    joined=before.merge(after,on=['month','index_code'],how='outer',suffixes=('_baseline','_current'),indicator=True,validate='one_to_one')
    rows=[]
    for _,r in joined.iterrows():
        both=r['_merge']=='both'
        old_volume=r['_merge']!='right_only' and r.volume_usable_baseline==True
        new_volume=r['_merge']!='left_only' and r.volume_usable_current==True
        old_cost=r['_merge']!='right_only' and r.cost_usable_baseline==True
        new_cost=r['_merge']!='left_only' and r.cost_usable_current==True
        volume_ok=both and not scope_issue and old_volume and new_volume
        cost_ok=volume_ok and old_cost and new_cost
        reason='ADDED GROUP' if r['_merge']=='right_only' else 'REMOVED GROUP' if r['_merge']=='left_only' else scope_issue if scope_issue else 'INVALID OR NON-DELIVERED BALANCE' if not volume_ok else 'INCOMPLETE OR INCONSISTENT COST' if not cost_ok else ''
        v0=r.net_open_purchase_dth_baseline if old_volume else None;v1=r.net_open_purchase_dth_current if new_volume else None
        p0=r.forward_usd_dth_baseline if old_cost else None;p1=r.forward_usd_dth_current if new_cost else None
        changes=[]
        for col in ['price_method','delivered_price_source','physical_balance_status']:
            old=txt(r.get(col+'_baseline'));new=txt(r.get(col+'_current'))
            if old!=new:changes.append(col+': '+old+' -> '+new)
        if bm['app_version']!=cm['app_version']:changes.append('App version changed')
        rows.append(dict(month=r['month'],index_code=r.index_code,
            baseline_net_open_purchase_dth=v0,current_net_open_purchase_dth=v1,
            net_open_purchase_change_dth=v1-v0 if volume_ok else None,
            baseline_forecast_usd_dth=p0,current_forecast_usd_dth=p1,
            forecast_price_change_usd_dth=p1-p0 if cost_ok else None,
            baseline_net_procurement_cost_usd=r.net_open_cost_usd_baseline if old_cost else None,
            current_net_procurement_cost_usd=r.net_open_cost_usd_current if new_cost else None,
            reported_cost_difference_usd=r.net_open_cost_usd_current-r.net_open_cost_usd_baseline if cost_ok else None,
            price_input_effect_usd=v0*(p1-p0) if cost_ok else None,
            volume_input_effect_usd=(v1-v0)*p1 if cost_ok else None,
            volume_comparison_status='COMPARABLE — reported volumes' if volume_ok else (reason+' — difference withheld'),
            comparison_status='COMPARABLE — reported procurement forecast' if cost_ok else (reason+' — cost difference withheld'),
            input_changes='; '.join(changes),baseline_status=r.get('status_baseline'),current_status=r.get('status_current')))
    result=pd.DataFrame(rows).sort_values(['month','index_code'])
    cost_complete=not result.empty and result.comparison_status.str.startswith('COMPARABLE').all()
    volume_complete=not result.empty and result.volume_comparison_status.str.startswith('COMPARABLE').all()
    metadata=dict(full_cost_difference=result.reported_cost_difference_usd.sum() if cost_complete else None,
        full_volume_difference=result.net_open_purchase_change_dth.sum() if volume_complete else None,
        price_input_effect=result.price_input_effect_usd.sum() if cost_complete else None,
        volume_input_effect=result.volume_input_effect_usd.sum() if cost_complete else None,
        baseline_price_date=bm['provenance'].get('latest_ice_trade_date','Not recorded'),current_price_date=cm['provenance'].get('latest_ice_trade_date','Not recorded'))
    return result,metadata


def saved_valuation_history(history, read_archive, reference_manifest, limit=5):
    if type(limit) is not int or limit<1 or limit>20:
        raise ValueError('Load between 1 and 20 saved snapshots.')
    scope=comparison_delivery_scope(reference_manifest)
    selected=sorted(history.items(),key=lambda pair:pd.Timestamp(pair[1]['provenance']['created_at_utc']),reverse=True)[:limit]
    rows=[]
    for run_id,manifest in selected:
        p=manifest['provenance']
        row=dict(run_id=run_id,snapshot_created_utc=p['created_at_utc'],
            ice_trade_date=p.get('latest_ice_trade_date'),run_profile=p.get('run_profile','interactive app'),
            app_version=manifest.get('app_version',''),delivery_start=p.get('delivery_start'),delivery_end=p.get('delivery_end'),
            reported_mtm=None,allocations=None,blocked_allocations=None,status='UNAVAILABLE')
        candidate=comparison_delivery_scope(manifest)
        if scope is None or candidate is None:
            row['status']='DELIVERY RANGE NOT RECORDED'
        elif candidate!=scope:
            row['status']='DIFFERENT DELIVERY RANGE'
        else:
            try:
                frame,checked=archived_monthly_values(read_archive(run_id))
                if snapshot_id(checked)!=run_id or comparison_delivery_scope(checked)!=scope:
                    raise ValueError('Saved snapshot identity or delivery scope mismatch.')
                row['allocations']=frame.trade_months.sum(min_count=1)
                row['blocked_allocations']=frame.blocked_allocations.sum(min_count=1)
                if frame.usable.all():
                    total=frame.complete_mtm.sum()
                    if math.isfinite(float(total)):
                        row['reported_mtm']=total
                        row['status']='COMPLETE REPORTED TOTAL'
                    else:
                        row['status']='INCOMPLETE — non-finite total withheld'
                else:
                    row['status']='INCOMPLETE — total withheld'
            except (ValueError,OSError,sqlite3.Error,KeyError):
                row['status']='UNAVAILABLE — archive or report could not be verified'
        rows.append(row)
    result=pd.DataFrame(rows,columns=['run_id','snapshot_created_utc','ice_trade_date','run_profile','app_version',
        'delivery_start','delivery_end','reported_mtm','allocations','blocked_allocations','status'])
    if not result.empty:
        result['snapshot_created_utc']=pd.to_datetime(result.snapshot_created_utc,utc=True)
        result=result.sort_values(['snapshot_created_utc','run_id']).reset_index(drop=True)
    return result


def render_valuation_history(path, history, current_manifest):
    st.subheader('Saved valuation history')
    st.caption('Load recent saved snapshots for the current delivery range. Each point is a complete reported commodity MTM total, not business P&L. Snapshot dates are UTC; the table separately records ICE pricing dates. Multiple snapshots can use the same pricing date.')
    count=st.selectbox('Recent snapshots to inspect',[5,10,20],key='valuation_history_count')
    # Results are intentionally outside the report collector. Viewing history must
    # never alter the snapshot content or trigger recursive snapshot creation.
    identity=hashlib.sha256(json.dumps(dict(target=path,ids=sorted(history),
        scope=str(comparison_delivery_scope(current_manifest)),count=count),sort_keys=True).encode()).hexdigest()
    if st.button('Load valuation history',key='load_valuation_history'):
        def reader(key):
            return read_history(path,key) if path else st.session_state['_snapshot_history'][key][1]
        with st.spinner('Reading and verifying saved valuations...'):
            try:
                table=saved_valuation_history(history,reader,current_manifest,count)
                st.session_state['_valuation_history_view']=(identity,table)
            except (ValueError,OSError,sqlite3.Error,KeyError):
                st.error('Valuation history could not be loaded. Check saved archive access and try again.')
                st.session_state.pop('_valuation_history_view',None)
    cached=st.session_state.get('_valuation_history_view')
    if cached is None or cached[0]!=identity:
        st.info('Select Load valuation history to inspect the most recent saved runs. Reload when the run list or delivery range changes.')
        return
    table=cached[1]
    points=table.loc[table.status.eq('COMPLETE REPORTED TOTAL')].copy()
    if points.empty:
        st.info('No complete, matching-range totals are available in this selection. See the status column for each run.')
    else:
        points['reported_mtm']=pd.to_numeric(points.reported_mtm)
        st.scatter_chart(points,x='snapshot_created_utc',y='reported_mtm')
    st.dataframe(table,hide_index=True,use_container_width=True)
    st.download_button('Download valuation history',table.to_csv(index=False).encode(),'saved_valuation_history.csv','text/csv',key='saved_valuation_history_download')
    st.caption('Incomplete, unavailable, and different-range runs stay in the table with blank MTM and are omitted from the chart. No missing run is treated as zero. Identical reruns may reuse a snapshot; use Background run health for the latest verified execution time. Changes can reflect prices, trades, terms, volumes, mappings, or app versions.')


def archived_book_groups(bundle):
    manifest=validate_run_archive(bundle)
    name='reports/complete_trade_ledger.csv'
    if name not in {r['file'] for r in manifest['reports']}:
        raise ValueError('A selected archive has no complete trade ledger.')
    with zipfile.ZipFile(BytesIO(bundle)) as archive:
        ledger=pd.read_csv(BytesIO(archive.read(name)),dtype=str,keep_default_na=False)
    terms=['source_sheet','trade_date','counterparty','index_code','location','start','end',
           'premium','fixed_price','contract_price','secondary_label','classification_code','pricing_label']
    if not set(terms+['daily_volume','source_id']).issubset(ledger.columns):
        raise ValueError('Archived trade ledger has an unsupported format.')
    if ledger.source_id.eq('').any() or ledger.source_id.duplicated().any():
        raise ValueError('Archived ledger has missing or duplicate source IDs.')
    ledger=ledger.copy()
    for col in terms:
        ledger[col]=ledger[col].str.strip()
    for col in ['premium','fixed_price','contract_price']:
        ledger[col]=ledger[col].map(lambda v:format(num(v),'.17g') if num(v) is not None else v)
    for col in ['trade_date','start','end']:
        # Normalize equivalent exported dates, retaining invalid source text.
        parsed=pd.to_datetime(ledger[col],errors='coerce',format='mixed')
        ledger[col]=parsed.dt.strftime('%Y-%m-%d').where(parsed.notna(),ledger[col])
    ledger['_volume']=ledger.daily_volume.map(num)
    groups={}
    for key,frame in ledger.groupby(terms,dropna=False,sort=False):
        key=tuple(key)
        complete=frame['_volume'].notna().all()
        try:
            total=math.fsum(frame['_volume']) if complete else None
        except OverflowError:
            total=None
        if total is not None and not math.isfinite(float(total)):
            total=None
        groups[key]=dict(terms=dict(zip(terms,key)),record_count=len(frame),
                         signed_daily_volume=total,source_ids=', '.join(sorted(frame.source_id)),
                         volume_entries=', '.join(sorted(format(num(v),'.17g') if num(v) is not None else 'INVALID:'+v for v in frame.daily_volume)))
    return groups


def compare_archived_books(baseline_bundle,current_bundle):
    before=archived_book_groups(baseline_bundle)
    after=archived_book_groups(current_bundle)
    rows=[]
    for key in sorted(set(before)|set(after)):
        old,new=before.get(key),after.get(key)
        old_volume=old['signed_daily_volume'] if old else None
        new_volume=new['signed_daily_volume'] if new else None
        complete=old_volume is not None and new_volume is not None
        if old is None:
            status='ADDED TERMS GROUP'
        elif new is None:
            status='REMOVED TERMS GROUP'
        elif not complete:
            status='REVIEW — missing or invalid volume'
        elif old_volume!=new_volume or old['record_count']!=new['record_count'] or old['volume_entries']!=new['volume_entries']:
            status='VOLUME OR RECORD COUNT CHANGED'
        elif old['source_ids']!=new['source_ids']:
            status='ROW REFERENCES ONLY'
        else:
            status='UNCHANGED'
        rows.append(dict(**(new or old)['terms'],change=status,
            baseline_records=old['record_count'] if old else None,
            current_records=new['record_count'] if new else None,
            baseline_signed_daily_volume=old_volume,current_signed_daily_volume=new_volume,
            signed_daily_volume_change=new_volume-old_volume if complete else None,
            baseline_volume_entries=old['volume_entries'] if old else '',current_volume_entries=new['volume_entries'] if new else '',
            baseline_source_ids=old['source_ids'] if old else '',current_source_ids=new['source_ids'] if new else ''))
    columns=['source_sheet','trade_date','counterparty','index_code','location','start','end','premium','fixed_price',
             'contract_price','secondary_label','classification_code','pricing_label','change','baseline_records',
             'current_records','baseline_signed_daily_volume','current_signed_daily_volume','signed_daily_volume_change',
             'baseline_volume_entries','current_volume_entries','baseline_source_ids','current_source_ids']
    return pd.DataFrame(rows,columns=columns)


def compare_run_manifests(baseline, current):
    """Compare report bytes and explicit provenance; never infer P&L attribution."""
    before={r['file']:r for r in baseline['reports']}
    after={r['file']:r for r in current['reports']}
    report_rows=[]
    for name in sorted(set(before)|set(after)):
        old,new=before.get(name),after.get(name)
        status=('ADDED' if old is None else 'REMOVED' if new is None else
                'UNCHANGED' if old['sha256']==new['sha256'] else 'CHANGED')
        report_rows.append(dict(report=name,change=status,
                                baseline_rows=None if old is None else old['rows'],
                                current_rows=None if new is None else new['rows'],
                                row_change=None if old is None or new is None else new['rows']-old['rows'],
                                baseline_status_counts=json.dumps(old.get('status_counts',{}),sort_keys=True) if old else '',
                                current_status_counts=json.dumps(new.get('status_counts',{}),sort_keys=True) if new else ''))
    def flatten(value,prefix=''):
        result={}
        for key,item in value.items():
            name=prefix+key
            if isinstance(item,dict):result.update(flatten(item,name+'.'))
            else:result[name]=json.dumps(item,sort_keys=True,ensure_ascii=False)
        return result
    old_source=flatten(baseline['provenance']);new_source=flatten(current['provenance'])
    old_source.pop('created_at_utc',None);new_source.pop('created_at_utc',None)
    old_source['app_version']=json.dumps(baseline['app_version'])
    new_source['app_version']=json.dumps(current['app_version'])
    source_rows=[dict(field=k,baseline=old_source.get(k,'(not recorded)'),
                      current=new_source.get(k,'(not recorded)'),
                      change='UNCHANGED' if old_source.get(k)==new_source.get(k) and (k in old_source)==(k in new_source) else 'CHANGED')
                 for k in sorted(set(old_source)|set(new_source))]
    return (pd.DataFrame(report_rows,columns=['report','change','baseline_rows','current_rows','row_change','baseline_status_counts','current_status_counts']),
            pd.DataFrame(source_rows,columns=['field','baseline','current','change']))


def validate_run_archive(content):
    """Validate a downloaded bundle in memory; never extract files or execute code."""
    if len(content)>50*1024*1024:
        raise ValueError('Archive exceeds the 50 MB upload limit.')
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            items=archive.infolist()
            names=[item.filename for item in items]
            if len(items)>250 or len(names)!=len(set(names)):
                raise ValueError('Archive has too many files or duplicate filenames.')
            if sum(item.file_size for item in items)>100*1024*1024:
                raise ValueError('Expanded archive exceeds 100 MB.')
            if 'manifest.json' not in names or archive.getinfo('manifest.json').file_size>1024*1024:
                raise ValueError('Archive manifest is missing or too large.')
            manifest=json.loads(archive.read('manifest.json'))
            if not isinstance(manifest,dict) or manifest.get('schema_version')!=1:
                raise ValueError('Unsupported run archive schema.')
            if not isinstance(manifest.get('app_version'),str) or not isinstance(manifest.get('provenance'),dict):
                raise ValueError('Archive version or provenance is missing.')
            created=pd.Timestamp(manifest['provenance'].get('created_at_utc'))
            if pd.isna(created) or created.tzinfo is None:
                raise ValueError('Archive creation time must include a timezone.')
            reports=manifest.get('reports')
            missing=manifest.get('missing_core_reports')
            if not isinstance(reports,list) or not isinstance(missing,list) or not all(isinstance(x,str) for x in missing):
                raise ValueError('Invalid report inventory.')
            expected={'manifest.json','README.txt'}
            for report in reports:
                if not isinstance(report,dict):raise ValueError('Invalid report entry.')
                name=report.get('file','')
                if not isinstance(name,str) or not re.fullmatch(r'reports/[A-Za-z0-9_-]+\.csv',name) or name in expected:
                    raise ValueError('Invalid or duplicate report path.')
                expected.add(name)
                if name not in names:raise ValueError('A listed report is missing from the archive.')
                if type(report.get('rows')) is not int or report['rows']<0:
                    raise ValueError('Invalid report row count.')
                if not isinstance(report.get('status_counts',{}),dict):
                    raise ValueError('Invalid report status summary.')
                if hashlib.sha256(archive.read(name)).hexdigest()!=report.get('sha256'):
                    raise ValueError('Report checksum mismatch: '+name)
            if set(names)-expected:
                raise ValueError('Archive contains unrecognized files.')
            # Also reject JSON non-finite values before persistence or comparison.
            json.dumps(manifest,allow_nan=False)
            return manifest
    except (zipfile.BadZipFile,KeyError,TypeError,UnicodeError,RuntimeError,NotImplementedError,RecursionError) as exc:
        raise ValueError('Archive is damaged or is not a supported Riverbend run archive.') from None


def render_run_history(bundle, manifest):
    st.subheader('Automatic run history')
    try:
        config = dict(st.secrets.get('history',{}))
    except (FileNotFoundError,KeyError):
        config = {}
    with st.expander('Deployment storage configuration'):
        st.write('Hosted PostgreSQL keeps run archives outside the app filesystem. Configure the connection in Streamlit Secrets and add the dependency to requirements.txt. Existing GitHub and workbook secrets stay unchanged.')
        st.code('psycopg[binary]>=3.2,<4',language='text')
        st.code('[history]\nbackend = "postgresql"\ndatabase_url = "postgresql://USER:PASSWORD@HOST:5432/DATABASE?sslmode=require"\nnamespace = "riverbend-risk"',language='toml')
        st.caption('Use the connection URL supplied by your database provider; paste it only into Secrets. The database user needs table creation, select, and insert permissions. Namespaces organize runs; they are not access controls. All users of this app can view its configured history. No automatic deletion is performed; the newest 100 runs are listed.')
        st.caption('Alternatively, keep history.database_path for SQLite on a persistent mounted disk. An ordinary temporary app disk is not durable storage. Existing archives can be restored into the configured store below.')
    try:
        path = history_target(config)
    except ValueError as exc:
        st.error(str(exc))
        return
    render_background_health(path)
    if isinstance(path,dict):
        st.caption('History backend: hosted PostgreSQL | Namespace: '+path['namespace'])
        if st.button('Test history connection',key='test_history_connection'):
            try:
                count = len(list_history(path))
                st.success(f'Connection and table access verified; {count} recent runs available. The automatic save below verifies insert access.')
            except (ValueError,OSError) as exc:
                st.error(str(exc) if isinstance(exc,ValueError) else 'History connection unavailable.')
    run_id = snapshot_id(manifest)
    with st.expander('Restore a downloaded run archive'):
        st.caption('Restore a prior Riverbend ZIP to history for comparison. Current workbook, pricing and manual entries are not replaced. Checksums detect corruption; they do not authenticate who created an archive.')
        uploaded=st.file_uploader('Prior run archive (.zip)',type=['zip'],key='restore_run_archive')
        if st.button('Validate and restore run',key='restore_run_button'):
            if uploaded is None:
                st.warning('Select a downloaded run archive first.')
            else:
                try:
                    restored_bytes=uploaded.getvalue()
                    restored=validate_run_archive(restored_bytes)
                    restored_id=snapshot_id(restored)
                    if path:
                        save_history(path,restored_bytes,restored)
                    else:
                        session=st.session_state.setdefault('_snapshot_history',{})
                        # Retain current and restored runs within the session budget.
                        if restored_id!=run_id and len(restored_bytes)+len(bundle)>50*1024*1024:
                            raise ValueError('Current and restored runs together exceed the 50 MB session budget. Configure durable history storage for larger comparisons.')
                        session.setdefault(restored_id,(restored,restored_bytes))
                    st.success('Archive validated and restored to history: '+restored_id[:12])
                except (ValueError,OSError,sqlite3.Error) as exc:
                    st.error('Run not restored: '+(str(exc) if isinstance(exc,ValueError) else 'History storage is unavailable. Check its location and permissions.'))
    if path:
        try:
            save_history(path,bundle,manifest)
            history = list_history(path)
        except (OSError,sqlite3.Error,ValueError) as exc:
            st.error(str(exc) if isinstance(exc,ValueError) else 'Run history could not be saved or read. Check the configured storage location and permissions. The current archive download remains available.')
            return
        st.success('Current run saved to configured storage. Identical reruns reuse the same snapshot.')
        st.caption('Hosted database history survives app restarts, subject to provider retention and availability.' if isinstance(path,dict) else 'Retention across app restarts depends on this location being a durable mounted disk. All users of this app share this history store; use only within the same authorized team.')
    else:
        session = st.session_state.setdefault('_snapshot_history',{})
        session.setdefault(run_id,(manifest,bundle))
        while len(session)>10 or (sum(len(v[1]) for v in session.values())>50*1024*1024 and len(session)>1):
            oldest=next(key for key in session if key!=run_id)
            del session[oldest]
        history = {key:value[0] for key,value in reversed(list(session.items()))}
        st.info('Session history active: up to 10 recent snapshots, limited to approximately 50 MB. Download archives to retain them after this session ends. A single larger archive is retained on its own.')
    rows=[]
    for key,item in history.items():
        p=item['provenance']
        rows.append(dict(run_id=key[:12],created_at_utc=p['created_at_utc'],
                         run_profile=p.get('run_profile','interactive app'),
                         workbook=p.get('workbook_name',''),ice_trade_date=p.get('latest_ice_trade_date'),
                         reports=len(item['reports']),missing_core_reports=len(item['missing_core_reports'])))
    st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
    selected=st.selectbox('Saved run',list(history),
                          format_func=lambda key:history[key]['provenance']['created_at_utc']+' | '+key[:12],key='saved_run')
    if selected:
        try:
            saved=read_history(path,selected) if path else st.session_state['_snapshot_history'][selected][1]
            st.download_button('Download saved run',saved,'riverbend_run_'+selected[:12]+'.zip',
                               'application/zip',key='saved_run_download')
        except (OSError,sqlite3.Error,ValueError):
            st.error('Saved archive could not be retrieved or failed its integrity check.')
    st.caption('The app captures snapshots when it runs with a loaded workbook. The separate GitHub workflow can capture background snapshots when configured and enabled. Missing or blocked reports remain explicitly recorded; saving does not approve results.')
    render_valuation_history(path,history,manifest)
    st.subheader('Compare saved run with current run')
    previous=[key for key in history if key!=run_id]
    if not previous:
        st.info('Comparison becomes available after a second distinct run is captured. Identical reruns do not create another snapshot.')
    else:
        baseline_id=st.selectbox('Baseline run',previous,
                                format_func=lambda key:history[key]['provenance']['created_at_utc']+' | '+key[:12],key='comparison_baseline')
        baseline=history[baseline_id]
        changes,sources=compare_run_manifests(baseline,manifest)
        a,b,c=st.columns(3)
        a.metric('Changed reports',int(changes.change.eq('CHANGED').sum()))
        b.metric('Added reports',int(changes.change.eq('ADDED').sum()))
        c.metric('Removed reports',int(changes.change.eq('REMOVED').sum()))
        st.caption('Baseline: '+baseline['provenance']['created_at_utc']+' | Current: '+manifest['provenance']['created_at_utc'])
        st.dataframe(sources,hide_index=True,use_container_width=True)
        st.dataframe(changes,hide_index=True,use_container_width=True)
        # Comparison views are not added to the snapshot report collector: choosing
        # a baseline must not change the underlying run or create another snapshot.
        for frame,label,filename in [(sources,'Download source changes','run_source_comparison.csv'),
                                      (changes,'Download report changes','run_report_comparison.csv')]:
            export=frame.copy()
            export.insert(0,'current_run_id',run_id)
            export.insert(0,'baseline_run_id',baseline_id)
            st.download_button(label,export.to_csv(index=False).encode(),filename,'text/csv',key=filename)
        st.caption('CHANGED means the exported report bytes differ. Changes may come from prices, trades, selected date ranges, filters, inputs, or app logic. A removed report was not generated in this run; it does not imply closed trades or zero exposure. This comparison does not attribute P&L or match trade identities across reordered workbooks.')
        st.subheader('Reported MTM change by month and market')
        try:
            baseline_bundle=read_history(path,baseline_id) if path else st.session_state['_snapshot_history'][baseline_id][1]
            mtm_changes,context=compare_archived_mtm(baseline_bundle,bundle)
            st.metric('Full reported MTM difference','Withheld — incomplete or different scope' if context['full_difference'] is None else f"${context['full_difference']:,.2f}")
            st.write('Baseline delivery range: '+context['baseline_delivery_scope']+' | Current: '+context['current_delivery_scope'])
            st.caption('ICE dates — baseline: '+str(context['baseline_price_date'])+'; current: '+str(context['current_price_date']))
            st.dataframe(mtm_changes,hide_index=True,use_container_width=True)
            export=mtm_changes.copy()
            export.insert(0,'current_run_id',run_id);export.insert(0,'baseline_run_id',baseline_id)
            st.download_button('Download MTM comparison',export.to_csv(index=False).encode(),'run_mtm_comparison.csv','text/csv',key='run_mtm_comparison')
            st.caption('Current reported MTM minus baseline reported MTM. Differences can include market prices, trades, volumes, terms, mappings, or code changes; this is not market-only P&L attribution. Missing groups are not treated as zero. Both runs must record the same exact delivery range; older interactive archives without that metadata show values but withhold differences. Comparison selections do not create new snapshots.')
        except (ValueError,OSError,sqlite3.Error):
            st.info('MTM comparison unavailable: a selected archive could not be retrieved or lacks a valid monthly valuation report. Source and report comparisons above remain available.')
        st.subheader('Delivered procurement changes by month and utility')
        try:
            old_delivery_bundle=read_history(path,baseline_id) if path else st.session_state['_snapshot_history'][baseline_id][1]
            delivery_changes,delivery_context=compare_archived_delivered(old_delivery_bundle,bundle)
            a,b=st.columns(2)
            a.metric('Net open purchase volume change','Withheld — incomplete or different scope' if delivery_context['full_volume_difference'] is None else f"{delivery_context['full_volume_difference']:,.0f} Dth")
            b.metric('Reported procurement cost change','Withheld — incomplete or different scope' if delivery_context['full_cost_difference'] is None else f"${delivery_context['full_cost_difference']:,.2f}")
            st.caption('DOM = DEO/Dominion delivered; TCO = Columbia delivered. Positive volume change means more net purchases required; positive cost change means higher forecast procurement cost. Surplus volumes and their values remain signed.')
            st.dataframe(delivery_changes,hide_index=True,use_container_width=True)
            export=delivery_changes.copy();export.insert(0,'current_run_id',run_id);export.insert(0,'baseline_run_id',baseline_id)
            st.download_button('Download delivered procurement comparison',export.to_csv(index=False).encode(),'run_delivered_procurement_comparison.csv','text/csv',key='run_delivered_procurement_comparison')
            st.caption('Arithmetic breakdown: price/input effect = baseline net volume × change in forecast price; volume/input effect = change in net volume × current forecast price. These explain the volume-times-price movement, subject to reported rounding. Price changes may reflect sources, overrides or methods, not only markets. Figures are procurement forecasts, not realized P&L, and are not added to hedge MTM. Provisional storage remains provisional. Missing groups are never zero-filled; full totals are withheld for incomplete or different delivery scope. Comparisons do not create snapshots.')
        except (ValueError,OSError,sqlite3.Error,pd.errors.EmptyDataError):
            st.info('Delivered procurement comparison unavailable: both runs need valid delivered balance and exposure reports. Older archives remain usable for the other comparisons.')
        st.subheader('Book changes between saved runs')
        st.caption('Compares the complete archived blotters, including historical trades. Groups use worksheet, trade date, counterparty, market, delivery location, delivery dates, and recorded pricing terms. Worksheet row numbers are references, not permanent trade IDs.')
        if st.button('Load book changes',key='load_book_changes'):
            try:
                with st.spinner('Comparing archived trade terms and volumes...'):
                    old_bundle=read_history(path,baseline_id) if path else st.session_state['_snapshot_history'][baseline_id][1]
                    book_changes=compare_archived_books(old_bundle,bundle)
                    st.session_state['_book_comparison_view']=(baseline_id,run_id,book_changes)
            except (ValueError,OSError,sqlite3.Error,KeyError):
                st.session_state.pop('_book_comparison_view',None)
                st.error('Book comparison unavailable. Both selected archives must contain valid complete trade ledgers.')
        cached=st.session_state.get('_book_comparison_view')
        if cached is not None and cached[:2]==(baseline_id,run_id):
            book_changes=cached[2]
            a,b,c=st.columns(3)
            a.metric('Added terms groups',int(book_changes.change.eq('ADDED TERMS GROUP').sum()))
            b.metric('Removed terms groups',int(book_changes.change.eq('REMOVED TERMS GROUP').sum()))
            c.metric('Volume / count changes',int(book_changes.change.eq('VOLUME OR RECORD COUNT CHANGED').sum()))
            include_unchanged=st.checkbox('Include unchanged groups and row-reference-only changes',key='book_comparison_unchanged')
            view=book_changes if include_unchanged else book_changes.loc[~book_changes.change.isin(['UNCHANGED','ROW REFERENCES ONLY'])]
            if view.empty:
                st.info('No changed terms groups, volumes, record counts, or invalid-volume review items in this selection.')
            else:
                st.dataframe(view,hide_index=True,use_container_width=True)
            export=book_changes.copy()
            export.insert(0,'current_run_id',run_id);export.insert(0,'baseline_run_id',baseline_id)
            st.download_button('Download complete book comparison',export.to_csv(index=False).encode(),'run_book_comparison.csv','text/csv',key='run_book_comparison_download')
            st.caption('An added or removed terms group can reflect an amendment or correction, not necessarily a new or closed trade. Volume changes are signed daily volumes, not total delivery quantities or P&L. Groups with missing volumes withhold volume differences. Saved worksheet prices and invoice amounts are excluded from the terms match. No records are removed or approved by this view.')
        else:
            st.info('Select Load book changes to compare the selected baseline with the current run.')


def load_github_workbook(config, fallback_token=''):
    repository = txt(config.get('repository'))
    path = txt(config.get('path'))
    branch = txt(config.get('branch'))
    token = txt(config.get('token')) or fallback_token
    if not repository or not path or not token:
        raise ValueError('Configure workbook repository and path, plus a read-access token in workbook.token or github.token.')
    if path.startswith('/') or any(part in ('','..','.') for part in path.split('/')) or '\\' in path:
        raise ValueError('Workbook path must be a repository-relative file path.')
    if not path.lower().endswith(('.xlsx','.xlsm')):
        raise ValueError('Workbook path must end in .xlsx or .xlsm.')
    inventory = github_inventory(repository,branch,token)
    if path not in inventory['paths']:
        raise ValueError('Configured workbook was not found in the selected repository revision.')
    content = github_bytes(repository,inventory['commit'],path,token)
    if not content:
        raise ValueError('Configured workbook is empty.')
    book = BytesIO(content)
    book.name = path.rsplit('/',1)[-1]
    source = dict(kind='github',repository=repository,path=path,
                  branch=inventory['branch'],commit=inventory['commit'])
    return book,source


def workbook_setup_status(config, fallback_token):
    repo=txt(config.get('repository'))
    path=txt(config.get('path'))
    valid_path=bool(path and not path.startswith('/') and '\\' not in path
                    and all(p not in ('','..','.') for p in path.split('/'))
                    and path.lower().endswith(('.xlsx','.xlsm')))
    return pd.DataFrame([
        dict(setting='Workbook repository',status='CONFIGURED' if re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+',repo) else 'NEEDED',
             detail=repo or 'Enter owner/repository for the private trade workbook.'),
        dict(setting='Workbook path',status='CONFIGURED' if valid_path else 'NEEDED',
             detail=path or 'Enter the repository-relative .xlsm or .xlsx path.'),
        dict(setting='Read-access token',status='PRESENT — ACCESS NOT YET VERIFIED' if txt(config.get('token')) or fallback_token else 'NEEDED',
             detail='Token values are read from app Secrets and are not displayed.'),
    ])


def workbook_connection_error(exc, config, fallback_token):
    message = str(exc) if isinstance(exc,ValueError) else 'GitHub returned an unexpected response. Check the repository and branch, then retry.'
    for token in (txt(config.get('token')),fallback_token):
        if token:
            message=message.replace(token,'[redacted]')
    return 'Workbook connection failed: '+message


def workbook_controls():
    try:
        config = dict(st.secrets.get('workbook',{}))
        fallback_token = txt(st.secrets.get('github',{}).get('token'))
    except (FileNotFoundError,KeyError):
        config = {}; fallback_token = ''
    # Session settings never contain credentials and never replace Secrets tokens.
    config.update(st.session_state.get('_workbook_connection',{}))
    tested_book = None
    connection_failed = False
    with st.expander('Workbook connection settings',expanded=True):
        st.caption('Connect the workbook once to avoid repeated uploads. The file must already exist in a private GitHub repository accessible to your token.')
        repo=st.text_input('Workbook repository (owner/repository)',value=txt(config.get('repository')),key='setup_workbook_repo').strip()
        branch=st.text_input('Workbook branch (blank = default)',value=txt(config.get('branch')),key='setup_workbook_branch').strip()
        path=st.text_input('Workbook file path',value=txt(config.get('path')) or 'data/poswksht.xlsm',key='setup_workbook_path').strip()
        proposed=dict(repository=repo,branch=branch,path=path)
        readiness=workbook_setup_status({**config,**proposed},fallback_token)
        st.dataframe(readiness,hide_index=True,use_container_width=True)
        if st.button('Use this workbook connection',key='apply_workbook_connection'):
            if readiness.status.eq('NEEDED').any():
                missing=', '.join(readiness.loc[readiness.status.eq('NEEDED'),'setting'])
                st.session_state['_workbook_test_result']=('error','Complete these settings first: '+missing)
                connection_failed=True
            else:
                candidate={**config,**proposed}
                try:
                    github_inventory.clear(); github_bytes.clear()
                    with st.spinner('Testing repository access and downloading workbook...'):
                        tested_book=load_github_workbook(candidate,fallback_token)
                    config=candidate
                    st.session_state['_workbook_connection']=proposed
                    st.session_state['workbook_source_mode']='GitHub — connected'
                    st.session_state['_workbook_test_result']=('success','Connection test passed: workbook downloaded. Workbook contents are checked next.')
                except (ValueError,KeyError,TypeError) as exc:
                    connection_failed=True
                    st.session_state['_workbook_test_result']=('error',workbook_connection_error(exc,candidate,fallback_token))
        result=st.session_state.get('_workbook_test_result')
        if result:
            st.caption('Last explicit connection test')
            if result[0]=='success': st.success(result[1])
            else: st.error(result[1])
        st.markdown('**Retain this connection after the session ends**')
        st.caption('Add or update this section in Streamlit Settings → Secrets. Keep your existing [github] pricing settings. Do not add a second [workbook] section if one already exists.')
        st.code('[workbook]\n'+'\n'.join(k+' = '+json.dumps(v) for k,v in proposed.items()),language='toml')
        st.caption('Uses github.token unless workbook.token is configured separately. The token must have read access to the workbook repository. Local Excel changes appear here only after the updated file reaches GitHub.')
    configured = bool(config.get('repository') and config.get('path'))
    source = st.radio('Trade workbook source',['GitHub — connected','Upload — validation'],index=0 if configured else 1,key='workbook_source_mode')
    if connection_failed:
        st.warning('The requested connection was not applied. Correct the settings above and click Use this workbook connection again.')
        return None,dict(kind='github')
    if source == 'Upload — validation':
        st.caption('Upload mode is selected. Automatic workbook loading is not active.')
        book = st.file_uploader('Trade workbook including cash, cash2, and wapp',type=['xlsm','xlsx'])
        return book,dict(kind='upload')
    if not configured:
        st.info('Configure the private workbook repository and path in Secrets to enable connected loading.')
        with st.expander('Workbook connection setup'):
            st.code('[workbook]\nrepository = "OWNER/PRIVATE_REPOSITORY"\nbranch = "main"\npath = "data/poswksht.xlsm"',language='toml')
            st.caption('The token needs read access to this repository. Reuse github.token if it has access, or set workbook.token separately. Keep the trade workbook in a private repository. This app reads the file; it does not upload or modify it.')
        return None,dict(kind='github')
    try:
        with st.spinner('Loading connected trade workbook...'):
            book,provenance = tested_book if tested_book is not None else load_github_workbook(config,fallback_token)
        st.caption('Workbook: '+provenance['repository']+'/'+provenance['path']+' @ '+provenance['commit'][:8])
        st.success('Workbook loaded from GitHub — no upload needed.')
        return book,provenance
    except (ValueError,KeyError,TypeError) as exc:
        st.error(workbook_connection_error(exc,config,fallback_token))
        return None,dict(kind='github')


def platform_checks(reports, ice_date, today):
    rows=[]
    def add(area,status,count,detail,location):
        rows.append(dict(area=area,status=status,affected_rows=count,detail=detail,open_tab=location))
    flags=reports.get('scoped_import_flags.csv')
    if flags is not None:
        flags=flags.loc[flags.requires_scheduled_review]
    else:
        flags=reports.get('import_flags.csv')
    add('Workbook import','NOT GENERATED' if flags is None else 'REVIEW' if len(flags) else 'CLEAR',
        None if flags is None else len(flags),'Reporting-period and unresolved-scope flags require review. Historical/future flags remain in the audit report. A clear check is not trade approval.','Flags')
    for area,name,column,prefix,numeric,location in [
        ('October cash pricing','october_cash_system_pricing.csv','status','SYSTEM PRICED','system_signed_amount','October cash'),
        ('Forward price coverage','forward_pricing_coverage.csv','price_status','AVAILABLE','ice_settlement_price','ICE forwards'),
        ('Forward valuation','forward_trade_valuation.csv','valuation_status','INDICATIVE','indicative_mtm','Forward valuation')]:
        frame=reports.get(name)
        if frame is None or frame.empty:
            add(area,'NOT GENERATED',None,'No report rows for the current selection; this does not mean zero exposure.',location)
            continue
        if column not in frame:
            add(area,'REVIEW',len(frame),'Expected status column is missing.',location)
            continue
        good=frame[column].fillna('').astype(str).str.startswith(prefix)
        if numeric:
            if numeric not in frame:
                good &= False
            else:
                good &= pd.to_numeric(frame[numeric],errors='coerce').map(lambda v:pd.notna(v) and math.isfinite(v))
        count=int((~good).sum())
        add(area,'REVIEW' if count else 'CALCULATED',count,
            'Rows with missing, blocked, or unrecognized results.' if count else 'Current report rows calculated; valuation and cash remain subject to their stated assumptions.',location)
    if ice_date is None or pd.isna(ice_date):
        add('ICE source date','MISSING',None,'No valid ICE trade date loaded.','ICE forwards')
    else:
        age=(pd.Timestamp(today).normalize()-pd.Timestamp(ice_date).normalize()).days
        add('ICE source date','REVIEW' if age<0 or age>4 else 'CURRENT',None,
            f'ICE trade date {pd.Timestamp(ice_date).date()}; age {age} calendar days. Review threshold: over 4 days or a future date.','ICE forwards')
    balance=reports.get('monthly_physical_balance.csv')
    if balance is None or balance.empty or 'status' not in balance:
        add('Physical balance','DEFERRED',None,'Manual demand and storage inputs are deferred.','Physical balance')
    else:
        incomplete=~balance.status.isin(['BALANCED','SURPLUS','SHORTFALL'])
        add('Physical balance','INPUTS / REVIEW' if incomplete.any() else 'CALCULATED',int(incomplete.sum()),
            'Incomplete inputs are retained as missing; no overall business P&L is implied.','Physical balance')
    return pd.DataFrame(rows)


def risk_action_list(reports,asof,ice_date=None):
    """Source-derived work list; waiting for an invoice is not an overdue invoice."""
    today=pd.Timestamp(asof).normalize();rows=[]
    columns=['task_id','priority','area','item','counterparty_or_account','delivery_month','due_date','source_status','affected_rows','next_action','open_tab','source_report','as_of']
    def add(priority,area,item,entity='',month=None,due=None,status='',count=None,action='',tab='',source=''):
        def date(value):return None if value is None or pd.isna(value) or not txt(value) else pd.Timestamp(value).date().isoformat()
        month=date(month);due=date(due)
        key='|'.join([area,item,txt(entity),month or '',source])
        rows.append(dict(task_id=hashlib.sha256(key.encode()).hexdigest()[:16],priority=priority,area=area,item=item,
            counterparty_or_account=txt(entity),delivery_month=month,due_date=due,source_status=status,affected_rows=count,
            next_action=action,open_tab=tab,source_report=source,as_of=today.date().isoformat()))
    source_map={'Workbook import':'scoped_import_flags.csv' if 'scoped_import_flags.csv' in reports else 'import_flags.csv','October cash pricing':'october_cash_system_pricing.csv','Forward price coverage':'forward_pricing_coverage.csv','Forward valuation':'forward_trade_valuation.csv','Physical balance':'monthly_physical_balance.csv'}
    for _,r in platform_checks(reports,ice_date,today).iterrows():
        if r.status in ['CLEAR','CALCULATED','CURRENT']:continue
        priority='SETUP / INPUTS' if r.status in ['DEFERRED','NOT GENERATED'] else 'ATTENTION NOW'
        add(priority,'Pricing and positions',r.area,status=r.status,count=r.affected_rows,action=r.detail,tab=r.open_tab,source=source_map.get(r.area,''))
    storage=storage_review_report(reports,today)
    for _,r in storage.iterrows():
        if r.check=='Update tasks' and r.status!='NOT CONFIGURED':continue
        if r.status in ['SCREEN COMPLETE','SOURCE COVERED','COMPLETE','MODELED COMPONENTS COMPLETE']:continue
        priority='BUILDOUT' if r.status in ['NOT IMPLEMENTED','OUTSIDE CURRENT MODEL'] else 'WAITING FOR INPUT' if r.status in ['WAITING FOR CIV','NOT STARTED'] else 'SETUP / INPUTS' if r.status=='NOT CONFIGURED' else 'UPCOMING' if r.status=='UPCOMING' else 'ATTENTION NOW'
        add(priority,'Storage',r.check,entity=r.book,status=r.status,count=r.items_requiring_review,action=r.next_action,tab=r.open_tab,source=r.source_report)
    for source,status_col,month_col,tab in [('storage_release_update_tasks.csv','task_status','delivery_month','Columbia storage'),('storage_deo_release_update_tasks.csv','status','month','DEO storage')]:
        for _,r in reports.get(source,pd.DataFrame()).iterrows():
            if r[status_col]=='COMPLETE':continue
            due=pd.Timestamp(r.due_date).normalize()
            state='OVERDUE' if due<today else 'DUE TODAY' if due==today else 'UPCOMING'
            add('UPCOMING' if due>today else 'ATTENTION NOW','Storage',r.task,entity=r.account,month=r[month_col],due=due,status=state,count=1,action=r.task,tab=tab,source=source)
    for source,area in [('counterparty_settlement_monthly.csv','Counterparty forecasts'),('physical_open_exposure_monthly.csv','Delivered supply pricing')]:
        f=reports.get(source,pd.DataFrame())
        if f.empty or 'status' not in f:continue
        failed=f.loc[f.status.fillna('').str.startswith(('INCOMPLETE','BLOCKED','REVIEW'))]
        for _,r in failed.iterrows():
            add('ATTENTION NOW',area,'Resolve incomplete forecast',entity=r.get('counterparty',r.get('book',r.get('index_code',''))),month=r.get('month'),status=r.status,count=1,action='Resolve missing prices or inputs in the source report; regenerate the forecast.',tab='Counterparty exposure' if area=='Counterparty forecasts' else 'Physical balance',source=source)
    source='counterparty_credit_monitor.csv';credit=reports.get(source,pd.DataFrame())
    if credit.empty:
        add('SETUP / INPUTS','Credit','Configure receivables monitoring',status='NOT CONFIGURED',action='Review approved limits and coverage in Counterparty exposure; export the combined profile for scheduled runs.',tab='Counterparty exposure',source=source)
    for _,r in credit.iterrows():
        if r.status.startswith('WITHIN') or r.status.startswith('PREPAY ONLY — no'):continue
        priority='SETUP / INPUTS' if r.status=='NOT CONFIGURED' else 'WAITING FOR INPUT' if r.status.startswith('REVIEW — confirm') else 'ATTENTION NOW'
        action='Review limit breach, recorded receivables and payment terms; do not assume collateral satisfies prepay terms.' if r.status=='BREACH' else 'Review zero remaining headroom before extending credit.' if r.status=='AT LIMIT' else 'Enter the approved limit or reconcile receivables coverage, collateral and source reference. Blank coverage is not a zero balance.'
        add(priority,'Credit','Receivables limit and coverage',entity=r.counterparty,status=r.status,count=1,action=action,tab='Counterparty exposure',source=source)
    source='counterparty_invoice_aging.csv'
    for _,r in reports.get(source,pd.DataFrame()).iterrows():
        if r.status!='OVERDUE':continue
        add('ATTENTION NOW','Invoices','Overdue recorded invoice '+str(r.invoice_id),entity=r.counterparty,month=r.get('delivery_month'),due=r.due_date,status=r.status,count=1,action='Review the recorded invoice and any payments; update the ledger with supporting records.',tab='Counterparty exposure',source=source)
    source='counterparty_invoice_unassigned.csv'
    for _,r in reports.get(source,pd.DataFrame()).iterrows():
        add('ATTENTION NOW','Invoices','Assign delivery month to invoice '+str(r.invoice_id),entity=r.counterparty,status='MISSING DELIVERY MONTH',count=1,action='Enter the delivery month from the received invoice; do not use its issue date as a substitute.',tab='Counterparty exposure',source=source)
    source='counterparty_invoice_reconciliation.csv';rec=reports.get(source,pd.DataFrame())
    for _,r in rec.iterrows():
        if r.status.startswith('AMOUNTS AGREE') or r.status=='REVIEW — assign delivery months':continue
        pending=r.status=='AWAITING INVOICE'
        future=pending and pd.Timestamp(r.month).to_period('M')>today.to_period('M')
        add('UPCOMING' if future else 'WAITING FOR INPUT' if pending else 'ATTENTION NOW','Invoices','Future invoice reconciliation' if future else 'Receive and compare invoice' if pending else 'Review invoice difference',entity=r.counterparty,month=r.month,status=r.status,count=1,
            action='Upload the actual invoice when received, enter its delivery month and compare amounts. Forecasts are not unpaid balances; no invoice due date is assumed.' if pending else 'Review invoice coverage, commodity forecast, fees, fuel and taxes. An amount difference is not automatically an error.',tab='Counterparty exposure',source=source)
    if rec.empty and reports.get('counterparty_invoice_aging.csv',pd.DataFrame()).empty:
        add('WAITING FOR INPUT','Invoices','Receive actual invoices',status='NO RECORDED INVOICES',action='Continue the buildout. Enter actual invoices only when received; do not create October invoices from forecasts.',tab='Counterparty exposure',source='counterparty_invoice_status.csv')
    source='pool_molecule_balance_monthly.csv'
    for _,r in reports.get(source,pd.DataFrame()).iterrows():
        if r.status=='PLANNING BALANCED':continue
        priority='ATTENTION NOW' if str(r.status).startswith('REVIEW') or pd.Timestamp(r.month)<=today.to_period('M').start_time else 'UPCOMING'
        add(priority,'Pool molecules','Review pool surplus / deficit',entity=r.pool_name,month=r.month,status=r.status,count=1,action='Resolve location/data reviews, then use daily quantities to buy deficit gas, sell surplus or assign confirmed transport. Pools and delivery demand are separate.',tab='Pool molecule balance',source=source)
    shadow_status=reports.get('parallel_register_status.csv',pd.DataFrame())
    if not shadow_status.empty:
        state=shadow_status.status.iloc[0]
        if state!='MATCHED':
            comparison=reports.get('parallel_deal_reconciliation.csv',pd.DataFrame())
            count=int(comparison.status.ne('MATCHED').sum()) if not comparison.empty else 0
            add('BUILDOUT' if state=='NOT STARTED' else 'ATTENTION NOW','Parallel run','Register or reconcile Excel deals',status=state,count=count,action='Review the Excel baseline, draft matches, ambiguous identities and amendments in the shadow register. Excel remains authoritative.',tab='Parallel deal register',source='parallel_deal_reconciliation.csv')
    add('BUILDOUT','Storage','Monthly statement reconciliation — Columbia 1, Columbia 2 and DEO',status='NOT IMPLEMENTED',action='Build separate account/month statement upload, opening + injections - withdrawals - losses + adjustments = closing comparison, cost differences, and first-of-month tasks. Awaiting statement, Review differences, Reconciled. Link existing injections without duplicating volumes or costs.',tab='Storage pricing & summary')
    result=pd.DataFrame(rows,columns=columns).drop_duplicates('task_id')
    order={'ATTENTION NOW':0,'SETUP / INPUTS':1,'UPCOMING':2,'WAITING FOR INPUT':3,'BUILDOUT':4}
    result['_order']=result.priority.map(order)
    return result.sort_values(['_order','due_date','area','counterparty_or_account','delivery_month','item'],na_position='last').drop(columns='_order').reset_index(drop=True)


def render_risk_actions(ice_date):
    st.subheader('Risk action list')
    st.caption('Derived from current reports. Update the source inputs to clear an item; no extra completion boxes. Upcoming releases and pending invoices remain separate from items needing attention now. This list refreshes with the app and is included in run archives; it does not send notifications.')
    actions=risk_action_list(st.session_state['_run_reports'],pd.Timestamp.now(tz='America/Chicago').tz_localize(None),ice_date)
    st.session_state['_run_reports']['risk_action_list.csv']=actions
    a,b,c=st.columns(3)
    a.metric('Attention now',int(actions.priority.eq('ATTENTION NOW').sum()))
    b.metric('Upcoming tasks',int(actions.priority.eq('UPCOMING').sum()))
    c.metric('Waiting for inputs',int(actions.priority.eq('WAITING FOR INPUT').sum()))
    selection=st.selectbox('Action list view',['All items','ATTENTION NOW','SETUP / INPUTS','UPCOMING','WAITING FOR INPUT','BUILDOUT'],key='risk_action_priority')
    view=actions if selection=='All items' else actions.loc[actions.priority.eq(selection)]
    horizon=st.selectbox('Planning horizon',['Current month and winter through March','All delivery months'],key='risk_action_horizon')
    if horizon.startswith('Current'):
        today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
        end=pd.Timestamp(year=today.year+1 if today.month>3 else today.year,month=3,day=31)
        months=pd.to_datetime(view.delivery_month,errors='coerce')
        visible=months.isna()|months.le(end)
        st.caption(str(int((~visible).sum()))+' longer-term tasks hidden in this view; full export and headline counts retain all tasks.')
        view=view.loc[visible]
    show(view,'risk_action_list_selected.csv')
    st.download_button('Download full action list',actions.to_csv(index=False),file_name='risk_action_list.csv',mime='text/csv',key='risk_actions_full')
    st.caption('Counts represent tasks, not unique trades, financial exposure or approval to transact. Missing reports stay visible; this list does not certify overall system completeness.')


TCO_FSS_SOURCE='https://hostedtariffs.com/tco/pdf/tariff.pdf'
EGTS_GSS_SOURCE='https://infopost.bhegts.com/docs/egts/tariff/full-tariff.pdf'
DEO_FSS_SOURCE='https://www.enbridgegas.com/-/media/Extranet-Pages/ohio/rates-tariffs/fss.pdf?hash=34EE8F28DD661CDF72E0B614F712F8A5&rev=8aa0770ddac74d2e9ff332397071eb4b'


def tco_fss_limits(month):
    m=int(month)
    injection={4:.15,5:.20,6:.20,7:.20,8:.18,9:.13,10:.09,11:.05,12:.10,1:.10,2:.10,3:.10}[m]
    return dict(injection_limit_fraction=injection,injection_divisor=30 if m in [11,12] else 25,
        min_net_withdrawal_fraction=.1 if m in [2,3] else 0,
        max_net_withdrawal_fraction={11:.4,12:.4,1:.4,2:.3,3:.2}.get(m),
        inventory_ceiling_fraction={1:.65,3:.25,6:.60,8:.85}.get(m,1.))


def tco_ratchet(inventory,capacity):
    # Inventory arithmetic can land a few trillionths below an exact threshold.
    f=(inventory+1e-7)/capacity
    return 1. if f>=.30 else .80 if f>=.20 else .65 if f>=.10 else .50


def storage_daily_screen(accounts,forecast,actual):
    """Screen uniform forecast; actual gaps explicitly prevent a complete inventory test."""
    out=[]
    for _,a in accounts.iterrows():
        cap=float(a.capacity_mcf);mdsq=storage_number(a.mdsq_mcf_day,'MDSQ',0,True)
        for _,r in forecast[forecast.account.eq(a.account)].iterrows():
            month=pd.Timestamp(r.month);n=month.days_in_month;balance=r.beginning_mcf
            for day in pd.date_range(month,periods=n):
                inj=r.injections_mcf/n;wd=r.withdrawals_mcf/n;begin=balance;balance+=inj-wd
                # Lowest of beginning/end is a conservative screen; operator's best estimate controls.
                multiplier=tco_ratchet(min(begin,balance),cap)
                limit=None if mdsq is None else mdsq*multiplier
                issues=[]
                if min(begin,balance)<-1e-7:issues.append('Inventory shortfall')
                if max(begin,balance)>cap+1e-7:issues.append('Inventory exceeds SCQ')
                if limit is None:issues.append('MDSQ missing')
                elif wd>limit+1e-7:issues.append('Withdrawal exceeds screened MDWQ')
                rule=tco_fss_limits(day.month);injlimit=cap*rule['injection_limit_fraction']/rule['injection_divisor']
                if inj>injlimit+1e-7:issues.append('Injection exceeds MDIQ')
                out.append(dict(account=a.account,date=day.date().isoformat(),basis='Uniform monthly forecast',beginning_mcf=begin,ending_mcf=balance,
                    injection_mcf=inj,withdrawal_mcf=wd,screened_ratchet=multiplier,mdwq_mcf=limit,mdiq_mcf=injlimit,
                    review='; '.join(issues) if issues else 'Within screened quantity limits; operator confirmation required'))
        if not actual.empty:
            rows=actual[actual.account.eq(a.account)].sort_values('date');last=None;complete=True
            start=pd.Timestamp(forecast[forecast.account.eq(a.account)].month.min())
            for _,r in rows.iterrows():
                day=pd.Timestamp(r.date)
                if day!=(start if last is None else last+pd.Timedelta(days=1)):complete=False
                last=day
                multiplier=tco_ratchet(min(r.beginning_mcf,r.ending_mcf),cap) if complete else None
                limit=None if mdsq is None or multiplier is None else mdsq*multiplier
                rule=tco_fss_limits(day.month);injlimit=cap*rule['injection_limit_fraction']/rule['injection_divisor']
                issues=[]
                if not complete:issues.append('Missing prior daily records; inventory ratchet not verified')
                if mdsq is None:issues.append('MDSQ missing')
                if complete and min(r.beginning_mcf,r.ending_mcf)<0:issues.append('Inventory shortfall')
                if complete and max(r.beginning_mcf,r.ending_mcf)>cap:issues.append('Inventory exceeds SCQ')
                if limit is not None and r.withdrawal_mcf>limit+1e-7:issues.append('Withdrawal exceeds screened MDWQ')
                if r.injection_mcf>injlimit+1e-7:issues.append('Injection exceeds MDIQ')
                out.append(dict(account=a.account,date=r.date,basis='Entered actual movements',beginning_mcf=r.beginning_mcf,ending_mcf=r.ending_mcf,
                    injection_mcf=r.injection_mcf,withdrawal_mcf=r.withdrawal_mcf,screened_ratchet=multiplier,mdwq_mcf=limit,mdiq_mcf=injlimit,
                    review='; '.join(issues) if issues else 'Within screened quantity limits; operator confirmation required'))
    return pd.DataFrame(out)


def egts_gss_capability(capacity,demand,inventory,month,interruptible=False):
    """GSS 7.4/7.5, 8.4–8.6. Same energy units required throughout."""
    if capacity<=0 or demand<=0 or inventory<0 or inventory>capacity:raise ValueError('Capacity, demand and inventory are outside valid ranges.')
    fraction=inventory/capacity
    exempt=interruptible or demand<=capacity/140
    ratchet=1. if exempt or fraction>.35 else .92 if fraction>.16 else .70 if fraction>.10 else .63
    minimum=.35 if month in [12,1] else .15 if month==2 else 0
    penalty=.90 if fraction<minimum else 1.
    daily=demand*ratchet*penalty
    monthly_factor=.70 if demand>=capacity/60 else .875
    divisor=214 if month in [11,12,1,2,3] or fraction>.5 else 180
    return dict(base_ratchet=ratchet,minimum_inventory_fraction=minimum,additional_reduction_factor=penalty,daily_entitlement=daily,
        monthly_average_factor=monthly_factor,monthly_average_at_current_inventory=demand*ratchet*monthly_factor,injection_entitlement=capacity/divisor,
        note='OFOs, service-agreement terms, releases and inventory credits may change entitlement. Summer: April–October; turnover deadline: April 15.')


def deo_fss_band(month):
    # End-of-month levels correspond to next month's tariff target date.
    return {4:(.04,.13),5:(.20,.30),6:(.34,.44),7:(.51,.57),8:(.65,.71),9:(.79,.85),10:(.95,1.),11:(.92,.95),12:(.55,.65),1:(.30,.35),2:(.12,.17),3:(0.,0.)}[month]


DEO_ECPS_SOURCE='https://www.enbridgegas.com/-/media/Extranet-Pages/ohio/rates-tariffs/ec-pooling-terms-conditions.pdf?hash=2AF605D9D7DB40A2E3ED14815B69FE4D&la=en&rev=6858d633d154499ba912a10e42b4ec12'


def deo_ecps_ratchet(inventory,cumulative_injection):
    inventory=storage_number(inventory,'Inventory',0)
    civ=storage_number(cumulative_injection,'End-of-summer cumulative injection volume',0)
    if civ==0:raise ValueError('Cumulative injection volume must be positive.')
    f=inventory/civ
    return 1. if f>.50 else .75 if f>=.30 else .60


def deo_ecps_schedule(civ,opening,rows):
    civ=storage_number(civ,'Cumulative injection volume',0)
    opening=storage_number(opening,'Opening inventory',0)
    if civ==0:raise ValueError('Cumulative injection volume must be positive.')
    months=pd.to_datetime(rows.month,errors='raise')
    if list(months)!=list(pd.date_range('2026-11-01',periods=5,freq='MS')):raise ValueError('DEO schedule must contain November 2026–March 2027 in order.')
    balance=opening;out=[]
    for (_,r),month in zip(rows.iterrows(),months):
        wd=storage_number(r.withdrawals,'Withdrawals',0);inj=storage_number(r.injections,'Injections',0)
        adjustment=storage_number(r.utility_adjustment,'Utility adjustment')
        begin=balance;balance+=inj-wd+adjustment
        low,high=deo_fss_band(month.month)
        issues=[]
        if balance < 0:issues.append('Inventory shortfall')
        if balance<low*civ-1e-7 or balance>high*civ+1e-7:issues.append('Outside ECPS target-date inventory band')
        demand=storage_number(r.assigned_storage_demand,'Assigned daily Storage Demand',0,True)
        multiplier=None if balance<0 else deo_ecps_ratchet(balance,civ)
        entitlement=None if demand is None or multiplier is None else demand*multiplier
        if inj>0:issues.append('Winter reinjection permission required; interruptible')
        if inj/month.days_in_month>civ/214+1e-7:issues.append('Average reinjection exceeds CIV/214 request limit')
        if demand is None:issues.append('Assigned Storage Demand missing')
        elif entitlement is not None and wd/month.days_in_month>entitlement+1e-7:issues.append('Average withdrawal exceeds end-inventory ratchet screen')
        out.append(dict(month=month.date().isoformat(),beginning_inventory=begin,injections=inj,withdrawals=wd,utility_adjustment=adjustment,ending_inventory=balance,
            target_date=(month+pd.offsets.MonthBegin(1)).date().isoformat() if month.month!=3 else '2027-03-31',
            minimum_inventory=low*civ,maximum_inventory=high*civ,ending_ratchet=multiplier,assigned_storage_demand=demand,screened_daily_demand=entitlement,
            review='; '.join(issues) if issues else 'Within monthly screen; daily nominations and utility notices still required'))
    return pd.DataFrame(out)


def deo_posted_daily(opening, rows):
    """November 2026 utility-posted screen; missing days never imply zero activity."""
    balance=storage_number(opening,'November opening inventory',0)
    if balance>332237:raise ValueError('Opening inventory exceeds posted November capacity.')
    data=rows.copy()
    data['date']=pd.to_datetime(data['date'],errors='raise')
    if data.date.isna().any() or data.date.duplicated().any():raise ValueError('Enter one valid row per date.')
    if any(d!=d.normalize() or d.year!=2026 or d.month!=11 for d in data.date):raise ValueError('Posted limits apply only to November 2026 dates.')
    data=data.sort_values('date');expected=pd.Timestamp('2026-11-01');complete=True;out=[]
    for _,r in data.iterrows():
        day=r.date
        complete=complete and day==expected
        inj=storage_number(r.injections,'Daily injections',0)
        wd=storage_number(r.withdrawals,'Daily withdrawals',0)
        adj=storage_number(r.utility_adjustment,'Daily utility adjustment')
        reported=storage_number(r.reported_balance,'Reported ending balance',0,True)
        begin=balance;balance+=inj-wd+adj
        issues=[]
        # End-of-day tier screen retains posted rounded thresholds exactly.
        limit=None
        if not complete:issues.append('Missing earlier activity; tier and statement reconciliation withheld')
        elif balance<0:issues.append('Inventory shortfall')
        elif balance>332237:issues.append('Above posted capacity')
        else:
            limit=6401. if balance>166119 else 4801. if balance>99671 else 3841. if balance>0 else 0.
            if wd>limit:issues.append('Withdrawal exceeds posted end-inventory tier screen')
        if inj>0:issues.append('Winter reinjection requires utility permission; injection limit not verified here')
        variance=reported-balance if complete and reported is not None else None
        if variance is not None and abs(variance)>1e-7:issues.append('Reported balance does not reconcile')
        out.append(dict(date=day.date().isoformat(),beginning_inventory=begin,injections=inj,withdrawals=wd,utility_adjustment=adj,
            ending_inventory=balance,history_complete=complete,posted_withdrawal_screen_dth_day=limit,
            reported_balance=reported,statement_variance=variance,
            review='; '.join(issues) if issues else 'Within posted daily screen; utility notices and approvals still apply'))
        expected=day+pd.Timedelta(days=1)
    return pd.DataFrame(out)


def deo_actual_reports(data,plan,asof):
    """Reconcile confirmed November actual inputs without promoting draft rows."""
    if data.get('schema')!='deo-daily-2' or data.get('units')!='Dth' or data.get('actuals_confirmed') is not True:
        raise ValueError('Actual inputs require Dth, deo-daily-2 and explicit confirmation.')
    if type(data.get('opening_confirmed')) is not bool:raise ValueError('Confirm opening status explicitly.')
    if not txt(data.get('source_reference')):raise ValueError('Enter the utility activity source reference.')
    rows=pd.DataFrame(data['rows'])
    if rows.empty:raise ValueError('Actual activity requires at least one row.')
    result=deo_posted_daily(data['opening'],rows)
    cutoff=min(pd.Timestamp(asof).normalize()-pd.Timedelta(days=1),pd.Timestamp('2026-11-30'))
    dates=pd.to_datetime(result.date)
    if (dates>cutoff).any():raise ValueError('Actual rows must be completed delivery days before the valuation date.')
    expected=pd.date_range('2026-11-01',cutoff,freq='D')
    missing=len(expected.difference(pd.DatetimeIndex(dates)))
    review=int(result.review.ne('Within posted daily screen; utility notices and approvals still apply').sum())
    contiguous=bool(result.history_complete.all()) and data['opening_confirmed']
    complete=contiguous and missing==0 and len(result)==30
    planned=pd.DataFrame(plan['plan']).iloc[0]
    adopted=plan.get('plan_adopted') is True
    ending=float(result.ending_inventory.iloc[-1]) if contiguous else None
    planned_ending=plan['opening']+planned.injections-planned.withdrawals+planned.utility_adjustment if adopted else None
    summary=dict(month='2026-11',last_entered_day=result.date.iloc[-1],expected_through=cutoff.date().isoformat(),
        missing_elapsed_days=missing,opening_confirmed=data['opening_confirmed'],entered_days=len(result),
        entered_injections_dth=float(result.injections.sum()),entered_withdrawals_dth=float(result.withdrawals.sum()),
        reconciled_ending_inventory_dth=ending,full_month_complete=complete,
        planned_month_end_inventory_dth=planned_ending,
        actual_minus_plan_month_end_dth=ending-planned_ending if complete and adopted else None,
        status='REVIEW' if missing or review or not data['opening_confirmed'] else 'Actual ledger complete through cutoff; month-end comparison pending' if not complete else 'Monthly actual ledger complete',
        source_reference=data['source_reference'],scope='November posted screen only; no automatic plan reset or additional fuel deduction')
    reports={'storage_deo_actual_inputs.csv':rows,'storage_deo_actual_daily_screen.csv':result,
        'storage_deo_actual_reconciliation.csv':pd.DataFrame([summary])}
    return reports,int(bool(missing or review or not data['opening_confirmed']))


def render_deo_daily(opening):
    st.subheader('DEO November daily tracker — Dth')
    st.caption('Contract 12663 / ECPS-S. Uses the November utility-posted capacity and withdrawal tiers. Opening inventory remains a forecast until confirmed. Enter a row for every elapsed day, including zero-activity days; missing days prevent a confirmed balance or tier check. A reported balance is a reconciliation check, not an automatic inventory reset.')
    upload=st.file_uploader('Restore DEO daily inputs (JSON)',type=['json'],key='deo_daily_restore')
    if upload is not None and st.button('Load DEO daily inputs'):
        try:
            saved=json.loads(upload.getvalue())
            if saved.get('schema') not in ['deo-daily-1','deo-daily-2'] or saved.get('units')!='Dth':raise ValueError('Invalid daily input format.')
            deo_posted_daily(saved['opening'],pd.DataFrame(saved['rows']))
            st.session_state['deo_daily_opening']=float(saved['opening'])
            st.session_state['deo_actual_confirmed']=saved.get('actuals_confirmed',False)
            st.session_state['deo_opening_confirmed']=saved.get('opening_confirmed',False)
            st.session_state['deo_actual_source']=saved.get('source_reference','')
            st.session_state['_deo_daily_rows']=saved['rows']
            st.session_state['_deo_daily_rev']=st.session_state.get('_deo_daily_rev',0)+1
        except (ValueError,KeyError,TypeError) as exc:st.error('Cannot restore daily inputs: '+str(exc))
    daily_open=st.number_input('Daily tracker November 1 opening inventory (Dth)',min_value=0.,value=float(opening),key='deo_daily_opening')
    st.caption('Review this opening after restoring inputs or changing the monthly forecast. The daily tracker retains its own saved opening.')
    defaults=[dict(date='2026-11-01',injections=0.,withdrawals=0.,utility_adjustment=0.,reported_balance=None)]
    entries=st.data_editor(pd.DataFrame(st.session_state.get('_deo_daily_rows',defaults)),num_rows='dynamic',hide_index=True,
        column_config={c:st.column_config.NumberColumn(c) for c in ['injections','withdrawals','utility_adjustment','reported_balance']},key='deo_daily_rows_'+str(st.session_state.get('_deo_daily_rev',0)))
    st.caption('Rows are draft inputs until you enter activity. This table is separate from the monthly plan and does not overwrite it. December–March require their own posted allocations. No CIV-based inventory-band certification is made by this screen.')
    confirmed=st.checkbox('These rows are actual completed daily activity',key='deo_actual_confirmed')
    opening_confirmed=st.checkbox('November opening inventory confirmed against utility records',key='deo_opening_confirmed')
    source=st.text_input('Utility activity source reference',key='deo_actual_source')
    if entries.empty:
        st.info('Enter at least one daily row to calculate or export.');return
    try:
        result=deo_posted_daily(daily_open,entries)
        show(entries,'storage_deo_daily_inputs.csv')
        show(result,'storage_deo_daily_screen.csv')
        show(pd.DataFrame([dict(opening_inventory=daily_open,units='Dth',contract='12663',effective_month='2026-11',opening_status='User input; initial default is forecast')]),'storage_deo_daily_account.csv')
        export=dict(schema='deo-daily-2',units='Dth',opening=daily_open,rows=json.loads(entries.to_json(orient='records')),
            actuals_confirmed=confirmed,opening_confirmed=opening_confirmed,source_reference=source)
        if confirmed:
            operations=st.session_state.get('_scheduled_storage_inputs',{}).get('deo_operations')
            if operations is not None:
                actual_reports,actual_review=deo_actual_reports(export,operations,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
                operations['actuals']=export
                for name,report in actual_reports.items():show(report,name)
                if actual_review:st.warning('Actual inventory requires review: check missing elapsed days, opening confirmation and statement differences.')
        st.download_button('Download DEO daily inputs',json.dumps(export,indent=2),'riverbend_deo_daily_inputs.json','application/json')
        st.caption('Download inputs to preserve edits across sessions. Confirmed actual rows are also included in the scheduled storage profile. Re-export and upload that profile after each activity update. Draft rows remain excluded from scheduled actuals. Month-end actual-versus-plan variance requires all 30 November days and a confirmed opening.')
    except (ValueError,TypeError,KeyError) as exc:st.error('Daily inputs need correction: '+str(exc))


def deo_injection_cost(total, rows):
    """DEO receipt-to-burner-tip conversion, once on new receipts; no withdrawal shrink."""
    total=storage_number(total,'Summer net injection volume (Dth)',0)
    months=list(pd.date_range('2026-04-01',periods=7,freq='MS'))
    if list(pd.to_datetime(rows.month,errors='raise'))!=months:raise ValueError('Cost schedule requires April–October 2026 in order.')
    weights=[.085,.165,.14,.15,.14,.14,.18];out=[]
    for (_,r),month,weight in zip(rows.iterrows(),months,weights):
        price=storage_number(r.spot_usd_dth,'Monthly average daily spot price',optional=True)
        net=total*weight;gross=net/.953;fuel=gross-net
        out.append(dict(injection_month=month.date().isoformat(),schedule_fraction=weight,net_injection_dth=net,
            gross_purchase_dth=gross,fuel_dth=fuel,fuel_fraction=.047,spot_usd_dth=price,
            conversion_basis='Received city-gate gas to burner-tip equivalent',conversion_side='Receipt only',withdrawal_conversion_fraction=0.,
            commodity_cost_usd=None if price is None else net*price,
            fuel_cost_usd=None if price is None else fuel*price,
            injection_month_expense_usd=None if price is None else gross*price,
            synthetic_usd_per_net_dth=None if price is None else price/.953,
            price_status='Missing price' if price is None else r.get('pricing_basis','Manual monthly price')))
    return pd.DataFrame(out)


def render_deo_injection_cost():
    with st.expander('DEO synthetic injection-month expense'):
        st.caption('Cost-only April–October 2026 schedule: 8.5%, 16.5%, 14%, 15%, 14%, 14%, 18%. The supplied 315,520 schedule is treated as net Dth for this DEO calculation; confirm that unit before using the expense. It does not set capacity, CIV, or opening inventory. Enter the DEO monthly average daily spot price per received city-gate Dth; no Columbia price is substituted. The user-confirmed 4.7% utility conversion applies once to new receipts, including receipts used for storage injection; already-adjusted storage withdrawals receive no second deduction.')
        uploaded=st.file_uploader('Restore DEO cost inputs (JSON)',type=['json'],key='deo_cost_restore')
        if uploaded is not None and st.button('Load DEO cost inputs'):
            try:
                saved=json.loads(uploaded.getvalue())
                if saved.get('schema')!='deo-cost-1' or saved.get('units')!='Dth':raise ValueError('Invalid DEO cost input format.')
                if saved.get('receipt_conversion_fraction',.047)!=.047 or saved.get('withdrawal_conversion_fraction',0.)!=0. or saved.get('conversion_side','Receipt only')!='Receipt only':raise ValueError('DEO requires 4.7% receipt-only conversion and no second withdrawal conversion.')
                deo_injection_cost(saved['net_injection_dth'],pd.DataFrame(saved['prices']))
                st.session_state['deo_cost_volume']=float(saved['net_injection_dth'])
                st.session_state['_deo_cost_prices']=saved['prices']
                st.session_state['_deo_cost_rev']=st.session_state.get('_deo_cost_rev',0)+1
            except (ValueError,TypeError,KeyError) as exc:st.error('Cannot restore cost inputs: '+str(exc))
        total=st.number_input('DEO summer cost-only net injections (Dth)',min_value=0.,value=315520.,key='deo_cost_volume')
        defaults=[dict(month=f'2026-{m:02}-01',spot_usd_dth=None) for m in range(4,11)]
        auto_price=st.session_state.get('_storage_auto_enabled',False)
        prices=st.data_editor(storage_price_input_frame('DEO',pd.DataFrame(st.session_state.get('_deo_cost_prices',defaults))),disabled=True if auto_price else ['month'],
            column_config={'spot_usd_dth':st.column_config.NumberColumn('Average daily spot ($/Dth)')},hide_index=True,key='deo_cost_prices_'+str(st.session_state.get('_deo_cost_rev',0))+('_auto' if auto_price else '_manual'))
        st.caption('Received city-gate Dth = net burner-tip-equivalent stored Dth ÷ 0.953. Utility shrink = received Dth × 4.7%. Expense = received Dth × that injection month’s city-gate price. This is the utility receipt conversion, not a separate storage fuel charge. Withdrawal conversion = 0%; stored gas is not purchased or expensed again on withdrawal. Zero and negative prices are valid; blanks remain missing. Calculations retain unrounded quantities; round dollars only for presentation.')
        try:
            result=deo_injection_cost(total,prices)
            show(result,'storage_deo_injection_month_cost.csv')
            complete=result.spot_usd_dth.notna().all()
            expense=float(result.injection_month_expense_usd.sum()) if complete else None
            summary=pd.DataFrame([dict(units='Dth',net_injection_dth=total,gross_purchase_dth=total/.953,
                fuel_dth=total/.953-total,fuel_fraction=.047,conversion_side='Receipt only',withdrawal_conversion_fraction=0.,rate_basis='User-confirmed utility receipt-to-burner-tip conversion',priced_months=int(result.spot_usd_dth.notna().sum()),
                total_injection_expense_usd=expense,synthetic_usd_per_net_dth=expense/total if complete and total>0 else None,
                status='All monthly prices entered' if complete else 'Total withheld: missing monthly prices')])
            show(summary,'storage_deo_injection_cost_summary.csv')
            show(prices,'storage_deo_cost_price_inputs.csv')
            if complete:
                st.metric('Summer injection expense',f'${expense:,.2f}')
                if total>0:st.metric('Synthetic cost per net stored Dth',f'${expense/total:.4f}')
            else:st.info('Enter all seven monthly prices for the full summer expense. Priced months are shown individually; missing months are not treated as zero.')
            export=dict(schema='deo-cost-1',units='Dth',receipt_conversion_fraction=.047,withdrawal_conversion_fraction=0.,conversion_side='Receipt only',net_injection_dth=total,prices=json.loads(prices.to_json(orient='records')))
            st.session_state.setdefault('_scheduled_storage_inputs',{})['deo']=export
            st.download_button('Download DEO cost inputs',json.dumps(export,indent=2),'riverbend_deo_cost_inputs.json','application/json')
            st.caption('Expense is assigned to the injection month, not the withdrawal month. This estimate excludes other charges and is not yet added to business P&L. Download inputs to preserve edits; scheduled snapshots use only the separately exported cost profile.')
        except (ValueError,TypeError,KeyError) as exc:st.error('DEO cost inputs need correction: '+str(exc))


def deo_saved_plan_reports(data,asof):
    if data.get('schema')!='deo-ecps-2' or data.get('units')!='Dth':raise ValueError('Saved DEO operations require Dth and schema deo-ecps-2.')
    if type(data.get('plan_adopted')) is not bool:raise ValueError('DEO plan adoption must be explicit.')
    rows=pd.DataFrame(data['plan']);months=list(pd.date_range('2026-11-01',periods=5,freq='MS'))
    if list(pd.to_datetime(rows.month,errors='raise'))!=months:raise ValueError('DEO requires November–March plan rows.')
    civ=storage_number(data['civ'],'CIV',0);opening=storage_number(data['opening'],'Opening inventory',0)
    confirmations=pd.DataFrame(data['allocations'])
    if list(pd.to_datetime(confirmations.month,errors='raise'))!=months:raise ValueError('DEO requires November–March allocation rows.')
    today=pd.Timestamp(asof).normalize();tasks=[];unconfirmed=0
    def task(month,label,complete):
        due=pd.Timestamp(month)
        state='COMPLETE' if complete else 'UPCOMING' if today<due else 'DUE TODAY' if today==due else 'OVERDUE'
        tasks.append(dict(account='DEO ECPS 12663',month=due.date().isoformat(),due_date=due.date().isoformat(),task=label,status=state))
    task(months[0],'Confirm end-of-summer CIV; do not substitute capacity or opening inventory',civ>0)
    task(months[0],'Review and adopt the winter movement plan',data['plan_adopted'])
    for i,m in enumerate(months):
        row=confirmations.iloc[i]
        if not isinstance(row.confirmed,(bool,type(pd.Series([True]).iloc[0]))):raise ValueError('Allocation confirmation must be true or false.')
        source_demand=storage_number(row.assigned_storage_demand,'Posted demand',0,True)
        demand=storage_number(rows.iloc[i].assigned_storage_demand,'Plan demand',0,True)
        complete=bool(row.confirmed) and source_demand is not None and source_demand>0 and same(source_demand,demand) and bool(txt(row.source_reference))
        if not complete:unconfirmed+=1
        task(m,'Obtain monthly utility allocation; update demand and source reference; export scheduled profile',complete)
    movement=[];balance=opening
    for _,r in rows.iterrows():
        inj=storage_number(r.injections,'Injections',0);wd=storage_number(r.withdrawals,'Withdrawals',0)
        adjustment=storage_number(r.utility_adjustment,'Adjustment');begin=balance;balance+=inj-wd+adjustment
        movement.append(dict(month=r.month,beginning_inventory_dth=begin,injections_dth=inj,withdrawals_dth=wd,
            utility_adjustment_dth=adjustment,ending_inventory_dth=balance,
            inventory_shortfall_dth=max(0.,-balance),inventory_status='REVIEW — inventory shortfall' if balance < -1e-7 else 'Nonnegative monthly ending inventory',
            basis='Adopted forecast' if data['plan_adopted'] else 'Draft; not adopted'))
    reports={'storage_deo_ecps_plan_inputs.csv':rows,'storage_deo_monthly_allocation_inputs.csv':confirmations,
        'storage_deo_release_update_tasks.csv':pd.DataFrame(tasks),'storage_deo_inventory_forecast.csv':pd.DataFrame(movement),
        'storage_deo_ecps_account_inputs.csv':pd.DataFrame([dict(units='Dth',civ=civ,opening_inventory=opening,plan_adopted=data['plan_adopted'],contract='12663')])}
    issues=0
    if civ>0:
        result=deo_ecps_schedule(civ,opening,rows)
        reports['storage_deo_ecps_schedule.csv']=result
        issues=int(result.review.ne('Within monthly screen; daily nominations and utility notices still required').sum())
    checks=dict(deo_civ_missing=int(civ<=0),deo_plan_not_adopted=int(not data['plan_adopted']),deo_unconfirmed_allocation_months=unconfirmed,deo_monthly_review_rows=issues)
    checks['deo_inventory_shortfall_months']=sum(r['ending_inventory_dth'] < -1e-7 for r in movement)
    checks['deo_actuals_review']=0
    if data.get('actuals') is not None:
        actual_reports,checks['deo_actuals_review']=deo_actual_reports(data['actuals'],data,asof)
        reports.update(actual_reports)
    reports['storage_deo_plan_summary.csv']=pd.DataFrame([dict(**checks,ending_inventory_dth=balance,
        planned_withdrawals_dth=sum(r['withdrawals_dth'] for r in movement) if data['plan_adopted'] else None,
        entered_plan_withdrawals_dth=sum(r['withdrawals_dth'] for r in movement),
        maximum_month_end_shortfall_dth=max(r['inventory_shortfall_dth'] for r in movement),
        first_shortfall_month=next((r['month'] for r in movement if r['ending_inventory_dth'] < -1e-7),None),
        status='REVIEW' if any(checks.values()) else 'Monthly forecast screen complete; posted daily limits and notices still require review',
        scope='Inventory forecast and CIV screen only; no new withdrawal fuel deduction or gas-purchase expense')])
    return reports,checks


def render_deo_ecps():
    st.subheader('DEO utility-allocated on-system storage — ECPS')
    render_deo_injection_cost()
    st.markdown('[Enbridge Gas Ohio ECPS §§5 and 13]('+DEO_ECPS_SOURCE+') govern this working profile based on your utility-allocation description. Optional purchased FSS and interstate EGTS GSS remain separate services.')
    st.caption('Winter ratchets and inventory bands use actual end-of-summer Cumulative Injection Volume (CIV). Summer injection limits use assigned Capacity. Neither is automatically replaced by opening inventory or the synthetic-cost schedule.')
    show(pd.DataFrame([dict(inventory_fraction='Above 50% of CIV',storage_demand_factor=1.),dict(inventory_fraction='30% through 50% of CIV',storage_demand_factor=.75),dict(inventory_fraction='Below 30% of CIV',storage_demand_factor=.60)]),'storage_deo_ecps_ratchets.csv')
    st.caption('§13.12 permits a further 25% Storage Demand reduction for missed inventory bands; utility enforcement is not assumed here. §13.13 winter reinjections require permission and are interruptible. OFOs can require movements or change effective demand. Obtain the utility-posted effective demand before nominations.')
    upload=st.file_uploader('Restore DEO ECPS inputs (JSON)',type=['json'],key='ecps_restore')
    if upload is not None and st.button('Load DEO ECPS inputs'):
        try:
            data=json.loads(upload.getvalue())
            if data.get('schema') not in ['deo-ecps-1','deo-ecps-2'] or data.get('units') not in ['Mcf','Dth']:raise ValueError('Invalid DEO input version or units.')
            if data.get('schema')=='deo-ecps-2':
                deo_saved_plan_reports(data,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
                st.session_state['_deo_allocations']=data['allocations']
                st.session_state['deo_plan_adopted']=data['plan_adopted']
                actuals=data.get('actuals')
                if actuals:
                    st.session_state['deo_daily_opening']=float(actuals['opening'])
                    st.session_state['_deo_daily_rows']=actuals['rows']
                    st.session_state['deo_actual_confirmed']=actuals['actuals_confirmed']
                    st.session_state['deo_opening_confirmed']=actuals['opening_confirmed']
                    st.session_state['deo_actual_source']=actuals['source_reference']
                    st.session_state['_deo_daily_rev']=st.session_state.get('_deo_daily_rev',0)+1
            else:
                deo_ecps_schedule(data['civ'],data['opening'],pd.DataFrame(data['plan']))
                st.session_state['deo_plan_adopted']=False
            for key,value in [('ecps_units',data['units']),('ecps_civ',float(data['civ'])),('ecps_opening',float(data['opening']))]:st.session_state[key]=value
            st.session_state['_ecps_rows']=data['plan'];st.session_state['_ecps_rev']=st.session_state.get('_ecps_rev',0)+1
        except (ValueError,KeyError,TypeError) as exc:st.error('Cannot restore DEO inputs: '+str(exc))
    units=st.selectbox('DEO inventory and allocation units',['Unconfirmed','Mcf','Dth'],index=2,key='ecps_units')
    civ=st.number_input('Utility-confirmed end-of-summer cumulative injection volume',min_value=0.,value=0.,key='ecps_civ')
    opening=st.number_input('November 1 opening inventory — projected Dth',min_value=0.,value=312430.,key='ecps_opening')
    st.caption('Opening forecast: reported October 26 balance 304,780 Dth + five days at 1,530 Dth/day = 312,430 Dth, assuming no withdrawals or adjustments. This supersedes the 319,826 estimate. The posted November zero activity ledger is not treated as actual zero opening inventory. CIV remains separate from capacity and opening inventory.')
    posted=pd.DataFrame([dict(effective_month='2026-11',storage_contract='12663',rate_schedule='ECPS-S',units='Dth',capacity=332237,inventory_above=166119,inventory_at_or_below=332237,withdrawal_dth_day=6401),dict(effective_month='2026-11',storage_contract='12663',rate_schedule='ECPS-S',units='Dth',capacity=332237,inventory_above=99671,inventory_at_or_below=166119,withdrawal_dth_day=4801),dict(effective_month='2026-11',storage_contract='12663',rate_schedule='ECPS-S',units='Dth',capacity=332237,inventory_above=0,inventory_at_or_below=99671,withdrawal_dth_day=3841)])
    show(posted,'storage_deo_posted_november_ratchets.csv')
    st.caption('Source: utility storage summary run October 26, 2026, as of November 30, 2026. Posted rounded thresholds are retained exactly. These November entitlements are not automatically extended to December–March. The generic CIV screen does not override utility-posted entitlements; reconcile their denominator before nominations.')
    profile_withdrawals={'2026-11-01':10289.,'2026-12-01':115583.,'2027-01-01':85646.,'2027-02-01':62178.,'2027-03-01':41824.}
    defaults=[dict(month=month,injections=0.,withdrawals=volume,utility_adjustment=0.,assigned_storage_demand=6401. if month=='2026-11-01' else None) for month,volume in profile_withdrawals.items()]
    if st.button('Apply supplied DEO monthly withdrawal profile'):
        current=st.session_state.get('_deo_current_plan',st.session_state.get('_ecps_rows',defaults))
        st.session_state['_ecps_rows']=[dict(r,withdrawals=profile_withdrawals[r['month']]) for r in current]
        st.session_state['_ecps_rev']=st.session_state.get('_ecps_rev',0)+1
        st.session_state['deo_plan_adopted']=False
    st.caption('Enter gross injections and withdrawals separately. Signed utility adjustments capture transfers or true-ups; they do not reset CIV. Withdrawal defaults use your supplied monthly profile: November 10,289; December 115,583; January 85,646; February 62,178; March 41,824, totaling 315,520. Negative source volumes are entered as positive withdrawals. Exact volumes govern; supplied percentages are rounded. These are planned movements, not actual activity. Use Apply supplied DEO monthly withdrawal profile to replace previously entered withdrawal values while retaining injections, adjustments and demands.')
    rows=st.data_editor(pd.DataFrame(st.session_state.get('_ecps_rows',defaults)),disabled=['month'],column_config={c:st.column_config.NumberColumn(c) for c in ['injections','withdrawals','utility_adjustment','assigned_storage_demand']},hide_index=True,key='ecps_plan_'+str(st.session_state.get('_ecps_rev',0)))
    st.session_state['_deo_current_plan']=json.loads(rows.to_json(orient='records'))
    projected_end=opening+pd.to_numeric(rows.injections).sum()-pd.to_numeric(rows.withdrawals).sum()+pd.to_numeric(rows.utility_adjustment).sum()
    if projected_end<0:st.warning(f'Planned withdrawals exceed opening inventory plus planned additions by {-projected_end:,.0f} {units}. Reconcile the opening inventory or movements; the model does not create balancing injections or adjustments.')
    adopted=st.checkbox('DEO winter movement plan reviewed and adopted',value=False,key='deo_plan_adopted')
    allocation_defaults=[dict(month=r['month'],assigned_storage_demand=6401. if i==0 else None,confirmed=i==0,
        source_reference='Utility posting run October 26, 2026; November 2026, contract 12663' if i==0 else '') for i,r in enumerate(defaults)]
    with st.expander('DEO monthly allocation confirmations and update tasks',expanded=True):
        st.caption('November demand is supported by the supplied posting. Confirm later months only after reviewing their utility allocation; enter the matching demand in the plan above. This confirmation does not replace posted daily ratchets, OFOs or reinjection permissions.')
        allocations=st.data_editor(pd.DataFrame(st.session_state.get('_deo_allocations',allocation_defaults)),disabled=['month'],hide_index=True,key='deo_allocations_'+str(st.session_state.get('_ecps_rev',0)))
        if units=='Dth':
            try:
                operations=dict(schema='deo-ecps-2',units=units,civ=civ,opening=opening,plan_adopted=adopted,
                    plan=json.loads(rows.to_json(orient='records')),allocations=json.loads(allocations.to_json(orient='records')))
                reports,checks=deo_saved_plan_reports(operations,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
                st.session_state.setdefault('_scheduled_storage_inputs',{})['deo_operations']=operations
                st.session_state.setdefault('_run_reports',{}).update(reports)
                for name in ['storage_deo_release_update_tasks.csv','storage_deo_inventory_forecast.csv','storage_deo_plan_summary.csv']:
                    show(reports[name],name)
                if any(checks.values()):st.warning('DEO inputs need review. The summary separates missing CIV, an unadopted plan, missing allocations, monthly plan exceptions and inventory shortfalls.')
                st.download_button('Download DEO operations inputs',json.dumps(operations,indent=2),'riverbend_deo_operations_inputs.json','application/json')
            except (ValueError,KeyError,TypeError) as exc:
                st.session_state.setdefault('_scheduled_storage_inputs',{}).pop('deo_operations',None)
                st.error('Cannot prepare DEO operations: '+str(exc))
        else:st.session_state.setdefault('_scheduled_storage_inputs',{}).pop('deo_operations',None)
    bands=pd.DataFrame([dict(target_date=t,minimum_fraction=lo,maximum_fraction=hi,denominator='End-of-summer CIV') for t,lo,hi in [('2026-12-01',.92,.95),('2027-01-01',.55,.65),('2027-02-01',.30,.35),('2027-03-01',.12,.17),('2027-03-31',0.,0.)]])
    show(bands,'storage_deo_ecps_inventory_bands.csv')
    st.caption('Summer §13.6: daily firm injection capacity = assigned Capacity/180 at inventory ≤50% of assigned Capacity; Capacity/214 above 50%. Summer inventory bands use assigned Capacity (§13.8). March 31 requires withdrawal of CIV and winter reinjections (§13.14).')
    st.session_state['_run_reports']['storage_deo_service_selection.csv']=pd.DataFrame([dict(service='Enbridge Gas Ohio ECPS on-system',basis='User confirms utility allocation; account quantities pending utility statement',units=units)])
    if units=='Dth':render_deo_daily(opening)
    else:st.info('The November posted daily tracker requires Dth units.')
    if units=='Unconfirmed' or civ<=0:
        st.info('CIV-based monthly results are withheld until units and CIV are entered. The separate posted daily screen does not require CIV. Use the monthly asset-allocation sheet and end-of-summer storage statement; an interstate capacity release is not required for this on-system account.')
        return
    try:
        result=deo_ecps_schedule(civ,opening,rows);result.insert(0,'units',units)
        show(result,'storage_deo_ecps_schedule.csv')
        show(rows,'storage_deo_ecps_plan_inputs.csv')
        if units!='Dth':
            show(pd.DataFrame([dict(civ=civ,opening_inventory=opening,units=units,source=DEO_ECPS_SOURCE)]),'storage_deo_ecps_account_inputs.csv')
        export=dict(schema='deo-ecps-1',units=units,civ=civ,opening=opening,plan=json.loads(rows.to_json(orient='records')))
        st.download_button('Download DEO ECPS inputs',json.dumps(export,indent=2),'riverbend_deo_ecps_inputs.json','application/json')
    except (ValueError,TypeError) as exc:st.error('DEO inputs need correction: '+str(exc))


def render_columbia_tariff_scope():
    st.subheader('Columbia tariff scope')
    st.markdown('[Columbia FSS §§3–4]('+TCO_FSS_SOURCE+')')
    st.caption('Reviewed October 26, 2026. Columbia FSS VI.9 version 7.0.0, effective April 1, 2025. These are quantity screens, not nomination approval. Live OFOs, waivers, contract entitlements, transport and retainage must also be checked.')
    st.write('Columbia rules apply separately to each account. Summer limits are included in the rule table; the current forecast remains November–March. February 1 inventory is capped at 65% of SCQ; it is not a minimum inventory requirement.')
    rules=pd.DataFrame([dict(month=m,**tco_fss_limits(m),source='Columbia FSS §§3(c)–(d), 4(b)–(d)') for m in range(1,13)])
    show(rules,'storage_tariff_rules_tco.csv')
    st.caption('Daily ratchets: 100% of MDSQ at 30% inventory and above; 80% at 20% to below 30%; 65% at 10% to below 20%; 50% below 10%. A preceding winter month’s excess reduces the next month’s maximum. Approved exceptions are not assumed or automatically netted between accounts.')


def render_deo_storage_tracker():
    st.subheader('DEO storage tracker')
    st.caption('DEO inventory, allocations, costs and update tasks. Shared pricing controls and the combined profile download are in Storage pricing & summary.')
    st.markdown('[Enbridge Ohio ECPS]('+DEO_ECPS_SOURCE+') · [East Ohio FSS §5.1]('+DEO_FSS_SOURCE+') · [EGTS GSS §§7–8]('+EGTS_GSS_SOURCE+')')
    service=st.selectbox('DEO governing service (confirm against your agreement)', ['Enbridge Ohio ECPS utility-allocated','Unconfirmed','East Ohio / Enbridge Ohio FSS seasonal','EGTS GSS interstate'],key='deo_tariff_service_v310')
    if service=='Enbridge Ohio ECPS utility-allocated':
        render_deo_ecps()
        return
    st.info('Other service references below do not replace the DEO ECPS utility-allocated working profile.')
    if service=='Unconfirmed':
        st.write('Confirm the operator and rate schedule from the storage agreement or capacity-release award before enabling DEO contract checks.')
        return
    st.caption('This is a separate reference calculator. It does not change demand, booked inventory or Columbia schedules. Enter capacity from the agreement, not the summer costing schedule.')
    if service=='East Ohio / Enbridge Ohio FSS seasonal':
        cap=st.number_input('DEO contractual maximum storage capacity (Mcf)',min_value=0.,value=0.,key='deo_fss_capacity')
        bands=pd.DataFrame([dict(month=m,end_month_min_fraction=deo_fss_band(m)[0],end_month_max_fraction=deo_fss_band(m)[1]) for m in range(1,13)])
        if cap>0:
            bands['minimum_mcf']=bands.end_month_min_fraction*cap;bands['maximum_mcf']=bands.end_month_max_fraction*cap
        show(bands,'storage_deo_fss_reference.csv')
        st.caption('Seasonal FSS §5.1 requires full withdrawal by March 31. Daily injection /214 and withdrawal /151 restrictions apply when East Ohio invokes daily management. In/Out service under §5.2 uses contract-specific limits and is not covered by these seasonal bands. Storage OFOs and waivers can alter requirements.')
    else:
        cap=st.number_input('EGTS contractual storage capacity (Dth)',min_value=0.,value=0.,key='egts_capacity')
        demand=st.number_input('EGTS contractual storage demand (Dth/day)',min_value=0.,value=0.,key='egts_demand')
        inventory=st.number_input('EGTS eligible inventory for entitlement (Dth)',min_value=0.,value=0.,key='egts_inventory')
        month=st.selectbox('EGTS winter month',[11,12,1,2,3],key='egts_month')
        interruptible=st.checkbox('Agreement expressly permits interruption (GSS §8.4.B exception)',key='egts_interruptible')
        if cap>0 and demand>0:
            try:show(pd.DataFrame([egts_gss_capability(cap,demand,inventory,month,interruptible)]),'storage_egts_gss_reference.csv')
            except ValueError as exc:st.error(str(exc))
        st.caption('GSS monthly average limit is separate from daily entitlement. Minimum turnover runs November 1–April 15 and begins with opening inventory less 35% of capacity, adjusted for releases and inventory transfers. Restricted winter injection credits, transport and OFOs need account-specific review. This calculator does not certify monthly turnover or daily nominations.')
    st.session_state['_run_reports']['storage_deo_service_selection.csv']=pd.DataFrame([dict(service=service,source='User-selected reference profile; executed agreement not independently verified')])


def storage_account_conversion_defaults(accounts):
    # Adopt the user-confirmed 1:1 default for missing factors, preserving explicit overrides.
    result=[]
    for account in accounts:
        account=dict(account)
        if account.get('dth_per_mcf') is None or pd.isna(account.get('dth_per_mcf')):
            account['dth_per_mcf']=1.
        known={'Columbia 1':('353571',210915.,3789.,3864.,'2026-10-01','2026-10-31','Provisional: October release carried through March 2027'),
               'Columbia 2':('341421',24402.,509.,519.,'2026-04-01','2027-03-31','Release covers winter through March 2027')}
        if account.get('account') in known and not account.get('release_defaults_version'):
            contract,capacity,mdsq,mdwq,start,end,status=known[account['account']]
            factor=storage_number(account['dth_per_mcf'],'Dth/Mcf',0)
            if factor<=0:raise ValueError('Positive Dth/Mcf is required for release conversion.')
            if account.get('mdsq_mcf_day') is None or pd.isna(account.get('mdsq_mcf_day')):
                account['mdsq_mcf_day']=mdsq/factor
            if account['account']=='Columbia 1' and account.get('capacity_mcf')==210804.:
                account['capacity_mcf']=capacity/factor
            account.update(release_contract=contract,release_capacity_dth=capacity,release_mdsq_dth_day=mdsq,
                release_displayed_mdwq_dth_day=mdwq,release_effective_start=start,release_effective_end=end,
                entitlement_basis=status,release_defaults_version='2026-10-26')
        result.append(account)
    return result


def storage_defaults():
    """Winter-priority forecast defaults; not actual movements or nomination approval."""
    months=['2026-11-01','2026-12-01','2027-01-01','2027-02-01','2027-03-01']
    accounts=[dict(account='Columbia 1',source_sheet='Winter storage',capacity_mcf=210804.,opening_mcf=196589.,mdsq_mcf_day=None,dth_per_mcf=1.),
              dict(account='Columbia 2',source_sheet='Winter storage 2',capacity_mcf=24402.,opening_mcf=16000.,mdsq_mcf_day=None,dth_per_mcf=1.)]
    rows=[]
    for account,loads in zip(accounts,[[800,1300,1800,1700,800],[0,0,400,250,0]]):
        for i,month in enumerate(months):
            rows.append(dict(account=account['account'],month=month,load_withdrawal_mcf_day=loads[i],replenish=i<2,target_fraction=[1,1,.39,.08,0][i],
                injection_limit_fraction=[.05,.1,.1,.1,.1][i],injection_divisor=[30,30,25,25,25][i],
                min_net_withdrawal_fraction=[0,0,0,.1,.1][i],max_net_withdrawal_fraction=[.4,.4,.4,.3,.2][i],inventory_ceiling_fraction=[1,1,.65,1,.25][i],
                injection_override_mcf=None,withdrawal_override_mcf=None))
    accounts=storage_account_conversion_defaults(accounts)
    proposal,_,_=priority_storage_plan(pd.DataFrame(accounts),pd.DataFrame(rows))
    rows=json.loads(proposal.to_json(orient='records'))
    return dict(schema=1,accounts=accounts,plan=rows,actuals=[],summer_prices=[dict(month=f'2026-{m:02}-01',weight=w,spot_usd_dth=None) for m,w in zip(range(4,11),[.085,.165,.14,.15,.14,.14,.18])])


def storage_number(value,label,minimum=None,optional=False):
    if value is None or pd.isna(value):
        if optional:return None
        raise ValueError(label+' is required.')
    try:value=float(value)
    except (TypeError,ValueError):raise ValueError(label+' must be numeric.')
    if not math.isfinite(value) or (minimum is not None and value<minimum):raise ValueError(label+' is invalid.')
    return value


def calculate_storage(accounts,plan,actuals):
    if accounts.account.duplicated().any() or accounts.account.isna().any():raise ValueError('Account names must be unique and populated.')
    if set(plan.account)!=set(accounts.account):raise ValueError('Every account needs a plan; unknown accounts are not allowed.')
    if plan.duplicated(['account','month']).any():raise ValueError('Duplicate account/month in storage plan.')
    results=[];daily=[]
    for _,a in accounts.iterrows():
        cap=storage_number(a.capacity_mcf,'Capacity',0)
        if cap==0:raise ValueError('Capacity must be positive.')
        opening=storage_number(a.opening_mcf,'Opening inventory',0)
        mdsq=storage_number(a.mdsq_mcf_day,'MDSQ',0,True)
        factor=storage_number(a.dth_per_mcf,'Dth/Mcf',0,True)
        if factor==0:raise ValueError('Dth/Mcf must be positive or blank.')
        rows=plan[plan.account.eq(a.account)].sort_values('month')
        dates=pd.to_datetime(rows.month,errors='raise')
        if any(d.day!=1 for d in dates) or list(dates)!=list(pd.date_range(dates.iloc[0],periods=len(rows),freq='MS')):raise ValueError('Plan months must be consecutive month starts.')
        balance=opening
        prior_excess=0.
        for (_,r),month in zip(rows.iterrows(),dates):
            n=month.days_in_month;begin=balance
            load=storage_number(r.load_withdrawal_mcf_day,'Load balancing withdrawal',0)*n
            target=storage_number(r.target_fraction,'Target fraction',0)
            rule=tco_fss_limits(month.month)
            inj_frac=rule['injection_limit_fraction'];divisor=rule['injection_divisor']
            minimum=rule['min_net_withdrawal_fraction']*cap
            base_maximum=None if rule['max_net_withdrawal_fraction'] is None else rule['max_net_withdrawal_fraction']*cap
            reduction=prior_excess if month.month in [12,1,2,3] else 0.
            maximum=None if base_maximum is None else max(0.,base_maximum-reduction)
            ceiling=rule['inventory_ceiling_fraction']*cap
            if not isinstance(r.replenish,(bool,type(pd.Series([True]).iloc[0]))):raise ValueError('Replenish must be true or false.')
            withdrawal=load+(0 if r.replenish else max(0,begin-load-cap*target))
            override=storage_number(r.withdrawal_override_mcf,'Withdrawal override',0,True)
            if override is not None:withdrawal=override
            required=max(0,cap-begin+withdrawal) if r.replenish else 0
            injmax=cap*inj_frac;dailymax=injmax/divisor
            injection=min(required,injmax,dailymax*n,max(0,cap-begin+withdrawal))
            override=storage_number(r.injection_override_mcf,'Injection override',0,True)
            if override is not None:injection=override
            balance=begin+injection-withdrawal;net=withdrawal-injection
            errors=[]
            if balance < -1e-7:errors.append('Negative ending inventory')
            if balance>ceiling+1e-7:errors.append('Ending inventory exceeds ceiling')
            if injection>injmax+1e-7:errors.append('Monthly injection maximum exceeded')
            if injection/n>dailymax+1e-7:errors.append('Average injection exceeds daily limit')
            if minimum>0 and net<minimum-1e-7:errors.append('Below minimum net withdrawal')
            if maximum is not None and net>maximum+1e-7:errors.append('Above maximum net withdrawal')
            prior_excess=0. if maximum is None else max(0.,net-maximum)
            ratchet=tco_ratchet(balance,cap)
            daily_limit=None if mdsq is None else mdsq*ratchet
            daily_review='MDSQ missing' if mdsq is None else ('Average exceeds ending-inventory limit' if withdrawal/n>daily_limit else 'Average screen only; daily schedule required')
            results.append(dict(account=a.account,month=month.date().isoformat(),beginning_mcf=begin,injections_mcf=injection,withdrawals_mcf=withdrawal,net_withdrawals_mcf=net,ending_mcf=balance,
                average_withdrawal_mcf_day=withdrawal/n,monthly_base_max_net_withdrawal_mcf=base_maximum,prior_month_excess_reduction_mcf=reduction,monthly_max_net_withdrawal_mcf=maximum,ending_ratchet=ratchet,ending_daily_limit_mcf=daily_limit,
                monthly_review='; '.join(errors) if errors else 'Within TCO FSS monthly quantity limits',daily_review=daily_review,
                net_withdrawals_dth=None if factor is None else net*factor))
        if actuals.empty:continue
        movements=actuals[actuals.account.eq(a.account)].copy()
        if movements.empty:continue
        movements['date']=pd.to_datetime(movements.date,errors='raise').dt.normalize()
        start=dates.iloc[0];end=dates.iloc[-1]+pd.offsets.MonthEnd(0)
        if ((movements.date<start)|(movements.date>end)).any():raise ValueError('Actual movement date is outside the plan period.')
        if movements.duplicated(['date']).any():raise ValueError('Enter one daily aggregate per account/date.')
        balance=opening
        for _,r in movements.sort_values('date').iterrows():
            inj=storage_number(r.injection_mcf,'Actual injection',0);wd=storage_number(r.withdrawal_mcf,'Actual withdrawal',0)
            begin=balance;balance+=inj-wd
            statement=storage_number(r.statement_inventory_mcf,'Statement inventory',0,True)
            daily.append(dict(account=a.account,date=r.date.date().isoformat(),beginning_mcf=begin,injection_mcf=inj,withdrawal_mcf=wd,ending_mcf=balance,
                statement_inventory_mcf=statement,statement_variance_mcf=None if statement is None else balance-statement,
                status='Inventory outside 0–capacity' if balance<0 or balance>cap else 'Recorded movements only'))
    if not actuals.empty and not set(actuals.account).issubset(set(accounts.account)):raise ValueError('Unknown account in actual movements.')
    return pd.DataFrame(results),pd.DataFrame(daily)


def priority_storage_plan(accounts,plan):
    """Monthly proposal with fixed injections; daily entitlement checks remain separate."""
    current,_=calculate_storage(accounts,plan,pd.DataFrame())
    proposed=plan.copy();changes=[]
    months=list(pd.date_range('2026-11-01',periods=5,freq='MS'))
    for _,account in accounts.iterrows():
        rows=current.loc[current.account.eq(account.account)].sort_values('month')
        if list(pd.to_datetime(rows.month))!=months:raise ValueError('Priority planning requires November 2026–March 2027 for every account.')
        cap=float(account.capacity_mcf);opening=float(account.opening_mcf)
        injections=list(rows.injections_mcf);lower=[];upper=[];ceilings=[]
        for m,inj in zip(months,injections):
            rule=tco_fss_limits(m.month)
            if inj>min(cap*rule['injection_limit_fraction'],cap*rule['injection_limit_fraction']/rule['injection_divisor']*m.days_in_month)+1e-7:
                raise ValueError('Existing injections exceed monthly or average daily limits; correct injections first.')
            lower.append(inj+cap*rule['min_net_withdrawal_fraction'] if rule['min_net_withdrawal_fraction']>0 else 0.)
            upper.append(inj+cap*rule['max_net_withdrawal_fraction'])
            ceilings.append(cap*rule['inventory_ceiling_fraction'])
        def feasible(lo,hi):
            a=b=opening
            for i in range(5):
                a=max(0.,a+injections[i]-hi[i]);b=min(ceilings[i],b+injections[i]-lo[i])
                if a>b+1e-9:return False
            return True
        if not feasible(lower,upper):raise ValueError('No monthly schedule can meet the limits with these opening inventories and fixed injections.')
        # Adjacent January/February are allocated first; January breaks ties.
        for i in [2,3,1,0,4]:
            left,right=lower[i],upper[i]
            for _ in range(60):
                mid=(left+right)/2;trial=lower.copy();trial[i]=mid
                if feasible(trial,upper):left=mid
                else:right=mid
            lower[i]=upper[i]=left
        for i,(_,r) in enumerate(rows.iterrows()):
            lower[i]=round(lower[i],6)
            key=proposed.account.eq(account.account)&pd.to_datetime(proposed.month).eq(months[i])
            proposed.loc[key,'injection_override_mcf']=injections[i]
            proposed.loc[key,'withdrawal_override_mcf']=lower[i]
            changes.append(dict(account=account.account,month=r.month,current_withdrawal_mcf=r.withdrawals_mcf,
                proposed_withdrawal_mcf=lower[i],withdrawal_change_mcf=lower[i]-r.withdrawals_mcf,
                fixed_injection_mcf=injections[i]))
    result,_=calculate_storage(accounts,proposed,pd.DataFrame())
    if result.monthly_review.ne('Within TCO FSS monthly quantity limits').any():
        raise ValueError('Proposed schedule did not pass monthly checks; proposal withheld.')
    return proposed,pd.DataFrame(changes),result


def columbia_cost_basis(basis):
    quantity=storage_number(basis.get('quantity',235595.),'Columbia summer cost quantity',0)
    units=basis.get('units','Mcf')
    if units not in ['Unconfirmed','Dth','Mcf']:raise ValueError('Select valid Columbia cost schedule units.')
    factor=basis.get('dth_per_mcf')
    # Owner-directed default also migrates the old unconfirmed/zero placeholder.
    if units=='Unconfirmed':
        units='Mcf'
        if factor is None or pd.isna(factor) or factor==0:factor=1.
    if factor is None or pd.isna(factor):factor=1.
    factor=storage_number(factor,'Columbia cost Dth/Mcf factor',0)
    net_dth=quantity if units=='Dth' else quantity*factor if units=='Mcf' and factor>0 else None
    return quantity,units,factor,net_dth


def columbia_tariff_cost(total,prices,rates):
    total=storage_number(total,'Net credited summer Dth',0)
    months=list(pd.date_range('2026-04-01',periods=7,freq='MS'))
    if list(pd.to_datetime(prices.month))!=months or list(pd.to_datetime(rates.month))!=months:raise ValueError('Cost inputs require April–October 2026 in order.')
    out=[]
    for (_,p),(_,r),weight in zip(prices.iterrows(),rates.iterrows(),[.085,.165,.14,.15,.14,.14,.18]):
        vals={c:storage_number(r[c],c,0,True) for c in ['transport_fuel_fraction','storage_loss_fraction','transport_usd_dth','injection_usd_dth','aca_usd_dth']}
        for c in ['transport_fuel_fraction','storage_loss_fraction']:
            if vals[c] is not None and vals[c]>=1:raise ValueError('Retainage fractions must be below 1.')
        price=storage_number(p.spot_usd_dth,'Spot price',optional=True);net=total*weight
        loss=vals['storage_loss_fraction'];fuel=vals['transport_fuel_fraction']
        inlet=None if loss is None else net/(1-loss)
        gross=None if inlet is None or fuel is None else inlet/(1-fuel)
        gas=None if gross is None or price is None else gross*price
        trans=None if inlet is None or vals['transport_usd_dth'] is None else inlet*vals['transport_usd_dth']
        inj=None if inlet is None or vals['injection_usd_dth'] is None else inlet*vals['injection_usd_dth']
        aca=None if inlet is None or vals['aca_usd_dth'] is None else inlet*vals['aca_usd_dth']
        complete=all(x is not None for x in [gas,trans,inj,aca])
        out.append(dict(injection_month=p.month,pricing_basis=p.get('pricing_basis','Manual monthly price'),net_credited_dth=net,storage_inlet_dth=inlet,gross_purchase_dth=gross,
            transport_fuel_dth=None if gross is None else gross-inlet,storage_loss_dth=None if inlet is None else inlet-net,
            gas_purchase_cost_usd=gas,transport_charge_usd=trans,injection_charge_usd=inj,aca_charge_usd=aca,
            variable_injection_expense_usd=gas+trans+inj+aca if complete else None,
            status='Variable-cost scenario; fixed charges excluded' if complete else 'Incomplete price or rate inputs'))
    return pd.DataFrame(out)


def columbia_outbound_cost(rows):
    out=[]
    if list(pd.to_datetime(rows.month))!=list(pd.date_range('2026-11-01',periods=5,freq='MS')):raise ValueError('Outbound inputs require November 2026–March 2027 in order.')
    for _,r in rows.iterrows():
        v={c:storage_number(r[c],c,0,True) for c in ['withdrawal_dth','transport_fuel_fraction','transport_usd_dth','withdrawal_usd_dth','aca_usd_dth']}
        q=v['withdrawal_dth'];fuel=v['transport_fuel_fraction']
        if fuel is not None and fuel>=1:raise ValueError('Outbound retainage must be below 1.')
        delivered=None if q is None or fuel is None else q*(1-fuel)
        wd=None if q is None or v['withdrawal_usd_dth'] is None else q*v['withdrawal_usd_dth']
        trans=None if delivered is None or v['transport_usd_dth'] is None else delivered*v['transport_usd_dth']
        aca=None if delivered is None or v['aca_usd_dth'] is None else delivered*v['aca_usd_dth']
        out.append(dict(month=r.month,withdrawal_dth=q,delivered_dth=delivered,transport_fuel_dth=None if delivered is None else q-delivered,
            withdrawal_charge_usd=wd,transport_charge_usd=trans,aca_charge_usd=aca,
            variable_outbound_charges_usd=wd+trans+aca if all(x is not None for x in [wd,trans,aca]) else None))
    return pd.DataFrame(out)


def columbia_tariff_defaults():
    return dict(inbound=[dict(month=f'2026-{m:02}-01',transport_fuel_fraction=.01943,storage_loss_fraction=.00483,
        transport_usd_dth=.0182,injection_usd_dth=.0242,aca_usd_dth=.0015 if m<10 else .0013) for m in range(4,11)],
        outbound=[dict(month=m,withdrawal_dth=None,transport_fuel_fraction=.01943,transport_usd_dth=.0182,
        withdrawal_usd_dth=.0242,aca_usd_dth=0.) for m in ['2026-11-01','2026-12-01','2027-01-01','2027-02-01','2027-03-01']])


def columbia_tariff_profile(profile):
    defaults=columbia_tariff_defaults()
    if profile is None:profile=defaults
    result=dict(profile)
    result['inbound']=[dict(row) for row in profile['inbound']]
    result['outbound']=[dict(row) for row in profile['outbound']]
    if profile.get('rate_defaults_version') is None:
        known={row['month']:row for row in defaults['inbound']}
        for row in result['inbound']:
            if row['month'] not in known:continue
            for field in ['transport_usd_dth','aca_usd_dth']:
                if row.get(field) is None or pd.isna(row.get(field)):
                    row[field]=known[row['month']][field]
    result['rate_defaults_version']='2026-10-26'
    return result


def columbia_planned_withdrawals(accounts,forecast):
    """Convert each account separately; never net withdrawals against reinjections."""
    if accounts.account.duplicated().any():raise ValueError('Duplicate Columbia accounts.')
    if forecast.duplicated(['account','month']).any():raise ValueError('Duplicate account/month in withdrawal forecast.')
    expected=set(accounts.account);detail=[];monthly=[]
    if not set(forecast.account).issubset(expected):raise ValueError('Unknown forecast account.')
    for month in pd.date_range('2026-11-01',periods=5,freq='MS'):
        key=month.date().isoformat();part=forecast.loc[pd.to_datetime(forecast.month)==month]
        values=[];issues=[]
        for _,a in accounts.iterrows():
            rows=part.loc[part.account==a.account]
            factor=storage_number(a.dth_per_mcf,'Account Dth/Mcf factor',0,True)
            wd=None if rows.empty else storage_number(rows.iloc[0].withdrawals_mcf,'Planned withdrawal Mcf',0)
            dth=None if wd is None or factor is None or factor<=0 else wd*factor
            review='Missing account/month' if rows.empty else str(rows.iloc[0].monthly_review)
            if dth is None:issues.append(str(a.account)+': missing withdrawal or positive Dth/Mcf factor')
            if not rows.empty and review!='Within TCO FSS monthly quantity limits':issues.append(str(a.account)+': '+review)
            detail.append(dict(account=a.account,month=key,withdrawals_mcf=wd,dth_per_mcf=factor,withdrawal_dth=dth,plan_review=review))
            values.append(dth)
        monthly.append(dict(month=key,withdrawal_dth=sum(values) if values and all(v is not None for v in values) else None,
            plan_review='; '.join(issues) if issues else 'Converted plan; daily billing quantities still require review'))
    return pd.DataFrame(monthly),pd.DataFrame(detail)


def render_columbia_injection_cost(prices,basis,rev,profile=None,accounts=None,plan=None):
    st.markdown('**Columbia FTS / FSS / SST cost components**')
    st.markdown('[Columbia tariff: Parts V.1, V.8, V.9, V.17; FSS §6; SST §4; GTC §34](https://hostedtariffs.com/tco/pdf/tariff.pdf)')
    st.warning('The former Columbia 4.7% assumption has been removed. These are tariff-reference scenarios, not confirmation of your negotiated or utility-allocated billing terms.')
    st.caption('Retainage effective April 1, 2026: transportation 1.943%; standard storage loss 0.483%; separate Ohio storage loss 0.848%. Standard FSS is the reference below; do not select the Ohio category solely from the customer location. FSS injection/withdrawal: $0.0242/Dth effective April 1. FTS/SST commodity: $0.0182 per delivered Dth for April–October 2026. Inbound ACA reference: $0.0015/Dth April–September and $0.0013/Dth from October 1. Apply only where the transportation route is ACA-chargeable; retain verified exemptions or negotiated overrides.')
    st.markdown('[April rate filing](https://hostedtariffs.com/tco/pending/RP24-1103-006.pdf) · [FERC ACA effective October 2025](https://www.ferc.gov/media/fy-2025-gas-annual-charges-correction-annual-charges-unit-charge) · [FERC ACA effective October 2026](https://www.ferc.gov/media/fy-2026-gas-annual-charges-correction-annual-charges-unit-charge)')
    st.caption('Older saved profiles with blank transportation/ACA inputs receive the published defaults once. Explicit zeros and other entered rates are preserved; review the route before treating reference costs as billed charges.')
    quantity,units,factor,_=columbia_cost_basis(basis)
    quantity=st.number_input('Columbia summer cost-only net injections',min_value=0.,value=float(quantity),key=f'coh_cost_quantity_{rev}')
    units=st.selectbox('Columbia summer schedule units',['Mcf','Dth'],index=['Mcf','Dth'].index(units),key=f'coh_cost_units_v326_{rev}')
    factor=st.number_input('Columbia summer cost conversion (Dth per Mcf)',min_value=0.,value=float(factor),key=f'coh_cost_factor_v326_{rev}',disabled=units!='Mcf')
    st.caption('Owner-directed default: Mcf with 1.000 Dth per Mcf. Positive custom conversions and explicit Dth selections are retained.')
    basis=dict(quantity=quantity,units=units,dth_per_mcf=factor);net=columbia_cost_basis(basis)[3]
    profile=columbia_tariff_profile(profile)
    st.caption('The 235,595 summer quantity is counted once across Columbia. It is treated as net credited inventory after storage loss. Cost-only calculations do not change the physical tracker. Rates are decimal fractions (0.01943 = 1.943%) and dollars/Dth, not cents.')
    inbound=st.data_editor(pd.DataFrame(profile['inbound']),disabled=['month'],hide_index=True,column_config={c:st.column_config.NumberColumn(c,format='%.5f') for c in ['transport_fuel_fraction','storage_loss_fraction','transport_usd_dth','injection_usd_dth','aca_usd_dth']},key=f'coh_inbound_rates_{rev}')
    st.caption('Inbound scenario: sequential transportation and storage retainage. Gross purchase = credited inventory / (1 − storage loss) / (1 − transportation retainage). Transportation commodity/ACA use delivered-to-storage Dth; injection fee uses daily net physical injections before storage loss. Summer scenario assumes no offsetting same-day withdrawals. Confirm these billing quantities against the statement. Enter a verified zero for an exemption; blank means unknown.')
    st.markdown('**Withdrawal and outbound transportation**')
    options=['Storage withdrawal plan','Manual volumes']
    fallback='Manual volumes' if any(r.get('withdrawal_dth') is not None for r in profile['outbound']) or accounts is None else 'Storage withdrawal plan'
    mode=st.selectbox('Columbia withdrawal volume source',options,index=options.index(profile.get('withdrawal_source',fallback)),key=f'coh_withdrawal_source_{rev}')
    outbound=st.data_editor(pd.DataFrame(profile['outbound']).drop(columns=['withdrawal_dth']) if mode==options[0] else pd.DataFrame(profile['outbound']),disabled=['month'],hide_index=True,column_config={c:st.column_config.NumberColumn(c) for c in ['withdrawal_dth','transport_fuel_fraction','transport_usd_dth','withdrawal_usd_dth','aca_usd_dth']},key=f'coh_outbound_rates_{rev}_{mode}')
    if mode==options[0]:
        outbound=outbound.copy();outbound['withdrawal_dth']=None
        try:
            if accounts is None or plan is None:raise ValueError('Storage account and plan inputs are unavailable.')
            planned,_=calculate_storage(accounts,plan,pd.DataFrame())
            linked,detail=columbia_planned_withdrawals(accounts,planned)
            outbound['withdrawal_dth']=linked.withdrawal_dth.to_numpy()
            show(detail,'storage_columbia_withdrawal_cost_account_bridge.csv')
            show(linked,'storage_columbia_withdrawal_cost_volume_bridge.csv')
            if linked.plan_review.str.contains('missing|Above|Below|Negative|Outside',case=False,regex=True).any():
                st.warning('The linked withdrawal plan has missing conversions or quantity-review flags. Costing does not make the plan compliant.')
        except (ValueError,TypeError,KeyError,AttributeError) as exc:st.error('Cannot link withdrawal plan: '+str(exc))
        st.caption('Calculated volumes below use each account’s gross planned withdrawals × its own Dth/Mcf factor. Reinjections are not subtracted. Linked quantities are shown in the reconciliation and cost tables; edit volumes through the storage plan. Monthly forecast withdrawals are a cost scenario; daily net billable quantities may differ when injections and withdrawals offset on the same day.')

    st.caption('Enter Dth withdrawn from inventory before outbound transport fuel, with withdrawal fees based on daily net withdrawals. Delivered gas = withdrawn gas × (1 − outbound retainage). SST primary receipt FSS Storage is exempt from ACA under GTC §34.1(b), hence zero ACA in this reference scenario. Secondary receipt routes can differ. Current rates are planning assumptions for future months and must be updated when effective rates change.')
    st.caption('Movements into TCO Pool can be exempt from transportation commodity charges and retainage; storage withdrawal fees can still apply. Do not charge both FTS and SST automatically to the same leg. Outbound fuel reduces delivered quantity; the gas purchase is not expensed again at withdrawal.')
    try:
        result=columbia_tariff_cost(0. if net is None else net,prices,inbound)
        show(inbound,'storage_columbia_inbound_rate_inputs.csv');show(outbound,'storage_columbia_outbound_inputs.csv')
        show(pd.DataFrame([basis]),'storage_columbia_cost_basis.csv')
        if net is None:st.info('Injection dollar results require confirmed schedule units and, for Mcf, a positive Dth/Mcf conversion.')
        else:
            show(result,'storage_columbia_injection_month_cost.csv')
            complete=result.variable_injection_expense_usd.notna().all()
            expense=float(result.variable_injection_expense_usd.sum()) if complete else None
            show(pd.DataFrame([dict(net_credited_dth=net,variable_injection_expense_usd=expense,
                status='Variable-cost scenario only' if complete else 'Total withheld: missing prices or rates')]),'storage_columbia_injection_cost_summary.csv')
            if complete:st.metric('Columbia summer variable injection expense',f'${expense:,.2f}')
            else:st.info('Full summer expense is withheld until all applicable monthly prices and rates are entered.')
        outbound_result=columbia_outbound_cost(outbound)
        show(outbound_result,'storage_columbia_outbound_cost.csv')
        monthly_flags=daily_flags=entitlement_flags=None
        try:
            if accounts is not None and plan is not None:
                reviewed,_=calculate_storage(accounts,plan,pd.DataFrame())
                screened=storage_daily_screen(accounts,reviewed,pd.DataFrame())
                monthly_flags=int(reviewed.monthly_review.ne('Within TCO FSS monthly quantity limits').sum())
                daily_flags=int(screened.review.ne('Within screened quantity limits; operator confirmation required').sum())
                entitlement_flags=int(storage_entitlement_report(accounts,reviewed).status.str.startswith('REVIEW').sum())
        except (ValueError,TypeError,KeyError,AttributeError):
            pass # Missing plan validation must remain visible in the summary.
        st.markdown('**Columbia outbound cost summary — November–March**')
        show(columbia_outbound_summary(outbound_result,mode,monthly_flags,daily_flags,entitlement_flags),'storage_columbia_outbound_cost_summary.csv')
    except (ValueError,TypeError,KeyError) as exc:st.error('Columbia cost inputs need correction: '+str(exc))
    st.caption('Fixed charges remain separate: published monthly FTS reservation $13.532 × contract demand; SST $13.332 × applicable seasonal demand; FSS reservation $3.865 × MDSQ plus capacity $0.0673 × SCQ. Contract discounts, allocation responsibility, taxes, penalties and other charges require confirmation. These variable-cost reports are not an all-in expense or business P&L. The old opening-cost estimate is withheld until the account cost allocation is resolved.')
    return basis,dict(inbound=json.loads(inbound.to_json(orient='records')),outbound=json.loads(outbound.to_json(orient='records')),withdrawal_source=mode,rate_defaults_version='2026-10-26')


def render_storage_tracker():
    st.subheader('Columbia storage tracker and withdrawal schedule')
    st.caption('Both Columbia inventory accounts are managed here. Shared pricing controls and the combined profile download are in Storage pricing & summary.')
    st.caption('Columbia FSS accounts are evaluated separately. Forecast inputs: Load4cast.xlsx, Winter storage and Winter storage 2. Quantity rules: TCO FSS §§3–4. Workbook units are Mcf equivalents; contract SCQ/MDSQ are in Dth. Confirm the conversion and equivalent entitlements. Forecasts are not approved nominations.')
    st.info('The reference accounts begin with 212,589 Mcf combined. The earlier Columbia book input remains 212,000 with units unconfirmed. No automatic demand offset or cross-account transfer is applied.')
    uploaded=st.file_uploader('Restore storage tracker inputs (JSON)',type=['json'],key='storage_restore')
    if uploaded is not None and st.button('Load storage tracker inputs'):
        try:
            data=json.loads(uploaded.getvalue())
            if data.get('schema')!=1:raise ValueError('Unsupported storage input version.')
            data['accounts']=storage_account_conversion_defaults(data['accounts'])
            calculate_storage(pd.DataFrame(data['accounts']),pd.DataFrame(data['plan']),pd.DataFrame(data['actuals']))
            columbia_cost_basis(data.get('columbia_cost_basis',{}))
            deo_injection_cost(0.,pd.DataFrame(data['summer_prices']))
            profile=columbia_tariff_profile(data.get('columbia_tariff_cost'))
            data['columbia_tariff_cost']=profile
            columbia_tariff_cost(0.,pd.DataFrame(data['summer_prices']),pd.DataFrame(profile['inbound']))
            columbia_outbound_cost(pd.DataFrame(profile['outbound']))
            if profile.get('withdrawal_source','Manual volumes') not in ['Storage withdrawal plan','Manual volumes']:raise ValueError('Invalid withdrawal volume source.')
            st.session_state['_storage_inputs']=data
            st.session_state['_storage_revision']=st.session_state.get('_storage_revision',0)+1
        except (ValueError,KeyError,TypeError,AttributeError) as exc:st.error('Cannot restore storage inputs: '+str(exc))
    data=st.session_state.get('_storage_inputs',storage_defaults());rev=st.session_state.get('_storage_revision',0)
    st.caption('Edits remain in this session. Download inputs to restore them later. Storage reports enter this interactive run archive; scheduled costs and the Columbia forecast use the separately exported version 2 cost profile; confirmed actual daily movements can also be included in the scheduled profile.')
    st.caption('User-selected planning default: 1.000 Dth per Mcf for each Columbia account. Factors remain editable; this is a modeling assumption, not a verified heat-content measurement. Existing explicit conversion overrides are preserved.')
    st.info('Winter planning: Columbia 1 uses October release 353571 (210,915 Dth capacity; MDSQ 3,789 Dth/day) provisionally through March. Columbia 2 release 341421 supplies 24,402 Dth and MDSQ 509 Dth/day through March 31. Posted MDWQ values 3,864 and 519 are retained separately and are not substituted for MDSQ. Replace planning quantities and source dates when each new release arrives. Opening inventory and injection-cost quantities are separate inputs.')
    accounts=st.data_editor(pd.DataFrame(storage_account_conversion_defaults(data['accounts'])),disabled=['account','source_sheet','release_defaults_version'],column_config={c:st.column_config.NumberColumn(c) for c in ['capacity_mcf','opening_mcf','mdsq_mcf_day','dth_per_mcf']},hide_index=True,key=f'storage_accounts_release_{rev}')
    with st.expander('Withdrawal plan and supplied limits',expanded=True):
        st.caption('New sessions default to the January–February priority plan. Saved plans and explicit edits take precedence. Blank overrides use the workbook strategy; zero means zero. Tariff limits are locked and recalculated from the month, including prior-month excess reductions. Overrides change proposed quantities, not tariff rights. Waivers and operational notices are not assumed.')
        plan=st.data_editor(pd.DataFrame(data['plan']),disabled=['account','month','injection_limit_fraction','injection_divisor','min_net_withdrawal_fraction','max_net_withdrawal_fraction','inventory_ceiling_fraction'],column_config={c:st.column_config.NumberColumn(c) for c in ['injection_override_mcf','withdrawal_override_mcf']},hide_index=True,key=f'storage_plan_{rev}')
    columns=['account','date','injection_mcf','withdrawal_mcf','statement_inventory_mcf']
    with st.expander('Actual daily movements and statement reconciliation'):
        st.caption('Enter one daily total per account/date, with both injection and withdrawal quantities (zero when none). Injections are net credited inventory. Blank statement inventory means no statement comparison. Missing dates are not certified zero activity.')
        actuals=st.data_editor(pd.DataFrame(data['actuals'],columns=columns),num_rows='dynamic',column_config={c:st.column_config.NumberColumn(c) for c in ['injection_mcf','withdrawal_mcf','statement_inventory_mcf']},hide_index=True,key=f'storage_actuals_{rev}')
    actual_meta=data.get('actual_confirmation',{})
    confirm_actuals=st.checkbox('Include entered Columbia rows as confirmed actual activity',value=actual_meta.get('confirmed',False),key=f'columbia_actual_confirm_{rev}')
    confirm_openings=st.checkbox('Both Columbia opening inventories confirmed against records',value=actual_meta.get('opening_confirmed',False),key=f'columbia_open_confirm_{rev}')
    actual_source=st.text_input('Columbia actual activity source reference',value=actual_meta.get('source_reference',''),key=f'columbia_actual_source_{rev}')
    st.caption('Leave actual activity unconfirmed before winter. Confirmed exports require completed delivery days and a source reference. Re-export the scheduled profile after each update.')
    with st.expander('Columbia synthetic storage cost'):
        st.caption('Enter injection-month TCO average daily spot prices in $/Dth. Columbia uses separate tariff components below; DEO 4.7% is not applied.')
        auto_price=st.session_state.get('_storage_auto_enabled',False)
        prices=st.data_editor(storage_price_input_frame('Columbia',pd.DataFrame(data['summer_prices'])),disabled=True if auto_price else ['month','weight'],column_config={'spot_usd_dth':st.column_config.NumberColumn('spot_usd_dth')},hide_index=True,key=f'storage_cost_{rev}_'+('auto' if auto_price else 'manual'))
        prices=prices.copy();prices['weight']=[.085,.165,.14,.15,.14,.14,.18]
        cost_basis,cost_profile=render_columbia_injection_cost(prices,data.get('columbia_cost_basis',{}),rev,data.get('columbia_tariff_cost'),accounts,plan)
    try:
        forecast,actual=calculate_storage(accounts,plan,actuals)
        needs=forecast.monthly_review.ne('Within TCO FSS monthly quantity limits')
        st.metric('Account-months requiring review',int(needs.sum()))
        show(forecast,'storage_withdrawal_schedule.csv')
        entitlements=storage_entitlement_report(accounts,forecast)
        show(entitlements,'storage_entitlement_assumptions.csv')
        st.subheader('Monthly storage release task list')
        tasks=storage_release_tasks(entitlements,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
        overdue=tasks.task_status.isin(['DUE TODAY','OVERDUE'])
        if overdue.any():st.error(f'{int(overdue.sum())} storage release updates are due. Obtain the new release and update capacity, MDSQ and effective dates before relying on daily capability.')
        elif tasks.task_status.eq('UPCOMING').any():st.warning('Upcoming monthly release updates: new volumes and MDSQ are due by the first of each delivery month. Carried-forward quantities remain provisional.')
        show(tasks,'storage_release_update_tasks.csv')
        st.caption('Task completion is based on release dates covering the delivery month and planning quantities matching that release. Update the account table above, then export and replace the scheduled profile. This flag appears when the app refreshes and in scheduled archives; no email notification is configured.')
        with st.expander('Daily ratchet and injection screens'):
            st.caption('Uniform daily forecast is a planning assumption. The lower of beginning/end inventory provides a conservative TCO screen; the pipeline’s best inventory estimate controls actual MDWQ. Missing actual dates prevent a verified inventory ratchet. Transport capacity, gas loss, OFOs and permissions are separate checks.')
            show(storage_daily_screen(accounts,forecast,actual),'storage_daily_tariff_screen.csv')
        st.line_chart(forecast.pivot(index='month',columns='account',values='ending_mcf'))
        if not actual.empty:show(actual,'storage_actual_inventory.csv')
        if needs.any():st.warning('Review the flagged account-months before using this plan for nominations. Negative balances remain visible; the model does not hide shortfalls by capping withdrawals.')
        st.session_state['_run_reports']['storage_account_inputs.csv']=accounts.copy()
        st.session_state['_run_reports']['storage_plan_inputs.csv']=plan.copy()
        st.session_state['_run_reports']['storage_actual_inputs.csv']=actuals.copy()
        st.session_state['_run_reports']['storage_summer_cost_inputs.csv']=prices.copy()
        clean=lambda df:json.loads(df.to_json(orient='records',date_format='iso'))
        export=dict(schema=1,accounts=clean(accounts),plan=clean(plan),actuals=clean(actuals),summer_prices=clean(prices),columbia_cost_basis=cost_basis,columbia_tariff_cost=cost_profile)
        st.session_state.setdefault('_scheduled_storage_inputs',{})['columbia']=dict(cost_basis=cost_basis,tariff=cost_profile,accounts=clean(accounts),plan=clean(plan))
        actual_confirmation=dict(confirmed=confirm_actuals,opening_confirmed=confirm_openings,source_reference=actual_source)
        export['actual_confirmation']=actual_confirmation
        if confirm_actuals:
            saved_actuals=dict(actual_confirmation,rows=clean(actuals))
            actual_reports,actual_review=columbia_actual_reports(accounts,plan,saved_actuals,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
            for name,report in actual_reports.items():show(report,name)
            st.session_state['_scheduled_storage_inputs']['columbia']['actual_activity']=saved_actuals
            if actual_review:st.warning('Columbia actual inventory requires review; inspect missing days, opening confirmation and statement differences.')
        st.download_button('Download storage tracker inputs',json.dumps(export,indent=2),'riverbend_storage_inputs.json','application/json')
        with st.expander('January–February priority withdrawal proposal'):
            st.caption('Maximize January–February withdrawals, then December, then November and March; January breaks ties within the first group. Existing monthly injections are fixed. This is a monthly quantity proposal, not daily nomination approval or a price optimization. Load-balancing withdrawals are comparison inputs, not hard minimums; reduced withdrawals require another supply source.')
            try:
                proposal,comparison,proposed_forecast=priority_storage_plan(accounts,plan)
                st.dataframe(comparison,use_container_width=True,hide_index=True)
                st.dataframe(proposed_forecast[['account','month','ending_mcf','monthly_review','daily_review']],use_container_width=True,hide_index=True)
                st.dataframe(storage_daily_screen(accounts,proposed_forecast,pd.DataFrame()),use_container_width=True,hide_index=True)
                st.caption('Apply preserves your account settings and costs, and replaces the plan’s injection/withdrawal overrides with this proposal. Export a new scheduled profile afterwards. Missing MDSQ and daily screen exceptions still require resolution.')
                if st.button('Apply winter-priority proposal'):
                    saved=dict(export);saved['plan']=clean(proposal)
                    st.session_state['_storage_inputs']=saved
                    st.session_state['_storage_revision']=rev+1
                    st.rerun()
            except (ValueError,TypeError,KeyError) as exc:st.warning('Cannot propose a monthly schedule: '+str(exc))
    except (ValueError,TypeError,KeyError) as exc:
        st.session_state.setdefault('_scheduled_storage_inputs',{}).pop('columbia',None)
        st.error('Storage inputs need correction: '+str(exc))


def storage_forward_reference(ice,code,month,mappings,asof):
    if ice is None or ice.empty:return None,None,'No ICE source'
    latest=ice['TRADE DATE'].max()
    if pd.isna(latest) or latest.normalize()>asof.normalize():return None,latest,'Invalid or future ICE snapshot'
    if (asof.normalize()-latest.normalize()).days>4:return None,latest,'ICE snapshot older than four calendar days'
    snapshot=ice[ice['TRADE DATE'].eq(latest)]
    code=market(code)
    mapping=None
    if mappings is not None:
        candidates=[v for k,v in mappings.items() if market(k)==code and v is not None]
        def signature(value):
            cols=['CONTRACT','PRODUCT','HUB','CONTRACT TYPE']
            return (value.get('_reference_type'),tuple(value.get(c) for c in cols),tuple((value.get('_adjustment') or {}).get(c) for c in cols))
        if len({signature(v) for v in candidates})>1:
            return None,latest,'Conflicting selected forward mappings for '+code+' aliases'
        if candidates:mapping=candidates[0]
    if mapping is None:
        # If mappings were explicitly constructed, respect missing selections.
        if mappings is not None:return None,latest,'No selected forward mapping for '+code
        pair={'DOM':('DSF','DIS'),'TCO':('TCF','TIS')}[code]
        parts=[]
        for contract in pair:
            catalog=snapshot.loc[snapshot.CONTRACT.eq(contract)&snapshot['CONTRACT TYPE'].eq('F'),['CONTRACT','PRODUCT','HUB','CONTRACT TYPE']].drop_duplicates()
            if len(catalog)!=1:return None,latest,'Missing or ambiguous product '+contract
            parts.append(catalog.iloc[0].to_dict())
        mapping=dict(parts[0],_reference_type='Daily cash proxy: IF + GDD adjustment',_adjustment=parts[1])
    if mapping.get('_reference_type')!='Daily cash proxy: IF + GDD adjustment' or not mapping.get('_adjustment'):
        return None,latest,'Daily cash proxy requires IF plus GDD adjustment'
    total=0.
    for part in [mapping,mapping['_adjustment']]:
        if part.get('CONTRACT TYPE')!='F':return None,latest,'Unsupported non-monthly forward mapping'
        hits=snapshot.loc[snapshot['STRIP'].eq(month)]
        for col in ['CONTRACT','PRODUCT','HUB','CONTRACT TYPE']:hits=hits.loc[hits[col].eq(part.get(col))]
        vals=hits['SETTLEMENT PRICE'].map(num)
        if hits.empty or vals.isna().any() or vals.nunique()!=1:return None,latest,'Missing, invalid or conflicting forward '+str(part.get('CONTRACT'))
        total+=float(vals.iloc[0])
    return total,latest,str(mapping['CONTRACT'])+' + '+str(mapping['_adjustment']['CONTRACT'])


def storage_daily_reference(ice,code,delivery,asof):
    """Current-month daily outright settlement; never add an IF/GDD component."""
    if ice is None or ice.empty:return None,None,'No ICE source'
    latest=ice['TRADE DATE'].max()
    if pd.isna(latest) or latest.normalize()>asof.normalize():return None,latest,'Invalid or future ICE snapshot'
    if (asof.normalize()-latest.normalize()).days>4:return None,latest,'ICE snapshot older than four calendar days'
    contract,hub={'DOM':('DSS','Eastern Gas-South'),'TCO':('CGS','TCO'),'REXZ3D':('RES','REX-Z3')}[code]
    snapshot=ice.loc[ice['TRADE DATE'].eq(latest)]
    hits=snapshot.loc[snapshot.CONTRACT.eq(contract)&snapshot.HUB.eq(hub)&
        snapshot.PRODUCT.eq('NG Swing GDD Futures')&snapshot['CONTRACT TYPE'].eq('D')&
        snapshot.STRIP.eq(delivery)]
    if hits.empty:return None,latest,'Missing daily '+contract+' for '+delivery.date().isoformat()
    vals=hits['SETTLEMENT PRICE'].map(num)
    if vals.isna().any():return None,latest,'Invalid daily '+contract+' price'
    if vals.nunique()!=1:return None,latest,'Conflicting daily '+contract+' prices'
    return float(vals.iloc[0]),latest,contract+' daily outright'


def storage_automatic_prices(legacy,newer,master,ice,mappings,cutoff,asof):
    cutoff=pd.Timestamp(cutoff).normalize();asof=pd.Timestamp(asof).normalize()
    if cutoff>asof:raise ValueError('Historical cutoff cannot be after the valuation date.')
    daily=[];monthly=[]
    for code,book in [('DOM','DEO'),('TCO','Columbia')]:
        historical=daily_prices(code,pd.Timestamp('2026-04-01'),min(cutoff,pd.Timestamp('2027-03-31')),legacy,newer,None)
        lookup={r.delivery_date:r for _,r in historical.iterrows()}
        for month in pd.date_range('2026-04-01',periods=12,freq='MS'):
            forward,snapshot,forward_note=storage_forward_reference(ice,code,month,mappings,asof) if month+pd.offsets.MonthEnd(0)>cutoff else (None,None,'')
            group=[]
            for day_ in pd.date_range(month,month+pd.offsets.MonthEnd(0)):
                if day_<=cutoff:
                    r=lookup.get(day_);price=None;source='';status='Missing historical cash'
                    if r is not None:
                        # cash2 is primary after transition; master fills absent cash2 values only.
                        if r.status=='MISSING' and day_>=TRANSITION and master is not None:
                            r=daily_prices(code,day_,day_,legacy,newer,master).iloc[0]
                        source=r.source
                        if r.status=='AVAILABLE':price=r.price;status='Historical cash'
                        else:status=str(r.status)+' historical cash'
                    kind='Historical'
                else:
                    if month.to_period('M')==asof.to_period('M'):
                        price,daily_snapshot,source=storage_daily_reference(ice,code,day_,asof)
                    else:
                        price=forward;daily_snapshot=snapshot;source=forward_note
                    status='Forward estimate' if price is not None else source;kind='Forward'
                entry=dict(book=book,index_code=code,month=month.date().isoformat(),delivery_date=day_.date().isoformat(),price=price,
                    price_kind=kind,status=status,source=source,ice_trade_date=daily_snapshot.date().isoformat() if kind=='Forward' and pd.notna(daily_snapshot) else None)
                daily.append(entry);group.append(entry)
            good=[r['price'] for r in group if r['price'] is not None]
            n_hist=sum(r['price_kind']=='Historical' and r['price'] is not None for r in group)
            n_fwd=sum(r['price_kind']=='Forward' and r['price'] is not None for r in group)
            missing=len(group)-len(good)
            monthly.append(dict(book=book,index_code=code,month=month.date().isoformat(),spot_usd_dth=sum(good)/len(group) if missing==0 else None,
                calendar_days=len(group),historical_days=n_hist,forward_days=n_fwd,missing_days=missing,
                pricing_basis='Incomplete — monthly price withheld' if missing else 'Historical cash average' if n_fwd==0 else 'Forward estimate' if n_hist==0 else 'Historical cash plus forward estimate'))
    return pd.DataFrame(monthly),pd.DataFrame(daily)


def storage_price_input_frame(book,manual):
    if st.session_state.get('_storage_auto_enabled',False):
        table=st.session_state.get('_storage_auto_monthly')
        result=manual.copy()
        values={} if table is None else table.loc[table.book.eq(book)].set_index('month').spot_usd_dth.to_dict()
        result['spot_usd_dth']=[values.get(pd.Timestamp(m).date().isoformat()) for m in result.month]
        bases={} if table is None else table.loc[table.book.eq(book)].set_index('month').pricing_basis.to_dict()
        result['pricing_basis']=[bases.get(pd.Timestamp(m).date().isoformat(),'Automatic source unavailable') for m in result.month]
        return result
    result=manual.copy();result['pricing_basis']='Manual monthly price'
    return result


def render_storage_price_sources(tables,master,ice,mappings):
    st.subheader('Storage price sources')
    mode=st.selectbox('Storage pricing mode',['Automatic cash and forward prices','Manual monthly prices'],key='storage_price_mode')
    enabled=mode=='Automatic cash and forward prices'
    st.session_state['_storage_auto_enabled']=enabled;st.session_state['_storage_auto_monthly']=None
    today=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
    cutoff=st.date_input('Historical cash delivery cutoff',value=today.date(),max_value=today.date(),key='storage_cash_cutoff')
    st.caption('Before June 1, 2026: workbook cash. From June 1: workbook cash2, with the connected cash master filling missing values. Days after the cutoff in the current valuation month use exact-date DSS (DEO) or CGS (Columbia) daily outright prices. Later months use the constructed IF + GDD forward proxy (DOM DSF + DIS; TCO TCF + TIS, or selected mappings). Monthly averages weight every calendar day equally. Historical gaps/conflicts are not filled with forwards. Latest ICE snapshot must be no more than four calendar days old.')
    if not enabled:
        st.info('Manual prices are active. Previously entered or restored monthly inputs are used.');return
    try:
        monthly,daily=storage_automatic_prices(tables['cash'],tables['cash2'],master,ice,mappings,cutoff,today)
        st.session_state['_storage_auto_monthly']=monthly
        show(monthly,'storage_automatic_monthly_prices.csv')
        with st.expander('Storage daily price provenance'):show(daily,'storage_automatic_daily_prices.csv')
        if monthly.missing_days.gt(0).any():st.warning('Some monthly prices are incomplete. Review missing delivery dates or forward products in the provenance report.')
        st.caption('Automatic prices feed the Columbia storage and DEO storage tabs. Forward-based months are estimates, not historical costs. Index-based gas prices do not replace transportation charges, tariff fees, or contract premiums. Current-month cash missing today can be reviewed by adjusting the cutoff to the last historical delivery date; only subsequent days use forwards.')
    except (ValueError,KeyError,TypeError,AttributeError) as exc:st.error('Automatic storage pricing unavailable: '+str(exc))


def storage_cost_summary(reports):
    """Combine modeled injection components only; no implicit zeros for missing reports."""
    def indexed(name):
        frame=reports.get(name)
        if frame is None or frame.empty:return {}
        if 'injection_month' not in frame:raise ValueError('Missing injection month in '+name)
        dates=pd.to_datetime(frame.injection_month,errors='raise')
        if dates.duplicated().any():raise ValueError('Duplicate injection months in '+name)
        return {d.date().isoformat():r for d,(_,r) in zip(dates,frame.iterrows())}
    deo=indexed('storage_deo_injection_month_cost.csv')
    coh=indexed('storage_columbia_injection_month_cost.csv')
    out=[]
    def value(row,key):
        if row is None:return None
        return storage_number(row.get(key),key,optional=True)
    for date in pd.date_range('2026-04-01',periods=7,freq='MS'):
        month=date.date().isoformat();d=deo.get(month);c=coh.get(month)
        gas_d=value(d,'injection_month_expense_usd');gas_c=value(c,'gas_purchase_cost_usd')
        trans=value(c,'transport_charge_usd');inj=value(c,'injection_charge_usd');aca=value(c,'aca_charge_usd')
        components=[gas_d,gas_c,trans,inj,aca]
        labels=['DEO received-gas cost','Columbia received-gas cost','Columbia transportation charge','Columbia injection charge','Columbia ACA']
        missing=[label for label,v in zip(labels,components) if v is None]
        out.append(dict(injection_month=month,deo_gas_including_receipt_shrink_usd=gas_d,
            columbia_gas_including_transport_and_storage_loss_usd=gas_c,columbia_transport_charge_usd=trans,
            columbia_injection_charge_usd=inj,columbia_aca_charge_usd=aca,
            combined_modeled_expense_usd=sum(components) if not missing else None,
            missing_components='; '.join(missing),status='Modeled components complete; exclusions apply' if not missing else 'Combined subtotal withheld'))
    result=pd.DataFrame(out)
    totals={c:(float(result[c].sum()) if result[c].notna().all() else None) for c in result.columns if c.endswith('_usd')}
    totals.update(period='April–October 2026',complete_months=int(result.combined_modeled_expense_usd.notna().sum()),
        scope='DEO received-gas cost plus Columbia received-gas and entered variable inbound charges; not all-in expense or P&L')
    return result,pd.DataFrame([totals])


def scheduled_storage_costs(profile,monthly):
    """Reprice explicitly exported injection inputs; never infer missing volumes or units."""
    if profile.get('schema') not in ['storage-cost-profile-1','storage-cost-profile-2','storage-cost-profile-3'] or profile.get('pricing_mode')!='Automatic':
        raise ValueError('Unsupported scheduled storage cost profile.')
    deo=profile['deo'];coh=profile['columbia']
    if deo.get('units')!='Dth' or deo.get('receipt_conversion_fraction')!=.047 or deo.get('withdrawal_conversion_fraction')!=0. or deo.get('conversion_side')!='Receipt only':
        raise ValueError('DEO requires Dth and 4.7% receipt-only conversion.')
    basis=coh['cost_basis']
    if not {'quantity','units','dth_per_mcf'}.issubset(basis):raise ValueError('Incomplete Columbia cost basis.')
    quantity,units,factor,net=columbia_cost_basis(basis)
    basis=dict(quantity=quantity,units=units,dth_per_mcf=factor)
    tariff=coh['tariff']
    if not tariff.get('rate_defaults_version'):raise ValueError('Export current tariff settings from the app.')
    rates=pd.DataFrame(tariff['inbound'])
    reports={}
    def prices(book):
        rows=monthly.loc[monthly.book.eq(book)&monthly.month.between('2026-04-01','2026-10-01'),['month','spot_usd_dth','pricing_basis']].copy()
        if rows.month.duplicated().any():raise ValueError('Duplicate storage price month.')
        return rows.sort_values('month').reset_index(drop=True)
    reports['storage_deo_injection_month_cost.csv']=deo_injection_cost(deo['net_injection_dth'],prices('DEO'))
    # Validate tariff structure even when the units are still unconfirmed.
    result=columbia_tariff_cost(0. if net is None else net,prices('Columbia'),rates)
    if net is not None:reports['storage_columbia_injection_month_cost.csv']=result
    reports['storage_columbia_cost_basis.csv']=pd.DataFrame([basis])
    reports['storage_columbia_inbound_rate_inputs.csv']=rates
    reports['storage_deo_scheduled_cost_inputs.csv']=pd.DataFrame([{k:deo[k] for k in ['units','net_injection_dth','receipt_conversion_fraction','withdrawal_conversion_fraction','conversion_side']}])
    combined,total=storage_cost_summary(reports)
    reports['storage_combined_injection_cost_monthly.csv']=combined
    reports['storage_combined_injection_cost_total.csv']=total
    return reports,int(combined.combined_modeled_expense_usd.isna().sum())


def storage_entitlement_report(accounts,forecast):
    rows=[]
    for _,a in accounts.iterrows():
        start=pd.to_datetime(a.get('release_effective_start'),errors='coerce')
        end=pd.to_datetime(a.get('release_effective_end'),errors='coerce')
        factor=storage_number(a.dth_per_mcf,'Dth/Mcf',0,True)
        for month in forecast.loc[forecast.account.eq(a.account),'month']:
            date=pd.Timestamp(month)
            covers=pd.notna(start) and pd.notna(end) and start<=date and end>=date+pd.offsets.MonthEnd(0)
            matches=factor is not None and same(a.capacity_mcf*factor,a.get('release_capacity_dth')) and same(a.mdsq_mcf_day*factor if pd.notna(a.mdsq_mcf_day) else None,a.get('release_mdsq_dth_day'))
            rows.append(dict(account=a.account,month=month,release_contract=a.get('release_contract'),
                release_effective_start=a.get('release_effective_start'),release_effective_end=a.get('release_effective_end'),
                planning_capacity_mcf=a.capacity_mcf,planning_mdsq_mcf_day=a.mdsq_mcf_day,
                posted_mdwq_dth_day=a.get('release_displayed_mdwq_dth_day'),
                status='Release covers month; planning quantities match' if covers and matches else 'REVIEW: provisional carry-forward or edited quantities'))
    return pd.DataFrame(rows)


def storage_release_tasks(entitlements,asof):
    today=pd.Timestamp(asof).normalize();rows=[]
    for _,r in entitlements.iterrows():
        due=pd.Timestamp(r.month)
        if due.to_period('M')<today.to_period('M'):continue
        complete=not str(r.status).startswith('REVIEW')
        state='COMPLETE' if complete else 'UPCOMING' if today<due else 'DUE TODAY' if today==due else 'OVERDUE'
        rows.append(dict(account=r.account,delivery_month=r.month,due_date=due.date().isoformat(),task_status=state,
            source_contract=r.release_contract,source_effective_end=r.release_effective_end,
            task='Obtain monthly release; update capacity, MDSQ and source effective dates; export and upload revised scheduled profile',
            planning_basis=r.status))
    return pd.DataFrame(rows,columns=['account','delivery_month','due_date','task_status','source_contract','source_effective_end','task','planning_basis'])


def columbia_outbound_summary(result,mode,monthly_flags=None,daily_flags=None,entitlement_flags=None):
    dates=list(pd.to_datetime(result.month,errors='raise'))
    if dates!=list(pd.date_range('2026-11-01',periods=5,freq='MS')):
        raise ValueError('Outbound summary requires all five November–March months in order.')
    incomplete=int(result.variable_outbound_charges_usd.isna().sum())
    totals={c:float(result[c].sum()) if result[c].notna().all() else None for c in ['withdrawal_dth','delivered_dth','transport_fuel_dth','withdrawal_charge_usd','transport_charge_usd','aca_charge_usd','variable_outbound_charges_usd']}
    totals.update(period='November 2026–March 2027',withdrawal_volume_source=mode,complete_months=5-incomplete,
        plan_months_requiring_review=monthly_flags,screened_days_requiring_review=daily_flags,provisional_entitlement_account_months=entitlement_flags,
        scope='Columbia variable outbound charges only; gas purchase remains in injection month; fixed charges and DEO excluded',
        status='REVIEW' if incomplete or monthly_flags is None or daily_flags is None or entitlement_flags is None or monthly_flags or daily_flags or entitlement_flags else 'Scenario calculated; operator confirmation required')
    return pd.DataFrame([totals])


def columbia_actual_reports(accounts,plan,data,asof):
    if data.get('confirmed') is not True or type(data.get('opening_confirmed')) is not bool:
        raise ValueError('Columbia actuals require explicit activity and opening-confirmation status.')
    if not txt(data.get('source_reference')):raise ValueError('Enter a Columbia actual activity source reference.')
    rows=pd.DataFrame(data['rows'])
    if rows.empty:raise ValueError('Confirmed Columbia actuals require daily rows.')
    dates=pd.to_datetime(rows.date,errors='raise');cutoff=min(pd.Timestamp(asof).normalize()-pd.Timedelta(days=1),pd.Timestamp('2027-03-31'))
    if dates.isna().any() or (dates!=dates.dt.normalize()).any() or (dates>cutoff).any():
        raise ValueError('Actual dates must be completed delivery days before the valuation date.')
    forecast,actual=calculate_storage(accounts,plan,rows)
    screen=storage_daily_screen(accounts,forecast,actual)
    screen=screen[screen.basis.eq('Entered actual movements')].copy()
    summaries=[]
    for _,account in accounts.iterrows():
        a=actual[actual.account.eq(account.account)].copy();days=pd.DatetimeIndex(pd.to_datetime(a.date))
        start=pd.Timestamp(forecast[forecast.account.eq(account.account)].month.min())
        missing=len(pd.date_range(start,cutoff,freq='D').difference(days))
        contiguous=bool(len(days)) and len(pd.date_range(start,days.max(),freq='D').difference(days))==0
        # Recorded movements remain auditable; incomplete history cannot reconcile a statement.
        history_valid=[]
        for day in days:
            history_valid.append(data['opening_confirmed'] and len(pd.date_range(start,day,freq='D').difference(days))==0)
        actual.loc[a.index,'statement_variance_mcf']=a.statement_variance_mcf.where(history_valid)
        variance=actual.loc[a.index,'statement_variance_mcf']
        mismatches=int(variance.abs().gt(1e-7).sum())
        limits=int(screen[screen.account.eq(account.account)].review.ne('Within screened quantity limits; operator confirmation required').sum())
        needs=bool(missing or mismatches or limits or not data['opening_confirmed'] or not contiguous)
        summaries.append(dict(account=account.account,expected_through=cutoff.date().isoformat(),entered_days=len(a),missing_elapsed_days=missing,
            opening_confirmed=data['opening_confirmed'],last_entered_day=None if a.empty else a.date.iloc[-1],
            reconciled_inventory_mcf=float(a.ending_mcf.iloc[-1]) if contiguous and data['opening_confirmed'] else None,
            statement_mismatches=mismatches,daily_screen_exceptions=limits,status='REVIEW' if needs else 'Entered actual ledger complete through cutoff',
            source_reference=data['source_reference']))
    summary=pd.DataFrame(summaries)
    return {'storage_columbia_actual_inputs.csv':rows,'storage_columbia_actual_inventory.csv':actual,
        'storage_columbia_actual_daily_screen.csv':screen,'storage_columbia_actual_reconciliation.csv':summary},int(summary.status.eq('REVIEW').sum())


def scheduled_storage_withdrawals(profile,asof=None):
    """Forecast gross withdrawals and variable outbound charges, without repurchasing gas."""
    if profile.get('schema') not in ['storage-cost-profile-2','storage-cost-profile-3']:raise ValueError('Withdrawal reports require a version 2 or 3 profile.')
    coh=profile['columbia'];tariff=coh['tariff']
    accounts=pd.DataFrame(storage_account_conversion_defaults(coh['accounts']));plan=pd.DataFrame(coh['plan'])
    if accounts.empty or plan.empty:raise ValueError('Withdrawal accounts and plan are required.')
    # Preserve explicit factors; do not silently fill missing inputs in scheduled runs.
    forecast,_=calculate_storage(accounts,plan,pd.DataFrame())
    screen=storage_daily_screen(accounts,forecast,pd.DataFrame())
    outbound=pd.DataFrame(tariff['outbound']).copy()
    mode=tariff['withdrawal_source']
    reports={'storage_account_inputs.csv':accounts,'storage_plan_inputs.csv':plan,
        'storage_withdrawal_schedule.csv':forecast,'storage_daily_tariff_screen.csv':screen}
    entitlements=storage_entitlement_report(accounts,forecast)
    reports['storage_entitlement_assumptions.csv']=entitlements
    reports['storage_release_update_tasks.csv']=storage_release_tasks(entitlements,pd.Timestamp.now(tz='America/Chicago').tz_localize(None) if asof is None else asof)
    if mode=='Storage withdrawal plan':
        linked,detail=columbia_planned_withdrawals(accounts,forecast)
        quantities=linked.set_index('month').withdrawal_dth
        outbound['withdrawal_dth']=[quantities.get(pd.Timestamp(m).date().isoformat()) for m in outbound.month]
        reports['storage_columbia_withdrawal_cost_account_bridge.csv']=detail
        reports['storage_columbia_withdrawal_cost_volume_bridge.csv']=linked
    elif mode!='Manual volumes':raise ValueError('Invalid withdrawal volume source.')
    result=columbia_outbound_cost(outbound)
    result['withdrawal_volume_source']=mode
    reports['storage_columbia_outbound_inputs.csv']=outbound
    reports['storage_columbia_outbound_cost.csv']=result
    incomplete=int(result.variable_outbound_charges_usd.isna().sum())
    monthly_flags=int(forecast.monthly_review.ne('Within TCO FSS monthly quantity limits').sum())
    daily_flags=int(screen.review.ne('Within screened quantity limits; operator confirmation required').sum())
    reports['storage_columbia_outbound_cost_summary.csv']=columbia_outbound_summary(result,mode,monthly_flags,daily_flags,int(entitlements.status.str.startswith('REVIEW').sum()))
    actual_review=0
    if coh.get('actual_activity') is not None:
        actual_reports,actual_review=columbia_actual_reports(accounts,plan,coh['actual_activity'],pd.Timestamp.now(tz='America/Chicago').tz_localize(None) if asof is None else asof)
        reports.update(actual_reports)
    return reports,dict(columbia_actual_review_accounts=actual_review,withdrawal_incomplete_months=incomplete,storage_plan_review_months=monthly_flags,storage_daily_review_rows=daily_flags,
        storage_entitlement_review_months=int(entitlements.status.str.startswith('REVIEW').sum()))


def storage_review_report(reports,asof):
    """Consolidate existing evidence; missing reports never imply a passing check."""
    today=pd.Timestamp(asof).normalize();out=[]
    def add(book,check,status,count,action,source):
        out.append(dict(book=book,check=check,status=status,items_requiring_review=count,next_action=action,
            open_tab='Columbia storage' if book=='Columbia' else 'DEO storage' if book=='DEO' else 'Storage pricing & summary',source_report=source))
    def frame(name):
        value=reports.get(name)
        return None if value is None or value.empty else value
    name='storage_withdrawal_schedule.csv';f=frame(name)
    n=None if f is None else int(f.monthly_review.ne('Within TCO FSS monthly quantity limits').sum())
    add('Columbia','Monthly plan','NOT CONFIGURED' if n is None else 'REVIEW' if n else 'SCREEN COMPLETE',n,'Review monthly plan quantity exceptions; daily nominations require separate checks.',name)
    name='storage_daily_tariff_screen.csv';f=frame(name)
    if f is not None:f=f[f.basis.eq('Uniform monthly forecast')]
    n=None if f is None or f.empty else int(f.review.ne('Within screened quantity limits; operator confirmation required').sum())
    add('Columbia','Daily forecast quantity screen','NOT CONFIGURED' if n is None else 'REVIEW' if n else 'SCREEN COMPLETE',n,'Review ratchets and injection limits; uniform daily forecasts are not nomination schedules.',name)
    name='storage_entitlement_assumptions.csv';f=frame(name)
    n=None if f is None else int(f.status.str.startswith('REVIEW').sum())
    add('Columbia','Release coverage','NOT CONFIGURED' if n is None else 'PROVISIONAL' if n else 'SOURCE COVERED',n,'Replace carried-forward quantities and source dates when new releases arrive.',name)
    for book,name,column in [('Columbia','storage_release_update_tasks.csv','task_status'),('DEO','storage_deo_release_update_tasks.csv','status')]:
        f=frame(name);due=None if f is None else int(f[column].isin(['DUE TODAY','OVERDUE']).sum());pending=None if f is None else int(f[column].eq('UPCOMING').sum())
        add(book,'Update tasks','NOT CONFIGURED' if due is None else 'DUE' if due else 'UPCOMING' if pending else 'COMPLETE',None if due is None else due+pending,'Complete due tasks and export the updated scheduled profile.',name)
    name='storage_deo_plan_summary.csv';f=frame(name)
    for key,label,action in [('deo_civ_missing','CIV','Enter utility-confirmed CIV; do not substitute capacity or opening inventory.'),('deo_plan_not_adopted','Plan adoption','Review monthly movements before adopting the plan.'),('deo_unconfirmed_allocation_months','Monthly allocations','Enter confirmed monthly demand and its source reference.'),('deo_monthly_review_rows','Monthly CIV screen','Resolve reported monthly exceptions; posted daily limits still apply.'),('deo_inventory_shortfall_months','Inventory shortfall','Reconcile opening inventory or movements; do not create a balancing adjustment without support.')]:
        n=None if f is None or key not in f else int(f.iloc[0][key])
        status='NOT CONFIGURED' if n is None else 'REVIEW' if n else 'SCREEN COMPLETE'
        if key=='deo_monthly_review_rows' and f is not None and f.iloc[0].get('deo_civ_missing',1):status='WAITING FOR CIV';n=None
        add('DEO',label,status,n,action,name)
    for book,name in [('Columbia','storage_columbia_actual_reconciliation.csv'),('DEO','storage_deo_actual_reconciliation.csv')]:
        f=frame(name);n=None if f is None else int(f.status.eq('REVIEW').sum())
        status=('NOT STARTED' if today<=pd.Timestamp('2026-11-01') else 'NOT CONFIGURED') if f is None else 'REVIEW' if n else 'SCREEN COMPLETE'
        add(book,'Actual activity',status,n,'Leave actuals empty before delivery; afterward enter completed days, confirm opening and retain source references.',name)
    for book,name in [('Both','storage_combined_injection_cost_total.csv'),('Columbia','storage_columbia_outbound_cost_summary.csv')]:
        f=frame(name);required=7 if book=='Both' else 5
        n=None if f is None else max(0,required-int(f.iloc[0].complete_months))
        add(book,'Injection cost completeness' if book=='Both' else 'Outbound cost completeness','NOT CONFIGURED' if n is None else 'INCOMPLETE' if n else 'MODELED COMPONENTS COMPLETE',n,'Review missing price/rate inputs. Complete modeled components do not represent all-in cost or operational approval.',name)
    add('DEO','December–March daily actual screen','NOT IMPLEMENTED',None,'Extend the daily screen using the relevant posted monthly allocations.','')
    add('Both','Unmodeled expenses and settlement','OUTSIDE CURRENT MODEL',None,'DEO withdrawal expenses, remaining fixed/other costs and invoice-based realized results remain outstanding.','')
    return pd.DataFrame(out)


def render_storage_review():
    st.subheader('Storage review — Columbia and DEO')
    st.caption('One checklist of current reports. SCREEN COMPLETE applies only to the named check; it is not nomination approval. Upcoming inputs and unimplemented work remain visible separately from calculated exceptions.')
    show(storage_review_report(st.session_state.get('_run_reports',{}),pd.Timestamp.now(tz='America/Chicago').tz_localize(None)),'storage_review_checklist.csv')


def storage_profile_restore_state(profile,asof):
    """Validate the complete export before preparing any session-state changes."""
    if profile.get('schema') not in ['storage-cost-profile-2','storage-cost-profile-3']:
        raise ValueError('Use a version 2 or 3 scheduled storage profile; version 1 contains costs only.')
    profile=json.loads(json.dumps(profile))
    missing_prices=pd.DataFrame([dict(book=book,month=d.date().isoformat(),spot_usd_dth=None,pricing_basis='Recalculate from connected sources') for book in ['DEO','Columbia'] for d in pd.date_range('2026-04-01',periods=7,freq='MS')])
    scheduled_storage_costs(profile,missing_prices)
    scheduled_storage_withdrawals(profile,asof)
    coh=profile['columbia'];defaults=storage_defaults();actual=coh.get('actual_activity')
    tracker=dict(schema=1,accounts=storage_account_conversion_defaults(coh['accounts']),plan=coh['plan'],
        actuals=actual['rows'] if actual else [],summer_prices=defaults['summer_prices'],
        columbia_cost_basis=coh['cost_basis'],columbia_tariff_cost=coh['tariff'],
        actual_confirmation={k:actual[k] for k in ['confirmed','opening_confirmed','source_reference']} if actual else {})
    updates={'_storage_inputs':tracker,'deo_cost_volume':float(profile['deo']['net_injection_dth']),
        'storage_price_mode':'Automatic cash and forward prices'}
    if profile['schema']=='storage-cost-profile-3':
        ops=profile['deo_operations'];deo_saved_plan_reports(ops,asof)
        updates.update(ecps_units=ops['units'],ecps_civ=float(ops['civ']),ecps_opening=float(ops['opening']),
            _ecps_rows=ops['plan'],_deo_current_plan=ops['plan'],_deo_allocations=ops['allocations'],
            deo_plan_adopted=ops['plan_adopted'],deo_tariff_service_v310='Enbridge Ohio ECPS utility-allocated')
        actual=ops.get('actuals')
        updates.update(deo_daily_opening=float(actual['opening'] if actual else ops['opening']),
            _deo_daily_rows=actual['rows'] if actual else [],deo_actual_confirmed=bool(actual),
            deo_opening_confirmed=actual['opening_confirmed'] if actual else False,
            deo_actual_source=actual['source_reference'] if actual else '')
    if profile.get('physical_balance') is not None:
        risk=profile['physical_balance'];checked=validate_physical_profile(risk)
        updates.update(_physical_saved_inputs=json.loads(checked.to_json(orient='records')),physical_link_storage=risk['link_storage'] if risk.get('defaults_version')=='linked-defaults-1' else True,physical_positive=risk['positive_volume_means_purchase'],
            physical_export_enabled=True,physical_dates=(pd.Timestamp(checked.month.min()).date(),(pd.Timestamp(checked.month.max())+pd.offsets.MonthEnd(0)).date()))
    if profile.get('counterparty_ledger') is not None:
        invoice_aging_reports(profile['counterparty_ledger'],asof)
        updates['_invoice_saved']=profile['counterparty_ledger']
    if profile.get('pool_molecule_balance') is not None:
        validate_pool_profile(profile['pool_molecule_balance'])
        updates['_pool_molecule_saved']=profile['pool_molecule_balance']
    if profile.get('monthly_market_scenario') is not None:
        validate_monthly_scenario(profile['monthly_market_scenario'])
        updates['_monthly_scenario_saved']=profile['monthly_market_scenario']
    if profile.get('counterparty_credit') is not None:
        counterparty_credit_reports(profile['counterparty_credit'],{},asof)
        updates['_credit_saved']=profile['counterparty_credit']
    return updates


def render_storage_profile_restore():
    with st.expander('Restore saved storage profile — both books'):
        st.caption('Upload riverbend_storage_cost_profile.json. Loading replaces Columbia accounts, plans, rates and included actuals plus DEO cost quantity. Version 3 also replaces DEO operations and included actuals; version 2 leaves existing DEO operations untouched. Actuals omitted from a restored book are cleared. Included physical-balance inputs and settings are restored too; older profiles without them leave those entries untouched. Automatic prices refresh from connected sources; draft daily rows and manual price tables are not included in this profile.')
        uploaded=st.file_uploader('Saved scheduled storage profile (JSON)',type=['json'],key='combined_storage_restore')
        if uploaded is not None and st.button('Load saved storage profile'):
            try:
                data=json.loads(uploaded.getvalue())
                updates=storage_profile_restore_state(data,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
                for key in ['_storage_revision','_deo_cost_rev','_physical_rev','_invoice_revision','_credit_revision','_monthly_scenario_revision','_pool_molecule_revision']+(['_ecps_rev','_deo_daily_rev'] if data['schema']=='storage-cost-profile-3' else []):
                    updates[key]=st.session_state.get(key,0)+1
                st.session_state.update(updates)
                st.success('Saved storage inputs loaded. Prices are refreshed; unresolved plan and allocation checks remain visible. Review both storage tabs before re-exporting.')
            except (ValueError,TypeError,KeyError,AttributeError) as exc:
                st.error('No storage inputs were replaced: '+str(exc))


def render_storage_cost_summary():
    st.subheader('Combined summer injection-cost summary')
    st.caption('Injection-month expenses only. DEO includes gas purchased with its 4.7% receipt conversion. Columbia includes gas purchased with its separately entered transportation/storage retainage, plus transportation, injection and ACA charges. Receipt loss is already in purchased gas cost; do not add its value again.')
    st.caption('Scope excludes DEO upstream transport not included in the city-gate gas price, fixed reservation/capacity charges, pooling fees, taxes and other unmodeled charges. Withdrawal charges remain in their withdrawal-month report. This is a subtotal of modeled costs, not an all-in storage cost or business P&L.')
    try:
        monthly,total=storage_cost_summary(st.session_state.get('_run_reports',{}))
        show(monthly,'storage_combined_injection_cost_monthly.csv')
        show(total,'storage_combined_injection_cost_total.csv')
        complete=int(total.iloc[0].complete_months)
        st.metric('Months with all modeled components',f'{complete} / 7')
        if complete==7:st.metric('Combined modeled summer expense',f"${total.iloc[0].combined_modeled_expense_usd:,.2f}")
        else:st.info('Combined summer total is withheld. Complete the missing components shown above in the DEO or Columbia cost sections. A missing value is not zero; a verified zero rate remains valid.')
        st.caption('These CSVs are included in the interactive Run archive. Save each book’s input JSON to retain edits; scheduled snapshots require the exported cost profile below.')
    except (ValueError,TypeError,KeyError) as exc:st.error('Cannot combine storage costs: '+str(exc))
    inputs=st.session_state.get('_scheduled_storage_inputs',{})
    if st.session_state.get('_storage_auto_enabled') and {'deo','columbia'}.issubset(inputs):
        profile=dict(schema='storage-cost-profile-2',pricing_mode='Automatic',
            deo={k:inputs['deo'][k] for k in ['units','net_injection_dth','receipt_conversion_fraction','withdrawal_conversion_fraction','conversion_side']},
            columbia=dict(cost_basis=inputs['columbia']['cost_basis'],accounts=inputs['columbia']['accounts'],plan=inputs['columbia']['plan'],
                tariff={k:inputs['columbia']['tariff'][k] for k in ['inbound','outbound','withdrawal_source','rate_defaults_version']}))
        if 'pool_molecule_balance' in inputs:
            profile['pool_molecule_balance']=inputs['pool_molecule_balance']
        if 'monthly_market_scenario' in inputs:
            profile['monthly_market_scenario']=inputs['monthly_market_scenario']
        if 'counterparty_credit' in inputs:
            profile['counterparty_credit']=inputs['counterparty_credit']
        if 'counterparty_ledger' in inputs:
            profile['counterparty_ledger']=inputs['counterparty_ledger']
        if 'physical_balance' in inputs:
            profile['physical_balance']=inputs['physical_balance']
        if 'actual_activity' in inputs['columbia']:
            profile['columbia']['actual_activity']=inputs['columbia']['actual_activity']
        if 'deo_operations' in inputs:
            profile['schema']='storage-cost-profile-3'
            profile['deo_operations']=inputs['deo_operations']
        st.download_button('Download scheduled storage cost profile',json.dumps(profile,indent=2),
            'riverbend_storage_cost_profile.json','application/json')
        st.caption('Upload this profile to the workbook repository using RISK_STORAGE_PROFILE_PATH. It includes injection inputs plus Columbia accounts, withdrawal plan, conversions, outbound rates and selected volume source. Scheduled runs refresh gas prices and retain plan review flags. Re-export whenever inputs change. Daily inventory is a uniform forecast screen, not actual nominations; Version 3 includes the DEO movement plan, CIV and allocation-confirmation tasks. Confirmed Columbia actual activity and November DEO actual activity are included when entered. Physical-balance inputs are included when enabled in the Physical balance tab. DEO withdrawal expenses and December–March actual tracking remain outside this profile.')


def main():
    st.session_state['_run_reports'] = {}
    st.session_state['_scheduled_storage_inputs'] = {}
    if st.session_state.get('_physical_defaults_version')!='linked-defaults-1':
        st.session_state['physical_link_storage']=True
        st.session_state['physical_export_enabled']=True
        st.session_state['_physical_defaults_version']='linked-defaults-1'
    st.set_page_config(page_title='Mercury Risk',layout='wide')
    st.title('Mercury'); st.caption(VERSION)
    with st.sidebar:
        if st.button('Refresh sources'):
            github_inventory.clear(); github_bytes.clear(); st.rerun()
        book,workbook_source=workbook_controls()
        mode=st.radio('Pricing source',['GitHub — automatic','Upload — validation'])
        remote=None; cash=None; ice=[]
        if mode=='GitHub — automatic':
            remote=github_pricing_controls()
        else:
            cash=st.file_uploader('Optional updated cash master',type=['csv'])
            ice=st.file_uploader('Full ICE download(s)',type=['xlsx'],accept_multiple_files=True)
        positive=st.checkbox('Positive volume means purchase',value=True)
        st.write('Cash transition: **June 1, 2026**')
        st.write('September 2022 DOM/EGTS: **7.053**')
        st.write('October 2026 cash: **system index plus premium**')
        if st.button('Clear processing cache'): st.cache_data.clear(); st.rerun()
        st.caption('GitHub is checked on app rerun with a five-minute cache. Refresh sources reloads workbook and pricing references. Dropdown choices are session-only unless configured in Secrets. Background refreshes use the separate GitHub workflow; see Run archive for runner health.')
    history=monthly_history()
    if book is None:
        st.subheader('Connect your trade workbook')
        st.info('The platform needs a workbook before it can generate trade reports. Pricing and workbook connections are separate.')
        st.markdown('1. Open **Workbook connection settings** in the sidebar.\n2. Enter the private repository, branch, and workbook path.\n3. Select **Use this workbook connection**.\n4. Save the generated [workbook] section in app Secrets to retain it across sessions.')
        st.caption('A token must already be present in app Secrets with read access to the workbook repository. This setup reads an existing GitHub file; it does not move your local workbook to GitHub. Upload — validation remains available.')
        with st.expander('Historical index reference'):
            show(history,'monthly_index_preview.csv')
        return
    if mode=='GitHub — automatic' and remote is None:
        st.warning('Complete the GitHub pricing setup in the sidebar, or select Upload — validation. GitHub mode does not fall back to saved worksheet prices.')
        return
    try:
        with st.spinner('Reading workbook and pricing sources...'):
            ledger,flags,tables,lookup,wapp=import_book(book.getvalue())
            master=remote['master'] if remote is not None else read_master(cash.getvalue()) if cash is not None else None
    except Exception as e:
        st.error(f'Import failed: {e}'); return
    if ledger.empty: st.error('No trade records found.'); return
    ledger,flags,expired_audit=acknowledge_expired_exceptions(ledger,flags)
    review_range=st.session_state.get('coverage_range',(date(2026,11,1),date(2027,3,31)))
    review_start,review_end=review_range if len(review_range)==2 else (None,None)
    scoped_flags=scope_import_flags(ledger,flags,review_start,review_end)
    st.caption('Workbook fingerprint: '+hashlib.sha256(book.getvalue()).hexdigest()[:12]+'. Source IDs follow worksheet rows.')
    if remote is not None:
        snapshot=pd.concat(remote['frames'],ignore_index=True)
        latest=snapshot['TRADE DATE'].max()
        st.write('GitHub pricing connected:', remote['repository'])
        st.caption('Cash file: '+remote['cash_path']+' | Last GitHub check: '+pd.Timestamp(remote['inventory']['checked_at']).tz_convert('America/Chicago').strftime('%Y-%m-%d %I:%M %p %Z'))
        if pd.notna(latest):
            st.write('Newest available ICE trade date:',latest.date())
            age=(pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()-latest.normalize()).days
            if age>4: st.warning(f'ICE trade date is {age} calendar days old. Confirm the upstream upload process.')
    a,b,c=st.columns(3)
    a.metric('Spot records',int(ledger.source_sheet.eq('cashblotter').sum()))
    b.metric('Forward records',int(ledger.source_sheet.eq('fwdblotter').sum()))
    c.metric('Reporting-period / unresolved flags',int(scoped_flags.requires_scheduled_review.sum()))
    coverage = None
    storage_ice=None;storage_mappings=None
    overview,*tabs=st.tabs(['Overview','October cash','Forward reconciliation','Trade ledger','Monthly forwards','Flags',
                  'Cash sources','Daily price check','Monthly indexes','ICE forwards','WAPP records','Forward valuation','Physical balance','Run archive','Columbia storage','DEO storage','Storage pricing & summary','Counterparty exposure','Pool molecule balance','Parallel deal register'])
    with tabs[1]:
        st.subheader('Forward trade classification reconciliation')
        if wapp.empty or not ledger.source_sheet.eq('fwdblotter').any(): st.warning('No forward trades or WAPP supporting records.')
        else:
            result,unused=reconcile(ledger,wapp)
            a,b,c=st.columns(3)
            a.metric('Unique classified matches',int(result.match_status.eq('MATCHED — REVIEW').sum()))
            b.metric('Shared consistent classifications',int(result.match_status.eq('SHARED — CONSISTENT CLASSIFICATION').sum()))
            c.metric('Unmatched/conflicting',int(result.match_status.isin(['UNMATCHED','CONFLICTING CLASSIFICATIONS']).sum()))
            st.info('Shared classification agreement retains every blotter trade. It does not prove one-to-one WAPP reconciliation, duplicate status, physical status, or hedge designation.')
            selected=st.multiselect('Match status',sorted(result.match_status.unique()),default=sorted(result.match_status.unique()))
            show(result[result.match_status.isin(selected)],'forward_classification_reconciliation.csv')
            with st.expander('WAPP records not used for a recognized classification'): show(unused,'wapp_unresolved.csv')
    with tabs[2]:
        choice=st.selectbox('Blotter',['All','cashblotter','fwdblotter'])
        show(ledger if choice=='All' else ledger[ledger.source_sheet.eq(choice)],'trade_ledger.csv')
    with tabs[3]:
        st.info('All blotter trades are physical, as confirmed by the owner. Hedge designation remains separate.')
        period=st.date_input('Delivery range',value=(date(2026,11,1),date(2027,3,31)))
        if len(period)==2:
            a,b=map(pd.Timestamp,period)
            if a>b: st.error('Reversed range.')
            else:
                f=monthly_volumes(ledger,a,b,positive)
                if f.empty: st.info('No allocatable trades.')
                else:
                    show(f.groupby(['month','index_code'],as_index=False)[['purchase_volume','sale_volume','net_volume']].sum(),'monthly_forward_summary.csv')
                    show(f,'monthly_forward_detail.csv')
        st.caption('Invalid dates/volumes are excluded and flagged; possible duplicates remain included.')
    with tabs[4]:
        st.subheader('Exception review')
        st.caption('Cash scope: October 1–31, 2026. Forward scope follows the delivery range selected under ICE forwards. Unknown dates or missing ledger matches always remain review items. No trades are removed by these filters.')
        st.write('Forward review range: '+str(review_start)+' to '+str(review_end))
        a,b,c,d=st.columns(4)
        a.metric('Needs current review',int(scoped_flags.requires_scheduled_review.sum()))
        b.metric('Historical warnings',int(scoped_flags.reporting_scope.str.startswith('HISTORICAL').sum()))
        c.metric('Future warnings',int(scoped_flags.reporting_scope.str.startswith('FUTURE').sum()))
        d.metric('All warnings',len(flags))
        st.session_state['_run_reports']['scoped_import_flags.csv']=scoped_flags.copy()
        view=st.radio('Exception scope',['Current / unresolved','Historical','Future','All'],horizontal=True,key='exception_scope_view')
        if view=='Current / unresolved':
            filtered=scoped_flags.loc[scoped_flags.requires_scheduled_review]
        elif view in {'Historical','Future'}:
            filtered=scoped_flags.loc[scoped_flags.reporting_scope.str.startswith(view.upper())]
        else:
            filtered=scoped_flags
        issues=st.multiselect('Filter issue types',sorted(scoped_flags.issue.dropna().unique()),key='exception_issue_filter')
        if issues:
            filtered=filtered.loc[filtered.issue.isin(issues)]
        if not scoped_flags.requires_scheduled_review.any():
            st.success('No import flags in the reporting periods or unresolved delivery scope. Historical warnings remain available below.')
        if filtered.empty:
            st.info('No flags match the selected filters.')
        else:
            st.dataframe(filtered,hide_index=True,use_container_width=True)
        st.download_button('Download filtered exception review',filtered.to_csv(index=False).encode(),'filtered_exception_review.csv','text/csv',key='filtered_exception_review')
        st.download_button('Download all scoped flags',scoped_flags.to_csv(index=False).encode(),'scoped_import_flags.csv','text/csv',key='all_scoped_flags')
        st.caption('Source IDs identify worksheet rows, for example cashblotter:13475. Owner-confirmed corrections are shown in the trade_correction and source_daily_volume ledger columns. Also correct the Excel source and refresh sources. Matching trade fields are a review signal, not permission to delete a trade.')
        with st.expander('Complete original import flags — audit trail'):
            show(flags,'import_flags.csv')
        _,location_audit=duplicate_trade_review(ledger)
        with st.expander('Duplicate check — distinct delivery locations'):
            st.caption('Delivery location is included in the duplicate check. The records below matched the earlier check but are distinguished by location; all remain in the ledger and calculations.')
            show(location_audit,'distinct_location_trade_audit.csv')
        st.caption(f'{len(expired_audit)} historical issues acknowledged as expired; excluded from active flags and balance blocking. Original records remain in the ledger.')
        with st.expander('Acknowledged expired exceptions — audit trail'):
            show(expired_audit,'acknowledged_expired_exceptions.csv')
        st.caption('October pricing failures are also shown in the October cash tab and its CSV.')
    with tabs[5]:
        sources={'Historical cash':tables['cash'],'Workbook cash2':tables['cash2']}
        if master is not None: sources['GitHub cash master' if remote is not None else 'Uploaded cash master']=master
        rows=[dict(source=n,rows=len(f),first_date=f.Date.min(),last_date=f.Date.max(),
                   duplicate_date_rows=int(f.Date.duplicated(keep=False).sum()),invalid_dates=f.attrs.get('invalid_dates',0)) for n,f in sources.items()]
        show(pd.DataFrame(rows),'cash_source_coverage.csv'); show(lookup,'historical_cash_mapping.csv')
        show(pd.DataFrame([dict(index_code=k,direct_market=v,historical_proxy=k in PROXIES) for k,v in DIRECT.items()]),'direct_cash_mapping.csv')
        st.info('Before June 1: historical cash/proxies. June 1 onward: GitHub master in automatic mode; uploaded master or cash2 in validation mode. No fallback from missing direct prices to historical proxies.')
    with tabs[6]:
        code=st.selectbox('Cash index',list(DIRECT))
        period=st.date_input('Price-check range',value=(date(2026,10,1),date(2026,10,31)),key='daily_range')
        if len(period)==2:
            a,b=map(pd.Timestamp,period)
            if b<a or (b-a).days>3660: st.error('Use a valid range of 10 years or less.')
            else:
                f=daily_prices(code,a,b,tables['cash'],tables['cash2'],master); show(f,'daily_price_check.csv')
                if f.status.eq('AVAILABLE').all(): st.metric('Calendar-day average',f'${f.price.mean():,.4f}/MMBtu')
                else: st.error('Complete-period average withheld: pricing exceptions exist.')
    with tabs[7]:
        show(history,'historical_monthly_indexes.csv')
        st.success('September 2022 DOM and EGTS approved at 7.053; original values retained.')
        st.caption('Blank REX settlements remain missing. Basis conversion is specific to this legacy monthly table; negative daily prices are valid.')
    with tabs[8]:
        frames=list(remote['frames']) if remote is not None else []
        for uploaded in ice:
            try:
                f=read_ice(uploaded.getvalue()).copy(); f['source_file']=uploaded.name; frames.append(f)
            except Exception as e: st.error(f'{uploaded.name}: {e}')
        if frames:
            f=pd.concat(frames,ignore_index=True); latest=f['TRADE DATE'].max()
            storage_ice=f
            if pd.isna(latest): st.error('No valid ICE pricing date.')
            else:
                st.write('Newest ICE pricing date:',latest.date())
                cols=[c for c in ['CONTRACT','HUB','PRODUCT','CONTRACT TYPE'] if c in f]
                show(f.loc[f['TRADE DATE'].eq(latest),cols].drop_duplicates(),'ice_product_catalog.csv')
                st.subheader('Forward pricing coverage')
                st.caption('Choose a monthly IF reference or a daily cash proxy built from IF plus its GDD adjustment. Component prices and their sum are exported. These references check price availability; no mark-to-market is calculated.')
                period = st.date_input('Forward coverage delivery range', value=(date(2026,11,1),date(2027,3,31)), key='coverage_range')
                if len(period) == 2:
                    beginning, ending = map(pd.Timestamp, period)
                    if ending < beginning or (ending-beginning).days > 3660:
                        st.error('Choose a valid coverage range of 10 years or less.')
                    else:
                        allocations = monthly_volumes(ledger, beginning, ending, positive)
                        if allocations.empty:
                            st.info('No forward trades overlap the selected range.')
                        elif wapp.empty:
                            st.warning('WAPP classifications are required for coverage review.')
                        else:
                            classifications, _ = reconcile(ledger,wapp)
                            catalog = f.loc[f['TRADE DATE'].eq(latest) & f['CONTRACT TYPE'].eq('F'), ['PRODUCT','HUB','CONTRACT','CONTRACT TYPE']].drop_duplicates().fillna('')
                            catalog = catalog.sort_values(['PRODUCT','HUB','CONTRACT','CONTRACT TYPE'])
                            records = catalog.to_dict('records')
                            labels = ['Select monthly product'] + [str(r['PRODUCT'])+' | '+str(r['HUB'])+' | '+str(r['CONTRACT'])+' | '+str(r['CONTRACT TYPE']) for r in records]
                            mappings = {}
                            with st.expander('Monthly ICE product mappings', expanded=True):
                                st.caption('Monthly IF uses an outright IF product. The daily cash proxy adds the monthly IF-to-GDD adjustment. Validated default pairs (editable): DOMSP/EGTS/DOM = DSF + DIS; TCO = TCF + TIS. These references remain subject to contract interpretation.')
                                for code in sorted(allocations.index_code.unique()):
                                    recipe = st.selectbox('Reference type for '+code,
                                                          ['Select reference type','Monthly IF reference','Daily cash proxy: IF + GDD adjustment'],
                                                          index=2 if code in {'DOMSP','DOM','EGTS','TCO'} else 0,
                                                          key='coverage_recipe_'+code)
                                    if recipe == 'Select reference type':
                                        continue
                                    suggested = {'DOMSP':('DSF','DIS'),'DOM':('DSF','DIS'),
                                                 'EGTS':('DSF','DIS'),'TCO':('TCF','TIS')}.get(code, ('',''))
                                    base_default = next((j+1 for j,r in enumerate(records) if r['CONTRACT']==suggested[0]),0)
                                    selected = st.selectbox('Monthly IF product for '+code, range(len(labels)), index=base_default,
                                                            format_func=lambda j: labels[j], key='coverage_base_'+code)
                                    adjustment = 0
                                    if recipe == 'Daily cash proxy: IF + GDD adjustment':
                                        adj_default = next((j+1 for j,r in enumerate(records) if r['CONTRACT']==suggested[1]),0)
                                        adjustment = st.selectbox('Monthly GDD adjustment for '+code, range(len(labels)), index=adj_default,
                                                                  format_func=lambda j: labels[j], key='coverage_adjustment_'+code)
                                    if selected:
                                        mappings[code] = dict(records[selected-1], _reference_type=recipe)
                                        if adjustment:
                                            mappings[code]['_adjustment'] = records[adjustment-1]
                            storage_mappings=mappings
                            if mappings:
                                flat=[]
                                for k,v in mappings.items():
                                    row=dict(index_code=k,reference_type=v['_reference_type'],
                                             **{c:v[c] for c in ['CONTRACT','PRODUCT','HUB','CONTRACT TYPE']})
                                    row.update({'adjustment_'+c:v.get('_adjustment',{}).get(c,'') for c in ['CONTRACT','PRODUCT','HUB','CONTRACT TYPE']})
                                    flat.append(row)
                                show(pd.DataFrame(flat),'forward_ice_product_mappings.csv')
                            coverage = forward_price_coverage(allocations,classifications,f,mappings)
                            priced = coverage.price_status.eq('AVAILABLE — mapped product')
                            a,b = st.columns(2)
                            a.metric('Trade-month allocations',len(coverage))
                            b.metric('Allocations missing usable mapped prices',int((~priced).sum()))
                            show(coverage,'forward_pricing_coverage.csv')
                            st.caption('DOMSP/EGTS/DOM and TCO use the validated daily-reference defaults after a restart. Custom changes are session-only; the download records your selections. Identical repeated prices across source files are accepted, conflicting prices are blocked. Older snapshots never fill missing prices. Coverage is not approval of pricing interpretation or hedge status.')

        else: st.info('Upload a full ICE file to inspect the product catalog.')
        st.caption('Open Forward valuation for indicative contract cash flows and MTM. Newest snapshot selected by TRADE DATE, not filename.')
    with tabs[9]:
        show(wapp,'wapp_classifications.csv')
        if not wapp.empty: show(wapp.groupby('pricing_classification',dropna=False).size().reset_index(name='record_count'),'wapp_classification_counts.csv')


    with tabs[10]:
        st.subheader('Indicative forward valuation')
        st.info('FIXED uses the selected market reference less the contract fixed price. GDD uses the daily cash proxy plus its contract premium for projected cash flows; INDEX uses monthly IF plus its premium. Floating MTM measures the premium against the same unadjusted index. All blotter trades are treated as physical, as confirmed by the owner. Hedge designation remains separate.')
        st.caption('All values are undiscounted commodity estimates. Positive MTM is value to Mercury; positive contract amounts are purchases/payables and negative amounts are sales/receivables. Fees, fuel, taxes, credit adjustments and other book costs are excluded. Partial-month trades use the selected monthly reference for their allocated days.')
        if coverage is None or coverage.empty:
            st.info('Open ICE forwards, choose the delivery range and reference type for each index, then return here.')
        else:
            try:
                values = value_forward_trades(coverage,ledger,positive)
                # Retain full reports for archive reconciliation and the delivered/hedge summary.
                full_reports={'forward_trade_valuation.csv':values,
                    'forward_valuation_monthly.csv':summarize_forward_values(values,['month','index_code']),
                    'forward_valuation_counterparty.csv':summarize_forward_values(values,['counterparty'])}
                st.session_state['_run_reports'].update(full_reports)
                st.markdown('**Filter forward valuation**')
                st.caption('Select multiple values in each field. Empty means all; filters combine across fields. The metrics, three tables and filtered downloads below use these selections. Full-book reports remain available for reconciliation.')
                filter_fields=[('counterparty','Counterparties'),('month','Delivery months'),('index_code','Indexes'),
                    ('candidate_classification','Pricing types'),('valuation_status','Valuation statuses')]
                def clear_forward_filters():
                    for field,_ in filter_fields:st.session_state['forward_filter_'+field]=[]
                st.button('Clear valuation filters',on_click=clear_forward_filters,key='forward_clear_filters')
                selections={}
                columns=st.columns(3)
                for i,(field,label) in enumerate(filter_fields):
                    labels=pd.to_datetime(values[field]).dt.strftime('%Y-%m') if field=='month' else values[field].map(lambda x:txt(x) or '(blank)')
                    options=sorted(labels.drop_duplicates().tolist());key='forward_filter_'+field
                    if key in st.session_state:st.session_state[key]=[v for v in st.session_state[key] if v in options]
                    selections[field]=columns[i%3].multiselect(label,options,key=key)
                selected_values=filter_forward_valuation(values,selections)
                st.caption(f'Showing {len(selected_values):,} of {len(values):,} trade-month allocations.')
                blocked=selected_values.valuation_status.eq('BLOCKED')
                a,b,c=st.columns(3)
                a.metric('Selected forward trades',selected_values.source_id.nunique())
                b.metric('Selected blocked allocations',int(blocked.sum()))
                c.metric('Selected complete indicative MTM','No matching rows' if selected_values.empty else
                    'Withheld — incomplete' if blocked.any() else f"${selected_values.indicative_mtm.sum():,.2f}")
                if selected_values.empty:st.info('No trades match this combination. Change a filter or clear all filters.')
                show(selected_values,'forward_trade_valuation_filtered.csv')
                show(summarize_forward_values(selected_values,['month','index_code']),'forward_valuation_monthly_filtered.csv')
                show(summarize_forward_values(selected_values,['counterparty']),'forward_valuation_counterparty_filtered.csv')
                st.download_button('Download full forward_trade_valuation.csv',values.to_csv(index=False).encode('utf-8'),
                    file_name='forward_trade_valuation.csv',mime='text/csv',key='forward_valuation_full_download')
                st.caption('Recognized WAPP classifications, including shared consistent matches, are applied as reconciled candidates. Missing prices, missing terms, unresolved matches and unsupported PHYSB trades block their allocations and complete group totals. Priced-subset values remain explicitly labeled.')

                st.subheader('Forward price risk — full selected ICE coverage range')
                st.caption('The valuation-table filters above do not change these full-book scenarios or the Platform overview hedge summary.')
                st.caption('Each shock changes the selected market reference across all months in the coverage range. Fixed-trade MTM changes with price. Under the same-index floating assumption, GDD/INDEX commodity MTM stays constant while projected contract cash flows move. Volumes, premiums and contract terms remain fixed; weather and basis divergence are not simulated.')
                exposure=[]
                for (month,code),g in values.groupby(['month','index_code']):
                    incomplete=g.valuation_status.eq('BLOCKED').any()
                    v=g.outright_price_exposure_volume.sum() if not incomplete else None
                    exposure.append(dict(month=month,index_code=code,
                                         fixed_price_exposure_volume=v,
                                         mtm_change_for_10_cent_increase=None if v is None else v*0.10,
                                         status='INCOMPLETE — full exposure withheld' if incomplete else 'INDICATIVE'))
                show(pd.DataFrame(exposure),'forward_price_exposure.csv')
                stress=[]
                codes=sorted(values.index_code.unique())
                for change in [-0.50,-0.25,0.0,0.25,0.50]:
                    scenario=forward_scenario(values,{code:change for code in codes})
                    incomplete=scenario.valuation_status.eq('BLOCKED').any()
                    stress.append(dict(parallel_price_change=change,
                                       indicative_mtm=None if incomplete else scenario.scenario_mtm.sum(),
                                       mtm_change=None if incomplete else scenario.scenario_mtm_change.sum(),
                                       projected_signed_contract_amount=None if incomplete else scenario.scenario_projected_signed_contract_amount.sum(),
                                       status='INCOMPLETE — full totals withheld' if incomplete else 'INDICATIVE'))
                show(pd.DataFrame(stress),'forward_parallel_price_scenarios.csv')
                st.markdown('**Custom price changes by index ($/MMBtu)**')
                shocks={code:st.number_input('Reference price change — '+code,value=0.0,step=0.05,format='%.4f',key='shock_'+code) for code in codes}
                custom=forward_scenario(values,shocks)
                show(summarize_scenario(custom,['month','index_code']),'forward_custom_scenario_monthly.csv')
                with st.expander('Custom scenario trade detail'):
                    show(custom,'forward_custom_scenario_detail.csv')

            except ValueError as exc:
                st.error(str(exc))


    # Set shared pricing before either book renders; combine only after both report sets exist.
    with tabs[15]:
        render_storage_profile_restore()
        render_storage_price_sources(tables,master,storage_ice,storage_mappings)
    with tabs[0]:
        st.subheader('October 2026 cash — system pricing')
        st.info('All cashblotter delivery days in October use the system index plus the column-L premium. Worksheet fixed and saved prices are comparison fields only. Trades crossing month boundaries are split to October.')
        st.caption('Uses the same dated price policy as Counterparty exposure: workbook cash caches first, master for missing entries, and the shared daily forward forecast where configured. The daily detail identifies each source.')
        st.caption('Prices remain estimates until you explicitly confirm a cutoff and indexes. Confirmation records your review; it does not independently validate the source. Fees, tax, fuel, and forward-blotter trades are excluded.')
        with st.expander('Confirm published daily settlements'):
            confirm = st.checkbox('I have verified published settlements for the selected indexes through the cutoff', value=False)
            cutoff = st.date_input('Verified settlement cutoff', value=max(date(2026,10,1), min(pd.Timestamp.now(tz='America/Chicago').date(), date(2026,10,31))),
                                   min_value=date(2026,10,1), max_value=date(2026,10,31))
            confirmed_markets = st.multiselect('Verified index codes', list(DIRECT), default=[])
            st.caption('A cutoff applies only to selected indexes. Future delivery days can have published weekend prices; verify the applicable delivery dates. These choices are session-only and included in downloads.')
        report,detail=october_cash(ledger,tables['cash'],tables['cash2'],master,positive,st.session_state['_run_reports'].get('storage_automatic_daily_prices.csv'),storage_ice,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
        if report.empty: st.info('No cash trades overlap October 2026. Check import flags for invalid dates.')
        else:
            report,detail,summary = settlement_reports(report,detail,cutoff if confirm else None, confirmed_markets if confirm else [])
            blocked=report.status.eq('BLOCKED')
            a,b,c=st.columns(3); a.metric('October trades',len(report)); b.metric('Blocked trades',int(blocked.sum()))
            c.metric('Unpriced volume (absolute)',f'{report.loc[blocked,"signed_volume"].abs().sum():,.0f}')
            st.caption('Positive system amounts = purchases/payables; negative = sales/receivables. Unpriced-volume metric excludes records whose volume itself is invalid.')
            show(report,'october_cash_system_pricing.csv')
            show(summary,'october_cash_counterparty_summary.csv')
            with st.expander('Daily pricing detail'): show(detail,'october_cash_daily_detail.csv')
            st.subheader('October invoice reconciliation')
            st.caption('Enter a signed commodity subtotal for the cashblotter trades shown above: purchases positive, sales negative. Combine multiple invoices per counterparty. Remove fees, tax, fuel, and forward trades from invoice totals before comparison. A blank amount means no invoice; zero is a valid amount.')
            seed = pd.DataFrame({'counterparty': summary.counterparty,
                                 'invoice_reference': '', 'invoice_signed_amount': float('nan')})
            upload = st.file_uploader('Restore invoice entries CSV (optional)', type=['csv'], key='invoice_entries_upload')
            if upload is not None:
                try:
                    loaded = pd.read_csv(upload)
                    required = ['counterparty','invoice_reference','invoice_signed_amount']
                    if not set(required).issubset(loaded.columns):
                        raise ValueError('Required columns: ' + ', '.join(required))
                    invoice_reconciliation(summary, loaded)
                    seed = seed.drop(columns=['invoice_reference','invoice_signed_amount']).merge(
                        loaded[required], on='counterparty', how='left', validate='one_to_one')
                except Exception as exc:
                    st.error(f'Invoice entries not imported: {exc}')
            entries = st.data_editor(seed, hide_index=True, use_container_width=True,
                                     disabled=['counterparty'], key='october_invoice_entries')
            show(entries,'october_invoice_entries.csv')
            try:
                comparison = invoice_reconciliation(summary, entries)
                show(comparison,'october_invoice_reconciliation.csv')
            except ValueError as exc:
                st.error(str(exc))
            st.caption('Download invoice entries to retain them and upload the CSV in a later session. Provisional differences include estimates; blocked counterparties have no complete variance. Invoice matching does not certify delivered volumes.')
    with tabs[13]:
        render_storage_tracker()
        render_columbia_tariff_scope()
    with tabs[14]:
        render_deo_storage_tracker()


    with tabs[11]:
        st.subheader('Monthly physical supply and demand')
        st.info('All cashblotter and fwdblotter trades are physical. WAPP records are supporting classifications only. Purchases supply gas; sales require gas.')
        bridge=storage_physical_bridge(st.session_state.get('_run_reports',{}))
        show(bridge,'storage_physical_supply_bridge.csv')
        with st.expander('Customer-demand reference — supplied winter forecast'):
            show(winter_customer_demand_reference(),'winter_customer_demand_reference.csv')
            st.caption('These are customer-demand volumes, not the Usage figures net of storage. DOM combines East and West; TCO uses Columbia. These full retail volumes are the editable defaults, because the owner confirmed booked sales are separate wholesale physical hedges.')
        link_storage=st.checkbox('Use saved storage plans for DOM and TCO physical balances',value=True,key='physical_link_storage')
        if link_storage:st.warning('Linked quantities use inventory-boundary Dth. Confirm compatibility with the blotter and account separately for transport/fuel losses before treating the result as a purchase requirement. DEO adjustments are not supply. Unresolved plans and provisional entitlements remain provisional.')
        st.caption('Enter retail demand and delivered supply in blotter volume units. Only explicitly identified delivered trades net against demand; pool purchases and sales remain separate. Other supply must be available for delivery and exclude booked delivered purchases and storage. Injections and withdrawals retain the provisional inventory boundary until delivery adjustments are established.')
        physical_positive=st.checkbox('Physical balance: positive volume means purchase',value=bool(positive),key='physical_positive')
        save_physical=st.checkbox('Include physical-balance inputs in saved profile and scheduled runs',value=True,key='physical_export_enabled')
        st.caption('Storage linking and inclusion in scheduled runs default to on. You can uncheck either for a deliberate alternative scenario. The combined JSON retains demand, other supply, storage-link selection and the purchase/sale sign convention. Blank demand remains incomplete; it is not treated as zero.')
        dates=st.date_input('Physical balance months',value=(date(2026,11,1),date(2027,3,31)),key='physical_dates')
        if len(dates)==2:
            start=pd.Timestamp(dates[0]).to_period('M').start_time
            end=pd.Timestamp(dates[1]).to_period('M').end_time.normalize()
            if end<start or (end-start).days>3660:
                st.error('Choose a valid range of 10 years or less.')
            else:
                allocations=monthly_volumes(ledger,start,end,physical_positive,all_physical=True)
                active=sorted(set(physical_delivery_classification(allocations).delivery_book)) if not allocations.empty else []
                options=sorted(set(active)|{market(k) for k in DIRECT})
                saved_physical=st.session_state.get('_physical_saved_inputs',[])
                saved_codes={market(r['index_code']) for r in saved_physical}
                options=sorted(set(options)|saved_codes)
                extra=st.multiselect('Additional markets with demand or storage but no booked trades', [k for k in options if k not in active],default=sorted(saved_codes-set(active)),key='physical_markets_'+str(st.session_state.get('_physical_rev',0)))
                codes=sorted(set(active+extra+(['DOM','TCO'] if link_storage or (start<=pd.Timestamp('2027-03-31') and end>=pd.Timestamp('2026-11-01')) else [])))
                seed=pd.DataFrame([dict(month=m.strftime('%Y-%m-01'),index_code=k,
                                        additional_customer_demand=float('nan'),other_supply=0.0,
                                        storage_withdrawals=0.0,storage_injections=0.0,delivered_price_usd_dth=float('nan'),delivery_differential_usd_dth=float('nan'),delivered_price_source='')
                                   for m in pd.date_range(start,end,freq='MS') for k in codes])
                if seed.empty:
                    st.info('Select a market to enter demand and storage plans.')
                else:
                    st.caption('November–March retail demand defaults to your supplied forecast: DOM = East + West; TCO = Columbia. Wholesale physical hedges remain separate purchases and sales, so no wholesale sale is subtracted from these defaults. Demand cells are editable; restored nonblank values, including zero, take precedence. Other markets/months remain blank. Other supply defaults to zero.')
                    if saved_physical:
                        loaded=pd.DataFrame(saved_physical).set_index(['month','index_code'])
                        seed=seed.set_index(['month','index_code']);overlap=seed.index.intersection(loaded.index)
                        common=seed.columns.intersection(loaded.columns)
                        seed.loc[overlap,common]=loaded.loc[overlap,common]
                        seed=seed.reset_index()
                    restored=st.file_uploader('Restore physical balance inputs CSV',type=['csv'],key='physical_restore')
                    if restored is not None:
                        try:
                            loaded=pd.read_csv(restored)
                            physical_balance(pd.DataFrame(),loaded)
                            loaded['month']=pd.to_datetime(loaded.month).dt.to_period('M').dt.strftime('%Y-%m-01')
                            loaded['index_code']=loaded.index_code.map(market)
                            keys=set(zip(seed.month,seed.index_code))
                            if not set(zip(loaded.month,loaded.index_code)).issubset(keys):
                                raise ValueError('Imported rows fall outside the selected months/markets. Adjust the selections first.')
                            seed=seed.set_index(['month','index_code'])
                            loaded=loaded.set_index(['month','index_code'])
                            common=seed.columns.intersection(loaded.columns)
                            seed.loc[loaded.index,common]=loaded[common]
                            seed=seed.reset_index()
                        except Exception as exc: st.error('Inputs not restored: '+str(exc))
                    seed=delivered_pricing_defaults(retail_demand_defaults(seed))
                    if link_storage:seed=link_storage_balance_inputs(seed,bridge)
                    st.caption('Delivery differentials default to $0.20/Dth for Dominion and $0.30/Dth for TCO, sourced to the applicable delivered Gas Daily index. Existing nonblank edits are preserved. Pricing: enter either delivery_differential_usd_dth (added to the refreshed pool reference) OR delivered_price_usd_dth (the full delivered forecast), plus delivered_price_source. Missing differentials use the approved defaults; enter zero explicitly if intended. Use a weighted delivery differential for the selected utility book; this forecast is not a supplier quote or a point-level allocation.')
                    entries=st.data_editor(seed,hide_index=True,use_container_width=True,disabled=['month','index_code']+(['storage_withdrawals','storage_injections'] if link_storage else []),key='physical_inputs_'+('linked' if link_storage else 'manual')+'_'+str(st.session_state.get('_physical_rev',0)))
                    if link_storage:entries=link_storage_balance_inputs(entries,bridge)
                    show(entries,'physical_balance_inputs.csv')
                    try:
                        risk_profile=dict(schema='physical-balance-1',defaults_version='linked-defaults-1',link_storage=link_storage,positive_volume_means_purchase=bool(physical_positive),inputs=json.loads(entries.to_json(orient='records')))
                        risk_reports,_=physical_profile_reports(ledger,flags,risk_profile,st.session_state.get('_run_reports',{}))
                        balance=risk_reports['monthly_physical_balance.csv']
                        st.session_state['_run_reports'].update(risk_reports)
                        if save_physical:st.session_state['_scheduled_storage_inputs']['physical_balance']=risk_profile
                        show(balance,'monthly_physical_balance.csv')
                        st.subheader('Delivered procurement — Gas Daily forecast cost and sensitivity')
                        show(risk_reports['pool_hedge_positions.csv'],'pool_hedge_positions.csv')
                        st.caption('Delivered purchases and sales net only against their delivery book. Pool hedges appear separately above. Delivered prices use either your full forecast or the refreshed pool reference plus your delivery differential, with a required source. DOM reference is DSF + DIS; TCO is TCF + TIS. Pool references must be within four calendar days and are never treated as delivered prices without an explicit differential. These monthly references estimate Gas Daily costs; they are not daily settlement prices. Storage remains provisional at the inventory boundary. Utility totals do not guarantee supply at each delivery point. This is not full P&L or hedge MTM.')
                        exposure_reports,blocked_exposure=physical_open_exposure(balance,storage_ice,storage_mappings,pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
                        for name,report in exposure_reports.items():show(report,name)
                        st.session_state['_run_reports']['monthly_delivery_hedge_summary.csv']=monthly_delivery_hedge_summary(st.session_state['_run_reports'])
                        render_monthly_market_scenario()
                        if blocked_exposure:st.warning(f'{blocked_exposure} month/market exposure rows cannot be priced. Review delivery classifications and enter delivered price forecasts with their sources.')
                        with st.expander('Underlying physical trade allocations'):
                            show(risk_reports['physical_trade_allocations.csv'],'physical_trade_allocations.csv')
                    except ValueError as exc: st.error(str(exc))


    with tabs[16]:
        st.subheader('Counterparty exposure — combined settlement forecast')
        st.caption('Current-month FIXED/GDD settlements are calculated directly from reconciled terms; current-month MTM is left blank. Combines cashblotter and fwdblotter by counterparty and delivery month. Purchases, sales, payables and receivables remain visible alongside their net. This is an unconfirmed commodity settlement forecast, not an unpaid balance or credit-limit report. Forward MTM stays separate; cash premiums come from booked trades, not delivery defaults.')
        now=pd.Timestamp.now(tz='America/Chicago').tz_localize(None).normalize()
        full_values=st.session_state['_run_reports'].get('forward_trade_valuation.csv',pd.DataFrame())
        last=now+pd.offsets.MonthEnd(0) if full_values.empty else max(now+pd.offsets.MonthEnd(0),pd.to_datetime(full_values.month).max()+pd.offsets.MonthEnd(0))
        period=st.date_input('Counterparty delivery range',value=(now.to_period('M').start_time.date(),last.date()),key='counterparty_delivery_range')
        if len(period)==2:
            try:
                cp_reports,cp_blocked=counterparty_settlement_reports(ledger,flags,full_values,
                    st.session_state['_run_reports'].get('storage_automatic_daily_prices.csv',pd.DataFrame()),
                    tables['cash'],tables['cash2'],master,period[0],period[1],now,positive,
                    ice=storage_ice,classifications=reconcile(ledger,wapp)[0] if not wapp.empty else pd.DataFrame())
                st.session_state['_run_reports'].update(cp_reports)
                summary=cp_reports['counterparty_settlement_monthly.csv'];detail_cp=cp_reports['counterparty_settlement_detail.csv']
                if summary.empty:st.info('No allocatable trades in this delivery range.')
                else:
                    selected_cp=st.multiselect('Counterparties — combined report',sorted(summary.counterparty.unique()),key='cp_report_parties')
                    selected_months=st.multiselect('Months — combined report',sorted(summary.month.unique()),key='cp_report_months')
                    for name,frame in cp_reports.items():
                        if name=='counterparty_index_price_checks.csv':st.caption('Index-price checks are shared source diagnostics for the full reporting range; counterparty/month display filters do not change this table. Look at delivery_date, expected_cash_column and issue for missing prices.')
                        view=frame
                        if selected_cp and 'counterparty' in view:view=view.loc[view.counterparty.isin(selected_cp)]
                        if selected_months and 'month' in view:view=view.loc[view.month.isin(selected_months)]
                        show(view,name.replace('.csv','_filtered.csv'))
                    st.download_button('Download full counterparty monthly report',summary.to_csv(index=False).encode(),
                        file_name='counterparty_settlement_monthly.csv',mime='text/csv',key='counterparty_full_export')
                    if cp_blocked:st.warning(f'{cp_blocked} counterparty/month totals withheld. Check the detail issues. Current-month FIXED/GDD settlements use reconciled terms and daily index prices. Other months require matching forward valuations. See detail issues; settlement completion does not imply MTM availability.')
            except ValueError as exc:st.error(str(exc))

    with tabs[16]:
        render_invoice_ledger()
        render_counterparty_credit()

    with tabs[18]:
        render_shadow_register(ledger,reconcile(ledger,wapp)[0] if not wapp.empty else pd.DataFrame(),book,positive)

    with tabs[17]:
        render_pool_molecules(ledger,flags)

    with tabs[15]:
        render_storage_cost_summary()
        render_storage_review()

    ice_date=pd.concat(frames,ignore_index=True)['TRADE DATE'].max() if frames else None
    with overview:
        st.subheader('Platform overview')
        render_risk_actions(ice_date)
        delivery_summary=st.session_state['_run_reports'].get('monthly_delivery_hedge_summary.csv',pd.DataFrame())
        if not delivery_summary.empty:
            st.subheader('Monthly delivered supply and pool hedges')
            st.caption('Delivered requirements and procurement forecasts appear alongside pool positions and matching forward MTM. These are separate measures, not a combined P&L. Missing or mismatched hedge valuations are withheld; cash pool trades are included in volumes but excluded from forward MTM.')
            show(delivery_summary,'monthly_delivery_hedge_summary.csv')
        st.caption('Operational checks for the currently loaded workbook and selected report ranges. Manual entry and business analysis can remain deferred.')
        st.write('Workbook source:', 'Connected GitHub repository' if workbook_source['kind']=='github' else 'Uploaded workbook')
        st.write('Pricing source:',mode)
        if workbook_source['kind']=='github':
            st.caption('Workbook commit: '+workbook_source['commit'])
        checks=platform_checks(st.session_state['_run_reports'],ice_date,
                               pd.Timestamp.now(tz='America/Chicago').tz_localize(None))
        review=checks.status.isin(['REVIEW','MISSING','NOT GENERATED'])
        a,b=st.columns(2)
        a.metric('Checks requiring attention',int(review.sum()))
        b.metric('Deferred or incomplete input areas',int(checks.status.isin(['DEFERRED','INPUTS / REVIEW']).sum()))
        show(checks,'platform_status.csv')
        st.caption('Use the Open tab column to locate detail. CALCULATED means the selected report rows passed these checks; it does not certify settlement, valuation, or total business profitability. Archives and saved runs are in Run archive.')

    with tabs[12]:
        st.session_state['_run_reports']['complete_trade_ledger.csv'] = ledger.copy()
        st.subheader('Run archive')
        st.caption('Download the current reports together with their source references and checksums. No additional entries are required. Each rerun rebuilds the package from current inputs; retain downloads for history.')
        provenance = dict(created_at_utc=pd.Timestamp.now(tz='UTC').isoformat(),
                          workbook_name=book.name,
                          workbook_source=workbook_source,
                          workbook_sha256=hashlib.sha256(book.getvalue()).hexdigest(),
                          pricing_mode=mode, positive_volume_means_purchase=bool(positive),
                          all_trades_physical=True)
        selected_range=st.session_state.get('coverage_range',())
        if coverage is not None and not coverage.empty and len(selected_range)==2:
            provenance['delivery_start']=pd.Timestamp(selected_range[0]).date().isoformat()
            provenance['delivery_end']=pd.Timestamp(selected_range[1]).date().isoformat()
        if remote is not None:
            provenance['github'] = dict(repository=remote['repository'],
                                         commit=remote['inventory']['commit'],
                                         branch=remote['inventory']['branch'],
                                         cash_path=remote['cash_path'], ice_paths=remote['ice_paths'])
        else:
            provenance['uploaded_prices'] = [dict(name=item.name, sha256=hashlib.sha256(item.getvalue()).hexdigest())
                                               for item in ([cash] if cash is not None else []) + list(ice)]
        if frames:
            pricing_date=pd.concat(frames,ignore_index=True)['TRADE DATE'].max()
            provenance['latest_ice_trade_date'] = None if pd.isna(pricing_date) else pricing_date.isoformat()
        bundle, manifest = build_run_archive(st.session_state['_run_reports'], provenance)
        a,b=st.columns(2)
        a.metric('Reports in archive',len(manifest['reports']))
        b.metric('Missing core reports',len(manifest['missing_core_reports']))
        if manifest['missing_core_reports']:
            st.warning('Not generated in this run: '+', '.join(manifest['missing_core_reports']))
        st.dataframe(pd.DataFrame([dict(report=r['file'],rows=r['rows']) for r in manifest['reports']]),hide_index=True,use_container_width=True)
        st.download_button('Download run archive',bundle,
                           'riverbend_run_'+pd.Timestamp.now(tz='UTC').strftime('%Y%m%d_%H%M%S')+'.zip',
                           'application/zip',key='run_archive_download')
        with st.expander('Run provenance'):
            st.json(provenance)
        st.caption('Archive creation does not certify report completeness. Manual-input reports retain their current defaults and missing values. Source workbooks and credentials are excluded.')
        render_run_history(bundle,manifest)


if __name__ == '__main__':
    main()
