"""Prefer dated ICE sources over the historically stale generic latest alias."""
from datetime import date
import re


def file_date(path):
    match=re.search(r'(20\d{2})[_-](\d{2})[_-](\d{2})',path.name)
    if not match:return None
    try:return date(*map(int,match.groups()))
    except ValueError:return None


def ordered(paths):
    return sorted(paths,key=lambda p:('ice' not in p.name.lower(),-(file_date(p).toordinal() if file_date(p) else 0),p.name.lower()))


def select_newest(state,name):
    """Same carry/restore path as an explicit pricing-file change; preserve deal inputs."""
    if '_current_snapshot' in state:state['_pricing_carry']=state['_current_snapshot']
    state['wk_pricing_file']=name;state['_saved_pricing_file']=name
    state.pop('_restore_saved_file_once',None)
    state['_restore_pricing']=True;state['_saved_trade_date']=None
    state.pop('wk_asof',None)
