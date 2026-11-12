"""Boltex supplied monthly history; latest observation for each calendar month."""
import re
HISTORY=[('2025-08-01',6055),('2025-09-01',6984),('2025-10-01',8959),('2025-11-01',6003),('2025-12-01',4376),('2026-01-01',5054),('2026-02-01',7508),('2026-03-01',8814),('2026-04-01',7641),('2026-05-01',7762),('2026-06-01',7237),('2026-07-01',6518),('2026-08-01',8246)]
PROFILE={int(month[5:7]):float(volume) for month,volume in HISTORY}

def is_boltex(customer):return bool(re.match(r'^boltex\b',str(customer).strip(),flags=re.I))

def seed(base,customer):
    result=base.copy()
    if is_boltex(customer):result['usage_mmbtu']=[PROFILE[m.month] for m in result.month]
    return result
