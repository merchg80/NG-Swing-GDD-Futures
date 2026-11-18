"""Per-session source refresh results and last verified values."""
from datetime import datetime
from zoneinfo import ZoneInfo


def read_source(state, scope, name, identity, load, version, *, refreshing=False):
    snapshots = state.setdefault(f'_{scope}_source_snapshots', {})
    previous = snapshots.get(name)
    if previous and previous['identity'] != identity:
        previous = None
        snapshots.pop(name, None)
    try:
        value = load()
        fingerprint = version(value)
    except Exception:
        if refreshing or previous is not None:
            record(state, scope, name, 'Failed', previous['checked'] if previous else '',
                   'Using last good data from this session.' if previous else 'No verified data in this session.')
        if previous is None:
            raise
        return previous['value']

    was_failed = state.get(f'_{scope}_source_status', {}).get(name, {}).get('Status') == 'Failed'
    if refreshing or was_failed or previous is None or previous['version'] != fingerprint:
        checked = datetime.now(ZoneInfo('America/Chicago')).strftime('%Y-%m-%d %I:%M:%S %p %Z')
        snapshots[name] = dict(identity=identity, value=value, version=fingerprint, checked=checked)
        if refreshing or was_failed:
            status = 'Unchanged' if previous and previous['version'] == fingerprint else 'Updated'
            record(state, scope, name, status, checked)
    return value


def record(state, scope, name, status, checked='', detail=''):
    state.setdefault(f'_{scope}_source_status', {})[name] = {
        'Source': name, 'Status': status, 'Last successful check (CT)': checked, 'Details': detail,
    }


def show_status(ui, state, scope):
    rows = state.get(f'_{scope}_source_status', {})
    if not rows:
        return
    failures = [name for name, row in rows.items() if row['Status'] == 'Failed']
    if failures:
        ui.warning('Refresh failed for ' + ', '.join(failures) + '. Check the last successful source times below.')
    ui.dataframe(list(rows.values()), hide_index=True, use_container_width=True)
