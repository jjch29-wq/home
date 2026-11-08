"""Read and update one vehicle slot without changing inspection or cost fields."""
import numbers
import re
import pandas as pd

VEHICLE_COLUMNS = ('차량번호', '주행거리', '차량점검', '차량비고')


def format_vehicle_numbers(values):
    """Print each vehicle once on its own line; keep commas inside model names."""
    if isinstance(values, str) or values is None or not hasattr(values, '__iter__'):
        values = [values]
    result = []
    for value in values:
        for label in re.split(r'\s*(?:\|\||\r?\n|,(?![^()]*\)))\s*', clean(value)):
            label = label.strip()
            if label and label not in result:
                result.append(label)
    return '\n'.join(result)


def clean(value):
    if value is None or pd.isna(value):
        return ''
    if isinstance(value, numbers.Number):
        return format(value, '.15g')
    return str(value).strip()


def fingerprint(row):
    return tuple(sorted((str(key), clean(value)) for key, value in row.items() if clean(value)))


def slots(row):
    columns = [[part.strip() for part in clean(row.get(key)).split('||')]
               for key in VEHICLE_COLUMNS]
    result = []
    for i in range(max(map(len, columns))):
        values = [column[i] if i < len(column) else '' for column in columns]
        if not any(values):
            continue
        data = dict(vehicle_info=values[0], mileage=values[1], remarks=values[3], photo_path='')
        pairs = values[2].split('|') if ':' in values[2] else values[2].split(',')
        for pair in pairs:
            if ':' in pair:
                key, value = pair.split(':', 1)
                data[key.strip()] = value.strip()
            elif pair.strip():
                key = pair.strip()
                data[key] = '잠금' if 'locking' in key else ('함' if 'cleaning' in key else '양호')
        result.append((i, data))
    return result


def record_choices(frame):
    """ORI/REP copies of the same timestamp and vehicle payload share one edit."""
    groups = {}
    for position, (_, row) in enumerate(frame.iterrows()):
        if clean(row.get('Active', 1)) == '0':
            continue
        stamp = clean(row.get('EntryTime'))
        date = clean(row.get('Date', row.get('날짜'))).split(' ')[0]
        site = clean(row.get('Site', row.get('현장')))
        key = (date, site, stamp or position, *(clean(row.get(k)) for k in VEHICLE_COLUMNS))
        if key not in groups:
            groups[key] = {'row': row.to_dict(), 'fingerprints': [], 'date': date, 'site': site}
        groups[key]['fingerprints'].append(fingerprint(row))
    result = []
    for group in groups.values():
        for slot, data in slots(group['row']):
            result.append({'date': group['date'], 'site': group['site'], 'data': data,
                           'method': clean(group['row'].get('검사방법')),
                           'source': {'fingerprints': group['fingerprints'], 'slot': slot,
                                      'date': group['date'], 'site': group['site']}})
    return sorted(result, key=lambda item: item['date'], reverse=True)


def update_vehicle(frame, source, data):
    """Return a new frame and refreshed source. Never mutate the caller's frame."""
    if not clean(data.get('vehicle_info')):
        raise ValueError('차량번호를 입력해주세요.')
    mileage = clean(data.get('mileage')).replace(',', '')
    if mileage:
        try:
            distance = float(mileage)
            if not 0 <= distance < float('inf'):
                raise ValueError
        except ValueError:
            raise ValueError('주행거리는 0 이상의 숫자로 입력해주세요.') from None
    expected = [tuple(tuple(pair) for pair in item) for item in source['fingerprints']]
    available = {}
    for position, (_, row) in enumerate(frame.iterrows()):
        available.setdefault(fingerprint(row), []).append(position)
    positions = []
    for identity in expected:
        matches = available.get(identity, [])
        if not matches:
            raise ValueError('원본 기록이 변경되었거나 삭제되었습니다. 기록 편집에서 다시 불러와주세요.')
        positions.append(matches.pop(0))
    result = frame.copy(deep=True)
    for key in VEHICLE_COLUMNS:
        if key not in result.columns:
            result[key] = ''
        result[key] = result[key].astype(object)
    checks = '|'.join(f'{k}:{v}' for k, v in data.items()
                      if k not in ('vehicle_info', 'mileage', '_raw_mileage', 'remarks') and v)
    replacement = (clean(data['vehicle_info']), mileage, checks, clean(data.get('remarks')))
    for position in positions:
        row = result.iloc[position]
        count = max(source['slot'] + 1, *(len(clean(row.get(k)).split('||')) for k in VEHICLE_COLUMNS))
        for key, value in zip(VEHICLE_COLUMNS, replacement):
            parts = [part.strip() for part in clean(row.get(key)).split('||')]
            parts += [''] * (count - len(parts))
            parts[source['slot']] = value
            result.iat[position, result.columns.get_loc(key)] = ' || '.join(parts)
    updated_source = dict(source, fingerprints=[fingerprint(result.iloc[p]) for p in positions],
                          _positions=positions)
    return result, updated_source


def persist_vehicle_edit(app, source, data):
    previous = app.daily_usage_df
    updated, next_source = update_vehicle(previous, source, data)
    app.daily_usage_df = updated
    try:
        if app.save_data() is False:
            raise RuntimeError('저장하지 못했습니다. 입력 내용은 유지됩니다. 다시 저장해주세요.')
    except Exception:
        app.daily_usage_df = previous
        raise
    positions = next_source.pop('_positions')
    next_source['fingerprints'] = [fingerprint(app.daily_usage_df.iloc[p]) for p in positions]
    return next_source
