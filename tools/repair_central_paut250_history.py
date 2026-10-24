"""Reclassify PAUT 250A day/night history, preserving source details and totals."""
import argparse
from copy import deepcopy
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'home/src'))
from site_apps.central.src.data_sync import get_history_path

DAY = 'PAUT_250A'
NIGHT = 'PAUT_250A-야간'
FIELDS = ('전일누계', '금일작업', '총누계', '공정률')


def number(value):
    return Decimal(str(value or '0').replace(',', ''))


def text(value):
    return f'{value:.4f}' if value else ''


def repair(original):
    result = deepcopy(original)
    active = False
    previous = None
    changes = []
    for date in sorted(original):
        old_data = original[date]
        quantities = old_data.get('qty_data', {})
        old_day, old_night = quantities.get(DAY, {}), quantities.get(NIGHT, {})
        rows = [row for row in old_data.get('ndt_results', [])
                if str(row.get('검사방법', '')).strip().upper() == 'PAUT'
                and re.fullmatch(r'250\s*[Aa]?', str(row.get('관경', '')).strip())]
        today = [number(old_day.get('금일작업')), number(old_night.get('금일작업'))]
        corrected = [Decimal(0), Decimal(0)]
        for row in rows:
            night = str(row.get('근무구분', '')).strip() == '야간'
            if not str(row.get('근무구분', '')).strip():
                night = str(row.get('규격', '')).strip().endswith('야간')
            corrected[int(night)] += number(row.get('PAUT'))
        if rows and sum(corrected) != sum(today):
            raise ValueError(f'{date}: source quantity differs from saved daily total; manual review required')
        if not rows:
            corrected = today
        if corrected != today:
            active = True
        if active:
            before = previous if previous is not None else [
                number(old_day.get('전일누계')), number(old_night.get('전일누계'))]
            total = [before[i] + corrected[i] for i in range(2)]
            for field, values in [('전일누계', before), ('금일작업', corrected), ('총누계', total)]:
                assert sum(values) == number(old_day.get(field)) + number(old_night.get(field)), (date, field)
                assert all(value >= 0 for value in values)
            new_qty = result[date].setdefault('qty_data', {})
            for i, key in enumerate((DAY, NIGHT)):
                values = new_qty.setdefault(key, {})
                for field, value in zip(FIELDS[:3], (before[i], corrected[i], total[i])):
                    values[field] = text(value)
                expected = number(values.get('예상량'))
                values['공정률'] = f'{total[i] / expected * 100:.1f}' if expected and total[i] else ''
            previous = total
            if result[date] != old_data:
                changes.append({'date': date, 'day_total': str(total[0]), 'night_total': str(total[1]),
                                'night_today': str(corrected[1]), 'night_pieces': sum(
                                    str(row.get('근무구분', '')).strip() == '야간' for row in rows)})
        else:
            previous = [number(old_day.get('총누계')), number(old_night.get('총누계'))]
    # No inspection rows, subtotal fields, or unrelated daily inputs may change.
    stripped = deepcopy(result)
    for date in original:
        for key in (DAY, NIGHT):
            if key in original[date].get('qty_data', {}):
                stripped[date]['qty_data'][key] = original[date]['qty_data'][key]
            else:
                stripped[date].get('qty_data', {}).pop(key, None)
    assert stripped == original
    return result, changes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()
    path = Path(get_history_path())
    if path.drive.upper() == 'C:':
        raise RuntimeError('Cloud history unavailable; refusing to alter the stale local fallback')
    raw = path.read_bytes()
    original = json.loads(raw.decode('utf-8-sig'))
    result, changes = repair(original)
    report = {'source': str(path), 'changes': changes, 'applied': False}
    if args.apply and changes:
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        backup = path.with_name(f'{path.stem}.before_paut250_repair_{stamp}.json')
        with backup.open('xb') as stream:
            stream.write(raw)
        assert backup.read_bytes() == raw
        if path.read_bytes() != raw:
            raise RuntimeError('History changed while preparing repair; retry after finishing app saves')
        temporary = path.with_name(f'{path.name}.repair_{stamp}.tmp')
        with temporary.open('x', encoding='utf-8') as stream:
            json.dump(result, stream, ensure_ascii=False, indent=4)
        assert json.loads(temporary.read_text(encoding='utf-8')) == result
        os.replace(temporary, path)
        assert json.loads(path.read_text(encoding='utf-8')) == result
        report.update(applied=True, backup=str(backup), backup_sha256=hashlib.sha256(raw).hexdigest())
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == '__main__':
    main()
