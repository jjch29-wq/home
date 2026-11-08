"""One OT rule for entry, inquiry, monthly totals and actual-budget labor."""
import re
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd

RATES = (4000, 5000, 7500)


def ot_breakdown(worktime, date, legacy_ot=''):
    text = str(worktime or '').strip()
    match = re.search(r'(\d{1,2}):(\d{2})\s*[~−-]\s*(?:익일\s*)?(\d{1,2}):(\d{2})', text)
    hours = [0.0, 0.0, 0.0]
    parsed_date = pd.to_datetime(date, errors='coerce')
    weekday = parsed_date.weekday() if not pd.isna(parsed_date) else 0
    if match:
        sh, sm, eh, em = map(int, match.groups())
        if sh > 23 or eh > 23 or sm > 59 or em > 59:
            raise ValueError('작업시간을 확인해주세요.')
        start, end = sh + sm / 60, eh + em / 60
        if end < start or ('익일' in text and end <= start): end += 24
        if '휴일' in text or weekday >= 5:
            hours[2] = end - start
        else:
            hours[0] = max(0, min(end, 22) - max(start, 18))
            hours[1] = max(0, min(end, 24) - max(start, 22))
            dawn = max(0, end - max(start, 24))
            hours[2 if weekday == 4 else 1] += dawn
    else:
        # Legacy OT duration, when no worktime range is available.
        duration = re.search(r'(\d+(?:\.\d+)?)\s*(?:시간|hr|h)', str(legacy_ot))
        if duration:
            amount = float(duration.group(1))
            if '휴일' in text or weekday >= 5: hours[2] = amount
            elif '야간' in text: hours[1] = amount
            else:
                hours[0], hours[1] = min(amount, 4), max(0, amount - 4)
    amounts = [int((Decimal(str(h)) * rate).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
               for h, rate in zip(hours, RATES)]
    return dict(day=hours[0], night=hours[1], holiday=hours[2], hours=sum(hours),
                day_amount=amounts[0], night_amount=amounts[1], holiday_amount=amounts[2], amount=sum(amounts))
