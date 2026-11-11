"""Company/method/unit quantities read from saved DailyUsage, never carried totals."""
import math
import pandas as pd
from site_apps.gyeongin.src.services.report_inspection import inspection_method_shift


def first_usage_date(frame, fallback):
    if frame is not None and not frame.empty and 'Date' in frame:
        dates = pd.to_datetime(frame['Date'], errors='coerce', format='mixed').dropna()
        if not dates.empty:
            return min(dates.min().date(), pd.Timestamp(fallback).date())
    return pd.Timestamp(fallback).date()


def summarize_company_usage(frame, start, selected):
    start, selected = pd.Timestamp(start).date(), pd.Timestamp(selected).date()
    if start > selected:
        raise ValueError('조회 시작일은 검사일자보다 늦을 수 없습니다.')
    groups, skipped = {}, 0
    if frame is None or frame.empty:
        return [], skipped

    def clean(value):
        return '' if pd.isna(value) else str(value).strip()

    for _, row in frame.iterrows():
        day = pd.to_datetime(row.get('Date'), errors='coerce')
        if pd.isna(day):
            skipped += 1
            continue
        day = day.date()
        if not start <= day <= selected:
            continue
        method, _ = inspection_method_shift(row)
        if not method or method in ('내근/대기', '휴무'):
            continue
        if method == 'RT' and clean(row.get('검사구분', '')).upper() == 'REP':
            continue
        # A genuine zero in Usage must not fall back to a repeated 검사량.
        raw = row.get('Usage', None)
        if pd.isna(raw):
            raw = row.get('검사량', row.get('수량', 0))
        try:
            quantity = float(clean(raw).replace(',', '') or 0)
            if not math.isfinite(quantity) or quantity < 0:
                raise ValueError
        except (ValueError, TypeError):
            skipped += 1
            continue
        if not quantity:
            continue
        company = clean(row.get('업체명', '')) or clean(row.get('Company', '')) or '(업체명 미입력)'
        unit = clean(row.get('Unit', '')) or clean(row.get('단위', '')) or ('매' if method == 'RT' else '(단위 미입력)')
        key = company, method, unit
        detail = groups.setdefault(key, dict(company=company, method=method, unit=unit,
                                             previous=0.0, today=0.0, total=0.0))
        detail['today' if day == selected else 'previous'] += quantity
        detail['total'] += quantity
    return [groups[key] for key in sorted(groups)], skipped
