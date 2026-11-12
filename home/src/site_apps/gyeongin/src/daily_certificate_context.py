"""Carry the original daily record's certificate fields through work-log saves."""
CONTEXT_FIELDS = ('_source_date', '장비명', '보조장비', '작업자', 'Site', '공사명', '발주처')


def daily_certificate_context(record):
    def first(*keys):
        for key in keys:
            value = str(record.get(key, '') or '').strip()
            if value.lower() not in ('', 'nan', 'none', 'nat', 'n/a', '-'):
                return value
        return ''
    project = first('공사명', '공사명/현장', '현장명', 'Site', '현장')
    return dict(zip(CONTEXT_FIELDS, (
        first('_source_date', 'Date', '날짜'), first('장비명', '장비', 'Equipment'),
        first('보조장비'), first('Workers', 'User', '작업자', '검사자'),
        project, project, first('발주처'),
    )))


def merge_daily_context(row, context):
    result = dict(row)
    for key in CONTEXT_FIELDS:
        if str(result.get(key, '') or '').strip().lower() in ('', 'n/a', 'nan', 'none', '-'):
            if context.get(key):
                result[key] = context[key]
    return result
