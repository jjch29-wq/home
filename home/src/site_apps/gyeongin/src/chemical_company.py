"""Chemical manufacturers must never be populated from inspection clients."""
import pandas as pd


def chemical_company_options(app):
    clients = set(str(v).strip() for v in getattr(app, 'companies', []) if v)
    result = []
    def add(value):
        if pd.isna(value):
            return
        value = str(value).strip()
        if value and value.lower() != 'nan' and value not in clients and value not in result:
            result.append(value)

    for value in getattr(app, 'co_code_list', []):
        add(value)
    inventory = getattr(app, 'materials_df', pd.DataFrame())
    for _, row in inventory.iterrows():
        description = ' '.join(str(row.get(c, '')) for c in ('품목명', 'Material', '분류', 'Category'))
        if any(word in description for word in ('약품', '침투', '세척', '현상', '자분', '페인트')):
            add(row.get('회사코드', ''))
            add(row.get('제조사', row.get('Manufacturer', '')))
    history = getattr(app, 'daily_usage_df', pd.DataFrame())
    for _, row in history.iterrows():
        chemical_columns = [c for c in history.columns if str(c).startswith('NDT_')]
        used = any(pd.to_numeric(row.get(c), errors='coerce') > 0 for c in chemical_columns)
        if used:
            add(row.get('회사코드', ''))
    return result
