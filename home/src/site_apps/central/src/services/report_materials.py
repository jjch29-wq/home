"""Chemical usage for the daily report; fluorescent powder is its own material."""
import pandas as pd


CHEMICALS = {
    'MT WHITE': ('백색페인트', ('NDT_백색', 'NDT_백색페인트', 'NDT_백색페인트_MT', 'DT_백색페인트'), ('백색페인트', '백색')),
    'MT 7C-BLACK': ('흑색자분', ('NDT_흑색', 'NDT_흑색자분', 'NDT_자분', 'DT_흑색자분'), ('흑색자분', '흑색')),
    'MT FLUORESCENT': ('형광자분', ('NDT_형광자분', 'DT_형광자분'), ('형광자분',)),
    'PT Penetrant': ('침투제', ('NDT_침투', 'NDT_침투제', 'DT_침투', 'DT_침투제'), ('침투제', '침투액')),
    'PT Cleaner': ('세척제', ('NDT_세척', 'NDT_세척제', 'DT_세척', 'DT_세척제'), ('세척제', '세척액')),
    'PT Developer': ('현상제', ('NDT_현상', 'NDT_현상제', 'DT_현상', 'DT_현상제'), ('현상제', '현상액')),
}


def numeric(value):
    value = pd.to_numeric(str(value).replace(',', ''), errors='coerce')
    return 0.0 if pd.isna(value) else float(value)


def _clean_co(raw):
    if raw is None or pd.isna(raw):
        return ''
    s = str(raw).strip()
    if s.lower() in ('', 'nan', 'none') or s in ('한국지역난방공사', '한국지역난방공사 중앙지사'):
        return ''
    return s


def _fmt_qty(val):
    return f"{int(val)}" if float(val).is_integer() else f"{val:g}"


def collect_chemical_usage(records, company_entries=()):
    totals = {key: 0.0 for key in CHEMICALS}
    company_totals = {key: {} for key in CHEMICALS}
    if not records.empty:
        for _, row in records.iterrows():
            normalized = {str(k).replace(' ', '').upper(): v for k, v in row.items()}
            name = str(row.get('Material', row.get('품목명', ''))).replace(' ', '')
            row_co = _clean_co(row.get('회사코드', ''))
            for key, (_, aliases, names) in CHEMICALS.items():
                values = [numeric(normalized.get(alias.upper())) for alias in aliases]
                # Historical column names are alternatives for the same usage.
                value = next((v for v in values if v), 0.0)
                if not value and any(token in name for token in names):
                    if key != 'PT Penetrant' or '형광침투' not in name:
                        value = numeric(row.get('Usage', row.get('수량', 0)))
                if value > 0:
                    totals[key] += value
                    if row_co:
                        company_totals[key][row_co] = company_totals[key].get(row_co, 0.0) + value
        # Fallback to UI company selection if DB record had no manufacturer code yet
        for entries in company_entries:
            co_widget = entries.get('_company') if isinstance(entries, dict) else None
            ui_co = _clean_co(co_widget.get() if co_widget is not None and hasattr(co_widget, 'get') else '')
            if not ui_co:
                continue
            for key, (ui_name, _, _) in CHEMICALS.items():
                if totals[key] > 0 and not company_totals[key]:
                    widget = entries.get(ui_name)
                    if widget is not None:
                        v_ui = numeric(widget.get())
                        if v_ui > 0:
                            company_totals[key][ui_co] = company_totals[key].get(ui_co, 0.0) + v_ui
    else:
        for entries in company_entries:
            co_widget = entries.get('_company') if isinstance(entries, dict) else None
            ui_co = _clean_co(co_widget.get() if co_widget is not None and hasattr(co_widget, 'get') else '')
            for key, (ui_name, _, _) in CHEMICALS.items():
                widget = entries.get(ui_name)
                if widget is not None:
                    val = numeric(widget.get())
                    if val > 0:
                        totals[key] += val
                        if ui_co:
                            company_totals[key][ui_co] = company_totals[key].get(ui_co, 0.0) + val
    result = {}
    for key, value in totals.items():
        if value:
            item = {'used': value}
            co_map = company_totals[key]
            if len(co_map) == 1:
                only_co, only_qty = next(iter(co_map.items()))
                if abs(only_qty - value) < 1e-6:
                    item['company'] = only_co
                else:
                    item['company'] = f"{only_co} {_fmt_qty(only_qty)}"
            elif len(co_map) > 1:
                item['company'] = ', '.join(f"{co} {_fmt_qty(q)}" for co, q in co_map.items())
            result[key] = item
    return result
