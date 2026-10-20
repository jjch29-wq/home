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


def collect_chemical_usage(records, company_entries=()):
    totals = {key: 0.0 for key in CHEMICALS}
    if not records.empty:
        for _, row in records.iterrows():
            normalized = {str(k).replace(' ', '').upper(): v for k, v in row.items()}
            name = str(row.get('Material', row.get('품목명', ''))).replace(' ', '')
            for key, (_, aliases, names) in CHEMICALS.items():
                values = [numeric(normalized.get(alias.upper())) for alias in aliases]
                # Historical column names are alternatives for the same usage.
                value = next((v for v in values if v), 0.0)
                if not value and any(token in name for token in names):
                    if key != 'PT Penetrant' or '형광침투' not in name:
                        value = numeric(row.get('Usage', row.get('수량', 0)))
                totals[key] += value
    else:
        for entries in company_entries:
            for key, (ui_name, _, _) in CHEMICALS.items():
                widget = entries.get(ui_name)
                if widget is not None:
                    totals[key] += numeric(widget.get())
    return {key: {'used': value} for key, value in totals.items() if value}
