"""Export the central PC catalog and a dated read-only inventory snapshot to GAS."""
import argparse
from datetime import datetime, timezone, timedelta
import hashlib
import json
from pathlib import Path
import re
import sys
from types import SimpleNamespace
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'home/src'))
from site_apps.central.src.models.stock_calculation import build_stock_lookups, stock_breakdown
from site_apps.central.src.models.material_model import _is_consumable_material_impl
from site_apps.central.src.services.ndt_calculator import calculate_billing
from site_apps.central.src.method_catalog import row_classification
from site_apps.central.src.aux_equipment import equipment_catalog


def prepare(inventory=None):
    central = ROOT/'home/src/site_apps/central'
    path = Path(inventory) if inventory else central/'data/Material_Inventory.xlsx'
    config = json.loads((central/'src/Material_Manager_Config.json').read_text(encoding='utf-8'))
    tables = pd.read_excel(path, sheet_name=None)
    clean = lambda v: '' if pd.isna(v) else str(v).strip()
    app = SimpleNamespace(materials_df=tables['Materials'], transactions_df=tables['Transactions'],
                          daily_usage_df=tables['DailyUsage'], clean_nan=clean,
                          equipments=config.get('equipments',[]),aux_equipments=config.get('aux_equipments',[]),
                          ndt_materials_all=['형광자분','흑색자분','백색페인트','침투제','세척제','현상제','형광침투제'],
                          normalize_id=lambda v: re.sub(r'\.0$', '', clean(v)))
    app._is_consumable_material = lambda name,method: _is_consumable_material_impl(app,name,method)
    def is_other_site_stationary_warehouse(wh: str) -> bool:
        w = clean(wh)
        return w.startswith('경인사무소-') and w != '경인사무소-본점'

    lookups = build_stock_lookups(app)
    items=[]
    central_mat_indices = []
    for idx, row in app.materials_df.iterrows():
        if not app.normalize_id(row.get('MaterialID')): continue
        amount = stock_breakdown(app,row,lookups)
        has_central_activity = any(amount[k] for k in ('in_qty','out_qty','daily_qty','ndt_usage'))
        if row.get('Active',1)==0 and not has_central_activity: continue
        if is_other_site_stationary_warehouse(row.get('창고')) and not has_central_activity:
            continue
        central_mat_indices.append(idx)
        methods,role=row_classification(row)
        items.append(dict(id=app.normalize_id(row['MaterialID']), name=clean(row.get('품목명')),
            model=clean(row.get('모델명')), sn=clean(row.get('SN')), warehouse=clean(row.get('창고')),
            company=clean(row.get('회사코드')), unit=clean(row.get('관리단위')),
            incoming=amount['total_incoming'],used=amount['total_used'],stock=amount['current_stock'],
            methods=sorted(methods),role=role,active=row.get('Active',1)!=0))
    app.materials_df = app.materials_df.loc[central_mat_indices].copy()
    fields=('companies','users','vehicles','equipments','aux_equipments','sites','daily_units','test_items','applied_codes','worktimes')
    managed={k:config[k] for k in fields if k in config}
    prices=json.loads((central/'src/config.json').read_text(encoding='utf-8'))

    # Harvest all companies from config, billing records, work history, and daily usage
    companies_set = set(config.get('companies', []))
    for r in prices.get('BILLING_RECORDS', []):
        if r.get('company'): companies_set.add(r['company'].strip())
    history_file = central/'src/daily_work_history.json'
    if history_file.exists():
        try:
            raw_hist = json.loads(history_file.read_text(encoding='utf-8'))
            for d, entry in raw_hist.items():
                if isinstance(entry, dict):
                    for r in entry.get('ndt_results', []):
                        if r.get('업체'): companies_set.add(r['업체'].strip())
        except Exception:
            pass
    if hasattr(app, 'daily_usage_df') and not app.daily_usage_df.empty:
        for c_col in ['업체명', 'Company', '업체']:
            if c_col in app.daily_usage_df.columns:
                for c in app.daily_usage_df[c_col].dropna().astype(str).str.strip().unique():
                    if c and c != 'nan': companies_set.add(c)
    managed['companies'] = sorted(c for c in companies_set if c)

    default_worktimes = [
        "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00", "09:00~22:00",
        "09:00~23:00", "09:00~24:00", "09:00~익일01:00", "09:00~익일02:00", "09:00~익일03:00",
        "09:00~익일04:00", "09:00~익일05:00", "09:00~익일06:00", "09:00~익일07:00", "09:00~익일08:00", "09:00~익일09:00",
        "10:00~17:00", "10:00~19:00", "12:00~익일08:00", "13:00~18:00", "13:00~22:00", "13:00~익일02:00",
        "18:00~익일01:00", "18:00~익일02:00", "18:00~익일03:00", "18:00~익일04:00", "18:00~익일05:00", "18:00~익일06:00", "18:00~익일07:00", "18:00~익일08:00",
        "20:00~익일02:00", "20:00~익일04:00", "20:00~익일05:00", "20:00~익일06:00",
        "21:00~익일03:00", "21:00~익일05:00", "21:00~익일06:00", "21:00~익일07:00",
        "22:00~익일01:00", "22:00~익일02:00", "22:00~익일03:00", "22:00~익일04:00", "22:00~익일05:00", "22:00~익일06:00", "22:00~익일07:00", "22:00~익일08:00", "22:00~익일09:00"
    ]
    wt_set = set(config.get('worktimes', [])) | set(default_worktimes)
    managed['worktimes'] = sorted(wt_set, key=lambda t: (1 if '익일' in t else 0, t.replace('익일', '')))

    managed['methods']=['PAUT','RT','UT','MT','PT']
    managed['base_prices']={}
    for method, material, shift in [('PAUT','PAUT_250A','야간'),('RT','RT_3 1/3 x 12"','야간'),('MT','MT_MT','일반'),('PT','PT_PT','일반')]:
        managed['base_prices'][method]=calculate_billing(quantity=1,adjusted_quantity=1,material_key=material,
            ndt_type=method,location='열배관',work_time=shift,material_costs=prices['MATERIAL_COST'],
            labor_costs=prices['LABOR_COST'],overhead_rate=1.1,technical_fee_rate=.2)['unit_price']
    managed['material_catalog']=[{k:x[k] for k in ('id','name','model','warehouse','company','unit','methods','role','active')} for x in items]
    managed['method_catalog']={}
    for method in managed['methods']:
        app.cb_daily_test_method=SimpleNamespace(get=lambda value=method:value)
        relevant=[x for x in managed['material_catalog'] if x['active'] and method in x['methods']]
        managed['method_catalog'][method]=dict(inventory=relevant,
            materials=sorted({f"{x['name']} ({x['model']})" if x['model'] else x['name']
                              for x in relevant if x['role']=='material'}),
            main=equipment_catalog(app,'main')[0],aux=equipment_catalog(app,'aux')[0])
    # Read daily work log progress history for mobile cumulative tracking
    progress_history = {}
    history_file = central/'src/daily_work_history.json'
    if history_file.exists():
        try:
            raw_hist = json.loads(history_file.read_text(encoding='utf-8'))
            for d, entry in raw_hist.items():
                if isinstance(entry, dict) and 'qty_data' in entry:
                    progress_history[d] = {
                        'qty_data': entry.get('qty_data', {}),
                        'weather': entry.get('weather', ''),
                        'remarks': entry.get('remarks', '')
                    }
        except Exception as e:
            print(f'Warning: Failed to load daily_work_history.json: {e}')

    default_expected_qty = {
        'PAUT_300A이상': 129, 'PAUT_300A이상-야간': 624,
        'PAUT_250A': 4, 'PAUT_250A-야간': 1, 'PAUT_200A': 4,
        'PAUT_200A-야간': 2, 'PAUT_150A~125A': 1, 'PAUT_150A~125A-야간': 1,
        'PAUT_100A이하': 1, 'PAUT_100A이하-야간': 1, 'PAUT_소계': 768,
        'RT_150A~100A': 293, 'RT_150A~100A-야간': 43,
        'RT_80A이하': 105, 'RT_80A이하-야간': 49, 'RT_소계': 490,
        'MT_전체(주간)': 26, 'MT_전체(야간)': 1,
        'PT_전체(주간)': 26, 'PT_전체(야간)': 1,
    }

    billing_fields = ('date', 'company', 'loc', 'ndt_type', 'work_time', 'material_type', 'qty', 'unit', 'unit_price', 'subtotal', 'section', 'line')
    billing_records = [
        {k: r.get(k) for k in billing_fields}
        for r in prices.get('BILLING_RECORDS', [])
    ]

    reference=dict(config=managed,inventory=items,
        daily_progress_history=progress_history,
        default_expected_qty=default_expected_qty,
        billing_records=billing_records,
        generated_at=datetime.now(timezone(timedelta(hours=9))).isoformat(timespec='seconds'),
        source_modified_at=datetime.fromtimestamp(path.stat().st_mtime,timezone(timedelta(hours=9))).isoformat(timespec='seconds'),
        total_master_rows=len(app.materials_df),
        source_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
        unassigned_chemicals=len(lookups[4].unassigned))
    out=central/'google_apps_script/Reference.gs'
    out.write_text('// Generated by tools/prepare_central_mobile_reference.py; read-only PC snapshot.\nvar CENTRAL_REFERENCE = '+json.dumps(reference,ensure_ascii=False,allow_nan=False,separators=(',',':'))+';\n',encoding='utf-8')
    print(f'Reference: {len(items)} materials, {len(managed.get("users",[]))} workers, {len(managed.get("vehicles",[]))} vehicles')
    return out


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--inventory')
    prepare(parser.parse_args().inventory)
