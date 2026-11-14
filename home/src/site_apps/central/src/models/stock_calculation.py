"""Central-site inventory calculations shared by stock rows and reorder alerts."""
import pandas as pd
import json

CHEMICALS = ('세척제', '침투제', '현상제', '백색페인트', '흑색자분', '형광자분', '형광침투제', '자분페인트')


class ChemicalUsage(dict):
    """Totals for diagnostics, plus consumption assigned to an exact material."""
    def __init__(self):
        super().__init__()
        self.by_material = {}
        self.unassigned = []


def chemical_key(name):
    """Prefer the most specific chemical name, so fluorescent penetrant is distinct."""
    text = str(name).replace(' ', '')
    return next((key for key in sorted(CHEMICALS, key=len, reverse=True) if key in text), None)


def build_stock_lookups(self):
    ndt_name_lookup = ChemicalUsage()
    # [OPTIMIZATION] Pre-calculate stock lookup to avoid O(N*M) performance hit
    stock_in_lookup = {}
    stock_out_lookup = {}
    if not self.transactions_df.empty:
        temp_trans = self.transactions_df.copy()
        # Use standardized normalization
        temp_trans['NormID'] = temp_trans['MaterialID'].apply(self.normalize_id)
        # [CRITICAL] Exclude "자동 차감" (Automatic Deduction) to avoid double-counting with Daily Usage sheet
        mask = ~temp_trans['Note'].astype(str).str.contains('자동 차감', na=False)
        
        in_mask = mask & (temp_trans['Type'] == 'IN')
        out_mask = mask & (temp_trans['Type'] == 'OUT')
        
        # Group by normalized ID and sum quantities
        stock_in_lookup = temp_trans[in_mask].groupby('NormID')['Quantity'].sum().to_dict()
        stock_out_lookup = temp_trans[out_mask].groupby('NormID')['Quantity'].sum().to_dict()

    # [NEW] Pre-calculate Daily Usage subtraction
    daily_usage_lookup = {}
    daily_name_lookup = {}
    
    # [NEW] Get a set of consumable MaterialIDs from master for robust deduction
    consumable_ids = set()
    if not self.materials_df.empty:
        for _, m in self.materials_df.iterrows():
            # Check if the master item itself is a consumable
            m_name = str(m.get('품목명', '')).strip()
            if self._is_consumable_material(m_name, ''):
                c_id = self.normalize_id(m.get('MaterialID'))
                if c_id: consumable_ids.add(c_id)
    
    if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
        temp_daily = self.daily_usage_df.copy()
        temp_daily['NormID'] = temp_daily['MaterialID'].apply(self.normalize_id)
        
        def _f(v):
            if pd.isna(v) or v is None: return 0.0
            try: return float(str(v).replace(',', '').strip()) if str(v).strip() else 0.0
            except: return 0.0
            
        # [ROBUST] Support multiple column names for usage and film counts
        # MT/PAUT/UT 등의 검사량(Meter)은 소모품 자재 ID가 부여되지 않은 한 소모량으로 합산하지 않음
        def _calc_daily_consumable(r):
            method = str(r.get('검사방법', '')).strip().upper()
            norm_id = r.get('NormID', '')
            if method in ['MT', 'PAUT', 'UT', 'PMI'] and norm_id not in consumable_ids:
                return 0.0
            for col in ['FilmCount', '매수', 'Usage', '검사량', '수량', 'Quantity']:
                v = _f(r.get(col, 0))
                if v > 0:
                    return v
            return 0.0

        temp_daily['TotalUsage'] = temp_daily.apply(_calc_daily_consumable, axis=1)
        
        # [NEW] Filter: Only include consumables in the site-usage deduction lookup
        # This ensures durable equipment (PAUT, MT-Yoke) reported in the site tab doesn't decrease stock
        # [ROBUST] Use MaterialID as primary signal, fallback to 품목명/장비명 keywords
        temp_daily['IsConsumable'] = temp_daily.apply(lambda r: 
            (r['NormID'] in consumable_ids) or 
            self._is_consumable_material(
                str(r.get('품목명', r.get('장비명', ''))).strip(), 
                str(r.get('검사방법', '')).strip()
            ), axis=1)
        
        temp_consumable = temp_daily[temp_daily['IsConsumable'] & (temp_daily['TotalUsage'] > 0)]
        
        # [CRITICAL FIX] 빈 ID('')나 결측치는 절대 룩업 키로 등록하지 않음 (미할당 품목 일괄 차감 방지)
        valid_id_consumables = temp_consumable[
            temp_consumable['NormID'].astype(str).str.strip().ne('') & 
            ~temp_consumable['NormID'].astype(str).str.lower().isin(['nan', 'none', 'null'])
        ]
        daily_usage_lookup = valid_id_consumables.groupby('NormID')['TotalUsage'].sum().to_dict()
        daily_usage_lookup.pop('', None)
        daily_usage_lookup.pop('nan', None)
        
        # [NEW] Also build a name-based lookup for fallback (빈 이름이나 'nan' 제외)
        temp_consumable['NormName'] = temp_consumable.apply(lambda r: str(r.get('품목명', r.get('장비명', ''))).strip(), axis=1)
        valid_name_consumables = temp_consumable[
            temp_consumable['NormName'].astype(str).str.strip().ne('') & 
            ~temp_consumable['NormName'].astype(str).str.lower().isin(['nan', 'none', 'null', '0', '0.0'])
        ]
        # Name fallback is only for unassigned usage and an unambiguous master.
        known_ids = {self.normalize_id(v) for v in self.materials_df['MaterialID']
                     if self.normalize_id(v) and self.normalize_id(v).lower() not in ('nan', 'none', 'null')}
        name_counts = self.materials_df['품목명'].map(self.clean_nan).value_counts()
        unassigned = valid_name_consumables[
            ~valid_name_consumables['NormID'].isin(known_ids) &
            valid_name_consumables['NormName'].map(name_counts).eq(1)]
        daily_name_lookup = unassigned.groupby('NormName')['TotalUsage'].sum().to_dict()
        print(f"DEBUG: Daily Consumable Lookup built: {daily_name_lookup}")
        
        # [NEW] Pre-calculate NDT chemical totals by name (Enhanced with fuzzy matching and JSON parsing)
        ndt_name_lookup = ChemicalUsage()
        ndt_keys = ['세척제', '침투제', '현상제', '백색페인트', '흑색자분', '형광자분', '형광침투제', '자분페인트']
        for k in ndt_keys: ndt_name_lookup[k] = 0.0
        chemical_materials = {key: [] for key in ndt_keys}
        for _, material in self.materials_df.iterrows():
            key = chemical_key(self.clean_nan(material.get('품목명', '')) +
                               self.clean_nan(material.get('모델명', '')))
            if key:
                chemical_materials[key].append(material)

        for _, row in temp_daily.iterrows():
            raw = row.get('ndt_data', row.get('ndtdata', ''))
            try:
                legacy = json.loads(raw) if isinstance(raw, str) and raw.strip().startswith('{') else raw
            except (ValueError, TypeError):
                legacy = {}
            if not isinstance(legacy, dict): legacy = {}
            for key in ndt_keys:
                # Positive direct usage wins; schema-default zero may still have legacy JSON usage.
                candidates = [f'NDT_{key}', key]
                candidates += [col for col in temp_daily.columns if chemical_key(col) == key and col not in candidates]
                value = None
                for col in candidates:
                    if col in row and not pd.isna(row[col]) and str(row[col]).strip():
                        value = _f(row[col])
                        break
                if value is None or value == 0:
                    value = sum(_f(amount) for name, amount in legacy.items() if chemical_key(name) == key)
                ndt_name_lookup[key] += value
                if value > 0:
                    candidates = list(chemical_materials[key])
                    row_id = self.normalize_id(row.get('MaterialID'))
                    exact = [m for m in candidates if row_id and self.normalize_id(m.get('MaterialID')) == row_id]
                    if exact:
                        candidates = exact
                    else:
                        for field in ('창고', '회사코드'):
                            selected = self.clean_nan(row.get(field, ''))
                            if selected:
                                candidates = [m for m in candidates if self.clean_nan(m.get(field, '')) == selected]
                    if len(candidates) == 1:
                        target = self.normalize_id(candidates[0].get('MaterialID'))
                        if target:
                            ndt_name_lookup.by_material[target] = ndt_name_lookup.by_material.get(target, 0) + value
                            continue
                    ndt_name_lookup.unassigned.append((key, value))

        # Filter out zero values for cleaner lookup
        for k in list(ndt_name_lookup):
            if ndt_name_lookup[k] <= 0:
                del ndt_name_lookup[k]
        print(f"DEBUG: Daily Usage lookup built. NDT totals: {ndt_name_lookup}")

    return stock_in_lookup, stock_out_lookup, daily_usage_lookup, daily_name_lookup, ndt_name_lookup


def stock_breakdown(self, mat, lookups):
    stock_in_lookup, stock_out_lookup, daily_usage_lookup, daily_name_lookup, ndt_name_lookup = lookups
    str_mat_id = self.normalize_id(mat.get('MaterialID'))
    mat_name_raw = self.clean_nan(mat.get('품목명', mat.get('ǰ', '')))
    model_name_raw = self.clean_nan(mat.get('모델명', mat.get('𵨸', '')))
    combined_name = (mat_name_raw + ' ' + model_name_raw).replace(' ', '')
    # Use optimized lookup (only if ID is valid)
    in_qty = stock_in_lookup.get(str_mat_id, 0.0) if str_mat_id else 0.0
    out_qty = abs(stock_out_lookup.get(str_mat_id, 0.0)) if str_mat_id else 0.0
    daily_qty = daily_usage_lookup.get(str_mat_id, 0.0) if str_mat_id else 0.0

    # The lookup contains only unassigned rows with a unique master match.
    if mat_name_raw:
        daily_qty += daily_name_lookup.get(mat_name_raw, 0.0)

    # [FINAL_STOCK_CALC] Current Stock = Master + In/Out - Site Usage
    ndt_usage = 0.0
    # If this is an NDT item, find matching usage by name
    if hasattr(self, 'ndt_name_lookup') or 'ndt_name_lookup' in locals():
        lookup_source = ndt_name_lookup if 'ndt_name_lookup' in locals() else self.ndt_name_lookup
        for ndt_key, ndt_val in lookup_source.items():
            if ndt_key == chemical_key(combined_name):
                ndt_usage = lookup_source.by_material.get(str_mat_id, 0)
                # If we found a specific NDT chemical (e.g. via model name), 
                # ensure we don't also subtract generic inspection quantity
                daily_qty = 0.0 
                break

    # Get stored quantity
    val = mat.get('수량', 0)
    try: stored_qty = float(str(val).replace(',', '')) if pd.notna(val) else 0.0
    except: stored_qty = 0.0

    total_incoming = stored_qty + in_qty
    total_used = out_qty + daily_qty + ndt_usage
    current_stock = total_incoming - total_used

    return dict(in_qty=in_qty, out_qty=out_qty, daily_qty=daily_qty,
                ndt_usage=ndt_usage, stored_qty=stored_qty,
                total_incoming=total_incoming, total_used=total_used,
                current_stock=current_stock)


def synchronize_legacy_chemicals(row, changes):
    """Mirror explicit edits into existing legacy JSON without touching unrelated usage."""
    chemicals = ('세척제', '침투제', '현상제', '백색페인트', '흑색자분', '형광자분', '형광침투제', '자분페인트')
    edited = {name: changes[f'NDT_{name}'] for name in chemicals if f'NDT_{name}' in changes}
    updates = {}
    if not edited:
        return updates
    for column in ('ndt_data', 'ndtdata'):
        if column not in row:
            continue
        raw = row[column]
        if isinstance(raw, dict):
            legacy = dict(raw)
        elif pd.isna(raw) or not str(raw).strip():
            continue
        else:
            legacy = json.loads(raw)
            if not isinstance(legacy, dict):
                raise ValueError('구형 약품 사용량 형식을 확인해주세요.')
        for name, amount in edited.items():
            for key in list(legacy):
                if chemical_key(key) == name:
                    del legacy[key]
            legacy[name] = float(amount or 0)
        updates[column] = json.dumps(legacy, ensure_ascii=False)
    return updates
