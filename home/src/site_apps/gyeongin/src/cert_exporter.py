"""Gyeongin customer test certificate (갑/을) and issue ledger Excel exporters."""
from __future__ import annotations

import datetime as dt
import os
import re
import zipfile
from copy import copy, deepcopy
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SUPPORTED_CERT_METHODS = ('MT', 'PT', 'UT', 'RT', 'PMI')

COMPANY_CODES = {
    '나라삼양감속기': 'SAMY',
    '제일푸리': 'JIFR',
    '경성기공': 'KSMC',
    '동해펌프': 'DHP',
    '태창기계공업': 'TMC',
    '모나스펌프': 'MNS',
    '현대건설': 'HD',
    '한국산업인력공단': 'HRDK',
    '서진인스텍': 'SJ',
    '현성펌프': 'HSP',
    '신성플랜트': 'SSP',
    '수성웹툰(주)': 'SS',
    '신성엔지니어링': 'SSEN',
    '한국종합기계': 'KHM',
    '유진건철': 'YJ',
    '현대산업': 'HDI',
    '함감속기제작소': 'HRM',
    'D&K켐텍': 'DNK',
    'K&M': 'KNM',
    'SG엔지니어링': 'SG',
    'SK에코플랜트': 'SKE',
    '글로벌 크레인': 'GC',
    '포스코인터내셔널': 'POSCO',
    '에버랜드': 'EVER',
    '스페이스 체인지': 'SC',
}

DEFAULT_METHOD_STANDARDS = {
    'MT': ('KS D 0213 / SIS H 301', '4'),
    'PT': ('KS B 0816 / SIS H 401', '2'),
    'UT': ('ASME Sec.Ⅴ (SA-388) / SIS H 247', '2'),
    'RT': ('KS B 0845', '0'),
    'PMI': ('ASTM E1476 / SIS', '0'),
}

_FORM_CODE_BY_METHOD_FAMILY = {
    ('MT', 'ASME'): 'H 341',
    ('MT', 'ISO'): 'H 363',
    ('MT', 'KS'): 'H 301',
    ('PT', 'ASME'): 'H 441',
    ('PT', 'KS'): 'H 401',
    ('UT', 'KS'): 'H 203',
    ('UT', 'ASME'): 'H 247',
    ('RT', 'ASME'): 'H 141',
    ('RT', 'KS'): 'H 101',
    ('PMI', 'ASTM'): 'H 801',
    ('PMI', 'KS'): 'H 801',
}


def classify_standard_family(code_text: str) -> str:
    """규격/절차 문자열에서 대표 규격군(KS, ASME, ISO, ASTM, AWS, API, JIS)을 판별합니다."""
    upper = _text(code_text).upper()
    if not upper:
        return ''
    if any(tok in upper for tok in ('ASME', 'SA-', 'SEC.', 'SEC ')):
        return 'ASME'
    if 'ASTM' in upper:
        return 'ASTM'
    if 'ISO' in upper:
        return 'ISO'
    if 'KS ' in upper or upper.startswith('KS'):
        return 'KS'
    if 'AWS' in upper:
        return 'AWS'
    if 'API' in upper:
        return 'API'
    if 'JIS' in upper:
        return 'JIS'
    return ''


def resolve_form_code(method: str, code_procedure: str = '') -> str:
    """검사방법과 적용규격/절차에서 성적서 바닥글 양식 코드(예: H 301, H 341)를 결정합니다."""
    m = re.search(r'\bH\s*(\d{3})\b', _text(code_procedure).upper())
    if m:
        return f"H {m.group(1)}"
    norm_method = normalize_cert_method(method)
    family = classify_standard_family(code_procedure)
    if (norm_method, family) in _FORM_CODE_BY_METHOD_FAMILY:
        return _FORM_CODE_BY_METHOD_FAMILY[(norm_method, family)]
    default_std = DEFAULT_METHOD_STANDARDS.get(norm_method, ('', ''))[0]
    default_family = classify_standard_family(default_std)
    return _FORM_CODE_BY_METHOD_FAMILY.get((norm_method, default_family), '')


def _sync_footer_form_code(wb, method: str, code_procedure: str = '') -> None:
    """성적서 각 시트의 왼쪽 바닥글 양식 번호(갑=-1, 을=-2)를 적용규격에 맞게 동기화합니다."""
    form_code = resolve_form_code(method, code_procedure)
    if not form_code:
        return
    for ws in wb.worksheets:
        suffix = '-2' if '(을)' in (ws.title or '') else '-1'
        left_footer = getattr(getattr(ws, 'oddFooter', None), 'left', None)
        if left_footer is not None and left_footer.text:
            left_footer.text = re.sub(
                r'H\s*\d+(?:-\d+)?',
                f"{form_code}{suffix}",
                left_footer.text,
            )


def code_procedure_candidates(method: str, company: str = '', current_value: str = '') -> list[str]:
    """Return ordered code/procedure dropdown candidates for (company, method)."""
    from site_apps.gyeongin.src.company_master import (
        load_company_ndt_presets,
        suggest_company_method_preset,
    )

    norm_method = normalize_cert_method(method) or _text(method).upper()
    presets = load_company_ndt_presets()
    ordered: list[str] = []
    seen: set[str] = set()

    def _add(val: str, *, require_balanced: bool = False) -> None:
        txt = _text(val)
        if not txt or txt in seen:
            return
        if require_balanced and txt.count('(') != txt.count(')'):
            return
        seen.add(txt)
        ordered.append(txt)

    _add(current_value)
    if company:
        co_preset = suggest_company_method_preset(company, norm_method, presets=presets)
        _add(co_preset.get('default_code', ''), require_balanced=True)

    lookup_methods = ('UT', 'PAUT') if norm_method == 'UT' else (norm_method,)
    for mmap in presets.values():
        if not isinstance(mmap, dict):
            continue
        for mk in lookup_methods:
            entry = mmap.get(mk)
            if isinstance(entry, dict):
                _add(entry.get('default_code', ''), require_balanced=True)

    if norm_method in DEFAULT_METHOD_STANDARDS:
        _add(DEFAULT_METHOD_STANDARDS[norm_method][0])
    if norm_method == 'RT':
        _add('KS B 0845 / SIS H 101')
        _add('ASME Sec.V / SIS H 141')
    return ordered

INSPECTOR_PROFILES = {
    '주진철': ('JIN CHEOL, JU', '기능사', 'II'),
    '권인경': ('IN KYUNG, KWON', '기 사', 'III'),
    '우명광': ('MYEONG GWANG, WOO', '기능사', 'II'),
    '진병학': ('BYUNG HAK, JIN', '기 사', 'III'),
    '이병국': ('BYUNG GUK, LEE', '기 사', 'II'),
    '박광복': ('KWANG BOK, PARK', '기 사', 'III'),
    '허남수': ('NAM SOO, HEO', '기능사', 'II'),
    '박원준': ('WON JUN, PARK', '기능사', 'II'),
    '형윤철': ('YUN CHUL, HYUNG', '기능사', 'II'),
    '유상훈': ('SANG HUN, YOO', '기능사', 'II'),
    '김성현': ('SUNG HYUN, KIM', '산업기사', 'II'),
}

EQUIPMENT_PROFILES = {
    'MT': ('NAWOO', 'MY-2', 'NW8064'),
    'PT': ('NAWOO', 'NPP-2 / NPR1-3 / NPD-4', '-'),
    'UT': ('SONATEST', 'MASTERSCAN 380', 'I005587'),
    'RT': ('RIGAKU', 'X-RAY / Se-75', '-'),
    'PMI': ('THERMO', 'NITON XL2 100G', '200983'),
}


def _text(value) -> str:
    if value is None:
        return ''
    s = str(value).strip()
    return '' if s.lower() in ('nan', 'none', 'nat') else s


def normalize_cert_method(method: str) -> str:
    m = _text(method).upper()
    if m == 'PAUT':
        return 'UT'
    return m if m in SUPPORTED_CERT_METHODS else ''


def clean_worker_name(raw: str) -> str:
    text = _text(raw)
    if not text:
        return ''
    first_part = re.split(r'[,/\n]+', text)[0].strip()
    first_part = re.sub(r'^(부장|차장|과장|대리|사원|주임|팀장|소장|이사)\s+', '', first_part)
    first_part = re.sub(r'\([^)]*\)', '', first_part).strip()
    return first_part


def format_korean_spaced_name(name: str) -> str:
    clean = _text(name).replace(' ', '')
    if re.fullmatch(r'[가-힣]{2,4}', clean):
        return (' ' * 5).join(list(clean))
    return _text(name)


def company_code(company_name: str) -> str:
    name = _text(company_name)
    if not name:
        return 'COMP'
    if name in COMPANY_CODES:
        return COMPANY_CODES[name]
    ascii_alnum = re.sub(r'[^A-Za-z0-9]', '', name).upper()
    if len(ascii_alnum) >= 2:
        return ascii_alnum[:5]
    return 'COMP'


def parse_date_value(date_str: str) -> dt.date:
    s = _text(date_str)
    for fmt in ('%Y-%m-%d', '%Y.%m.%d', '%Y/%m/%d', '%Y%m%d'):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return dt.date.today()


def default_report_no(date_str: str, company: str, method: str, seq: int = 1) -> str:
    d = parse_date_value(date_str)
    code = company_code(company)
    m = _text(method).upper() or 'NDT'
    return f"SIS/GI-{code}-{m}-{d.strftime('%y%m')}-{seq:03d}"


def _parse_single_equipment(raw: str, method: str) -> tuple[str, str, str]:
    default_maker, _, default_sn = EQUIPMENT_PROFILES.get(method, ('-', '-', '-'))
    sn_match = re.search(r'\(([^)]+)\)', raw)
    sn = sn_match.group(1).strip() if sn_match else default_sn
    upper_raw = raw.upper()
    for yoke_model in ('MY-2', 'MP-A3', 'MP-A2D'):
        if yoke_model in upper_raw:
            parts = sn.split()
            return 'NAWOO', yoke_model, parts[-1] if parts else default_sn
    if 'XL2' in upper_raw:
        return 'THERMO', 'NITON XL2 100G', sn
    if any(k in upper_raw for k in ('MASTERSCAN', 'SITESCAN', 'EPOCH', 'USM', 'OMNI')):
        model_head = raw.split('(')[0].strip()
        return 'OLYMPUS/SONATEST', model_head or raw, sn
    model_head = raw.split('(')[0].strip()
    return default_maker, model_head or raw, sn


def match_equipment_profile(equipment_text: str, method: str) -> tuple[str, str, str]:
    raw = _text(equipment_text)
    if not raw or raw == '해당없음':
        return '', '', ''
    items = [p.strip() for p in re.split(r'[;,]+', raw) if p.strip() and p.strip() != '해당없음']
    if not items:
        return '', '', ''
    makers, models, sns = [], [], []
    for item in items:
        mk, md, sn = _parse_single_equipment(item, method)
        if mk and mk not in makers:
            makers.append(mk)
        if md and md not in models:
            models.append(md)
        if sn and sn not in sns:
            sns.append(sn)
    return ' / '.join(makers), ' / '.join(models), ', '.join(sns)


def match_black_light_profile(aux_text: str) -> tuple[str, str]:
    """Extract (maker, model) from Black Light entries in 보조장비 text."""
    raw = _text(aux_text)
    if not raw:
        return '', ''
    items = [p.strip() for p in re.split(r'[;,]+', raw) if p.strip()]
    bl_items = [it for it in items if 'BLACK LIGHT' in it.upper() or 'UV-LED' in it.upper() or '자외선' in it]
    if not bl_items:
        return '', ''
    makers, models = [], []
    for it in bl_items:
        mk = 'NAWOO'
        if 'EE-365' in it.upper():
            mk = 'SPECTROLINE'
        if mk not in makers:
            makers.append(mk)
        cleaned = re.sub(r'^Black Light\s*(?:\([^)]*\))?\s*', '', it, flags=re.IGNORECASE).strip()
        if cleaned and cleaned not in models:
            models.append(cleaned)
    return ' / '.join(makers), ' / '.join(models)


def find_daily_usage_context(daily_usage_df, date_str: str, company: str, method: str, report_no: str = '') -> dict:
    """Lookup matching DailyUsage fields for (date, company, method) without mutating df."""
    empty = {
        'site': '',
        'customer': '',
        'equipment': '',
        'aux_equipment': '',
        'workers': '',
        'report_no': '',
        'applied_code': '',
        'test_item': '',
    }
    if daily_usage_df is None or getattr(daily_usage_df, 'empty', True):
        return empty

    df = daily_usage_df
    target_date = parse_date_value(date_str)
    target_comp = _text(company).casefold()
    target_method = _text(method).upper()

    matched_rows = []
    for _, row in df.iterrows():
        raw_d = _text(row.get('Date') or row.get('날짜'))
        if not raw_d:
            continue
        if parse_date_value(raw_d) != target_date:
            continue
        row_comp = _text(row.get('업체명') or row.get('업체')).casefold()
        row_site = _text(row.get('Site') or row.get('현장')).casefold()
        if target_comp and row_comp != target_comp and row_site != target_comp:
            continue
        row_method = _text(row.get('검사방법') or row.get('Method')).upper()
        if target_method and row_method and not row_method.startswith(target_method):
            continue
        if report_no and _text(row.get('성적서번호')) not in ('', report_no):
            continue
        matched_rows.append(row)

    # Relax method match if exact method row was not recorded yet for that date+company
    if not matched_rows and target_comp:
        for _, row in df.iterrows():
            raw_d = _text(row.get('Date') or row.get('날짜'))
            if not raw_d or parse_date_value(raw_d) != target_date:
                continue
            row_comp = _text(row.get('업체명') or row.get('업체')).casefold()
            row_method = _text(row.get('검사방법') or row.get('Method')).upper()
            if row_comp == target_comp and not row_method and (not report_no or _text(row.get('성적서번호')) in ('', report_no)):
                matched_rows.append(row)

    if not matched_rows:
        return empty

    def first_non_empty(*cols):
        for r in matched_rows:
            for c in cols:
                v = _text(r.get(c))
                if v:
                    return v
        return ''

    site = first_non_empty('공사명', '공사명/현장', '현장명', 'Site', '현장')
    if site == '경인사무소':
        site = ''
    return {
        'site': site,
        'customer': first_non_empty('발주처'),
        'equipment': first_non_empty('장비명', '장비', 'Equipment'),
        'aux_equipment': first_non_empty('보조장비'),
        'workers': clean_worker_name(first_non_empty('Workers', 'User', 'Worker1', '작업자', '검사자')),
        'report_no': first_non_empty('성적서번호'),
        'applied_code': first_non_empty('적용코드'),
        'test_item': first_non_empty('검사품명'),
    }


def row_identifier(row: dict) -> str:
    """Return primary item identifier across all 5 NDT target types."""
    for key in ('식별번호', '부품번호', '시험편번호', '설비번호', '태그번호', 'Joint No.', '검사품명'):
        val = _text(row.get(key))
        if val:
            return val
    return ''


def row_material(row: dict) -> str:
    for key in ('재질', '판정재질', '요구재질'):
        val = _text(row.get(key))
        if val:
            return val
    return ''


def row_remarks(row: dict) -> str:
    """Combine dimension text, quantity+unit, and location cleanly."""
    dim = _text(row.get('치수원문'))
    qty = _text(row.get('수량'))
    unit = _text(row.get('단위'))
    qty_text = f"{qty}{unit}" if qty and unit else qty

    parts = []
    if dim:
        parts.append(dim)
    if qty_text and qty_text.casefold() not in dim.casefold():
        parts.append(qty_text)
    loc = _text(row.get('검사위치'))
    if loc:
        parts.append(loc)
    return ' / '.join(parts)


def is_rejected(result_text: str) -> bool:
    s = _text(result_text).upper()
    return any(tok in s for tok in ('불합격', 'REJ', 'FAIL', 'NG'))


def row_interpretation(row: dict, default_interp: str) -> str:
    """Return NDT interpretation/finding text (not inspection scope like ALL SURFACE)."""
    explicit = _text(row.get('판정')) or _text(row.get('판정내용')) or _text(row.get('결함내용'))
    if explicit:
        return explicit
    result_txt = _text(row.get('결과'))
    if is_rejected(result_txt):
        return result_txt
    return default_interp


def build_cert_batches(
    ndt_results: list[dict],
    target_date: str = '',
    default_site: str = '',
    daily_usage_df=None,
) -> list[dict]:
    """Group daily ndt_results into printable customer certificate batches."""
    grouped: dict[tuple[str, str, str, str], list[dict]] = {}
    order_keys: list[tuple[str, str, str, str]] = []

    for row in ndt_results or []:
        if not any(_text(v) for v in row.values()):
            continue
        method_raw = _text(row.get('검사방법')).upper()
        cert_method = normalize_cert_method(method_raw)
        if not cert_method:
            continue
        company = _text(row.get('업체')) or '경인사무소'
        explicit_no = _text(row.get('성적서번호'))
        row_spec = _text(row.get('적용규격')) or _text(row.get('규격')) or _text(row.get('검사절차서'))
        default_family = classify_standard_family(DEFAULT_METHOD_STANDARDS[cert_method][0]) or 'KS'
        std_family = classify_standard_family(row_spec) or default_family
        key = (company, cert_method, explicit_no, std_family)
        if key not in grouped:
            grouped[key] = []
            order_keys.append(key)
        grouped[key].append(dict(row))

    batches: list[dict] = []
    seq_by_pair: dict[tuple[str, str], int] = {}

    for company, cert_method, explicit_no, std_family in order_keys:
        rows = grouped[(company, cert_method, explicit_no, std_family)]
        pair = (company, cert_method)
        seq_by_pair[pair] = seq_by_pair.get(pair, 0) + 1
        seq = seq_by_pair[pair]

        first_non_empty = lambda k: next((_text(r.get(k)) for r in rows if _text(r.get(k))), '')
        ctx_date = first_non_empty('_source_date') or target_date
        ctx = find_daily_usage_context(daily_usage_df, ctx_date, company, cert_method, explicit_no)
        report_no = explicit_no or ctx['report_no'] or default_report_no(ctx_date, company, cert_method, seq)
        customer = first_non_empty('발주처') or ctx['customer'] or company
        source_project = next((value for value in (first_non_empty('공사명'), first_non_empty('Site'), ctx['site'])
                               if value and value.upper() not in ('N/A', '-') and value != '경인사무소'), '')
        project_name = source_project or (_text(default_site) if _text(default_site) != '경인사무소' else '') or 'N/A'

        item_names = list(dict.fromkeys(filter(None, (_text(r.get('검사품명')) for r in rows))))
        if not item_names and ctx['test_item']:
            item_names = [ctx['test_item']]
        item_name = item_names[0] if len(item_names) == 1 else (f"{item_names[0]} 외" if item_names else 'N/A')

        materials = list(dict.fromkeys(filter(None, (row_material(r) for r in rows))))
        material_summary = ', '.join(materials[:3]) if materials else 'N/A'
        if len(materials) > 3:
            material_summary += ' 외'

        thicknesses = list(dict.fromkeys(filter(None, (_text(r.get('두께')) for r in rows))))
        thickness = ', '.join(thicknesses) if thicknesses else 'N/A'

        dwg_no = first_non_empty('도면번호') or first_non_empty('라인번호') or 'N/A'
        rev_no = first_non_empty('개정번호') or DEFAULT_METHOD_STANDARDS[cert_method][1]

        spec = first_non_empty('적용규격') or first_non_empty('규격') or ctx['applied_code']
        proc = first_non_empty('검사절차서')
        default_std, default_rev = DEFAULT_METHOD_STANDARDS[cert_method]
        if spec and proc:
            code_proc = f"{spec} / {proc}"
        elif spec:
            code_proc = f"{spec} / SIS" if '/' not in spec else spec
        else:
            code_proc = default_std
        if not first_non_empty('개정번호'):
            rev_no = default_rev

        batches.append({
            'date': target_date or dt.date.today().strftime('%Y-%m-%d'),
            'company': company,
            'customer': customer,
            'method': cert_method,
            'raw_method': _text(rows[0].get('검사방법')).upper() or cert_method,
            'report_no': report_no,
            'project_name': project_name,
            'item_name': item_name,
            'material': material_summary,
            'thickness': thickness,
            'drawing_no': dwg_no,
            'rev_no': rev_no,
            'code_procedure': code_proc,
            'standard_family': classify_standard_family(code_proc) or std_family,
            'standard_short': spec or ('ASME' if 'ASME' in code_proc.upper() else (std_family or 'KS')),
            'surface': first_non_empty('표면상태') or 'AS MACHINED',
            'requester': first_non_empty('의뢰자') or '담당자',
            'inspector': clean_worker_name(first_non_empty('작업자')) or ctx['workers'] or '주진철',
            'approver': '권인경',
            'equipment': first_non_empty('장비명') or ctx['equipment'],
            'aux_equipment': first_non_empty('보조장비') or ctx['aux_equipment'],
            'target_type': first_non_empty('대상유형') or '일반',
            'rows': rows,
        })
        # The daily record is the source for shared header fields. Saved PMI
        # criteria may supplement these fields, but must not replace them with
        # an older empty header or another equipment selection.
        daily_header = {
            'project_name': source_project,
            'equipment': first_non_empty('장비명') or ctx['equipment'],
            'aux_equipment': first_non_empty('보조장비') or ctx['aux_equipment'],
            'customer': first_non_empty('발주처') or ctx['customer'],
            'inspector': clean_worker_name(first_non_empty('작업자')) or ctx['workers'],
        }
        if spec or proc:
            daily_header['code_procedure'] = code_proc
            daily_header['standard_family'] = classify_standard_family(code_proc) or std_family
        batches[-1]['_daily_context'] = {key: val for key, val in daily_header.items()
                                        if val and str(val).strip().upper() not in ('N/A', '-')}

    return batches


def _write_cell(ws, coord: str, value) -> None:
    """Safely write value into coord, resolving MergedCell to its top-left anchor."""
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if coord in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    cell.value = value


def _template_dir() -> Path:
    return Path(__file__).resolve().parent.parent / 'resources' / 'cert_templates'


class CertificateExporter:
    """Populate method-specific Gyeongin test certificate templates from a batch dict."""

    def __init__(self, template_dir: Path | str | None = None):
        self.template_dir = Path(template_dir) if template_dir else _template_dir()

    def export_certificate(self, batch: dict, output_path: str | Path) -> Path:
        method = normalize_cert_method(batch.get('method') or batch.get('raw_method', ''))
        if not method:
            raise ValueError(f"지원하지 않는 검사방법입니다: {batch.get('method')}")

        std_family = classify_standard_family(
            batch.get('standard_family')
            or batch.get('code_procedure')
            or batch.get('standard_short')
            or ''
        )
        family_tmpl = self.template_dir / f"Template_Cert_{method}_{std_family}.xlsx" if std_family else None
        tmpl_path = family_tmpl if (family_tmpl is not None and family_tmpl.exists()) else (self.template_dir / f"Template_Cert_{method}.xlsx")
        if not tmpl_path.exists():
            raise FileNotFoundError(f"성적서 템플릿을 찾을 수 없습니다: {tmpl_path}")

        wb = openpyxl.load_workbook(tmpl_path)
        dispatch = {
            'MT': self._populate_mt,
            'PT': self._populate_pt,
            'UT': self._populate_ut,
            'RT': self._populate_rt,
            'PMI': self._populate_pmi,
        }
        dispatch[method](wb, batch)
        _sync_footer_form_code(wb, method, _text(batch.get('code_procedure')))

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out)
        footer_logo = self.template_dir.parent / "바닥글 PMI.png"
        if footer_logo.exists():
            self._inject_footer_logo_hf(out, footer_logo)
        return out

    @staticmethod
    def _inject_footer_logo_hf(xlsx_path: Path, logo_path: Path) -> None:
        """Attach Excel VML header/footer graphic (&G) to sheets with an &G oddFooter."""
        vml_template = (
            '<xml xmlns:v="urn:schemas-microsoft-com:vml"\n'
            ' xmlns:o="urn:schemas-microsoft-com:office:office"\n'
            ' xmlns:x="urn:schemas-microsoft-com:office:excel">\n'
            ' <o:shapelayout v:ext="edit">\n'
            '  <o:idmap v:ext="edit" data="1"/>\n'
            ' </o:shapelayout>\n'
            ' <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"\n'
            '  o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">\n'
            '  <v:stroke joinstyle="miter"/>\n'
            '  <v:formulas>\n'
            '   <v:f eqn="if lineDrawn pixelLineWidth 0"/>\n'
            '   <v:f eqn="sum @0 1 0"/>\n'
            '   <v:f eqn="sum 0 0 @1"/>\n'
            '   <v:f eqn="prod @2 1 2"/>\n'
            '   <v:f eqn="prod @3 21600 pixelWidth"/>\n'
            '   <v:f eqn="prod @3 21600 pixelHeight"/>\n'
            '   <v:f eqn="sum @0 0 1"/>\n'
            '   <v:f eqn="prod @6 1 2"/>\n'
            '   <v:f eqn="prod @7 21600 pixelWidth"/>\n'
            '   <v:f eqn="sum @8 21600 0"/>\n'
            '   <v:f eqn="prod @7 21600 pixelHeight"/>\n'
            '   <v:f eqn="sum @10 21600 0"/>\n'
            '  </v:formulas>\n'
            '  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>\n'
            '  <o:lock v:ext="edit" aspectratio="t"/>\n'
            ' </v:shapetype>\n'
            ' <v:shape id="RF" o:spid="_x0000_s1025" type="#_x0000_t75"\n'
            '  style="position:absolute;margin-left:0;margin-top:0;width:162pt;height:15.75pt;z-index:1">\n'
            '  <v:imagedata o:relid="rId1" o:title="SITCO 서울검사 주식회사"/>\n'
            ' </v:shape>\n'
            '</xml>\n'
        ).encode("utf-8")
        vml_rels_bytes = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            ' <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            'Target="../media/footer_sitco.png"/>\n'
            '</Relationships>\n'
        ).encode("utf-8")
        vml_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

        with zipfile.ZipFile(xlsx_path, "r") as zin:
            entries = {name: zin.read(name) for name in zin.namelist()}

        sheet_names = sorted(
            name for name in entries
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)
        )
        modified = False
        for s_name in sheet_names:
            xml_text = entries[s_name].decode("utf-8")
            if "&amp;G" not in xml_text or "<legacyDrawingHF" in xml_text:
                continue

            sheet_num = re.search(r"sheet(\d+)\.xml$", s_name).group(1)
            rels_name = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
            rels_text = entries.get(
                rels_name,
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            ).decode("utf-8")

            used_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_text)]
            next_rid = f"rId{(max(used_ids) + 1) if used_ids else 1}"
            vml_name = f"xl/drawings/vmlDrawingHF{sheet_num}.vml"
            vml_rels_name = f"xl/drawings/_rels/vmlDrawingHF{sheet_num}.vml.rels"

            rel_tag = f'<Relationship Id="{next_rid}" Type="{vml_rel_type}" Target="../drawings/vmlDrawingHF{sheet_num}.vml"/>'
            if "</Relationships>" in rels_text:
                rels_text = rels_text.replace("</Relationships>", f"{rel_tag}</Relationships>", 1)
            else:
                rels_text = rels_text.replace("/>", f">{rel_tag}</Relationships>", 1)
            entries[rels_name] = rels_text.encode("utf-8")
            entries[vml_name] = vml_template
            entries[vml_rels_name] = vml_rels_bytes

            if 'xmlns:r=' not in xml_text.split('>', 1)[0]:
                xml_text = xml_text.replace(
                    '<worksheet ',
                    '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                    1,
                )
            hf_tag = f'<legacyDrawingHF r:id="{next_rid}"/>'
            insert_match = re.search(r"<(?:picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|</worksheet>", xml_text)
            if not insert_match:
                continue
            pos = insert_match.start()
            xml_text = xml_text[:pos] + hf_tag + xml_text[pos:]
            entries[s_name] = xml_text.encode("utf-8")
            modified = True

        if not modified:
            return

        entries["xl/media/footer_sitco.png"] = logo_path.read_bytes()
        ct_text = entries["[Content_Types].xml"].decode("utf-8")
        if 'Extension="vml"' not in ct_text:
            ct_text = ct_text.replace(
                "</Types>",
                '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>',
                1,
            )
        if 'Extension="png"' not in ct_text:
            ct_text = ct_text.replace(
                "</Types>",
                '<Default Extension="png" ContentType="image/png"/></Types>',
                1,
            )
        entries["[Content_Types].xml"] = ct_text.encode("utf-8")

        tmp_path = xlsx_path.with_suffix(".tmp.xlsx")
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        os.replace(tmp_path, xlsx_path)

    @staticmethod
    def _chunk_rows(rows: list[dict], page_cap: int) -> list[list[dict]]:
        if not rows:
            return [[]]
        return [rows[i:i + page_cap] for i in range(0, len(rows), page_cap)]

    @staticmethod
    def _inspector_meta(name: str) -> tuple[str, str, str, str]:
        clean = clean_worker_name(name) or '주진철'
        eng, kor_grade, level_roman = INSPECTOR_PROFILES.get(clean, (clean, '기사', 'II'))
        return format_korean_spaced_name(clean), eng, kor_grade, level_roman

    def _populate_mt(self, wb, batch: dict) -> None:
        base = wb['MT(갑)']
        chunks = self._chunk_rows(batch.get('rows', []), 7)
        total_pages = len(chunks)
        d = parse_date_value(batch.get('date', ''))
        insp_kor, _, insp_grade, _ = self._inspector_meta(batch.get('inspector', ''))
        app_kor, _, app_grade, _ = self._inspector_meta(batch.get('approver', '권인경'))
        maker, model, sn = match_equipment_profile(batch.get('equipment', ''), 'MT')
        bl_maker, bl_model = match_black_light_profile(batch.get('aux_equipment', ''))
        is_asme = 'ASME' in _text(batch.get('code_procedure')).upper()
        default_interp = 'NO RECORDABLE INDICATION' if is_asme else '관련지시 없음'

        for page_idx, chunk in enumerate(chunks, start=1):
            ws = base if page_idx == 1 else wb.copy_worksheet(base)
            if page_idx > 1:
                ws.title = f"MT(갑) ({page_idx})"
            _write_cell(ws, 'V3', str(page_idx))
            _write_cell(ws, 'X3', str(total_pages))
            _write_cell(ws, 'S4', batch.get('customer', ''))
            _write_cell(ws, 'S6', batch.get('report_no', ''))
            _write_cell(ws, 'B9', batch.get('project_name', 'N/A'))
            _write_cell(ws, 'J9', batch.get('item_name', 'N/A'))
            _write_cell(ws, 'Q9', batch.get('drawing_no', 'N/A'))
            _write_cell(ws, 'X9', batch.get('rev_no', '0'))
            _write_cell(ws, 'B11', batch.get('material', 'N/A'))
            _write_cell(ws, 'E11', batch.get('thickness', 'N/A'))
            _write_cell(ws, 'J11', batch.get('surface', 'AS MACHINED'))
            _write_cell(ws, 'Q11', batch.get('code_procedure', DEFAULT_METHOD_STANDARDS['MT'][0]))
            _write_cell(ws, 'X11', batch.get('rev_no', '4'))
            _write_cell(ws, 'D13', maker)
            _write_cell(ws, 'D14', model)
            _write_cell(ws, 'D15', sn)
            if bl_maker:
                _write_cell(ws, 'O17', bl_maker)
            if bl_model:
                _write_cell(ws, 'O18', bl_model)
            _write_cell(ws, 'L39', d)
            _write_cell(ws, 'D40', insp_kor)
            _write_cell(ws, 'K40', insp_grade)
            _write_cell(ws, 'D43', app_kor)
            _write_cell(ws, 'K43', app_grade)

            header_mat = _text(batch.get('material'))
            for offset, row in enumerate(chunk):
                r = 29 + offset
                ident = row_identifier(row)
                mat = row_material(row)
                if ident and (not mat or mat == header_mat or mat in ident):
                    id_mat = ident
                elif ident and mat:
                    id_mat = f"{ident} / {mat}"
                else:
                    id_mat = ident or mat
                _write_cell(ws, f'B{r}', id_mat)
                if is_rejected(row.get('결과', '')):
                    _write_cell(ws, f'F{r}', '')
                    _write_cell(ws, f'H{r}', 'V')
                else:
                    _write_cell(ws, f'F{r}', 'V')
                    _write_cell(ws, f'H{r}', '')
                _write_cell(ws, f'K{r}', '1')
                interp = row_interpretation(row, default_interp)
                _write_cell(ws, f'L{r}', interp)
                _write_cell(ws, f'U{r}', row_remarks(row))

            blank_row = 29 + len(chunk)
            if blank_row <= 36:
                _write_cell(ws, f'L{blank_row}', '~  이  하  여  백  ~')

    def _populate_pt(self, wb, batch: dict) -> None:
        base = wb['PT(갑)']
        has_eul = 'PT(을)' in wb.sheetnames
        chunks = self._chunk_rows(batch.get('rows', []), 7)
        total_pages = len(chunks) + (1 if has_eul else 0)
        d = parse_date_value(batch.get('date', ''))
        insp_kor, insp_eng, insp_grade, _ = self._inspector_meta(batch.get('inspector', ''))
        app_kor, app_eng, app_grade, _ = self._inspector_meta(batch.get('approver', '진병학'))
        is_asme = 'ASME' in _text(batch.get('code_procedure')).upper()
        default_interp = 'NO RECORDABLE INDICATION' if is_asme else '관련지시 없음'

        gap_sheets = []
        for page_idx, chunk in enumerate(chunks, start=1):
            ws = base if page_idx == 1 else wb.copy_worksheet(base)
            if page_idx > 1:
                ws.title = f"PT(갑) ({page_idx})"
            _write_cell(ws, 'R3', f"{page_idx}    of    {total_pages}")
            _write_cell(ws, 'P4', batch.get('customer', ''))
            _write_cell(ws, 'P6', batch.get('report_no', ''))
            _write_cell(ws, 'A9', batch.get('project_name', 'N/A'))
            _write_cell(ws, 'F9', batch.get('item_name', 'N/A'))
            _write_cell(ws, 'L9', batch.get('drawing_no', 'N/A'))
            _write_cell(ws, 'S9', batch.get('rev_no', '0'))
            _write_cell(ws, 'A11', batch.get('material', 'N/A'))
            _write_cell(ws, 'D11', batch.get('thickness', 'N/A'))
            _write_cell(ws, 'F11', '20')
            _write_cell(ws, 'L11', batch.get('code_procedure', DEFAULT_METHOD_STANDARDS['PT'][0]))
            _write_cell(ws, 'S11', batch.get('rev_no', '2'))
            _write_cell(ws, 'J40', d)
            _write_cell(ws, 'D39', insp_eng)
            _write_cell(ws, 'D40', insp_kor)
            _write_cell(ws, 'H40', insp_grade)
            _write_cell(ws, 'D41', insp_eng)
            _write_cell(ws, 'D42', insp_kor)
            _write_cell(ws, 'H42', insp_grade)
            _write_cell(ws, 'D43', app_eng)
            _write_cell(ws, 'D44', app_kor)
            _write_cell(ws, 'H44', app_grade)

            header_mat = _text(batch.get('material'))
            for offset, row in enumerate(chunk):
                r = 30 + offset
                ident = row_identifier(row)
                mat = row_material(row)
                if ident and (not mat or mat == header_mat or mat in ident):
                    id_mat = ident
                elif ident and mat:
                    id_mat = f"{ident} ({mat})"
                else:
                    id_mat = ident or mat
                _write_cell(ws, f'A{r}', id_mat)
                if is_rejected(row.get('결과', '')):
                    _write_cell(ws, f'E{r}', '')
                    _write_cell(ws, f'F{r}', 'V')
                else:
                    _write_cell(ws, f'E{r}', 'V')
                    _write_cell(ws, f'F{r}', '')
                _write_cell(ws, f'G{r}', '1')
                interp = row_interpretation(row, default_interp)
                _write_cell(ws, f'H{r}', interp)
                _write_cell(ws, f'R{r}', row_remarks(row))

            blank_row = 30 + len(chunk)
            if blank_row <= 37:
                _write_cell(ws, f'H{blank_row}', '~ 이   하  여  백 ~')
            gap_sheets.append(ws)

        if has_eul:
            ws_eul = wb['PT(을)']
            _write_cell(ws_eul, 'Q3', f"{total_pages}    of")
            _write_cell(ws_eul, 'R3', str(total_pages))
            _write_cell(ws_eul, 'O4', batch.get('customer', ''))
            _write_cell(ws_eul, 'O6', batch.get('report_no', ''))

    def _populate_ut(self, wb, batch: dict) -> None:
        ws_gap = wb['UT(갑)']
        base_eul = wb['UT(을)']
        chunks = self._chunk_rows(batch.get('rows', []), 11)
        total_pages = 1 + len(chunks)
        d = parse_date_value(batch.get('date', ''))
        insp_kor, insp_eng, _, insp_lvl = self._inspector_meta(batch.get('inspector', ''))
        app_kor, app_eng, _, app_lvl = self._inspector_meta(batch.get('approver', '진병학'))
        _, model, sn = match_equipment_profile(batch.get('equipment', ''), 'UT')
        is_asme = 'ASME' in _text(batch.get('code_procedure')).upper()
        default_interp = 'NO RECORDABLE INDICATION' if is_asme else '관련지시 없음'

        _write_cell(ws_gap, 'L3', f"Page     1    of    {total_pages}    .")
        _write_cell(ws_gap, 'L4', batch.get('customer', ''))
        _write_cell(ws_gap, 'L6', batch.get('report_no', ''))
        _write_cell(ws_gap, 'A9', batch.get('project_name', 'N/A'))
        _write_cell(ws_gap, 'E9', batch.get('item_name', 'N/A'))
        _write_cell(ws_gap, 'J9', 'N/A')
        _write_cell(ws_gap, 'A11', batch.get('drawing_no', 'N/A'))
        _write_cell(ws_gap, 'D11', batch.get('rev_no', '0'))
        _write_cell(ws_gap, 'E11', batch.get('surface', 'AS ROUGH MACHINED'))
        _write_cell(ws_gap, 'J11', batch.get('code_procedure', DEFAULT_METHOD_STANDARDS['UT'][0]))
        _write_cell(ws_gap, 'A13', f"{batch.get('material', 'N/A')}    /")
        _write_cell(ws_gap, 'B13', batch.get('thickness', 'N/A'))
        _write_cell(ws_gap, 'E13', 'N/A')
        _write_cell(ws_gap, 'B15', model)
        _write_cell(ws_gap, 'F15', sn)
        _write_cell(ws_gap, 'B38', insp_eng)
        _write_cell(ws_gap, 'B39', insp_kor)
        _write_cell(ws_gap, 'G38', insp_lvl)
        _write_cell(ws_gap, 'B40', insp_eng)
        _write_cell(ws_gap, 'B41', insp_kor)
        _write_cell(ws_gap, 'G40', insp_lvl)
        _write_cell(ws_gap, 'B42', app_eng)
        _write_cell(ws_gap, 'B43', app_kor)
        _write_cell(ws_gap, 'G42', app_lvl)
        _write_cell(ws_gap, 'H39', d)

        for page_offset, chunk in enumerate(chunks):
            page_no = 2 + page_offset
            ws_eul = base_eul if page_offset == 0 else wb.copy_worksheet(base_eul)
            if page_offset > 0:
                ws_eul.title = f"UT(을) ({page_offset + 1})"
            _write_cell(ws_eul, 'L3', f"Page     {page_no}    of    {total_pages}   .")
            _write_cell(ws_eul, 'L4', batch.get('customer', ''))
            _write_cell(ws_eul, 'L6', batch.get('report_no', ''))

            for idx, row in enumerate(chunk):
                r = 10 + idx
                _write_cell(ws_eul, f'A{r}', row_identifier(row) or batch.get('item_name', ''))
                if is_rejected(row.get('결과', '')):
                    _write_cell(ws_eul, f'C{r}', '')
                    _write_cell(ws_eul, f'D{r}', 'V')
                else:
                    _write_cell(ws_eul, f'C{r}', 'V')
                    _write_cell(ws_eul, f'D{r}', '')
                interp = row_interpretation(row, default_interp)
                _write_cell(ws_eul, f'E{r}', interp)
                mat = row_material(row)
                rem = row_remarks(row)
                full_rem = f"{rem} / {mat}" if rem and mat and mat.casefold() not in rem.casefold() else (rem or mat)
                _write_cell(ws_eul, f'K{r}', full_rem)

            blank_row = 10 + len(chunk)
            if blank_row <= 21:
                _write_cell(ws_eul, f'E{blank_row}', '~ B  L  A  N  K ~')

    def _populate_rt(self, wb, batch: dict) -> None:
        ws_gap = wb['RT(갑)']
        base_eul = wb['RT(을)']
        chunks = self._chunk_rows(batch.get('rows', []), 23)
        total_pages = 1 + len(chunks)
        d = parse_date_value(batch.get('date', ''))
        insp_kor, _, insp_grade, _ = self._inspector_meta(batch.get('inspector', ''))
        app_kor, _, app_grade, _ = self._inspector_meta(batch.get('approver', '권인경'))

        _write_cell(ws_gap, 'N3', f"Page    1    of    {total_pages}    .")
        _write_cell(ws_gap, 'N4', batch.get('customer', ''))
        _write_cell(ws_gap, 'N6', batch.get('report_no', ''))
        _write_cell(ws_gap, 'C8', batch.get('project_name', 'N/A'))
        _write_cell(ws_gap, 'L8', batch.get('item_name', 'N/A'))
        _write_cell(ws_gap, 'C9', batch.get('material', 'N/A'))
        _write_cell(ws_gap, 'L9', batch.get('drawing_no', 'N/A'))
        _write_cell(ws_gap, 'P9', batch.get('rev_no', '0'))
        _write_cell(ws_gap, 'L31', batch.get('code_procedure', DEFAULT_METHOD_STANDARDS['RT'][0]))
        _write_cell(ws_gap, 'P30', f"Rev.{batch.get('rev_no', '0')}")
        _write_cell(ws_gap, 'C44', insp_kor)
        _write_cell(ws_gap, 'H44', insp_grade)
        _write_cell(ws_gap, 'C46', app_kor)
        _write_cell(ws_gap, 'H46', app_grade)
        _write_cell(ws_gap, 'C48', app_kor)
        _write_cell(ws_gap, 'H48', app_grade)
        _write_cell(ws_gap, 'K44', d)

        for page_offset, chunk in enumerate(chunks):
            page_no = 2 + page_offset
            ws_eul = base_eul if page_offset == 0 else wb.copy_worksheet(base_eul)
            if page_offset > 0:
                ws_eul.title = f"RT(을) ({page_offset + 1})"
            _write_cell(ws_eul, 'S2', f"Page      {page_no}      of      {total_pages}     .")
            _write_cell(ws_eul, 'R3', batch.get('customer', ''))
            _write_cell(ws_eul, 'R5', batch.get('report_no', ''))

            for idx, row in enumerate(chunk):
                r = 11 + idx
                _write_cell(ws_eul, f'A{r}', row_identifier(row) or str(idx + 1))
                _write_cell(ws_eul, f'B{r}', _text(row.get('구간정보')) or _text(row.get('구간')) or 'ALL')
                if is_rejected(row.get('결과', '')):
                    _write_cell(ws_eul, f'C{r}', '')
                    _write_cell(ws_eul, f'D{r}', 'V')
                else:
                    _write_cell(ws_eul, f'C{r}', 'V')
                    _write_cell(ws_eul, f'D{r}', '')
                _write_cell(ws_eul, f'U{r}', _text(row.get('용접사')))
                _write_cell(ws_eul, f'X{r}', _text(row.get('결과')) or '무결함')

            blank_row = 11 + len(chunk)
            if blank_row <= 34:
                _write_cell(ws_eul, f'E{blank_row}', '~ 이 하 여 백 ~')

    @staticmethod
    def _format_pmi_eul(ws) -> None:
        """Use one body style and the full Item No. column on every PMI page."""
        groups = ((1, 4), (5, 6), (7, 7), (8, 8), (9, 9), (10, 10),
                  (11, 11), (12, 12), (13, 13), (14, 14), (15, 15),
                  (16, 19), (20, 20))
        for merged in list(ws.merged_cells.ranges):
            if 20 <= merged.min_row <= merged.max_row <= 39 and merged.max_col <= 20:
                ws.unmerge_cells(str(merged))
        thin = Side(style='thin', color='000000')
        outer = Side(style='medium', color='000000')
        font = Font(name='바탕', size=10, color='000000')
        for r in range(20, 40):
            ws.row_dimensions[r].height = 28
            for first, last in groups:
                for c in range(first, last + 1):
                    cell = ws.cell(r, c)
                    cell.font = copy(font)
                    cell.alignment = Alignment(horizontal='center', vertical='center',
                                               wrap_text=True, shrink_to_fit=False)
                    cell.border = Border(
                        left=(outer if first == 1 else thin) if c == first else Side(),
                        right=(outer if last == 20 else thin) if c == last else Side(),
                        top=thin, bottom=outer if r == 39 else thin,
                    )
                if first != last:
                    ws.merge_cells(start_row=r, start_column=first,
                                   end_row=r, end_column=last)
            ws.cell(r, 16).number_format = 'yyyy-mm-dd'
        ws.column_dimensions['D'].hidden = False
        ws.column_dimensions['G'].width = 7
        ws.column_dimensions['N'].width = 13
        ws.column_dimensions['P'].width = 5.5
        ws['E19'].number_format = 'yyyy-mm-dd'
        ws['P19'].number_format = 'yyyy-mm-dd'
        ws.print_area = 'A1:T39'
        ws.print_options.gridLines = False
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        ws.sheet_view.showGridLines = False
        # Let Excel/printer metrics fit the entire print area, including T.
        # Column widths depend on the template's default font, so a fixed
        # percentage calculated from screen pixels can split off the right edge.
        ws.page_margins.left = ws.page_margins.right = 0.25
        ws.page_margins.top = ws.page_margins.bottom = 0.25
        ws.page_margins.header = ws.page_margins.footer = 0.15
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = 'portrait'
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_setup.scale = None

    @staticmethod
    def _align_pmi_page_frames(cover, detail) -> None:
        """Give both page types identical print width, height and header origin."""
        widths = []
        for col in range(1, 21):
            dimension = next(
                (item for item in detail.column_dimensions.values()
                 if (item.min or openpyxl.utils.column_index_from_string(item.index)) <= col
                 <= (item.max or openpyxl.utils.column_index_from_string(item.index))),
                None,
            )
            widths.append(dimension.width if dimension else detail.sheet_format.defaultColWidth or 13)
        # Expand grouped dimensions into individual columns on both sheets so
        # Excel resolves exactly the same physical columns, rather than merely
        # giving two differently sized tables the same page-margin settings.
        for sheet in (cover, detail):
            for key in list(sheet.column_dimensions):
                if openpyxl.utils.column_index_from_string(key) <= 20:
                    del sheet.column_dimensions[key]
            for col, width in enumerate(widths, 1):
                dimension = sheet.column_dimensions[get_column_letter(col)]
                dimension.width = width
                dimension.hidden = False

        def height(sheet, row):
            return sheet.row_dimensions[row].height or sheet.sheet_format.defaultRowHeight

        # Titles, page numbers and the company header share rows 1:10.
        for row in range(1, 11):
            cover.row_dimensions[row].height = height(detail, row)
        target = sum(height(detail, row) for row in range(11, 40))
        original = sum(height(cover, row) for row in range(11, 51))
        for row in range(11, 51):
            cover.row_dimensions[row].height = round(height(cover, row) * target / original, 4)
        residual = target - sum(height(cover, row) for row in range(11, 51))
        cover.row_dimensions[50].height += residual
        cover.print_area = 'A1:T50'
        cover.print_options.gridLines = False
        cover.sheet_view.showGridLines = False
        cover.page_setup.paperSize = detail.page_setup.paperSize
        cover.page_setup.orientation = detail.page_setup.orientation

    @staticmethod
    def _copy_pmi_eul(wb, source):
        """Worksheet copying omits images, print area and header/footer settings."""
        target = wb.copy_worksheet(source)
        target.print_area = 'A1:T39'
        target.HeaderFooter = deepcopy(source.HeaderFooter)
        for original in source._images:
            # Give every image its own stream: saving consumes the image data.
            payload = original._data()
            original.ref = BytesIO(payload)
            image = openpyxl.drawing.image.Image(BytesIO(payload))
            image.width, image.height = original.width, original.height
            image.anchor = deepcopy(original.anchor)
            target.add_image(image)
        return target

    def _populate_pmi(self, wb, batch: dict) -> None:
        ws_gap = wb['PMI(갑)']
        base_eul = wb['PMI(을)']
        # Use the same right-footer graphic as MT/PT/UT. Set this before
        # copying detail sheets so every continuation page inherits it.
        for sheet in (ws_gap, base_eul):
            sheet.oddFooter.right.text = '&G'
            sheet.HeaderFooter.alignWithMargins = False
        ws_gap.page_margins.left = ws_gap.page_margins.right = 0.25
        ws_gap.page_margins.top = ws_gap.page_margins.bottom = 0.25
        ws_gap.page_margins.header = ws_gap.page_margins.footer = 0.15
        ws_gap.print_options.horizontalCentered = True
        ws_gap.print_options.verticalCentered = True
        ws_gap.sheet_properties.pageSetUpPr.fitToPage = True
        ws_gap.page_setup.fitToWidth = ws_gap.page_setup.fitToHeight = 1
        ws_gap.page_setup.scale = None
        self._format_pmi_eul(base_eul)
        self._align_pmi_page_frames(ws_gap, base_eul)
        measurements = list(batch.get('pmi_measurements') or [])
        element_cols = choose_pmi_element_columns(measurements, batch.get('pmi_elements')) if measurements or batch.get('pmi_elements') else []
        criteria = batch.get('pmi_criteria') or []
        criteria_elements = choose_pmi_element_columns(measurements, batch.get('pmi_elements'))
        if criteria:
            from site_apps.gyeongin.src.pmi_criteria import validate_criteria
            validate_criteria(criteria, criteria_elements)
        source_rows = measurements if measurements else batch.get('rows', [])
        chunks = self._chunk_rows(source_rows, 20)
        criteria_pages = (max(0, len(criteria) - 4) + 3) // 4
        total_pages = 1 + len(chunks) + criteria_pages
        d = parse_date_value(batch.get('date', ''))
        insp_kor, insp_eng, _, _ = self._inspector_meta(batch.get('inspector', ''))
        _, model, sn = match_equipment_profile(batch.get('equipment', ''), 'PMI')

        _write_cell(ws_gap, 'N3', f"Page        1       of       {total_pages}        .")
        _write_cell(ws_gap, 'N5', batch.get('customer', ''))
        _write_cell(ws_gap, 'N8', batch.get('report_no', ''))
        _write_cell(ws_gap, 'E11', batch.get('project_name', 'N/A'))
        _write_cell(ws_gap, 'N11', batch.get('drawing_no', 'N/A'))
        _write_cell(ws_gap, 'E13', batch.get('item_name', 'N/A'))
        _write_cell(ws_gap, 'E15', batch.get('thickness', 'N/A'))
        _write_cell(ws_gap, 'E17', batch.get('material', 'N/A'))
        _write_cell(ws_gap, 'N17', batch.get('code_procedure', DEFAULT_METHOD_STANDARDS['PMI'][0]))
        _write_cell(ws_gap, 'E19', f"Alloy Analyzer 합금분석기 (XRF) / {model} (Serial No:{sn})")
        _write_cell(ws_gap, 'D42', insp_eng)
        _write_cell(ws_gap, 'D47', insp_kor)
        _write_cell(ws_gap, 'K42', d.strftime('%Y. %m. %d'))

        for page_offset, chunk in enumerate(chunks):
            page_no = 2 + page_offset
            ws_eul = base_eul if page_offset == 0 else self._copy_pmi_eul(wb, base_eul)
            if page_offset > 0:
                ws_eul.title = f"PMI(을) ({page_offset + 1})"
            # Clear copied measurements, including the previous blank marker.
            for cells in ws_eul.iter_rows(min_row=20, max_row=39, min_col=1, max_col=20):
                for cell in cells:
                    if not isinstance(cell, MergedCell):
                        cell.value = None
            _write_cell(ws_eul, 'O3', f"Page       {page_no}        of        {total_pages}     .")
            _write_cell(ws_eul, 'O5', batch.get('customer', ''))
            _write_cell(ws_eul, 'O8', batch.get('report_no', ''))
            _write_cell(ws_eul, 'E19', d)
            _write_cell(ws_eul, 'O19', clean_worker_name(batch.get('inspector', '')) or '주진철')
            _write_cell(ws_eul, 'P19', d)
            if element_cols:
                for i, slot in enumerate(_PMI_ELEMENT_SLOTS):
                    _write_cell(ws_eul, f'{slot}17', element_cols[i] if i < len(element_cols) else '-')

            for idx, row in enumerate(chunk):
                r = 20 + idx
                if measurements:
                    seq = page_offset * 20 + idx + 1
                    tag = row.get('tag') or batch.get('item_name', '')
                    comp_part = row.get('component') or f"Point-{seq}"
                    judged_mat = row.get('alloy') or 'MATCH'
                    _write_cell(ws_eul, f'A{r}', tag)
                    _write_cell(ws_eul, f'E{r}', comp_part)
                    _write_cell(ws_eul, f'G{r}', str(seq))
                    elems = row.get('elements') or {}
                    for slot, sym in zip(_PMI_ELEMENT_SLOTS, element_cols):
                        _write_cell(ws_eul, f'{slot}{r}', elems.get(sym, '') or '-')
                    _write_cell(ws_eul, f'N{r}', judged_mat)
                    _write_cell(ws_eul, f'O{r}', clean_worker_name(batch.get('inspector', '')) or '주진철')
                    _write_cell(ws_eul, f'P{r}', d)
                    continue
                tag = _text(row.get('태그번호')) or _text(row.get('식별번호')) or batch.get('item_name', '')
                comp_part = _text(row.get('구성부품')) or _text(row.get('검사품명')) or f"Point-{idx + 1}"
                judged_mat = _text(row.get('판정재질')) or _text(row.get('재질')) or _text(row.get('결과')) or 'MATCH'
                _write_cell(ws_eul, f'A{r}', tag)
                _write_cell(ws_eul, f'E{r}', comp_part)
                _write_cell(ws_eul, f'G{r}', str(page_offset * 20 + idx + 1))
                _write_cell(ws_eul, f'N{r}', judged_mat)
                _write_cell(ws_eul, f'O{r}', clean_worker_name(batch.get('inspector', '')) or '주진철')
                _write_cell(ws_eul, f'P{r}', d)
                _write_cell(ws_eul, f'T{r}', row_remarks(row))

            blank_row = 20 + len(chunk)
            if blank_row <= 39:
                ws_eul.merge_cells(f'H{blank_row}:M{blank_row}')
                _write_cell(ws_eul, f'H{blank_row}', '~ B  L  A  N  K ~')

        if criteria:
            from site_apps.gyeongin.src.pmi_criteria import render_criteria
            # Create continuation covers before changing the original blank area.
            extra_pages = []
            for offset in range(4, len(criteria), 4):
                page = self._copy_pmi_eul(wb, ws_gap)
                page.title = f'PMI(기준 {offset // 4 + 1})'
                page.print_area = 'A1:T50'
                extra_pages.append((page, offset))
            render_criteria(ws_gap, criteria[:4], criteria_elements)
            for index, (page, offset) in enumerate(extra_pages, 1):
                _write_cell(page, 'N3', f'Page        {1 + len(chunks) + index}       of       {total_pages}        .')
                _write_cell(page, 'E21', 'ACCEPTANCE CRITERIA (CONTINUED)')
                render_criteria(page, criteria[offset:offset + 4], criteria_elements, offset)


PMI_ELEMENT_SYMBOLS = (
    'Ni', 'Cr', 'Mo', 'Mn', 'Cu', 'Ti', 'Nb', 'V', 'W', 'Co', 'Fe', 'Si', 'Al',
    'Zr', 'Ta', 'Hf', 'Re', 'Sn', 'Zn', 'Pb', 'Bi', 'Sb', 'Ag', 'Au', 'Pd', 'Pt',
    'Mg', 'S', 'P', 'Se', 'As', 'Cd', 'Ga', 'Ge', 'B', 'C', 'N',
)
_PMI_FIXED_ELEMENTS = ('Ni', 'Cr', 'Mo')
_PMI_ELEMENT_SLOTS = ('H', 'I', 'J', 'K', 'L', 'M')
_PMI_TAG_KEYS = ('tag', 'sample', 'name', 'reading', '태그', '시료', '식별', '번호', 'id')
_PMI_ALLOY_KEYS = ('alloy', 'grade', 'match', 'material', '합금', '재질', '판정')
_PMI_COMPONENT_KEYS = ('component', 'description', 'part', 'location', 'note', '부품', '품명', '위치', '비고')


def _pmi_element_of(header) -> str:
    """'Ni', 'Ni %', 'Ni(%)', 'Ni Concentration' 같은 열 이름에서 원소 기호를 돌려줍니다. 오차 열은 제외."""
    text = _text(header)
    if not text:
        return ''
    low = text.lower()
    if any(tok in low for tok in ('+/-', '±', 'err', 'sigma', '2σ', 'σ', 'lod', 'unc')):
        return ''
    m = re.match(r'^([A-Z][a-z]?)(?:\s*(?:%|\(%\)|\[%\]|wt%?|concentration|conc\.?))?$', text.strip())
    if m and m.group(1) in PMI_ELEMENT_SYMBOLS:
        return m.group(1)
    return ''


def _pmi_value_text(value) -> str:
    text = _text(value)
    if not text or text.lower() in ('nan', 'none'):
        return ''
    if text.startswith('<') or 'lod' in text.lower() or text.upper() in ('ND', 'N.D', 'N/D'):
        return 'ND'
    try:
        num = float(text.replace('%', '').replace(',', ''))
    except ValueError:
        return text
    return f"{num:.2f}"


def read_pmi_measurements(path: str | Path) -> list[dict]:
    """PMI 장비가 내보낸 엑셀/CSV에서 측정 건별 태그·부품·판정합금·원소 함량을 읽습니다.

    머리글 행은 원소 기호 열이 3개 이상인 첫 행으로 자동 감지합니다(Vanta/Niton/XL2 등 공통).
    """
    import pandas as pd

    src = Path(path)
    suffix = src.suffix.lower()
    if suffix == '.csv':
        raw = None
        for enc in ('utf-8-sig', 'cp949', 'latin-1'):
            try:
                raw = pd.read_csv(src, header=None, dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        sheets = [raw] if raw is not None else []
    else:
        sheets = list(pd.read_excel(src, header=None, dtype=str, sheet_name=None).values())

    for raw in sheets:
        if raw is None or raw.empty:
            continue
        header_idx = None
        for idx in range(min(len(raw), 40)):
            if sum(1 for v in raw.iloc[idx].tolist() if _pmi_element_of(v)) >= 3:
                header_idx = idx
                break
        if header_idx is None:
            continue

        headers = [_text(v) for v in raw.iloc[header_idx].tolist()]
        element_cols: dict[str, int] = {}
        for col, head in enumerate(headers):
            sym = _pmi_element_of(head)
            if sym and sym not in element_cols:
                element_cols[sym] = col

        def find_col(keys) -> int | None:
            for key in keys:
                for col, head in enumerate(headers):
                    low = head.lower()
                    if col in element_cols.values() or not low:
                        continue
                    if key in low:
                        return col
            return None

        tag_col = find_col(_PMI_TAG_KEYS)
        alloy_col = find_col(_PMI_ALLOY_KEYS)
        comp_col = find_col(_PMI_COMPONENT_KEYS)

        results: list[dict] = []
        for idx in range(header_idx + 1, len(raw)):
            values = raw.iloc[idx].tolist()
            elements = {
                sym: _pmi_value_text(values[col]) for sym, col in element_cols.items() if col < len(values)
            }
            if not any(v and v != 'ND' for v in elements.values()):
                continue
            cell = lambda c: _text(values[c]) if c is not None and c < len(values) else ''
            results.append({
                'tag': '' if cell(tag_col).lower() == 'nan' else cell(tag_col),
                'component': '' if cell(comp_col).lower() == 'nan' else cell(comp_col),
                'alloy': '' if cell(alloy_col).lower() == 'nan' else cell(alloy_col),
                'elements': elements,
            })
        if results:
            return results
    return []


def available_pmi_elements(measurements: list[dict]) -> list[str]:
    """측정파일에 값이 있는 원소를 평균 함량 높은 순으로 돌려줍니다."""
    totals: dict[str, float] = {}
    seen: list[str] = []
    for m in measurements or []:
        for sym, val in (m.get('elements') or {}).items():
            if sym not in seen:
                seen.append(sym)
            try:
                totals[sym] = totals.get(sym, 0.0) + float(val)
            except (TypeError, ValueError):
                totals.setdefault(sym, 0.0)
    return sorted(seen, key=lambda s: -totals.get(s, 0.0))


def choose_pmi_element_columns(measurements: list[dict], preferred: list[str] | None = None) -> list[str]:
    """성적서 H~M 6칸에 넣을 원소.

    preferred(사용자 선택)가 있으면 그 순서대로 최대 6개, 없으면 Ni/Cr/Mo 고정 + 평균 함량 높은 순(Fe 제외).
    """
    if preferred:
        picked = [s for s in preferred if s]
        return picked[: len(_PMI_ELEMENT_SLOTS)]
    totals: dict[str, float] = {}
    for m in measurements or []:
        for sym, val in (m.get('elements') or {}).items():
            try:
                totals[sym] = totals.get(sym, 0.0) + float(val)
            except (TypeError, ValueError):
                continue
    extra = [
        sym for sym, _ in sorted(totals.items(), key=lambda kv: -kv[1])
        if sym not in _PMI_FIXED_ELEMENTS and sym != 'Fe' and totals[sym] > 0
    ]
    return list(_PMI_FIXED_ELEMENTS) + extra[: len(_PMI_ELEMENT_SLOTS) - len(_PMI_FIXED_ELEMENTS)]


def collect_ledger_records_from_history(
    history: dict,
    daily_usage_df=None,
    company_filter: str = '',
) -> list[dict]:
    """Build issue-ledger rows chronologically across all saved work_history dates."""
    target_company = _text(company_filter).casefold()
    records: list[dict] = []

    for date_key in sorted(history.keys()):
        day = history.get(date_key, {})
        ndt_rows = day.get('ndt_results', [])
        batches = build_cert_batches(ndt_rows, target_date=date_key, daily_usage_df=daily_usage_df)
        for batch in batches:
            if target_company and batch['company'].casefold() != target_company:
                continue
            project = batch['project_name']
            item = batch['item_name']
            proj_item = f"{project} / {item}" if project not in ('', 'N/A') and item not in ('', 'N/A') else (item if item not in ('', 'N/A') else project)
            records.append({
                'date': date_key,
                'company': batch['company'],
                'report_no': batch['report_no'],
                'customer': batch['customer'],
                'project_item': proj_item,
                'requester': batch['requester'],
                'issuer': batch['inspector'],
                'reissue_date': '',
                'reissue_check': '',
                'remark': f"{batch['standard_short']} ({batch['method']})",
            })
    return records


class CertificateLedgerExporter:
    """Create or append customer issue ledger workbooks (성적서 발행대장)."""

    def __init__(self, template_dir: Path | str | None = None):
        self.template_dir = Path(template_dir) if template_dir else _template_dir()

    def export_ledger(
        self,
        records: list[dict],
        output_path: str | Path,
        company_name: str = '',
        year: str = '2026',
        existing_ledger_path: str | Path | None = None,
    ) -> Path:
        out = Path(output_path)
        source_path = Path(existing_ledger_path) if existing_ledger_path else (self.template_dir / 'Template_Cert_Ledger.xlsx')
        if not source_path.exists():
            raise FileNotFoundError(f"발행대장 템플릿을 찾을 수 없습니다: {source_path}")

        wb = openpyxl.load_workbook(source_path)
        ws = wb.active

        if not existing_ledger_path:
            title_prefix = f"{year}년 {company_name}".strip() if company_name else f"{year}년"
            _write_cell(ws, 'A1', f"{title_prefix} 시험성적서 발행대장")

        thin = Side(style='thin', color='000000')
        border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        font_norm = Font(name='맑은 고딕', size=10)
        font_month = Font(name='맑은 고딕', size=10, bold=True)
        fill_month = PatternFill(start_color='E8EEF5', end_color='E8EEF5', fill_type='solid')
        align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_l = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # Read existing signatures when appending so we do not duplicate existing report rows
        existing_keys: set[tuple[str, str, str]] = set()
        max_row = ws.max_row
        start_scan = 5
        last_data_row = 4
        for r in range(start_scan, max_row + 1):
            row_vals = [ws.cell(r, c).value for c in range(1, 11)]
            if any(_text(v) for v in row_vals):
                last_data_row = r
                b_val = _text(row_vals[1])[:10]
                c_val = _text(row_vals[2])
                e_val = _text(row_vals[4])
                if c_val:
                    existing_keys.add((b_val, c_val, e_val))

        new_records = []
        for rec in records:
            d_txt = parse_date_value(rec.get('date', '')).strftime('%Y-%m-%d')
            key = (d_txt, _text(rec.get('report_no')), _text(rec.get('project_item')))
            if key not in existing_keys:
                existing_keys.add(key)
                new_records.append((d_txt, rec))

        # Group cleanly by month if starting fresh, or continue sequentially after last row
        row_cursor = last_data_row + 1
        current_month_label = ''
        month_seq = 0
        if not existing_ledger_path:
            # New workbook: write month divider headers (1월, 2월, ...)
            for d_txt, rec in sorted(new_records, key=lambda item: (item[0], _text(item[1].get('report_no')))):
                month_label = f"{int(d_txt[5:7])}월"
                if month_label != current_month_label:
                    current_month_label = month_label
                    month_seq = 0
                    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=10)
                    mcell = ws.cell(row_cursor, 1, value=month_label)
                    mcell.font = font_month
                    mcell.alignment = Alignment(horizontal='left', vertical='center')
                    for c in range(1, 11):
                        cell = ws.cell(row_cursor, c)
                        cell.fill = fill_month
                        cell.border = border_thin
                    ws.row_dimensions[row_cursor].height = 20
                    row_cursor += 1

                month_seq += 1
                self._write_ledger_row(ws, row_cursor, month_seq, d_txt, rec, font_norm, border_thin, align_c, align_l)
                row_cursor += 1
        else:
            # Appending to existing company workbook
            seq = sum(1 for r in range(5, last_data_row + 1) if str(_text(ws.cell(r, 1).value)).rstrip('.0').isdigit())
            for d_txt, rec in new_records:
                seq += 1
                self._write_ledger_row(ws, row_cursor, seq, d_txt, rec, font_norm, border_thin, align_c, align_l)
                row_cursor += 1

        out.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out)
        return out

    @staticmethod
    def _write_ledger_row(ws, row_idx: int, seq: int, d_txt: str, rec: dict, font, border, align_c, align_l) -> None:
        values = [
            seq,
            parse_date_value(d_txt),
            _text(rec.get('report_no')),
            _text(rec.get('customer') or rec.get('company')),
            _text(rec.get('project_item')),
            _text(rec.get('requester')),
            _text(rec.get('issuer')),
            _text(rec.get('reissue_date')),
            _text(rec.get('reissue_check')),
            _text(rec.get('remark')),
        ]
        ws.row_dimensions[row_idx].height = 22
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row_idx, col_idx, value=val)
            cell.font = font
            cell.border = border
            cell.alignment = align_l if col_idx in (4, 5, 10) else align_c
            if col_idx == 2:
                cell.number_format = 'yyyy-mm-dd'

