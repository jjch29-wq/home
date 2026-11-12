"""Gyeongin customer test certificate (갑/을) and issue ledger Excel exporters."""
from __future__ import annotations

import datetime as dt
import os
import re
import zipfile
from copy import copy
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


def match_equipment_profile(equipment_text: str, method: str) -> tuple[str, str, str]:
    raw = _text(equipment_text)
    default_maker, default_model, default_sn = EQUIPMENT_PROFILES.get(method, ('-', '-', '-'))
    if not raw or raw == '해당없음':
        return default_maker, default_model, default_sn
    sn_match = re.search(r'\(([^)]+)\)', raw)
    sn = sn_match.group(1).strip() if sn_match else default_sn
    if 'MY-2' in raw.upper():
        parts = sn.split()
        return 'NAWOO', 'MY-2', parts[-1] if parts else default_sn
    if 'XL2' in raw.upper():
        return 'THERMO', 'NITON XL2 100G', sn
    if 'MASTERSCAN' in raw.upper() or 'OMNI' in raw.upper():
        return 'OLYMPUS/SONATEST', raw, sn
    return default_maker, raw, sn


def find_daily_usage_context(daily_usage_df, date_str: str, company: str, method: str) -> dict:
    """Lookup matching DailyUsage fields for (date, company, method) without mutating df."""
    empty = {
        'site': '',
        'customer': '',
        'equipment': '',
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
        matched_rows.append(row)

    # Relax method match if exact method row was not recorded yet for that date+company
    if not matched_rows and target_comp:
        for _, row in df.iterrows():
            raw_d = _text(row.get('Date') or row.get('날짜'))
            if not raw_d or parse_date_value(raw_d) != target_date:
                continue
            row_comp = _text(row.get('업체명') or row.get('업체')).casefold()
            if row_comp == target_comp:
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

    site = first_non_empty('Site', '현장')
    if site == '경인사무소':
        site = ''
    return {
        'site': site,
        'customer': first_non_empty('발주처'),
        'equipment': first_non_empty('장비명'),
        'workers': clean_worker_name(first_non_empty('Workers', 'Worker1', '작업자')),
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
    grouped: dict[tuple[str, str, str], list[dict]] = {}
    order_keys: list[tuple[str, str, str]] = []

    for row in ndt_results or []:
        if not any(_text(v) for v in row.values()):
            continue
        method_raw = _text(row.get('검사방법')).upper()
        cert_method = normalize_cert_method(method_raw)
        if not cert_method:
            continue
        company = _text(row.get('업체')) or '경인사무소'
        explicit_no = _text(row.get('성적서번호'))
        key = (company, cert_method, explicit_no)
        if key not in grouped:
            grouped[key] = []
            order_keys.append(key)
        grouped[key].append(dict(row))

    batches: list[dict] = []
    seq_by_pair: dict[tuple[str, str], int] = {}

    for company, cert_method, explicit_no in order_keys:
        rows = grouped[(company, cert_method, explicit_no)]
        pair = (company, cert_method)
        seq_by_pair[pair] = seq_by_pair.get(pair, 0) + 1
        seq = seq_by_pair[pair]

        ctx = find_daily_usage_context(daily_usage_df, target_date, company, cert_method)
        report_no = explicit_no or ctx['report_no'] or default_report_no(target_date, company, cert_method, seq)

        first_non_empty = lambda k: next((_text(r.get(k)) for r in rows if _text(r.get(k))), '')
        customer = first_non_empty('발주처') or ctx['customer'] or company
        project_name = first_non_empty('공사명') or ctx['site'] or (_text(default_site) if _text(default_site) != '경인사무소' else '') or 'N/A'

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
            'standard_short': spec or ('ASME' if 'ASME' in code_proc.upper() else 'KS'),
            'surface': first_non_empty('표면상태') or 'AS MACHINED',
            'requester': first_non_empty('의뢰자') or '담당자',
            'inspector': ctx['workers'] or '주진철',
            'approver': '권인경',
            'equipment': ctx['equipment'],
            'target_type': first_non_empty('대상유형') or '일반',
            'rows': rows,
        })

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

        tmpl_path = self.template_dir / f"Template_Cert_{method}.xlsx"
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

    def _populate_pmi(self, wb, batch: dict) -> None:
        ws_gap = wb['PMI(갑)']
        base_eul = wb['PMI(을)']
        chunks = self._chunk_rows(batch.get('rows', []), 15)
        total_pages = 1 + len(chunks)
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
            ws_eul = base_eul if page_offset == 0 else wb.copy_worksheet(base_eul)
            if page_offset > 0:
                ws_eul.title = f"PMI(을) ({page_offset + 1})"
            _write_cell(ws_eul, 'O3', f"Page       {page_no}        of        {total_pages}     .")
            _write_cell(ws_eul, 'O5', batch.get('customer', ''))
            _write_cell(ws_eul, 'O8', batch.get('report_no', ''))
            _write_cell(ws_eul, 'E19', d)
            _write_cell(ws_eul, 'O19', clean_worker_name(batch.get('inspector', '')) or '주진철')
            _write_cell(ws_eul, 'P19', d)

            for idx, row in enumerate(chunk):
                r = 20 + idx
                tag = _text(row.get('태그번호')) or _text(row.get('식별번호')) or batch.get('item_name', '')
                comp_part = _text(row.get('구성부품')) or _text(row.get('검사품명')) or f"Point-{idx + 1}"
                judged_mat = _text(row.get('판정재질')) or _text(row.get('재질')) or _text(row.get('결과')) or 'MATCH'
                _write_cell(ws_eul, f'A{r}', tag)
                _write_cell(ws_eul, f'E{r}', comp_part)
                _write_cell(ws_eul, f'G{r}', str(idx + 1))
                _write_cell(ws_eul, f'N{r}', judged_mat)
                _write_cell(ws_eul, f'O{r}', clean_worker_name(batch.get('inspector', '')) or '주진철')
                _write_cell(ws_eul, f'P{r}', d)
                _write_cell(ws_eul, f'T{r}', row_remarks(row))

            blank_row = 20 + len(chunk)
            if blank_row <= 36:
                _write_cell(ws_eul, f'H{blank_row}', '~ B  L  A  N  K ~')


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

