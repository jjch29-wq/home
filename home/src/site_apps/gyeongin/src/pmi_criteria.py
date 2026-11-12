"""PMI material catalogue, report snapshots and printable acceptance tables."""
from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from copy import deepcopy
from pathlib import Path

ELEMENT_NAMES = dict(zip(
    'Ni Cr Mo Mn Cu Ti Nb V W Co Fe Si Al C N P S Zr Ta'.split(),
    ['Nickel', 'Chromium', 'Molybdenum', 'Manganese', 'Copper', 'Titanium',
     'Niobium', 'Vanadium', 'Tungsten', 'Cobalt', 'Iron', 'Silicon', 'Aluminium',
     'Carbon', 'Nitrogen', 'Phosphorus', 'Sulfur', 'Zirconium', 'Tantalum']))


def material_catalogue():
    """Names are searchable candidates; no unverified composition is inferred."""
    groups = [
        ('스테인리스', 'KS / 제품규격 확인', '판·봉·PIPE',
         'STS304 STS304L STS316 STS316L STS321 STS347 STS310S STS410 STS420J2 STS430'),
        ('스테인리스', 'JIS / 제품규격 확인', '판·봉',
         'SUS304 SUS304L SUS316 SUS316L SUS321 SUS347 SUS310S SUS410 SUS430'),
        ('스테인리스', 'JIS G 3459', 'PIPE',
         'SUS304TP SUS304LTP SUS316TP SUS316LTP SUS321TP SUS347TP'),
        ('스테인리스', 'ASTM A312', 'PIPE',
         'TP304 TP304L TP304H TP316 TP316L TP316H TP321 TP347 TP310S'),
        ('스테인리스', 'ASTM A403', '맞대기 피팅',
         'WP304 WP304L WP316 WP316L WP321 WP347'),
        ('스테인리스', 'ASTM A182', '단조 피팅·플랜지',
         'F304 F304L F316 F316L F321 F347'),
        ('스테인리스', 'ASTM A351', '주강', 'CF8 CF8M CF3 CF3M'),
        ('탄소강', 'KS D 3503', '구조용', 'SS275 SS315 SS410'),
        ('탄소강', 'KS D 3515', '용접 구조용',
         'SM275A SM275B SM275C SM355A SM355B SM355C'),
        ('탄소강', 'KS D 3507', 'PIPE', 'SPP'),
        ('탄소강', 'KS D 3566', '원형 강관', 'SGT275 SGT355'),
        ('탄소강', 'KS D 3568', '각형 강관', 'SRT275 SRT355'),
        ('탄소강', 'JIS G 3101', '구조용', 'SS400 SS490'),
        ('탄소강', 'JIS G 3106', '용접 구조용', 'SM490A SM490B'),
        ('탄소강', 'ASTM A106', 'PIPE', 'Gr.B Gr.C'),
        ('탄소강', 'ASTM A53', 'PIPE', 'Gr.A Gr.B'),
        ('탄소강', 'API 5L / PSL 확인', 'PIPE', 'Gr.B X42 X52 X60 X65 X70'),
        ('탄소강', 'ASTM A234', '맞대기 피팅', 'WPB WPC'),
        ('탄소강', 'ASTM A105', '단조 피팅·플랜지', 'A105'),
        ('탄소강', 'ASTM A333', '저온 PIPE', 'Gr.6'),
        ('탄소강', 'ASTM A420', '저온 피팅', 'WPL6'),
        ('탄소강', 'ASTM A350', '저온 단조품', 'LF2'),
        ('저합금강', 'ASTM A335', 'PIPE', 'P5 P9 P11 P12 P22 P91 P92'),
        ('저합금강', 'ASTM A234', '피팅', 'WP5 WP9 WP11 WP22 WP91'),
        ('저합금강', 'ASTM A182', '단조품', 'F5 F9 F11 F22 F91'),
        ('듀플렉스', 'ASTM A790', 'PIPE', 'S31803 S32205 S32750 S32760'),
        ('듀플렉스', 'ASTM A815', '피팅', 'S31803 S32205 S32750 S32760'),
        ('듀플렉스', 'ASTM A182', '단조품', 'F51 F53 F55 F60'),
        ('니켈합금', '제품규격 확인', 'PIPE·피팅·기타',
         'N04400 N06600 N06625 N08800 N08810 N08825 N10276'),
        ('기타', '제품규격 확인', '기타', 'Titanium Copper Aluminium'),
    ]
    result = []
    for family, standard, product, grades in groups:
        for grade in grades.split():
            label = f'{standard} {grade} — {product}'
            result.append(dict(id=label, label=label, family=family,
                               standard=standard, grade=grade, product=product,
                               ranges={}, source=''))
    # These are the user's supplied report criteria, not inferred equivalents
    # or a claim that every edition of A351 uses the same project limits.
    supplied = {
        'CF3M': {'Ni': '9.00~13.00', 'Cr': '17.00~21.00', 'Mo': '2.00~3.00'},
        'CF8': {'Ni': '8.00~11.00', 'Cr': '18.00~21.00', 'Mo': '-'},
    }
    for item in result:
        if item['standard'] == 'ASTM A351' and item['grade'] in supplied:
            item['ranges'] = deepcopy(supplied[item['grade']])
            item['source'] = '사용자 제공 합격기준표 / 프로젝트 적용규격·판본 확인'
    from site_apps.gyeongin.src.pmi_reference_data import public_material_references
    indexed = {item['id']: item for item in result}
    indexed.update({item['id']: item for item in public_material_references()})
    return list(indexed.values())


def normalize_ranges(ranges):
    """Accept numeric limits, an explicitly unspecified dash, or Fe balance."""
    out = {}
    for symbol, raw in ranges.items():
        text = str(raw).strip().replace('～', '~').replace('≤', '<=').replace('≥', '>=')
        text = text.replace('%', '').strip()
        if not text:
            continue
        if text in ('-', '—'):
            out[symbol] = '-'
            continue
        if text.lower() in ('bal', 'balance', '잔부'):
            out[symbol] = 'Balance'
            continue
        match = re.fullmatch(r'(<=|>=)?\s*(\d+(?:\.\d+)?)\s*(?:~\s*(\d+(?:\.\d+)?))?\s*(이하|이상)?', text)
        if not match:
            raise ValueError(f'{symbol}: 0.90~1.50, <=1.50, >=0.90, -, Balance 형식으로 입력하세요.')
        op, lo, hi, suffix = match.groups()
        if not hi and not op and not suffix:
            raise ValueError(f'{symbol}: 단일 값은 <= 또는 >=로 한계를 지정하세요.')
        if hi and (op or suffix or float(lo) > float(hi)):
            raise ValueError(f'{symbol}: 하한과 상한을 확인하세요.')
        if any(float(v) > 100 for v in (lo, hi) if v):
            raise ValueError(f'{symbol}: 함량은 100% 이하여야 합니다.')
        if op and suffix:
            raise ValueError(f'{symbol}: 한계 기호를 하나만 사용하세요.')
        out[symbol] = f'{lo}~{hi}' if hi else f'{op or ("<=" if suffix == "이하" else ">=")}{lo}'
    return out


REPORT_FIELDS = (
    'report_no', 'customer', 'project_name', 'code_procedure', 'standard_family',
    'inspector', 'equipment', 'aux_equipment', 'pmi_criteria', 'pmi_elements',
    'pmi_measurements', 'pmi_source',
)
PROJECT_FIELDS = tuple(k for k in REPORT_FIELDS if k not in
                       ('report_no', 'pmi_measurements', 'pmi_source'))


def state_key(batch):
    identity = [str(batch.get(k, '')) for k in ('date', 'company', 'report_no')]
    identity += [str(batch.get('project_name', '')).strip(), str(batch.get('method', 'PMI'))]
    return hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()


def project_key(batch):
    name = str(batch.get('project_name', '')).strip()
    if name.upper() in ('', 'N/A', 'NA', '-'):
        return None
    identity = [str(batch.get('company', '')).strip(), name, str(batch.get('method', 'PMI'))]
    return hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()


def read_state(path):
    path = Path(path)
    if not path.exists():
        return {'reports': {}, 'presets': {}, 'projects': {}}
    data = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(data, dict) or not all(isinstance(data.get(k), dict) for k in ('reports', 'presets')):
        raise ValueError('PMI 기준 저장 파일 형식을 확인하세요.')
    data.setdefault('projects', {})
    if not isinstance(data['projects'], dict):
        raise ValueError('PMI 현장 저장 파일 형식을 확인하세요.')
    return data


def write_state(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=path.stem, suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def restore_report(batch, path):
    data = read_state(path)
    snapshot = data['reports'].get(state_key(batch))
    if snapshot is None:
        identity = [str(batch.get(k, '')) for k in ('date', 'company', 'report_no')]
        legacy_key = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()
        legacy = data['reports'].get(legacy_key, {})
        # Old records had no project identity. Do not assume an unnamed legacy
        # record belongs to a specifically named project.
        if not project_key(batch) or legacy.get('project_name') == batch.get('project_name'):
            snapshot = legacy or None
    if snapshot is None:
        restore_project(batch, path)
        batch.update(deepcopy(batch.get('_daily_context', {})))
        return
    for key in REPORT_FIELDS:
        if key in snapshot:
            batch[key] = deepcopy(snapshot[key])
    batch.update(deepcopy(batch.get('_daily_context', {})))


def saved_projects(batch, path):
    return sorted({item['project_name'] for item in read_state(path)['projects'].values()
                   if item.get('company') == batch.get('company')
                   and item.get('method', 'PMI') == batch.get('method', 'PMI')})


def restore_project(batch, path, name=None):
    target = dict(batch)
    if name is not None:
        target['project_name'] = name
    snapshot = read_state(path)['projects'].get(project_key(target), {})
    if not snapshot:
        return False
    for key in PROJECT_FIELDS:
        if key in snapshot:
            batch[key] = deepcopy(snapshot[key])
        elif key in ('pmi_criteria', 'pmi_elements'):
            batch.pop(key, None)
    return bool(snapshot)


def save_report(batch, path, original_key=None):
    data = read_state(path)
    snapshot = {key: deepcopy(batch[key]) for key in REPORT_FIELDS if key in batch}
    data['reports'][state_key(batch)] = snapshot
    if original_key:
        data['reports'][original_key] = deepcopy(snapshot)
    project = project_key(batch)
    if project:
        profile = {key: deepcopy(batch[key]) for key in PROJECT_FIELDS if key in batch}
        profile.update(company=batch.get('company'), method=batch.get('method', 'PMI'))
        data['projects'][project] = profile
    for item in batch.get('pmi_criteria', []):
        data['presets'][item['id']] = deepcopy(item)
    write_state(path, data)


def validate_criteria(criteria, elements):
    if not elements or len(elements) > 6 or len(set(elements)) != len(elements):
        raise ValueError('출력 원소를 중복 없이 1~6개 선택하세요.')
    for item in criteria:
        if not item.get('label', '').strip() or not item.get('source', '').strip():
            raise ValueError('재질명과 기준 출처/적용 판본을 입력하세요.')
        ranges = normalize_ranges(item.get('ranges', {}))
        missing = [s for s in elements if s not in ranges]
        if missing:
            raise ValueError(f'{item["label"]}: {", ".join(missing)} 기준을 입력하세요. 규정 없음은 -로 지정하세요.')


def render_criteria(ws, criteria, elements, start_index=0):
    """Replace only the cover's blank criteria region; keep its page frame."""
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils import get_column_letter
    from copy import copy
    # Equal-height table rows, preserving the cover's overall printed height.
    heights = [ws.row_dimensions[r].height or ws.sheet_format.defaultRowHeight
               for r in range(24, 38)]
    average = sum(heights) / len(heights)
    for rr in range(24, 38):
        ws.row_dimensions[rr].height = average
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 23 and merged.max_row <= 38 and merged.max_col <= 20:
            ws.unmerge_cells(str(merged))
    ws._images = [img for img in ws._images
                  if not (hasattr(img.anchor, '_from') and 22 <= img.anchor._from.row < 38
                          and img.anchor._from.col < 20)]
    for row in ws.iter_rows(min_row=23, max_row=38, max_col=20):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.border = Border()
    dense = len(criteria) > 2
    panel_rows = 3 + 1 + len(elements) + 1 + 4
    end_region = 38
    body_size = 11
    if dense:
        # Subdivide the blank area into finer native Excel rows. Its physical
        # height stays unchanged; signatures and the print boundary move together.
        area_height = sum(ws.row_dimensions[r].height or ws.sheet_format.defaultRowHeight
                          for r in range(23, 39))
        extra = 2 * panel_rows - 16
        remaining_merges = [copy(rng) for rng in ws.merged_cells.ranges]
        dimensions = {r: copy(dim) for r, dim in ws.row_dimensions.items() if r >= 23}
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        ws.insert_rows(23, extra)
        for rng in remaining_merges:
            if rng.min_row >= 23:
                rng.shift(row_shift=extra)
            ws.merge_cells(str(rng))
        for r in list(ws.row_dimensions):
            if r >= 23:
                del ws.row_dimensions[r]
        for r, dim in dimensions.items():
            if r >= 39:
                dim.index = r + extra
                ws.row_dimensions[r + extra] = dim
        end_region = 38 + extra
        row_height = area_height / (2 * panel_rows)
        for rr in range(23, end_region + 1):
            ws.row_dimensions[rr].height = row_height
        for img in ws._images:
            if hasattr(img.anchor, '_from') and img.anchor._from.row >= 22:
                img.anchor._from.row += extra
                if hasattr(img.anchor, 'to'):
                    img.anchor.to.row += extra
        ws.print_area = f'A1:T{50 + extra}'
        body_size = max(9, min(11, int(row_height - 0.5)))
    thin = Side(style='thin', color='000000')
    for rr in range(23, end_region + 1):
        ws.cell(rr, 1).border = Border(left=thin)
        ws.cell(rr, 20).border = Border(right=thin)
    for idx, item in enumerate(criteria):
        first, last = (1, 10) if idx % 2 == 0 else (11, 20)
        title_row = 23 + (idx // 2) * panel_rows if dense else 24
        header_row = title_row + 3
        first_element_row = header_row + 1
        panel_end = title_row + panel_rows - 1 if dense else 38
        widths = {c: ws.column_dimensions[get_column_letter(c)].width or 13
                  for c in range(first, last + 1)}
        half = sum(widths.values()) / 2
        split = min(range(first, last), key=lambda c: abs(
            sum(widths[k] for k in range(first, c + 1)) - half))

        def block(row, end_row, left, right, text, bold=False, size=None):
            ws.merge_cells(start_row=row, end_row=end_row, start_column=left, end_column=right)
            for rr in range(row, end_row + 1):
                for cc in range(left, right + 1):
                    ws.cell(rr, cc).border = Border(left=thin if cc == left else Side(),
                        right=thin if cc == right else Side(), top=thin if rr == row else Side(),
                        bottom=thin if rr == end_row else Side())
            cell = ws.cell(row, left, text)
            cell.data_type = 's'
            cell.font = Font(name='Times New Roman', size=size or body_size, bold=bold)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        block(title_row, title_row + 2, first, last,
              f'{start_index + idx + 1}. Acceptance Criteria (%)\n{item["label"]}', True)
        block(header_row, header_row, first, split, 'Element', True)
        block(header_row, header_row, split + 1, last, 'Composition, %', True)
        normalized = normalize_ranges(item['ranges'])
        for offset, symbol in enumerate(elements):
            rr = first_element_row + offset
            block(rr, rr, first, split, ELEMENT_NAMES.get(symbol, symbol))
            block(rr, rr, split + 1, last, normalized[symbol])
        bottom = first_element_row + len(elements)
        block(bottom, bottom, first, last, '* Tolerance ±10%')
        source_text = item['source']
        if item.get('user_modified'):
            source_text += ' / 사용자 수정'
        if item.get('notes'):
            source_text += '\n' + item['notes']
        block(bottom + 1, panel_end, first, last, source_text, size=8)
        if item.get('source_url', '').startswith(('https://', 'http://')):
            ws.cell(bottom + 1, first).hyperlink = item['source_url']


def edit_criteria_dialog(parent, batch, state_path):
    """Edit a report's independent material/limit snapshots."""
    import tkinter as tk
    from tkinter import ttk, messagebox
    from site_apps.gyeongin.src.cert_exporter import choose_pmi_element_columns
    elements = choose_pmi_element_columns(batch.get('pmi_measurements', []), batch.get('pmi_elements'))
    elements = elements[:6]
    saved = read_state(state_path)
    builtin_catalogue = {i['id']: i for i in material_catalogue()}
    catalogue = deepcopy(builtin_catalogue)
    catalogue.update(saved['presets'])
    chosen = deepcopy(batch.get('pmi_criteria', []))
    win = tk.Toplevel(parent)
    win.title('PMI 재질·합격기준')
    win.transient(parent)
    win.grab_set()
    frm = ttk.Frame(win, padding=12)
    frm.pack(fill='both', expand=True)
    family = tk.StringVar(value='전체')
    query = tk.StringVar()
    ttk.Label(frm, text='분류').grid(row=0, column=0, sticky='w')
    families = ['전체'] + sorted({i['family'] for i in catalogue.values()})
    ttk.Combobox(frm, textvariable=family, values=families, state='readonly', width=14).grid(row=0, column=1, sticky='w')
    ttk.Label(frm, text='검색 / 재질 직접 입력').grid(row=1, column=0, sticky='w')
    combo = ttk.Combobox(frm, textvariable=query, width=70)
    combo.grid(row=1, column=1, columnspan=2, sticky='ew')
    source = tk.StringVar()
    notes = tk.StringVar()
    source_url = tk.StringVar()
    ttk.Label(frm, text='적용규격·판본 / 기준 출처').grid(row=2, column=0, sticky='w')
    ttk.Entry(frm, textvariable=source, width=70).grid(row=2, column=1, columnspan=2, sticky='ew')
    ttk.Label(frm, text='적용 조건 / 주의사항').grid(row=3, column=0, sticky='w')
    ttk.Entry(frm, textvariable=notes, width=70).grid(row=3, column=1, columnspan=2, sticky='ew')
    ttk.Label(frm, text='자료 링크').grid(row=4, column=0, sticky='w')
    ttk.Entry(frm, textvariable=source_url, width=70).grid(row=4, column=1, columnspan=2, sticky='ew')
    ttk.Label(frm, textvariable=notes, wraplength=680, foreground='#8a4b08').grid(row=5, column=0, columnspan=3, sticky='w', pady=4)
    ttk.Label(frm, text='범위: 0.90~1.50 / <=1.50 / >=0.90 / -（규정 없음）/ Balance').grid(row=6, column=0, columnspan=3, sticky='w', pady=6)
    fields = {}
    for idx, symbol in enumerate(elements):
        fields[symbol] = tk.StringVar()
        ttk.Label(frm, text=f'{symbol} / {ELEMENT_NAMES.get(symbol, symbol)}').grid(row=7 + idx, column=0, sticky='w')
        ttk.Entry(frm, textvariable=fields[symbol], width=28).grid(row=7 + idx, column=1, sticky='w')
    row = 7 + len(elements)
    material_status = tk.StringVar(value='재질을 선택하면 등록된 기준 원소를 표시합니다. ±10%는 문구만 출력.')
    ttk.Label(frm, textvariable=material_status, wraplength=680).grid(row=row, column=0, columnspan=3, sticky='w', pady=6)
    listing = tk.Listbox(frm, width=92, height=6, exportselection=False)
    listing.grid(row=row + 2, column=0, columnspan=3, sticky='ew', pady=6)
    editing = [None]
    active_base = [None]

    def filter_items(*_):
        # A selected material is a value, not a search term. Keep all items
        # in its family available when the dropdown is opened again.
        needle = '' if query.get() in catalogue else query.get().casefold().replace('sus', '').replace('sts', '')
        combo['values'] = [i['label'] for i in sorted(catalogue.values(), key=lambda item: not bool(item.get('ranges')))
                           if (family.get() == '전체' or i['family'] == family.get())
                           and needle in i['label'].casefold().replace('sus', '').replace('sts', '')]

    def fill(item):
        active_base[0] = deepcopy(item)
        query.set(item['label'])
        source.set(item.get('source', ''))
        notes.set(item.get('notes', ''))
        source_url.set(item.get('source_url', ''))
        for sym, var in fields.items():
            var.set(item.get('ranges', {}).get(sym, ''))
        available = item.get('ranges', {})
        missing = [sym for sym in fields if sym not in available]
        if available:
            text = '등록 기준 원소: ' + ', '.join(available)
            if missing:
                text += '\n현재 선택 중 미확인: ' + ', '.join(missing) + ' — 원소 선택을 변경하거나 프로젝트 기준을 입력하세요.'
        else:
            text = '이 재질은 기준 미등록입니다. 성분 범위를 직접 입력하거나 기준이 등록된 재질을 선택하세요.'
        material_status.set(text)

    def pick(_=None):
        editing[0] = None
        item = catalogue.get(query.get())
        if item:
            fill(item)

    def change_family(*_):
        # Do not filter the new family using the previous material's name.
        # Also leave list-edit mode so adding a new material cannot replace
        # an already selected report criterion.
        editing[0] = None
        active_base[0] = None
        listing.selection_clear(0, 'end')
        for var in (query, source, notes, source_url, *fields.values()):
            var.set('')
        filter_items()
        candidates = [catalogue[label] for label in combo['values']]
        if candidates:
            fill(next((item for item in candidates if item.get('ranges')), candidates[0]))

    def redraw():
        listing.delete(0, 'end')
        for item in chosen:
            listing.insert('end', item['label'])

    def add():
        try:
            label = query.get().strip()
            ranges = normalize_ranges({s: v.get() for s, v in fields.items()})
            original = (active_base[0] if active_base[0] and active_base[0]['label'] == label
                        else catalogue.get(label, {}))
            base = deepcopy(original)
            complete_ranges = deepcopy(base.get('ranges', {}))
            for symbol in fields:
                complete_ranges.pop(symbol, None)
            complete_ranges.update(ranges)
            base.update(id=base.get('id', label), label=label,
                        family=base.get('family', '기타'), source=source.get().strip(),
                        notes=notes.get().strip(), source_url=source_url.get().strip(),
                        ranges=complete_ranges)
            if base.get('reference_ranges'):
                base['user_modified'] = base['ranges'] != base['reference_ranges']
            validate_criteria([base], elements)
            if editing[0] is None:
                if any(i['id'] == base['id'] for i in chosen):
                    raise ValueError('이미 추가된 재질입니다. 아래 목록에서 선택해 수정하세요.')
                chosen.append(base)
            else:
                chosen[editing[0]] = base
            editing[0] = None
            redraw()
        except ValueError as exc:
            messagebox.showwarning('합격기준 확인', str(exc), parent=win)

    def select(_=None):
        if listing.curselection():
            editing[0] = listing.curselection()[0]
            fill(chosen[editing[0]])

    def remove():
        if listing.curselection():
            chosen.pop(listing.curselection()[0])
            editing[0] = None
            redraw()

    def reload_reference():
        item = builtin_catalogue.get(query.get())
        if not item or not item.get('ranges'):
            messagebox.showinfo('공개 기준', '이 재질에는 공개 기본 기준이 등록되어 있지 않습니다.', parent=win)
            return
        fill(item)

    def apply():
        try:
            validate_criteria(chosen, elements)
            updated = deepcopy(batch)
            updated['pmi_criteria'] = chosen
            updated['pmi_elements'] = elements
            save_report(updated, state_path, batch.get('_pmi_state_key'))
            batch['pmi_criteria'] = chosen
            batch['pmi_elements'] = elements
            win.destroy()
        except (ValueError, OSError) as exc:
            messagebox.showerror('기준 저장', str(exc), parent=win)

    family.trace_add('write', change_family)
    combo.bind('<KeyRelease>', filter_items)
    combo.bind('<<ComboboxSelected>>', pick)
    listing.bind('<<ListboxSelect>>', select)
    bar = ttk.Frame(frm)
    bar.grid(row=row + 1, column=0, columnspan=3, sticky='ew')
    ttk.Button(bar, text='재질 추가 / 선택 항목 수정', command=add).pack(side='left')
    ttk.Button(bar, text='선택 항목 삭제', command=remove).pack(side='left', padx=6)
    ttk.Button(bar, text='공개 기본값 불러오기', command=reload_reference).pack(side='left')
    buttons = ttk.Frame(frm)
    buttons.grid(row=row + 3, column=0, columnspan=3, sticky='e')
    ttk.Button(buttons, text='취소', command=win.destroy).pack(side='right')
    ttk.Button(buttons, text='목록 적용·저장', command=apply).pack(side='right', padx=6)
    filter_items()
    redraw()
    parent.wait_window(win)
