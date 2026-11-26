# -*- coding: utf-8 -*-
"""한국지역난방 배관 관경(A) 기준 표준 재질 매핑 규칙.

규격 기준:
- 20A ~ 600A : SPPS380
- 650A ~ 900A(이상 대관경) : SPW400
- 소·대관경 혼합 : SPPS380 / SPW400
- 핫태핑(HOT TAPPING) : SPPS250
"""
from __future__ import annotations

import re

KDHC_PIPE_MATERIALS = (
    'SPPS380',
    'SPW400',
    'SPPS380 / SPW400',
    'SPPS250',
)

KDHC_NOMINAL_A_SIZES = frozenset({
    20, 25, 32, 40, 50, 65, 80, 100, 125, 150,
    200, 250, 300, 350, 400, 450, 500, 550, 600,
    650, 700, 750, 800, 850, 900, 1000, 1100, 1200,
})

_SIZE_FIELDS = ('Size', 'NPS', '관경')
_LINE_FIELDS = ('Line No.', 'ISO', 'Dwg', 'ISO/Dwg', 'ISO/DWG', 'Drawing No.', '라인번호')
_EXPLICIT_A_RE = re.compile(r'(?<![A-Za-z0-9])(\d{2,4})\s*[Aa](?![A-Za-z0-9])')
_PLAIN_NUM_RE = re.compile(r'(?<![A-Za-z0-9."\'])(\d{2,4})(?![A-Za-z0-9."\'])')


def extract_nominal_pipe_sizes_a(item: dict) -> list[int]:
    """단일 행에서 호칭관경(A) 숫자 목록을 추출한다."""
    if not isinstance(item, dict):
        return []

    sizes: list[int] = []
    for key in _SIZE_FIELDS:
        raw = str(item.get(key, '') or '').strip()
        if not raw or raw.lower() == 'nan':
            continue
        found = [int(m) for m in _EXPLICIT_A_RE.findall(raw)]
        if not found:
            found = [
                int(m) for m in _PLAIN_NUM_RE.findall(raw)
                if int(m) in KDHC_NOMINAL_A_SIZES
            ]
        for sz in found:
            if sz not in sizes:
                sizes.append(sz)

    if sizes:
        return sizes

    for key in _LINE_FIELDS:
        raw = str(item.get(key, '') or '').strip()
        if not raw or raw.lower() == 'nan':
            continue
        for m in _EXPLICIT_A_RE.findall(raw):
            sz = int(m)
            if sz in KDHC_NOMINAL_A_SIZES or 20 <= sz <= 1200:
                if sz not in sizes:
                    sizes.append(sz)
        if sizes:
            break

    return sizes


def material_for_size_a(size_a: int) -> str | None:
    """관경(A)에 대응하는 표준 KS 배관 재질을 반환한다."""
    if size_a >= 650:
        return 'SPW400'
    if 20 <= size_a <= 600:
        return 'SPPS380'
    return None


def is_row_active_for_report(item: dict, mode: str = '') -> bool:
    """미리보기 표에서 현재 선택·필터 활성 상태인 행인지 판별한다."""
    if not item.get('selected', True):
        return False
    if not item.get('date_filtered', True):
        return False
    if mode == 'PAUT' and not (
        item.get('line_filtered', True) and item.get('size_filtered', True)
    ):
        return False
    if mode in ('MT', 'RT', 'KOGAS') and not item.get('drawing_filtered', True):
        return False
    return True


def _combine_materials(found: set[str]) -> str | None:
    if {'SPPS380', 'SPW400'}.issubset(found):
        return 'SPPS380 / SPW400'
    if 'SPW400' in found:
        return 'SPW400'
    if 'SPPS380' in found:
        return 'SPPS380'
    return None


def infer_row_pipe_material(item: dict, item_name: str = '') -> str | None:
    """단일 행의 품명·관경으로부터 배관 재질을 추론한다."""
    if str(item_name or '').strip().upper() == 'HOT TAPPING':
        return 'SPPS250'
    found = {
        mat for sz in extract_nominal_pipe_sizes_a(item)
        if (mat := material_for_size_a(sz)) is not None
    }
    return _combine_materials(found)


def infer_kdhc_pipe_material(items, mode: str = '', item_name: str = '') -> str | None:
    """현재 활성(선택·필터 통과) 행들의 관경 구성으로부터 갑지 재질을 추론한다."""
    if str(item_name or '').strip().upper() == 'HOT TAPPING':
        return 'SPPS250'
    if not items:
        return None

    active_rows = [r for r in items if isinstance(r, dict) and is_row_active_for_report(r, mode=mode)]
    target_rows = active_rows if active_rows else [r for r in items if isinstance(r, dict)]

    found: set[str] = set()
    for row in target_rows:
        for sz in extract_nominal_pipe_sizes_a(row):
            mat = material_for_size_a(sz)
            if mat:
                found.add(mat)

    return _combine_materials(found)

