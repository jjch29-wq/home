"""선택한 검사 기록의 소모품 수정량과 재고 반영 차이 계산."""

import math


def chemical_changes(record, method, values):
    allowed = {'MT': ('형광자분', '흑색자분', '백색페인트'),
               'PT': ('침투제', '세척제', '현상제', '형광침투제')}.get(method.strip().upper(), ())
    updates, differences = {}, {}
    for name in allowed:
        quantity = float(str(values.get(name, '')).replace(',', '').strip() or 0)
        if not math.isfinite(quantity) or quantity < 0:
            raise ValueError('소모품 수량은 0 이상의 숫자로 입력해주세요.')
        previous = record.get('NDT_' + name, 0)
        previous = float(previous or 0)
        if not math.isfinite(previous):
            previous = 0
        if quantity != previous:
            updates['NDT_' + name] = quantity
            differences[name] = quantity - previous
    return updates, differences
