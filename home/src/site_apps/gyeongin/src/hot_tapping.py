"""사용자 제공 MT/PT 검사길이표 (m)."""
from decimal import Decimal

# 관경(A): 연장, 원주, END CAP. 보강링 내측=.22, 외측=.377.
LENGTHS = {
    850: (3.048, 5.4262, 2.667), 800: (2.846, 5.107, 2.51),
    750: (2.846, 4.7878, 2.349), 700: (2.438, 4.4686, 2.191),
    650: (2.388, 4.1494, 2.029), 600: (2.184, 3.8302, 1.866),
    550: (2.032, 3.511, 1.702), 500: (1.828, 3.1918, 1.547),
    450: (1.68, 2.8726, 1.388), 400: (1.52, 2.5534, 1.231),
    350: (1.32, 2.2334, 1.068), 300: (1.12, 2.0012, .964),
    250: (1, 1.6802, .804), 200: (.82, 1.359, .65),
    150: (.7, 1.038, .493), 125: (.5, .8784, .418),
    100: (.42, .7182, .336), 80: (.37, .5598, .265),
}
STAGES = {'1차': 'Stopple Fitting 연장·원주 + 보강링 내측·외측',
          '2차': 'END CAP + 보강링 cap'}

def inspection_length(size, stage):
    try:
        diameter = int(str(size).strip().upper().removesuffix('A'))
        longitudinal, circumferential, end_cap = LENGTHS[diameter]
    except (ValueError, KeyError):
        raise ValueError('핫태핑 관경을 표의 규격에서 선택하세요.') from None
    if stage not in STAGES:
        raise ValueError('검사 차수를 선택하세요.')
    parts = (longitudinal, circumferential, .22, .377) if stage == '1차' else (end_cap, .239 if diameter <= 150 else .319)
    return float(sum((Decimal(str(v)) for v in parts), Decimal(0)))
