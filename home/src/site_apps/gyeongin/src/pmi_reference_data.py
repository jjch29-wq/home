"""Public supplier chemistry references, never labelled as verified ASME editions.

All values are mass percent. Missing elements remain missing (not zero or '-').
Tables are frozen references retrieved on 2026-11-13; report snapshots retain
their source, conditions and original values when users override future presets.
"""
from copy import deepcopy

RETRIEVED = '2026-11-13'
APP = 'https://www.amerpipe.com/products/'


def public_material_references():
    records = []

    def add(number, grade, family, product, ranges, url, publisher='APP',
            edition='미기재', notes='', asme=True, source_section='Chemical Requirements'):
        for standard in [f'ASTM A{number}'] + ([f'ASME SA-{number}'] if asme else []):
            label = f'{standard} {grade} — {product}'
            records.append(dict(
                id=label, label=label, standard=standard, grade=grade,
                family=family, product=product, ranges=deepcopy(ranges),
                source=f'{publisher} 공개참고 / 판본 {edition}',
                source_url=url, source_section=source_section, retrieved=RETRIEVED,
                edition=edition, reference_status='public_reference',
                official_edition_verified=False, notes=notes,
                reference_ranges=deepcopy(ranges),
            ))

    # APP's published table has historical limits; preserve its actual values
    # and label them as an undated reference rather than silently modernizing.
    url = APP + 'stainless-steel-pipe/a312-specifications-page/'
    for grade, c, ni, cr, mo in [
        ('TP304', '<=0.08', '8.00~11.00', '18.00~20.00', '-'),
        ('TP304L', '<=0.035', '8.00~13.00', '18.00~20.00', '-'),
        ('TP316', '<=0.08', '10.00~14.00', '16.00~18.00', '2.00~3.00'),
        ('TP316L', '<=0.035', '10.00~14.00', '16.00~18.00', '2.00~3.00'),
    ]:
        add('312', grade, '스테인리스', 'PIPE',
            dict(C=c, Mn='<=2.00', P='<=0.045', S='<=0.030', Si='<=1.00',
                 Ni=ni, Cr=cr, Mo=mo), url,
            notes='공급사 참고표 값. L등급 C·Ni 범위 등은 적용 원문 판본과 대조 필요.')

    url = APP + 'carbon-pipe/a106/a106-specifications/'
    for grade, carbon, mn in [('Gr.A', '<=0.25', '0.27~0.93'),
                              ('Gr.B', '<=0.30', '0.29~1.06'),
                              ('Gr.C', '<=0.35', '0.29~1.06')]:
        add('106', grade, '탄소강', 'PIPE',
            dict(C=carbon, Mn=mn, P='<=0.035', S='<=0.035', Si='>=0.10',
                 Cr='<=0.40', Cu='<=0.40', Mo='<=0.15', Ni='<=0.40', V='<=0.08'), url,
            notes='Cr+Cu+Mo+Ni+V≤1.00%. C 감소에 따른 Mn 증가는 구매조건·판본 확인; 자동 가산 안 함.')
        for item in records[-2:]:
            cap = '1.35' if item['standard'].startswith('ASME') else '1.65'
            item['notes'] += f' 공급사 조건부 Mn 상한 {cap}%.'

    url = APP + 'carbon-pipe/a53/a53-specifications/'
    for pipe_type in ('S', 'E'):
        for grade, carbon, mn in [('Gr.A', '<=0.25', '<=0.95'), ('Gr.B', '<=0.30', '<=1.20')]:
            add('53', f'{grade} Type {pipe_type}', '탄소강', 'PIPE',
                dict(C=carbon, Mn=mn, P='<=0.05', S='<=0.045', Cu='<=0.40',
                     Ni='<=0.40', Cr='<=0.40', Mo='<=0.15', V='<=0.08'), url,
                notes=f'Type {pipe_type} 전용. Mn은 상한값. ASTM 조건부 Mn 증가 규정은 SA53에 자동 적용하지 않음.')

    # The supplier chart explicitly separates A105 and SA105 and identifies
    # its historical addendum. This is NOT the ASME 2025 original text.
    url = 'https://www.amerpipe.com/wp-content/uploads/2023/03/carbon-fittings.pdf'
    add('105', 'A105', '탄소강', '단조 피팅·플랜지',
        dict(C='<=0.350', Mn='0.600~1.050', P='<=0.035', S='<=0.040',
             Si='0.100~0.350', Cu='<=0.400', Ni='<=0.400', Cr='<=0.300',
             Mo='<=0.120', V='<=0.080'), url, edition='공급사 표 Addendum 2013',
        source_section='p.2 Carbon fittings / A105 and SA105 columns',
        notes='Cu+Ni+Cr+Mo+V≤1.00%; Cr+Mo≤0.32%. 조건부 Mn 증가는 C·구매조건 확인.')
    # Keep the ASME designation natural in its grade field, with the original
    # catalogue A105 identity retained separately for backwards compatibility.
    for item in records[-2:]:
        cap = '1.35' if item['standard'].startswith('ASME') else '1.65'
        item['notes'] += f' 공급사 조건부 Mn 상한 {cap}%.'

    for grade, carbon in [('WPB', '<=0.300'), ('WPC', '<=0.350')]:
        add('234', grade, '탄소강', '맞대기 피팅',
            dict(C=carbon, Mn='0.290~1.060', P='<=0.050', S='<=0.058', Si='0.100~0.350',
                 Cu='<=0.400', Ni='<=0.400', Cr='<=0.400', Mo='<=0.150', V='<=0.080'),
            url, edition='공급사 표 Addendum 2013',
            source_section=f'p.2 A234-{grade} / SA234-{grade} columns',
            notes='Cu+Ni+Cr+Mo≤1.00%. WPB 원소재·제조방법별 C·Si 예외와 조건부 Mn 증가 확인.')
        for item in records[-2:]:
            cap = '1.35' if item['standard'].startswith('ASME') else '1.65'
            item['notes'] += f' 공급사 조건부 Mn 상한 {cap}%.'

    add('350', 'LF2 CL1/CL2', '탄소강', '저온 단조품',
        dict(C='<=0.300', Mn='0.600~1.350', P='<=0.035', S='<=0.040',
             Si='0.150~0.300', Cu='<=0.400', Ni='<=0.400', Cr='<=0.300',
             Mo='<=0.120', V='<=0.080'), url, edition='공급사 표 Addendum 2011',
        source_section='p.2 A/SA350 LF2 CL1/CL2 column',
        notes='화학성분 참고값. Class별 기계적 성질을 의미하지 않음. 합계·Nb 조건은 원표 확인.')

    url = APP + 'chrome-moly-pipe/a335-spec-sheet/'
    for grade, carbon, si, cr, mo in [
        ('P5', '<=0.15', '<=0.50', '4.00~6.00', '0.45~0.65'),
        ('P9', '<=0.15', '0.25~1.00', '8.00~10.00', '0.90~1.10'),
        ('P11', '0.05~0.15', '0.50~1.00', '1.00~1.50', '0.44~0.65'),
        ('P22', '0.05~0.15', '<=0.50', '1.90~2.60', '0.87~1.13'),
    ]:
        add('335', grade, '저합금강', 'PIPE',
            dict(C=carbon, Mn='0.30~0.60', P='<=0.025', S='<=0.025', Si=si, Cr=cr, Mo=mo), url,
            notes='기재 없는 원소는 미확인. 다른 제품규격의 동계열 재질로 치환하지 않음.')
    add('335', 'P91 Type 1', '저합금강', 'PIPE',
        dict(C='0.08~0.12', Mn='0.30~0.60', P='<=0.020', S='<=0.010', Si='0.20~0.50',
             Cr='8.00~9.50', Mo='0.85~1.05', V='0.18~0.25', N='0.030~0.070',
             Ni='<=0.40', Al='<=0.02', Nb='0.06~0.10', Ti='<=0.01', Zr='<=0.01'), url,
        notes='P91 Type 1 전용. Type 2 및 heat/product별 허용범위를 혼용하지 않음.')

    # Manufacturer explicitly publishes SA/A403. Dual-certified grades are
    # separate rows, not silently mapped to single-grade WP304/WP316 limits.
    url = 'https://www.taylorforgestnls.com/productspec.html'
    for grade, c, ni, cr, mo in [
        ('WP304/304L', '<=0.030', '8.00~11.00', '18.00~20.00', '-'),
        ('WP304L', '<=0.030', '8.00~12.00', '18.00~20.00', '-'),
        ('WP316/316L', '<=0.030', '10.00~14.00', '16.00~18.00', '2.00~3.00'),
        ('WP316L', '<=0.030', '10.00~14.00', '16.00~18.00', '2.00~3.00'),
    ]:
        add('403', grade, '스테인리스', '맞대기 피팅',
            dict(C=c, Mn='<=2.00', P='<=0.045', S='<=0.030', Si='<=1.00', Ni=ni, Cr=cr, Mo=mo),
            url, publisher='Taylor Forge', notes='SA/A403 제조사 공개참고. 이중등급과 단일등급을 구분; 적용 판본 확인.')

    for grade, c, ni, cr, mo, filename in [
        ('F304', '<=0.08', '8.00~11.00', '18.00~20.00', None, 'F304-304L'),
        ('F304L', '<=0.03', '8.00~13.00', '18.00~20.00', None, 'F304-304L'),
        ('F316', '<=0.08', '10.00~14.00', '16.00~18.00', '2.00~3.00', 'F316-316L'),
        ('F316L', '<=0.035', '10.00~14.00', '16.00~18.00', '2.00~3.00', 'F316-316L'),
    ]:
        ranges = dict(C=c, Mn='<=2.00', P='<=0.045', S='<=0.030', Si='<=1.00', Ni=ni, Cr=cr)
        if mo:
            ranges.update(Mo=mo, Fe='Balance')
        add('182', grade, '스테인리스', '단조 피팅·플랜지', ranges,
            f'https://semetalgroup.com/wp-content/uploads/ASTM-A182-{filename}-Forging-specification.pdf',
            publisher='SE Metal', source_section='p.1 Chemical Composition',
            notes='공급사 ASTM A182 / ASME SA182 참고값. L등급 C·Ni는 적용 원문 판본 대조 필요.')
    url = 'https://www.dongkuk.com/upload/promotion/print/0ed42c37-7ff6-4c6b-a5da-f6b7e1cc2b92.pdf'
    structural = [
        ('KS D 3503', 'SS275', '구조용', dict(C='<=0.25', Si='<=0.45', Mn='<=1.40', P='<=0.050', S='<=0.050'), 90),
        ('KS D 3503', 'SS315', '구조용', dict(C='<=0.28', Si='<=0.50', Mn='<=1.50', P='<=0.050', S='<=0.050'), 90),
        ('KS D 3503', 'SS410', '구조용', dict(C='<=0.30', Si='<=0.55', Mn='<=1.60', P='<=0.040', S='<=0.040'), 90),
        ('KS D 3515', 'SM275A', '용접 구조용', dict(C='<=0.23', Si='-', P='<=0.035', S='<=0.035'), 92),
        ('KS D 3515', 'SM275B', '용접 구조용', dict(C='<=0.20', Si='<=0.35', Mn='0.50~1.40', P='<=0.030', S='<=0.030'), 92),
        ('KS D 3515', 'SM275C', '용접 구조용', dict(C='<=0.20', Si='<=0.35', Mn='<=1.40', P='<=0.025', S='<=0.025'), 92),
        ('KS D 3515', 'SM355A', '용접 구조용', dict(C='<=0.20', Si='<=0.55', Mn='<=1.60', P='<=0.035', S='<=0.035'), 92),
        ('KS D 3515', 'SM355B', '용접 구조용', dict(C='<=0.18', Si='<=0.55', Mn='<=1.60', P='<=0.030', S='<=0.030'), 92),
        ('KS D 3515', 'SM355C', '용접 구조용', dict(C='<=0.18', Si='<=0.55', Mn='<=1.60', P='<=0.025', S='<=0.025'), 92),
    ]
    for standard, grade, product, ranges, page in structural:
        label = f'{standard} {grade} — {product}'
        notes = '형강·용강분석 참고표 전용. 판재·강관에 자동 적용하지 않음. 미기재 원소는 0% 또는 규정 없음으로 간주하지 않음.'
        if standard == 'KS D 3515':
            notes += ' C는 t≤50mm 표; 두께·원문 판본 확인.'
        if grade == 'SM275A':
            notes += ' Mn≥2.5×C 조건은 고정 수치로 환산하지 않음.'
        records.append(dict(
            id=label, label=label, standard=standard, grade=grade, family='탄소강',
            product=product, ranges=ranges, reference_ranges=deepcopy(ranges),
            source=f'동국제강 형강 공개참고 p.{page} / 판본 미기재',
            source_url=url, source_section=f'p.{page} 형강 Chemical Composition (Ladle)',
            retrieved=RETRIEVED, edition='미기재', reference_status='public_reference',
            official_edition_verified=False, notes=notes,
        ))
    return records
