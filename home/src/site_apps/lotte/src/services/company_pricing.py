"""업체별 계약 단가. 사진에서 확인된 조건에만 적용한다."""

import re
import math

COMPANY_PRICES = {
    "함감속기": {
        "UT": {"D200 이하": 11000, "D200 초과~D400 이하": 12000, "D400 초과": 15000},
        "MT": {"D200 이하": 13000, "D200 초과~D400 이하": 22000, "D400 초과": 33000},
    },
    "일신기계": {"관경 650 이하": 275000, "관경 650 초과": 375000},
}


def normalize_company(value):
    return re.sub(r"\s+", "", value).replace("(주)", "").replace("㈜", "").replace("주식회사", "")


def company_conditions(company, method):
    company = normalize_company(company)
    method = method.strip().upper()
    if company == "함감속기" and method in ("UT", "MT"):
        return tuple(COMPANY_PRICES[company][method])
    if company == "일신기계" and method == "UT":
        return tuple(COMPANY_PRICES[company])
    if company == "세영정공" and method in ("UT", "MT"):
        return ("길이 없음, D100 이상", "길이 1m, D100 이상", "길이 2m 이상")
    if company == "한은이앤씨" and method == "MT":
        return ("1m", "1식")
    if company == "피테크" and method == "MT":
        return ("EA당",)
    return ()


def get_company_price(company, method, condition, quantity):
    conditions = company_conditions(company, method)
    if condition not in conditions:
        return None
    company = normalize_company(company)
    quantity = float(quantity)
    if not math.isfinite(quantity) or quantity < 0:
        return None
    if company == "함감속기":
        price = COMPANY_PRICES[company][method.strip().upper()][condition]
        return {"price": price, "travel": 100000, "fee": quantity * price + 100000, "unit": "EA"}
    if company == "일신기계":
        price = COMPANY_PRICES[company][condition]
        return {"price": price, "travel": 0, "fee": quantity * price, "unit": "EA"}
    if company == "세영정공":
        price = dict(zip(conditions, (20000, 25000, 30000)))[condition]
        return {"price": price, "travel": 0, "fee": quantity * price, "unit": "EA"}
    if company == "한은이앤씨":
        # 검사비가 130,000원을 넘으면 m 단가로 청구한다.
        price = 11000 if condition == "1m" else 130000
        fee = max(quantity * price, 130000) if condition == "1m" else quantity * price
        return {"price": price, "travel": 0, "fee": fee, "unit": "M" if condition == "1m" else "식"}
    return {"price": 13000, "travel": 150000, "fee": quantity * 13000 + 150000, "unit": "EA"}


def resolve_item_price(company, method, diameter, length_condition, quantity):
    """관경은 mm 숫자이며 해당 업체의 유효 조건을 선택한다."""
    options = company_conditions(company, method)
    if not options:
        return None
    name = normalize_company(company)
    condition = length_condition
    if name in ('함감속기', '일신기계', '세영정공'):
        try:
            dia = float(str(diameter).strip())
        except ValueError:
            raise ValueError('관경(mm)을 숫자로 입력해주세요.') from None
        if not math.isfinite(dia) or dia <= 0:
            raise ValueError('관경(mm)은 0보다 큰 숫자여야 합니다.')
        if name == '함감속기':
            condition = options[0] if dia <= 200 else options[1] if dia <= 400 else options[2]
        elif name == '일신기계':
            condition = options[0] if dia <= 650 else options[1]
        elif condition in options[:2] and dia < 100:
            raise ValueError('세영정공의 해당 길이 조건은 D100 이상에만 적용됩니다.')
    elif name == '피테크':
        condition = options[0]
    if condition not in options:
        raise ValueError('길이/단위 조건을 선택해주세요.')
    result = get_company_price(company, method, condition, quantity)
    if result is None:
        raise ValueError('수량은 0 이상의 유효한 숫자여야 합니다.')
    return dict(result, condition=condition)
