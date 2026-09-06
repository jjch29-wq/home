"""Pure NDT billing calculations shared by the billing views."""

from __future__ import annotations

from typing import Any, Mapping


def calculate_billing(
    *, quantity: float, adjusted_quantity: float, material_key: str,
    ndt_type: str, unit_prices: Mapping[str, Any], vat_rate: float = 0.1,
) -> dict[str, int]:
    unit_price = float(unit_prices.get(material_key, unit_prices.get(ndt_type, 0)))
    subtotal = int(adjusted_quantity * unit_price)
    vat = int(subtotal * vat_rate)
    return {
        "unit_price": int(unit_price),
        "subtotal": subtotal,
        "vat": vat,
        "total_amount": subtotal + vat,
    }

