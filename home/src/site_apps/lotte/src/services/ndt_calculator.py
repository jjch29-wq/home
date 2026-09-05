"""Pure NDT billing calculations shared by the billing views."""

from __future__ import annotations

from typing import Any, Mapping


def calculate_billing(
    *, quantity: float, adjusted_quantity: float, material_key: str,
    ndt_type: str, location: str, work_time: str,
    material_costs: Mapping[str, Any], labor_costs: Mapping[str, Any],
    overhead_rate: float, technical_fee_rate: float, vat_rate: float = 0.1,
) -> dict[str, int]:
    material_unit = material_costs.get(material_key, material_costs.get(ndt_type, 0))
    location_costs = labor_costs.get(location, labor_costs)
    time_costs = location_costs.get(work_time, {})
    labor_unit = time_costs.get(material_key, time_costs.get(ndt_type, 0))

    material = int(quantity * float(material_unit or 0))
    labor = int(adjusted_quantity * float(labor_unit or 0))
    overhead = int(labor * overhead_rate)
    technical_fee = int((labor + overhead) * technical_fee_rate)
    subtotal = material + labor + overhead + technical_fee
    vat = int(subtotal * vat_rate)
    return {
        "mat_cost": material,
        "lab_cost": labor,
        "overhead": overhead,
        "tech": technical_fee,
        "subtotal": subtotal,
        "vat": vat,
        "total_amount": subtotal + vat,
    }

