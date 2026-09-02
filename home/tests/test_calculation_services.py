from services.budget_calculator import row_margin, summarize_budget_rows
from services.ndt_calculator import calculate_billing


def test_budget_summary_handles_formatted_values():
    rows = [
        {"Revenue": "1,000", "LaborCost": 100, "MaterialCost": 200,
         "Expense": 50, "OutsourceCost": 50, "Profit": 600},
        {"Revenue": 500, "LaborCost": 100, "Profit": 400},
    ]
    result = summarize_budget_rows(rows)
    assert result.revenue == 1500
    assert result.cost == 500
    assert result.profit == 1000
    assert round(result.margin, 2) == 66.67
    assert row_margin(rows[0]) == 60


def test_ndt_billing_calculation():
    result = calculate_billing(
        quantity=2,
        adjusted_quantity=3,
        material_key="RT_FILM",
        ndt_type="RT",
        location="plant",
        work_time="day",
        material_costs={"RT_FILM": 100},
        labor_costs={"plant": {"day": {"RT_FILM": 200}}},
        overhead_rate=0.1,
        technical_fee_rate=0.1,
    )
    assert result == {
        "mat_cost": 200,
        "lab_cost": 600,
        "overhead": 60,
        "tech": 66,
        "subtotal": 926,
        "vat": 92,
        "total_amount": 1018,
    }
