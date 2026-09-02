"""Pure project-budget calculations, independent from Tkinter widgets."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Mapping


def number(value: Any) -> float:
    if value is None:
        return 0.0
    try:
        text = str(value).replace(",", "").replace("%", "").strip()
        return float(text) if text else 0.0
    except (TypeError, ValueError):
        return 0.0


@dataclass(frozen=True)
class BudgetSummary:
    revenue: float
    cost: float
    profit: float
    margin: float


def summarize_budget_rows(rows: Iterable[Mapping[str, Any]]) -> BudgetSummary:
    revenue = 0.0
    cost = 0.0
    profit = 0.0
    for row in rows:
        revenue += number(row.get("Revenue"))
        cost += sum(number(row.get(key)) for key in (
            "LaborCost", "MaterialCost", "Expense", "OutsourceCost"
        ))
        profit += number(row.get("Profit"))
    margin = profit / revenue * 100 if revenue else 0.0
    return BudgetSummary(revenue, cost, profit, margin)


def row_margin(row: Mapping[str, Any]) -> float:
    revenue = number(row.get("Revenue"))
    return number(row.get("Profit")) / revenue * 100 if revenue else 0.0

