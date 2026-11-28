"""Flexible contract billing calculator & Excel statement exporter.

Supports arbitrary custom rate schedules (annual unit-price contracts) as well as
per-job trip fees and lump-sum (일식) billing without KDHC-specific hardcoded items.
"""
from __future__ import annotations

from datetime import datetime
import os
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def _to_float(val: Any) -> float:
    try:
        return float(str(val or 0).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _to_int(val: Any) -> int:
    return int(round(_to_float(val)))


def calculate_flexible_billing(
    contract: dict[str, Any],
    usage_entries: list[dict[str, Any]],
    *,
    prev_cum_by_code: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Calculate flexible billing lines, carry-over progress, and supply/VAT/grand totals.

    Each entry in ``usage_entries`` may reference a ``code`` from ``contract['rate_items']``
    or provide custom ``name`` / ``unit_price`` / ``qty`` / ``trip_fee`` / ``lump_sum``.
    """
    prev_map = {str(k): _to_float(v) for k, v in (prev_cum_by_code or {}).items()}
    rate_map = {str(item.get("code") or ""): item for item in (contract.get("rate_items") or [])}

    curr_qty_by_code: dict[str, float] = {}
    extra_trip_by_code: dict[str, int] = {}
    extra_lump_by_code: dict[str, int] = {}
    adhoc_lines: list[dict[str, Any]] = []

    for entry in usage_entries or []:
        code = str(entry.get("code") or "").strip()
        qty = _to_float(entry.get("qty"))
        trip_fee = _to_int(entry.get("trip_fee"))
        lump_sum = _to_int(entry.get("lump_sum"))

        if code and code in rate_map:
            curr_qty_by_code[code] = curr_qty_by_code.get(code, 0.0) + qty
            extra_trip_by_code[code] = extra_trip_by_code.get(code, 0) + trip_fee
            extra_lump_by_code[code] = extra_lump_by_code.get(code, 0) + lump_sum
        else:
            unit_price = _to_int(entry.get("unit_price"))
            base_amt = int(round(qty * unit_price))
            line_total = base_amt + trip_fee + lump_sum
            if line_total <= 0 and qty <= 0:
                continue
            adhoc_lines.append(
                {
                    "code": code or "ADHOC",
                    "name": str(entry.get("name") or "건별 검사").strip(),
                    "method": str(entry.get("method") or "").strip().upper(),
                    "spec": str(entry.get("spec") or "").strip(),
                    "unit": str(entry.get("unit") or "식").strip(),
                    "unit_price": unit_price,
                    "contract_qty": 0.0,
                    "prev_qty": 0.0,
                    "curr_qty": qty,
                    "total_qty": qty,
                    "remain_qty": 0.0,
                    "progress_pct": 0.0,
                    "base_amount": base_amt,
                    "trip_fee": trip_fee,
                    "lump_sum": lump_sum,
                    "line_total": line_total,
                }
            )

    lines: list[dict[str, Any]] = []
    for item in contract.get("rate_items") or []:
        code = str(item.get("code") or "")
        contract_qty = _to_float(item.get("contract_qty"))
        prev_qty = prev_map.get(code, 0.0)
        curr_qty = curr_qty_by_code.get(code, 0.0)
        total_qty = prev_qty + curr_qty
        remain_qty = max(0.0, contract_qty - total_qty) if contract_qty > 0 else 0.0
        progress_pct = round((total_qty / contract_qty) * 100.0, 1) if contract_qty > 0 else 0.0

        unit_price = _to_int(item.get("unit_price"))
        base_amt = int(round(curr_qty * unit_price))
        trip_fee = extra_trip_by_code.get(code, 0)
        lump_sum = extra_lump_by_code.get(code, 0)
        line_total = base_amt + trip_fee + lump_sum

        lines.append(
            {
                "code": code,
                "name": str(item.get("name") or ""),
                "method": str(item.get("method") or ""),
                "spec": str(item.get("spec") or ""),
                "unit": str(item.get("unit") or "EA"),
                "unit_price": unit_price,
                "contract_qty": contract_qty,
                "prev_qty": prev_qty,
                "curr_qty": curr_qty,
                "total_qty": total_qty,
                "remain_qty": remain_qty,
                "progress_pct": progress_pct,
                "base_amount": base_amt,
                "trip_fee": trip_fee,
                "lump_sum": lump_sum,
                "line_total": line_total,
            }
        )

    lines.extend(adhoc_lines)
    supply_total = sum(line["line_total"] for line in lines)
    vat_total = int(round(supply_total * 0.1))
    return {
        "lines": lines,
        "supply_total": supply_total,
        "vat_total": vat_total,
        "grand_total": supply_total + vat_total,
    }


def export_flexible_billing_excel(
    contract: dict[str, Any],
    client: dict[str, Any] | None,
    billing_result: dict[str, Any],
    output_path: str,
    *,
    period_label: str = "",
    round_no: int = 1,
) -> dict[str, Any]:
    """Export a 2-sheet Excel billing statement (청구서 요약 + 기성·거래내역)."""
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "청구서(요약)"
    ws_det = wb.create_sheet("기성·거래내역")

    title_font = Font(name="Malgun Gothic", size=16, bold=True, color="1B365D")
    label_font = Font(name="Malgun Gothic", size=10, bold=True, color="333333")
    val_font = Font(name="Malgun Gothic", size=10)
    hdr_font = Font(name="Malgun Gothic", size=10, bold=True, color="FFFFFF")
    tot_font = Font(name="Malgun Gothic", size=11, bold=True, color="0D47A1")

    hdr_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sub_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
    tot_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    client_info = client or {}
    ws_sum.merge_cells("A1:F1")
    ws_sum["A1"] = f"비파괴검사 기성·청구서 ({round_no}회차)"
    ws_sum["A1"].font = title_font
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    summary_rows = [
        ("거래처명", client_info.get("name", ""), "사업자번호", client_info.get("biz_no", "")),
        ("대표자", client_info.get("ceo", ""), "담당자 / 연락처", f"{client_info.get('contact_name', '')} {client_info.get('contact_phone', '')}".strip()),
        ("계약명", contract.get("contract_name", ""), "계약기간", f"{contract.get('start_date', '')} ~ {contract.get('end_date', '')}".strip(" ~")),
        ("청구기간", period_label or datetime.now().strftime("%Y-%m"), "결제조건", client_info.get("payment_terms", "")),
        ("작성일자", datetime.now().strftime("%Y-%m-%d"), "계약총액", contract.get("budget_amount", 0)),
    ]
    for offset, (k1, v1, k2, v2) in enumerate(summary_rows, start=3):
        ws_sum.cell(row=offset, column=1, value=k1).font = label_font
        ws_sum.cell(row=offset, column=1).fill = sub_fill
        ws_sum.merge_cells(start_row=offset, start_column=2, end_row=offset, end_column=3)
        ws_sum.cell(row=offset, column=2, value=v1).font = val_font

        ws_sum.cell(row=offset, column=4, value=k2).font = label_font
        ws_sum.cell(row=offset, column=4).fill = sub_fill
        ws_sum.merge_cells(start_row=offset, start_column=5, end_row=offset, end_column=6)
        c2 = ws_sum.cell(row=offset, column=5, value=v2)
        c2.font = val_font
        if isinstance(v2, int):
            c2.number_format = "#,##0"
            c2.alignment = Alignment(horizontal="right", vertical="center")
        ws_sum.row_dimensions[offset].height = 22
        for col in range(1, 7):
            ws_sum.cell(row=offset, column=col).border = border

    money_rows = [
        ("공급가액", billing_result["supply_total"]),
        ("부가세 (10%)", billing_result["vat_total"]),
        ("총 청구금액", billing_result["grand_total"]),
    ]
    for offset, (label, amount) in enumerate(money_rows, start=10):
        ws_sum.merge_cells(start_row=offset, start_column=1, end_row=offset, end_column=3)
        ws_sum.merge_cells(start_row=offset, start_column=4, end_row=offset, end_column=6)
        lc = ws_sum.cell(row=offset, column=1, value=label)
        vc = ws_sum.cell(row=offset, column=4, value=amount)
        lc.font = tot_font if offset == 12 else label_font
        vc.font = tot_font if offset == 12 else val_font
        vc.number_format = "#,##0"
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc.alignment = Alignment(horizontal="right", vertical="center")
        fill = tot_fill if offset == 12 else sub_fill
        for col in range(1, 7):
            cell = ws_sum.cell(row=offset, column=col)
            cell.border = border
            if col <= 3 or offset == 12:
                cell.fill = fill
        ws_sum.row_dimensions[offset].height = 26

    for col_idx, width in enumerate([16, 18, 16, 16, 18, 18], 1):
        ws_sum.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    # Sheet 2: Detailed lines
    headers = [
        "품목코드", "품목명", "검사방법", "규격", "단위", "단가",
        "계약수량", "전회누계", "금회기성", "총누계", "잔여수량", "진도율(%)",
        "금회금액", "출장비", "일식", "합계금액",
    ]
    ws_det.row_dimensions[1].height = 24
    for col_idx, text in enumerate(headers, 1):
        c = ws_det.cell(row=1, column=col_idx, value=text)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    row_num = 2
    for line in billing_result["lines"]:
        vals = [
            line["code"], line["name"], line["method"], line["spec"], line["unit"],
            line["unit_price"], line["contract_qty"], line["prev_qty"], line["curr_qty"],
            line["total_qty"], line["remain_qty"], line["progress_pct"],
            line["base_amount"], line["trip_fee"], line["lump_sum"], line["line_total"],
        ]
        for col_idx, val in enumerate(vals, 1):
            c = ws_det.cell(row=row_num, column=col_idx, value=val)
            c.font = val_font
            c.border = border
            if col_idx in (1, 3, 5):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (6, 13, 14, 15, 16):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0"
            elif col_idx in (7, 8, 9, 10, 11, 12):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.##"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
        ws_det.row_dimensions[row_num].height = 20
        row_num += 1

    ws_det.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=15)
    ws_det.cell(row=row_num, column=1, value="공급가액 합계").font = tot_font
    ws_det.cell(row=row_num, column=1).alignment = Alignment(horizontal="center", vertical="center")
    tc = ws_det.cell(row=row_num, column=16, value=billing_result["supply_total"])
    tc.font = tot_font
    tc.number_format = "#,##0"
    tc.alignment = Alignment(horizontal="right", vertical="center")
    for col_idx in range(1, 17):
        cell = ws_det.cell(row=row_num, column=col_idx)
        cell.fill = tot_fill
        cell.border = border
    ws_det.row_dimensions[row_num].height = 24

    det_widths = [12, 20, 10, 14, 8, 12, 10, 10, 10, 10, 10, 10, 13, 12, 12, 14]
    for col_idx, width in enumerate(det_widths, 1):
        ws_det.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return billing_result

