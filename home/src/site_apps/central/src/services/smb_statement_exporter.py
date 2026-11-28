"""Monthly client billing statement (거래명세서/청구서) Excel exporter for SMB & annual contracts."""
from __future__ import annotations

from datetime import datetime
import os
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd


def _to_float(val: Any) -> float:
    if val is None:
        return 0.0
    try:
        if pd.isna(val):
            return 0.0
    except Exception:
        pass
    text = str(val).replace(",", "").strip()
    if not text or text.lower() in ("nan", "none"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _clean_str(val: Any) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    text = str(val).strip()
    return "" if text.lower() in ("nan", "none", "nat") else text


def build_statement_rows(
    daily_usage_df: pd.DataFrame | None,
    *,
    year: int | str | None = None,
    month: int | str | None = None,
    site: str | None = None,
    company: str | None = None,
) -> dict[str, Any]:
    """Filter daily usage records and compute per-row & aggregate billing totals."""
    if daily_usage_df is None or daily_usage_df.empty:
        return {"rows": [], "supply_total": 0, "vat_total": 0, "grand_total": 0}

    df = daily_usage_df.copy()
    if "Date" in df.columns:
        df["_dt"] = pd.to_datetime(df["Date"], errors="coerce")
    else:
        df["_dt"] = pd.NaT

    if year not in (None, "", "전체"):
        try:
            y = int(year)
            df = df[df["_dt"].dt.year == y]
        except (TypeError, ValueError):
            pass

    if month not in (None, "", "전체"):
        try:
            m = int(month)
            df = df[df["_dt"].dt.month == m]
        except (TypeError, ValueError):
            pass

    if site not in (None, "", "전체") and "Site" in df.columns:
        target_site = str(site).strip()
        df = df[df["Site"].astype(str).str.strip() == target_site]

    if company not in (None, "", "전체"):
        target_company = str(company).strip()
        if "업체명" in df.columns:
            df = df[df["업체명"].astype(str).str.strip() == target_company]

    df = df.sort_values(by="_dt", kind="stable")

    rows: list[dict[str, Any]] = []
    supply_sum = 0.0
    for _, rec in df.iterrows():
        qty = _to_float(rec.get("검사량"))
        unit_price = _to_float(rec.get("단가"))
        trip_fee = _to_float(rec.get("출장비"))
        lump_sum = _to_float(rec.get("일식"))
        explicit_fee = _to_float(rec.get("검사비"))

        if explicit_fee > 0:
            supply = explicit_fee
        else:
            supply = (qty * unit_price) + trip_fee + lump_sum

        method = _clean_str(rec.get("검사방법"))
        item_name = _clean_str(rec.get("검사품명")) or _clean_str(rec.get("적용코드"))
        if supply <= 0 and not method and qty <= 0:
            continue

        dt_val = rec.get("_dt")
        date_str = dt_val.strftime("%Y-%m-%d") if pd.notna(dt_val) else _clean_str(rec.get("Date"))
        supply_int = int(round(supply))
        supply_sum += supply_int
        rows.append(
            {
                "date": date_str,
                "site": _clean_str(rec.get("Site")),
                "company": _clean_str(rec.get("업체명")) or _clean_str(rec.get("발주처")),
                "method": method,
                "item": item_name,
                "qty": qty,
                "unit_price": int(round(unit_price)),
                "trip_fee": int(round(trip_fee)),
                "lump_sum": int(round(lump_sum)),
                "supply": supply_int,
                "note": _clean_str(rec.get("Note")),
            }
        )

    supply_total = int(round(supply_sum))
    vat_total = int(round(supply_total * 0.1))
    return {
        "rows": rows,
        "supply_total": supply_total,
        "vat_total": vat_total,
        "grand_total": supply_total + vat_total,
    }


def export_smb_statement_excel(
    daily_usage_df: pd.DataFrame | None,
    output_path: str,
    *,
    year: int | str | None = None,
    month: int | str | None = None,
    site: str | None = None,
    company: str | None = None,
    profile_name: str = "",
) -> dict[str, Any]:
    """Generate a formatted Excel billing statement (거래명세서/청구서) from daily usage records."""
    summary = build_statement_rows(
        daily_usage_df, year=year, month=month, site=site, company=company
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "거래명세서"

    title_font = Font(name="Malgun Gothic", size=16, bold=True, color="1B365D")
    label_font = Font(name="Malgun Gothic", size=10, bold=True, color="333333")
    value_font = Font(name="Malgun Gothic", size=10)
    header_font = Font(name="Malgun Gothic", size=10, bold=True, color="FFFFFF")
    total_font = Font(name="Malgun Gothic", size=10, bold=True, color="0D47A1")

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sub_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
    total_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:K1")
    ws["A1"] = "비파괴검사 거래명세서 (월간 청구내역)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    period_parts = []
    if year not in (None, "", "전체"):
        period_parts.append(f"{year}년")
    if month not in (None, "", "전체"):
        period_parts.append(f"{month}월")
    period_label = " ".join(period_parts) if period_parts else "전체 기간"
    target_label = " / ".join(
        p for p in (
            company if company not in (None, "", "전체") else "",
            site if site not in (None, "", "전체") else "",
        ) if p
    ) or "전체 거래처·현장"

    meta = [
        ("A3", "청구 대상", "B3:E3", target_label),
        ("F3", "청구 기간", "G3:K3", period_label),
        ("A4", "사업 프로필", "B4:E4", profile_name or "공통"),
        ("F4", "작성 일자", "G4:K4", datetime.now().strftime("%Y-%m-%d")),
        ("A5", "공급가액", "B5:C5", summary["supply_total"]),
        ("D5", "부가세(10%)", "E5:F5", summary["vat_total"]),
        ("G5", "총 청구금액", "H5:K5", summary["grand_total"]),
    ]
    for lbl_cell, lbl_text, val_range, val in meta:
        ws[lbl_cell] = lbl_text
        ws[lbl_cell].font = label_font
        ws[lbl_cell].fill = sub_fill
        ws[lbl_cell].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(val_range)
        top_left = val_range.split(":")[0]
        ws[top_left] = val
        ws[top_left].font = total_font if isinstance(val, int) else value_font
        ws[top_left].alignment = Alignment(horizontal="right" if isinstance(val, int) else "left", vertical="center")
        if isinstance(val, int):
            ws[top_left].number_format = "#,##0"

    for r in (3, 4, 5):
        ws.row_dimensions[r].height = 22
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = border

    headers = [
        "일자", "현장명", "업체명", "검사방법", "검사품명/규격",
        "검사량", "단가", "출장비", "일식", "공급가액", "비고",
    ]
    header_row = 7
    ws.row_dimensions[header_row].height = 24
    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    row_num = header_row + 1
    for item in summary["rows"]:
        values = [
            item["date"],
            item["site"],
            item["company"],
            item["method"],
            item["item"],
            item["qty"],
            item["unit_price"],
            item["trip_fee"],
            item["lump_sum"],
            item["supply"],
            item["note"],
        ]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font = value_font
            cell.border = border
            if col_idx in (1, 4):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.##"
            elif col_idx in (7, 8, 9, 10):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row_num].height = 20
        row_num += 1

    # Total row
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=9)
    ws.cell(row=row_num, column=1, value="합계 (공급가액)").font = total_font
    ws.cell(row=row_num, column=1).alignment = Alignment(horizontal="center", vertical="center")
    tot_cell = ws.cell(row=row_num, column=10, value=summary["supply_total"])
    tot_cell.font = total_font
    tot_cell.number_format = "#,##0"
    tot_cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=row_num, column=11, value=f"부가세 포함 {summary['grand_total']:,}원").font = total_font
    for col_idx in range(1, 12):
        c = ws.cell(row=row_num, column=col_idx)
        c.fill = total_fill
        c.border = border
    ws.row_dimensions[row_num].height = 24

    widths = [12, 18, 18, 10, 18, 10, 12, 12, 12, 14, 24]
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return summary

