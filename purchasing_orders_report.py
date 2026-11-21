#!/usr/bin/env python3
"""Build a four-sheet purchasing analysis workbook from a raw purchasing export.

Requirements:
    python -m pip install openpyxl

Usage:
    python purchasing_orders_report.py input.xlsx
    python purchasing_orders_report.py input.xlsx --output purchasing_analysis.xlsx

The original first worksheet is preserved. Existing generated worksheets named
"Recurring Purchases", "By TO", and "Monthly Spending" are replaced.
"""

from __future__ import annotations

import argparse
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


GENERATED_SHEETS = ("Recurring Purchases", "By TO", "Monthly Spending")
NAVY = "17365D"
BLUE = "2F75B5"
TEAL = "1F6D7A"
PALE_BLUE = "D9EAF7"
PALE_TEAL = "DDEBF7"
PALE_GOLD = "FFF2CC"
PALE_GREEN = "E2F0D9"
PALE_RED = "FCE4D6"
GRAY = "667085"
LIGHT_GRAY = "E7EAF0"
WHITE = "FFFFFF"
CURRENCY = '$#,##0.00;[Red]($#,##0.00);-'


def normalized_header(value: Any) -> str:
    return str(value or "").replace("\xa0", " ").strip()


def find_header_row(ws) -> tuple[int, dict[str, int]]:
    """Find the row containing the purchasing export's real column headers."""
    for row_number in range(1, min(ws.max_row, 30) + 1):
        mapping = {
            normalized_header(ws.cell(row_number, column).value): column
            for column in range(1, min(ws.max_column, 40) + 1)
            if normalized_header(ws.cell(row_number, column).value)
        }
        if "Request #" in mapping and "Total $" in mapping and "Date Created" in mapping:
            return row_number, mapping
    raise ValueError("Could not find a header row containing Request #, Total $, and Date Created.")


def as_date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = str(value or "").strip()
    for pattern in ("%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            pass
    return None


def as_number(value: Any) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value or "0").replace("$", "").replace(",", ""))
    except ValueError:
        return 0.0


def clean_vendor(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    return text.split("(", 1)[0].strip()


def normalize_vendor(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_vendor(value).lower()).strip()


def normalize_item(value: Any) -> str:
    text = str(value or "").lower()
    text = re.sub(r"https?://\S+", " ", text)
    text = re.sub(r"\b\d+\s*\[ea\]", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def spend_stage(status: Any) -> str:
    status_text = str(status or "").strip().lower()
    if status_text == "complete":
        return "Completed"
    if status_text in {"cancelled", "denied"}:
        return "Cancelled/Denied"
    return "Pending"


def to_code(number: int) -> str:
    return f"TO-{number}"


def extract_tos(value: Any) -> list[int]:
    text = str(value or "").upper()
    found = {int(match) for match in re.findall(r"\bTO[\s-]*(\d{1,6})\b", text)}
    return sorted(number for number in found if number > 0)


def to_labels(tos: Iterable[int]) -> str:
    values = list(tos)
    return ", ".join(f"TO {number}" for number in values) if values else "Unmapped"


def most_common(values: Iterable[str]) -> str:
    counts = Counter(value for value in values if value)
    return sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0] if counts else ""


def read_records(ws) -> list[dict[str, Any]]:
    header_row, columns = find_header_row(ws)
    records: list[dict[str, Any]] = []
    for row_number in range(header_row + 1, ws.max_row + 1):
        record = {header: ws.cell(row_number, column).value for header, column in columns.items()}
        created = as_date(record.get("Date Created"))
        if not record.get("Request #") or not created:
            continue
        tos = extract_tos(record.get("Account Numbers"))
        record.update(
            {
                "_date": created,
                "_month": datetime(created.year, created.month, 1),
                "_amount": as_number(record.get("Total $")),
                "_vendor": clean_vendor(record.get("Vendor")),
                "_tos": tos,
                "_to_labels": to_labels(tos),
                "_stage": spend_stage(record.get("Status")),
            }
        )
        records.append(record)
    if not records:
        raise ValueError("No dated purchasing records were found below the detected header row.")
    return records


def add_table(ws, ref: str, name: str, style: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name=style,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def title(ws, cell_range: str, text: str) -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = text
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def note(ws, cell_range: str, text: str) -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = text
    cell.fill = PatternFill("solid", fgColor=PALE_BLUE)
    cell.font = Font(name="Aptos", size=10, italic=True, color=GRAY)
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def card(ws, label_range: str, value_range: str, label: str, formula: str, number_format: str = "#,##0") -> None:
    ws.merge_cells(label_range)
    ws.merge_cells(value_range)
    label_cell = ws[label_range.split(":")[0]]
    value_cell = ws[value_range.split(":")[0]]
    label_cell.value = label
    label_cell.fill = PatternFill("solid", fgColor=BLUE)
    label_cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.value = formula
    value_cell.fill = PatternFill("solid", fgColor=PALE_TEAL)
    value_cell.font = Font(name="Aptos", size=15, bold=True, color=NAVY)
    value_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.number_format = number_format
    value_cell.border = Border(
        left=Side(style="thin", color=LIGHT_GRAY),
        right=Side(style="thin", color=LIGHT_GRAY),
        top=Side(style="thin", color=LIGHT_GRAY),
        bottom=Side(style="thin", color=LIGHT_GRAY),
    )


def style_header(ws, row: int, start_col: int, end_col: int) -> None:
    for column in range(start_col, end_col + 1):
        cell = ws.cell(row, column)
        cell.fill = PatternFill("solid", fgColor=TEAL)
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_widths(ws, widths: dict[str, float]) -> None:
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def add_stage_formatting(ws, cell_range: str, first_row: int, stage_column: str) -> None:
    ws.conditional_formatting.add(
        cell_range,
        FormulaRule(formula=[f'${stage_column}{first_row}="Completed"'], fill=PatternFill("solid", fgColor=PALE_GREEN)),
    )
    ws.conditional_formatting.add(
        cell_range,
        FormulaRule(formula=[f'${stage_column}{first_row}="Pending"'], fill=PatternFill("solid", fgColor=PALE_GOLD)),
    )
    ws.conditional_formatting.add(
        cell_range,
        FormulaRule(formula=[f'${stage_column}{first_row}="Cancelled/Denied"'], fill=PatternFill("solid", fgColor=PALE_RED)),
    )


def write_rows(ws, start_row: int, rows: list[list[Any]]) -> int:
    for row_offset, values in enumerate(rows):
        for column, value in enumerate(values, start=1):
            ws.cell(start_row + row_offset, column, value)
    return start_row + max(len(rows), 1) - 1


def build_monthly_sheet(ws, records: list[dict[str, Any]]) -> None:
    ws.sheet_view.showGridLines = False
    title(ws, "A1:O1", "Monthly Purchasing & Spending")
    note(
        ws,
        "A2:O2",
        "Choose Month 1 and Month 2 in the Combined Month View at right. The top KPI bar and combined result row "
        "update together; selecting the same month twice counts it only once. Use the detail-table filters for the matching transactions.",
    )
    headers = ["Month", "Date", "Request #", "PO #", "Vendor", "What Was Ordered", "Amount", "Status", "Spend Stage", "Order Type", "TO(s)", "Created By", "Purchaser", "Has Splits", "Paid"]
    for column, value in enumerate(headers, start=1):
        ws.cell(8, column, value)
    sorted_records = sorted(records, key=lambda record: (record["_date"], str(record.get("Request #"))), reverse=True)
    rows = [
        [
            record["_month"], record["_date"], record.get("Request #", ""), record.get("PO #", ""), record["_vendor"],
            record.get("Line Item Summary", ""), record["_amount"], record.get("Status", ""), record["_stage"],
            record.get("Order Type", ""), record["_to_labels"], record.get("Created By", ""), record.get("Purchaser", ""),
            record.get("Has Splits", ""), record.get("Paid", ""),
        ]
        for record in sorted_records
    ]
    last = write_rows(ws, 9, rows)
    add_table(ws, f"A8:O{last}", "MonthlySpendingDetailTable", "TableStyleMedium2")
    style_header(ws, 8, 1, 15)
    for row in range(9, last + 1):
        ws.cell(row, 1).number_format = "mmm yyyy"
        ws.cell(row, 2).number_format = "yyyy-mm-dd"
        ws.cell(row, 7).number_format = CURRENCY
        ws.cell(row, 6).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
    add_stage_formatting(ws, f"I9:I{last}", 9, "I")

    months = sorted({record["_month"] for record in records})
    summary_headers = ["Month", "Orders / Requests", "Total Requested", "Completed Spend", "Pending Amount", "Cancelled / Denied", "Average Request"]
    for column, value in enumerate(summary_headers, start=17):
        ws.cell(8, column, value)
    for offset, month in enumerate(months, start=9):
        ws.cell(offset, 17, month)
        ws.cell(offset, 18, f'=COUNTIF($A$9:$A${last},Q{offset})')
        ws.cell(offset, 19, f'=SUMIF($A$9:$A${last},Q{offset},$G$9:$G${last})')
        ws.cell(offset, 20, f'=SUMIFS($G$9:$G${last},$A$9:$A${last},Q{offset},$I$9:$I${last},"Completed")')
        ws.cell(offset, 21, f'=SUMIFS($G$9:$G${last},$A$9:$A${last},Q{offset},$I$9:$I${last},"Pending")')
        ws.cell(offset, 22, f'=SUMIFS($G$9:$G${last},$A$9:$A${last},Q{offset},$I$9:$I${last},"Cancelled/Denied")')
        ws.cell(offset, 23, f'=IFERROR(S{offset}/R{offset},0)')
        ws.cell(offset, 17).number_format = "mmm yyyy"
        for column in range(19, 24):
            ws.cell(offset, column).number_format = CURRENCY
    summary_last = 8 + len(months)
    add_table(ws, f"Q8:W{summary_last}", "MonthlySpendingSummaryTable", "TableStyleMedium4")
    style_header(ws, 8, 17, 23)

    ws.merge_cells("Q1:X1")
    ws["Q1"] = "Combined Two-Month View"
    for row in ws["Q1:X1"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(name="Aptos Display", size=13, bold=True, color=WHITE)
            cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("Q2:R2")
    ws["Q2"] = "Month 1"
    for row in ws["Q2:R2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=TEAL)
            cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("S2:T2")
    ws["S2"] = months[-1]
    for row in ws["S2:T2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=PALE_GOLD)
            cell.font = Font(name="Aptos", size=12, bold=True, color=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style="thin", color=BLUE), right=Side(style="thin", color=BLUE),
                top=Side(style="thin", color=BLUE), bottom=Side(style="thin", color=BLUE),
            )
    ws["S2"].number_format = "mmm yyyy"
    ws.merge_cells("U2:V2")
    ws["U2"] = "Month 2"
    for row in ws["U2:V2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=TEAL)
            cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("W2:X2")
    ws["W2"] = months[-2] if len(months) > 1 else months[-1]
    for row in ws["W2:X2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=PALE_GOLD)
            cell.font = Font(name="Aptos", size=12, bold=True, color=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style="thin", color=BLUE), right=Side(style="thin", color=BLUE),
                top=Side(style="thin", color=BLUE), bottom=Side(style="thin", color=BLUE),
            )
    ws["W2"].number_format = "mmm yyyy"
    month_validation = DataValidation(type="list", formula1=f"$Q$9:$Q${summary_last}", allow_blank=False)
    month_validation.promptTitle = "Select Month"
    month_validation.prompt = "Choose either month to update the combined spending view."
    month_validation.errorTitle = "Invalid Month"
    month_validation.error = "Choose a month from the dropdown list."
    month_validation.errorStyle = "stop"
    month_validation.showErrorMessage = True
    month_validation.showInputMessage = True
    ws.add_data_validation(month_validation)
    month_validation.add(ws["S2"])
    month_validation.add(ws["W2"])

    selected_headers = ["Month 1", "Month 2", "Orders / Requests", "Total Requested", "Completed Spend", "Pending Amount", "Cancelled / Denied", "Average Request"]
    for column, value in enumerate(selected_headers, start=17):
        ws.cell(4, column, value)
    style_header(ws, 4, 17, 24)
    ws["Q5"] = "=$S$2"
    ws["R5"] = "=$W$2"
    ws["S5"] = f'=COUNTIF($A$9:$A${last},$S$2)+IF($W$2=$S$2,0,COUNTIF($A$9:$A${last},$W$2))'
    ws["T5"] = f'=SUMIF($A$9:$A${last},$S$2,$G$9:$G${last})+IF($W$2=$S$2,0,SUMIF($A$9:$A${last},$W$2,$G$9:$G${last}))'
    ws["U5"] = f'=SUMIFS($G$9:$G${last},$A$9:$A${last},$S$2,$I$9:$I${last},"Completed")+IF($W$2=$S$2,0,SUMIFS($G$9:$G${last},$A$9:$A${last},$W$2,$I$9:$I${last},"Completed"))'
    ws["V5"] = f'=SUMIFS($G$9:$G${last},$A$9:$A${last},$S$2,$I$9:$I${last},"Pending")+IF($W$2=$S$2,0,SUMIFS($G$9:$G${last},$A$9:$A${last},$W$2,$I$9:$I${last},"Pending"))'
    ws["W5"] = f'=SUMIFS($G$9:$G${last},$A$9:$A${last},$S$2,$I$9:$I${last},"Cancelled/Denied")+IF($W$2=$S$2,0,SUMIFS($G$9:$G${last},$A$9:$A${last},$W$2,$I$9:$I${last},"Cancelled/Denied"))'
    ws["X5"] = "=IFERROR(T5/S5,0)"
    for cell in ws[5][16:24]:
        cell.fill = PatternFill("solid", fgColor=PALE_TEAL)
        cell.font = Font(name="Aptos", size=11, bold=True, color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin", color=LIGHT_GRAY), right=Side(style="thin", color=LIGHT_GRAY),
            top=Side(style="thin", color=LIGHT_GRAY), bottom=Side(style="thin", color=LIGHT_GRAY),
        )
    ws["Q5"].number_format = "mmm yyyy"
    ws["R5"].number_format = "mmm yyyy"
    ws["S5"].number_format = "#,##0"
    for cell in ws[5][19:24]:
        cell.number_format = CURRENCY

    card(ws, "A4:B4", "A5:B5", "Selected Orders / Requests", "=S5")
    card(ws, "D4:E4", "D5:E5", "Selected Total Requested", "=T5", CURRENCY)
    card(ws, "G4:H4", "G5:H5", "Selected Completed Spend", "=U5", CURRENCY)
    card(ws, "J4:K4", "J5:K5", "Selected Pending Amount", "=V5", CURRENCY)
    card(ws, "M4:N4", "M5:N5", "Selected Cancelled / Denied", "=W5", CURRENCY)
    ws.freeze_panes = "A9"
    set_widths(ws, {"A": 12, "B": 12, "C": 17, "D": 15, "E": 26, "F": 48, "G": 14, "H": 15, "I": 18, "J": 22, "K": 20, "L": 20, "M": 18, "N": 12, "O": 10, "P": 3, "Q": 12, "R": 14, "S": 16, "T": 16, "U": 16, "V": 18, "W": 15, "X": 15})
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[8].height = 34


def build_to_sheet(ws, records: list[dict[str, Any]]) -> None:
    ws.sheet_view.showGridLines = False
    title(ws, "A1:S1", "Purchasing by Task Order")
    note(
        ws,
        "A2:S2",
        "Task orders are recognized from entries such as TO-12. For orders listing multiple TOs, "
        "Allocated Amount is an equal-split estimate; Original Order Total remains visible. Unmapped rows flag source accounts that could not be matched.",
    )
    headers = ["TO", "Account Code", "Date", "Month", "Request #", "PO #", "Vendor", "What Was Ordered", "Original Order Total", "Allocated Amount", "Allocation Basis", "Status", "Spend Stage", "Order Type", "Has Splits", "TOs on Order", "Created By", "Purchaser", "Account Numbers"]
    for column, value in enumerate(headers, start=1):
        ws.cell(8, column, value)
    detail: list[tuple[int | None, dict[str, Any], float, str]] = []
    for record in records:
        destinations: list[int | None] = record["_tos"] or [None]
        allocation = record["_amount"] / len(destinations)
        for number in destinations:
            if number is None:
                basis = "Unmapped account"
            elif len(destinations) > 1:
                basis = f"Even split estimate ({len(destinations)} TOs)"
            elif str(record.get("Has Splits", "")).lower() == "yes":
                basis = "Source marked split; one TO recognized"
            else:
                basis = "Direct assignment"
            detail.append((number, record, allocation, basis))
    detail.sort(key=lambda item: (item[0] if item[0] is not None else 99, -item[1]["_date"].timestamp(), str(item[1].get("Request #"))))
    rows = []
    for number, record, allocation, basis in detail:
        rows.append(
            [
                "Unmapped" if number is None else f"TO {number}", "(not found)" if number is None else to_code(number),
                record["_date"], record["_month"], record.get("Request #", ""), record.get("PO #", ""), record["_vendor"],
                record.get("Line Item Summary", ""), record["_amount"], allocation, basis, record.get("Status", ""),
                record["_stage"], record.get("Order Type", ""), record.get("Has Splits", ""), record["_to_labels"],
                record.get("Created By", ""), record.get("Purchaser", ""), record.get("Account Numbers", ""),
            ]
        )
    last = write_rows(ws, 9, rows)
    add_table(ws, f"A8:S{last}", "TaskOrderDetailTable", "TableStyleMedium2")
    style_header(ws, 8, 1, 19)
    for row in range(9, last + 1):
        ws.cell(row, 3).number_format = "yyyy-mm-dd"
        ws.cell(row, 4).number_format = "mmm yyyy"
        ws.cell(row, 9).number_format = CURRENCY
        ws.cell(row, 10).number_format = CURRENCY
        ws.cell(row, 8).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
    add_stage_formatting(ws, f"M9:M{last}", 9, "M")

    summary_headers = ["TO", "Account Code", "Records", "Requested Allocation", "Completed Allocation", "Pending Allocation", "Cancelled / Denied", "Est. Split Records"]
    for column, value in enumerate(summary_headers, start=21):
        ws.cell(8, column, value)
    labels = [f"Even split estimate ({number} TOs)" for number in range(2, 22)]
    task_orders = sorted({number for record in records for number in record["_tos"]})
    summary = [(f"TO {number}", to_code(number)) for number in task_orders] + [("Unmapped", "(not found)")]
    for offset, (label, code) in enumerate(summary, start=9):
        ws.cell(offset, 21, label)
        ws.cell(offset, 22, code)
        ws.cell(offset, 23, f'=COUNTIF($A$9:$A${last},U{offset})')
        ws.cell(offset, 24, f'=SUMIF($A$9:$A${last},U{offset},$J$9:$J${last})')
        ws.cell(offset, 25, f'=SUMIFS($J$9:$J${last},$A$9:$A${last},U{offset},$M$9:$M${last},"Completed")')
        ws.cell(offset, 26, f'=SUMIFS($J$9:$J${last},$A$9:$A${last},U{offset},$M$9:$M${last},"Pending")')
        ws.cell(offset, 27, f'=SUMIFS($J$9:$J${last},$A$9:$A${last},U{offset},$M$9:$M${last},"Cancelled/Denied")')
        split_parts = [f'COUNTIFS($A$9:$A${last},U{offset},$K$9:$K${last},"{basis}")' for basis in labels]
        ws.cell(offset, 28, "=" + "+".join(split_parts))
        for column in range(24, 28):
            ws.cell(offset, column).number_format = CURRENCY
    summary_last = 8 + len(summary)
    add_table(ws, f"U8:AB{summary_last}", "TaskOrderSummaryTable", "TableStyleMedium4")
    style_header(ws, 8, 21, 28)
    all_split_parts = [f'COUNTIF($K$9:$K${last},"{basis}")' for basis in labels]
    card(ws, "A4:B4", "A5:B5", "Total Requested", f'=SUM($J$9:$J${last})', CURRENCY)
    card(ws, "D4:E4", "D5:E5", "Completed Spend", f'=SUMIFS($J$9:$J${last},$M$9:$M${last},"Completed")', CURRENCY)
    card(ws, "G4:H4", "G5:H5", "Estimated Split Rows", "=" + "+".join(all_split_parts))
    card(ws, "J4:K4", "J5:K5", "Unmapped Requested", f'=SUMIF($A$9:$A${last},"Unmapped",$J$9:$J${last})', CURRENCY)
    ws.freeze_panes = "C9"
    set_widths(ws, {"A": 11, "B": 16, "C": 12, "D": 12, "E": 17, "F": 15, "G": 26, "H": 48, "I": 17, "J": 16, "K": 28, "L": 15, "M": 18, "N": 22, "O": 12, "P": 22, "Q": 20, "R": 18, "S": 42, "T": 3, "U": 11, "V": 16, "W": 10, "X": 18, "Y": 18, "Z": 17, "AA": 18, "AB": 16})
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 40
    ws.row_dimensions[8].height = 36


def build_recurring_sheet(ws, records: list[dict[str, Any]]) -> None:
    ws.sheet_view.showGridLines = False
    title(ws, "A1:M1", "Monthly Recurring Purchases & Requests")
    note(
        ws,
        "A2:M2",
        "A recurring pattern is the same normalized vendor + order description appearing in at least two different calendar months. "
        "Quantities, punctuation, and web links are ignored for matching. Review the detail before treating a pattern as a contractual recurrence.",
    )
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        key = f'{normalize_vendor(record.get("Vendor"))}|{normalize_item(record.get("Line Item Summary"))}'
        if key != "|":
            grouped[key].append(record)
    groups = []
    for key, group_records in grouped.items():
        months = {(record["_date"].year, record["_date"].month) for record in group_records}
        if len(group_records) >= 2 and len(months) >= 2:
            groups.append((key, group_records, months))
    groups.sort(key=lambda item: (-len(item[2]), -len(item[1]), item[1][0]["_vendor"]))

    detail_rows: list[list[Any]] = []
    summary_rows: list[list[Any]] = []
    for index, (_, group_records, months) in enumerate(groups, start=1):
        pattern_id = f"RP-{index:03d}"
        ordered = sorted(group_records, key=lambda record: record["_date"])
        tos = sorted({number for record in ordered for number in record["_tos"]})
        summary_rows.append(
            [
                pattern_id, ordered[0]["_vendor"], ordered[0].get("Line Item Summary", ""), None, len(months),
                ordered[0]["_date"], ordered[-1]["_date"], None, None, ordered[-1].get("Status", ""),
                to_labels(tos), most_common(str(record.get("Order Type", "")) for record in ordered),
            ]
        )
        for record in sorted(group_records, key=lambda record: record["_date"], reverse=True):
            detail_rows.append(
                [
                    pattern_id, record["_month"], record["_date"], record.get("Request #", ""), record["_vendor"],
                    record.get("Line Item Summary", ""), record["_amount"], record.get("Status", ""), record.get("Order Type", ""),
                    record["_to_labels"], record.get("Created By", ""), record.get("Purchaser", ""), record["_stage"],
                ]
            )

    detail_headers = ["Pattern ID", "Month", "Date", "Request #", "Vendor", "What Was Ordered", "Amount", "Status", "Order Type", "TO(s)", "Created By", "Purchaser", "Spend Stage"]
    for column, value in enumerate(detail_headers, start=1):
        ws.cell(8, column, value)
    if not detail_rows:
        detail_rows = [[None] * len(detail_headers)]
    detail_last = write_rows(ws, 9, detail_rows)
    add_table(ws, f"A8:M{detail_last}", "RecurringPurchaseDetailTable", "TableStyleMedium2")
    style_header(ws, 8, 1, 13)
    for row in range(9, detail_last + 1):
        ws.cell(row, 2).number_format = "mmm yyyy"
        ws.cell(row, 3).number_format = "yyyy-mm-dd"
        ws.cell(row, 7).number_format = CURRENCY
        ws.cell(row, 6).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
    add_stage_formatting(ws, f"M9:M{detail_last}", 9, "M")

    summary_headers = ["Pattern ID", "Vendor", "Recurring Description", "Occurrences", "Active Months", "First Date", "Latest Date", "Total Requested", "Average Request", "Latest Status", "TO(s)", "Primary Order Type"]
    for column, value in enumerate(summary_headers, start=15):
        ws.cell(8, column, value)
    if not summary_rows:
        summary_rows = [[None] * len(summary_headers)]
    for offset, values in enumerate(summary_rows, start=9):
        for column, value in enumerate(values, start=15):
            ws.cell(offset, column, value)
        if ws.cell(offset, 15).value:
            ws.cell(offset, 18, f'=COUNTIF($A$9:$A${detail_last},O{offset})')
            ws.cell(offset, 22, f'=SUMIF($A$9:$A${detail_last},O{offset},$G$9:$G${detail_last})')
            ws.cell(offset, 23, f'=IFERROR(V{offset}/R{offset},0)')
        ws.cell(offset, 20).number_format = "yyyy-mm-dd"
        ws.cell(offset, 21).number_format = "yyyy-mm-dd"
        ws.cell(offset, 22).number_format = CURRENCY
        ws.cell(offset, 23).number_format = CURRENCY
        ws.cell(offset, 17).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[offset].height = 30
    summary_last = 8 + len(summary_rows)
    add_table(ws, f"O8:Z{summary_last}", "RecurringPurchaseSummaryTable", "TableStyleMedium4")
    style_header(ws, 8, 15, 26)
    card(ws, "A4:B4", "A5:B5", "Recurring Patterns", f'=COUNTA($O$9:$O${summary_last})')
    card(ws, "D4:E4", "D5:E5", "Recurring Requests", f'=COUNTA($A$9:$A${detail_last})')
    card(ws, "G4:H4", "G5:H5", "Total Requested", f'=SUM($G$9:$G${detail_last})', CURRENCY)
    card(ws, "J4:K4", "J5:K5", "Completed Spend", f'=SUMIFS($G$9:$G${detail_last},$M$9:$M${detail_last},"Completed")', CURRENCY)
    ws.freeze_panes = "B9"
    set_widths(ws, {"A": 13, "B": 12, "C": 12, "D": 17, "E": 26, "F": 48, "G": 14, "H": 15, "I": 22, "J": 20, "K": 20, "L": 18, "M": 18, "N": 3, "O": 13, "P": 26, "Q": 48, "R": 12, "S": 13, "T": 12, "U": 12, "V": 16, "W": 15, "X": 15, "Y": 20, "Z": 22})
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 40
    ws.row_dimensions[8].height = 36


def build_report(input_path: Path, output_path: Path) -> None:
    workbook = load_workbook(input_path)
    for name in GENERATED_SHEETS:
        if name in workbook.sheetnames:
            workbook.remove(workbook[name])
    raw_sheet = workbook.worksheets[0]
    records = read_records(raw_sheet)

    recurring_sheet = workbook.create_sheet("Recurring Purchases")
    to_sheet = workbook.create_sheet("By TO")
    monthly_sheet = workbook.create_sheet("Monthly Spending")
    build_monthly_sheet(monthly_sheet, records)
    build_to_sheet(to_sheet, records)
    build_recurring_sheet(recurring_sheet, records)

    workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build recurring, TO, and monthly purchasing analysis sheets.")
    parser.add_argument("input", type=Path, help="Raw purchasing .xlsx export")
    parser.add_argument("--output", "-o", type=Path, help="Output .xlsx path; defaults to <input>_analysis.xlsx")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = args.input.expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")
    output_path = args.output.expanduser().resolve() if args.output else input_path.with_name(f"{input_path.stem}_analysis.xlsx")
    build_report(input_path, output_path)
    print(f"Created: {output_path}")


if __name__ == "__main__":
    main()
