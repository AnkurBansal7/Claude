#!/usr/bin/env python3
"""
Build the Petpooja Daily Discount Dashboard workbook from Item Wise Report
(with Bill No.) and Payment Wise Order Summary xlsx exports.

Both reports follow Petpooja's standard notification-email template:
    Row 1: "Date:"             , "<start> to <end>"
    Row 2: "Name:"             , "<report name>"
    Row 3: "Restaurant Name:"  , "<outlet>"
    Row 4: (blank)
    Row 5: column headers
    Row 6+: data

Item Wise Report With Bill No. columns (as observed):
    Date, Timestamp, Server Name, Table No., Invoice No., hsn_code,
    Category, Item, Variation, Price, Qty., Sub Total, Discount, Tax,
    Final Total
  -> one row per line item; grouped by Invoice No. to get per-order totals.

Payment Wise Order Summary columns vary by account configuration, so the
parser looks for a bill/invoice column and a payment-mode/order-type column
by header keyword instead of a fixed position. Until that report is enabled,
pass payment_wise_xlsx=None for an outlet and every order is marked
"Pending (awaiting Payment Wise report)" in the workbook, ready to resolve
automatically once real data is supplied (see PaymentMap sheet).

Usage:
    python3 build_dashboard.py <manifest.json> <output.xlsx>

Manifest schema:
{
  "report_date": "2026-09-22",
  "outlets": [
    {"name": "Tuskin Coffee Andheri West",
     "item_wise_xlsx": "/path/Item_bill_report_....xlsx",
     "payment_wise_xlsx": "/path/Payment_wise_....xlsx" | null}
  ]
}
"""
import sys
import json
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color="1F2937")
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, size=10, color="6B7280")
LABEL_FONT = Font(name=FONT_NAME, bold=True, size=10)
BODY_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_STAFF = PatternFill("solid", fgColor="FCA5A5")      # 100% discount dine-in
FILL_HIGH_ONLINE = PatternFill("solid", fgColor="FDBA74")  # >50% Swiggy/Zomato
FILL_TIER1 = PatternFill("solid", fgColor="FEF3C7")       # dine-in >15%
FILL_TIER2 = PatternFill("solid", fgColor="FDE68A")       # dine-in >30%
FILL_TIER3 = PatternFill("solid", fgColor="FCA5A5")       # dine-in >50%

PENDING_LABEL = "Pending (awaiting Payment Wise report)"


def _find_header_row(ws, max_scan=10):
    """Petpooja report exports start with 3 metadata rows, a blank row, then headers."""
    for r in range(1, max_scan + 1):
        row_vals = [c.value for c in ws[r]]
        non_null = [v for v in row_vals if v not in (None, "")]
        if len(non_null) >= 4 and all(isinstance(v, str) for v in non_null):
            return r
    raise ValueError(f"Could not locate header row in sheet {ws.title!r}")


def parse_item_wise(path):
    """Return list of {invoice, sub_total, discount, final_total} per line item."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header_row = _find_header_row(ws)
    headers = [str(c.value).strip() if c.value else "" for c in ws[header_row]]
    idx = {h.lower(): i for i, h in enumerate(headers)}

    def col(*keywords):
        for h, i in idx.items():
            if all(k in h for k in keywords):
                return i
        raise KeyError(f"No column matching {keywords} in {headers}")

    c_invoice = col("invoice")
    c_subtotal = col("sub", "total")
    c_discount = col("discount")
    c_final = col("final", "total")

    rows = []
    for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if r[c_invoice] in (None, ""):
            continue
        rows.append({
            "invoice": str(r[c_invoice]),
            "sub_total": float(r[c_subtotal] or 0),
            "discount": float(r[c_discount] or 0),
            "final_total": float(r[c_final] or 0),
        })
    return rows


def parse_payment_wise(path):
    """Return {invoice: payment_mode} using flexible header matching."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header_row = _find_header_row(ws)
    headers = [str(c.value).strip() if c.value else "" for c in ws[header_row]]
    idx = {h.lower(): i for i, h in enumerate(headers)}

    def find(*keywords):
        for h, i in idx.items():
            if all(k in h for k in keywords):
                return i
        return None

    c_invoice = find("invoice") or find("bill")
    c_mode = find("payment", "mode") or find("payment", "type") or find("order", "type")
    if c_invoice is None or c_mode is None:
        raise KeyError(
            f"Could not find invoice/payment-mode columns in headers: {headers}"
        )

    mapping = {}
    for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if r[c_invoice] in (None, ""):
            continue
        mapping[str(r[c_invoice])] = str(r[c_mode] or "").strip()
    return mapping


def classify_platform(payment_mode: str) -> str:
    m = (payment_mode or "").lower()
    if "zomato" in m:
        return "Zomato"
    if "swiggy" in m:
        return "Swiggy"
    return "Dine-in"


def aggregate_orders(item_rows):
    """Group line items by invoice -> per-order sub_total/discount/final_total."""
    agg = defaultdict(lambda: {"sub_total": 0.0, "discount": 0.0, "final_total": 0.0})
    for row in item_rows:
        a = agg[row["invoice"]]
        a["sub_total"] += row["sub_total"]
        a["discount"] += row["discount"]
        a["final_total"] += row["final_total"]
    return agg


def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r"[:\\/?*\[\]]", "-", name)
    return name[:31]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_outlet_sheet(wb, outlet_name, report_date, order_agg, payment_map):
    sheet_name = sanitize_sheet_name(outlet_name)
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = outlet_name
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Item Wise + Payment Wise Discount Analysis — {report_date}"
    ws["A2"].font = SUBTITLE_FONT

    # ---- Summary block (formulas reference the order table below) ----
    ws["A4"] = "Summary"
    ws["A4"].font = Font(name=FONT_NAME, bold=True, size=12)

    labels = [
        "Orders — Dine-in", "Orders — Swiggy", "Orders — Zomato", "Orders — Pending platform",
        "Avg Discount % — Dine-in", "Avg Discount % — Swiggy", "Avg Discount % — Zomato",
        "Orders >50% Discount — Swiggy", "Orders >50% Discount — Zomato",
        "100% Discount Dine-in Orders (Staff)",
        "Dine-in Orders >15% Discount", "Dine-in Orders >30% Discount", "Dine-in Orders >50% Discount",
    ]
    for i, lbl in enumerate(labels):
        ws.cell(row=5 + i, column=1, value=lbl).font = LABEL_FONT

    TABLE_HEADER_ROW = 20
    headers = ["Invoice No.", "Sub Total", "Discount", "Discount %", "Payment Mode",
               "Platform", "Flag: 100% Dine-in (Staff)", "Flag: >50% Online Discount",
               "Flag: Dine-in Discount Tier"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=TABLE_HEADER_ROW, column=c, value=h)
    style_header(ws, TABLE_HEADER_ROW, len(headers))

    invoices = sorted(order_agg.keys())
    first_data_row = TABLE_HEADER_ROW + 1
    for i, invoice in enumerate(invoices):
        r = first_data_row + i
        a = order_agg[invoice]
        mode = payment_map.get(invoice, "")
        platform = classify_platform(mode) if invoice in payment_map else PENDING_LABEL

        ws.cell(row=r, column=1, value=invoice).font = BODY_FONT
        ws.cell(row=r, column=2, value=round(a["sub_total"], 2)).number_format = "#,##0.00"
        ws.cell(row=r, column=3, value=round(a["discount"], 2)).number_format = "#,##0.00"
        disc_pct_cell = ws.cell(row=r, column=4,
                                 value=f"=IF(B{r}=0,0,C{r}/B{r})")
        disc_pct_cell.number_format = "0.0%"
        ws.cell(row=r, column=5, value=mode).font = BODY_FONT
        ws.cell(row=r, column=6, value=platform).font = BODY_FONT
        ws.cell(row=r, column=7,
                value=f'=IF(AND(F{r}="Dine-in",D{r}>=0.999),"STAFF ORDER","")')
        ws.cell(row=r, column=8,
                value=f'=IF(AND(OR(F{r}="Swiggy",F{r}="Zomato"),D{r}>0.5),"HIGH DISCOUNT","")')
        ws.cell(row=r, column=9,
                value=(f'=IF(F{r}<>"Dine-in","",'
                       f'IF(D{r}>0.5,">50%",IF(D{r}>0.3,">30%",IF(D{r}>0.15,">15%",""))))'))
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    last_data_row = first_data_row + len(invoices) - 1 if invoices else first_data_row
    rng = lambda col: f"{col}{first_data_row}:{col}{last_data_row}"

    if invoices:
        ws["B5"] = f'=COUNTIF({rng("F")},"Dine-in")'
        ws["B6"] = f'=COUNTIF({rng("F")},"Swiggy")'
        ws["B7"] = f'=COUNTIF({rng("F")},"Zomato")'
        ws["B8"] = f'=COUNTIF({rng("F")},"{PENDING_LABEL}")'
        ws["B9"] = f'=IFERROR(AVERAGEIF({rng("F")},"Dine-in",{rng("D")}),0)'
        ws["B10"] = f'=IFERROR(AVERAGEIF({rng("F")},"Swiggy",{rng("D")}),0)'
        ws["B11"] = f'=IFERROR(AVERAGEIF({rng("F")},"Zomato",{rng("D")}),0)'
        ws["B12"] = f'=COUNTIFS({rng("F")},"Swiggy",{rng("D")},">0.5")'
        ws["B13"] = f'=COUNTIFS({rng("F")},"Zomato",{rng("D")},">0.5")'
        ws["B14"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">=0.999")'
        ws["B15"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.15")'
        ws["B16"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.3")'
        ws["B17"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.5")'
    else:
        for r in range(5, 18):
            ws.cell(row=r, column=2, value=0)

    for r in (9, 10, 11):
        ws.cell(row=r, column=2).number_format = "0.0%"

    # Conditional formatting highlights on the order table
    if invoices:
        data_range = f"A{first_data_row}:I{last_data_row}"
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'$G{first_data_row}="STAFF ORDER"'], fill=FILL_STAFF))
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'$H{first_data_row}="HIGH DISCOUNT"'], fill=FILL_HIGH_ONLINE))
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'$I{first_data_row}=">50%"'], fill=FILL_TIER3))
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'$I{first_data_row}=">30%"'], fill=FILL_TIER2))
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'$I{first_data_row}=">15%"'], fill=FILL_TIER1))

    ws.freeze_panes = f"A{first_data_row}"
    autosize(ws, [14, 12, 12, 12, 24, 12, 22, 22, 20])
    return sheet_name, (first_data_row, last_data_row if invoices else None)


def build_payment_map_sheet(wb, all_rows):
    ws = wb.create_sheet("PaymentMap")
    ws["A1"] = "Raw Payment Wise Order Summary rows (used to resolve Platform on each outlet sheet)"
    ws["A1"].font = SUBTITLE_FONT
    headers = ["Outlet", "Invoice No.", "Payment Mode", "Platform"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, 3, len(headers))
    for i, (outlet, invoice, mode) in enumerate(all_rows):
        r = 4 + i
        ws.cell(row=r, column=1, value=outlet)
        ws.cell(row=r, column=2, value=invoice)
        ws.cell(row=r, column=3, value=mode)
        ws.cell(row=r, column=4, value=classify_platform(mode))
    autosize(ws, [28, 14, 22, 12])
    if not all_rows:
        ws["A5"] = "No Payment Wise Order Summary data received yet for any outlet."
        ws["A5"].font = BODY_FONT


def build_dashboard_sheet(wb, outlet_sheets, report_date):
    ws = wb.create_sheet("Dashboard", 0)
    ws["A1"] = "Petpooja Daily Discount Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Report date: {report_date}"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["Outlet", "Avg Disc % Dine-in", "Avg Disc % Swiggy", "Avg Disc % Zomato",
               ">50% Disc Swiggy", ">50% Disc Zomato", "100% Disc Dine-in (Staff)",
               "Dine-in >15%", "Dine-in >30%", "Dine-in >50%", "Orders Pending Platform"]
    header_row = 4
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    style_header(ws, header_row, len(headers))

    for i, sheet_name in enumerate(outlet_sheets):
        r = header_row + 1 + i
        q = f"'{sheet_name}'"
        ws.cell(row=r, column=1, value=sheet_name)
        ws.cell(row=r, column=2, value=f"={q}!B9")
        ws.cell(row=r, column=3, value=f"={q}!B10")
        ws.cell(row=r, column=4, value=f"={q}!B11")
        ws.cell(row=r, column=5, value=f"={q}!B12")
        ws.cell(row=r, column=6, value=f"={q}!B13")
        ws.cell(row=r, column=7, value=f"={q}!B14")
        ws.cell(row=r, column=8, value=f"={q}!B15")
        ws.cell(row=r, column=9, value=f"={q}!B16")
        ws.cell(row=r, column=10, value=f"={q}!B17")
        ws.cell(row=r, column=11, value=f"={q}!B8")
        for c in (2, 3, 4):
            ws.cell(row=r, column=c).number_format = "0.0%"
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    last_row = header_row + len(outlet_sheets)
    autosize(ws, [28, 16, 16, 16, 14, 14, 18, 12, 12, 12, 18])

    if outlet_sheets:
        chart = BarChart()
        chart.title = "Average Discount % by Platform"
        chart.y_axis.title = "Avg Discount %"
        chart.y_axis.numFmt = "0%"
        chart.x_axis.title = "Outlet"
        data = Reference(ws, min_col=2, max_col=4, min_row=header_row, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=header_row + 1, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 8
        chart.width = 18
        ws.add_chart(chart, f"A{last_row + 3}")

    notes_row = last_row + 22
    ws.cell(row=notes_row, column=1,
            value="Legend: red = 100% discount dine-in (staff order) / dine-in >50% tier · "
                  "orange = >50% discount on Swiggy or Zomato · "
                  "amber/yellow = dine-in >15% / >30% discount tiers.").font = SUBTITLE_FONT
    ws.cell(row=notes_row + 1, column=1,
            value='Orders show as "Pending (awaiting Payment Wise report)" until that report '
                  "is enabled in Petpooja's Notification tab — see PaymentMap sheet.").font = SUBTITLE_FONT


def main(manifest_path, output_path):
    with open(manifest_path) as f:
        manifest = json.load(f)

    report_date = manifest["report_date"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    outlet_sheets = []
    all_payment_rows = []

    for outlet in manifest["outlets"]:
        name = outlet["name"]
        item_rows = parse_item_wise(outlet["item_wise_xlsx"])
        order_agg = aggregate_orders(item_rows)

        payment_map = {}
        pw_path = outlet.get("payment_wise_xlsx")
        if pw_path:
            payment_map = parse_payment_wise(pw_path)
            for invoice, mode in payment_map.items():
                all_payment_rows.append((name, invoice, mode))

        sheet_name, _ = build_outlet_sheet(wb, name, report_date, order_agg, payment_map)
        outlet_sheets.append(sheet_name)

    build_payment_map_sheet(wb, all_payment_rows)
    build_dashboard_sheet(wb, outlet_sheets, report_date)

    wb.save(output_path)
    print(json.dumps({"output": output_path, "outlets": outlet_sheets}))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: build_dashboard.py <manifest.json> <output.xlsx>", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
