#!/usr/bin/env python3
"""
Build the Petpooja Daily Discount Dashboard workbook from Item Wise Report
(with Bill No.) and Payment Wise Summary report exports.

Item Wise Report With Bill No. arrives as a real .xlsx (OOXML). Columns
observed:
    Date, Timestamp, Server Name, Table No., Invoice No., hsn_code,
    Category, Item, Variation, Price, Qty., Sub Total, Discount, Tax,
    Final Total
  -> one row per line item; grouped by Invoice No. to get per-order totals.

Payment Wise Summary arrives with an .xls extension but is actually an HTML
table (Excel's classic "save as HTML, name it .xls" export) — NOT a real
xlsx/xls binary. Columns observed:
    Invoice No., Date, Payment Type, Order Type, Status, Persons, Area,
    Assign To, Not Paid, Cash, Card, Due Payment, Other, Wallet, UPI, Online
  -> one row per bill. Platform is derived from Order Type + Area:
    - Order Type contains "dine"        -> Dine-in
    - Area contains "zomato"            -> Zomato
    - Area contains "swiggy"            -> Swiggy
    - otherwise (cash/card, no area)    -> Dine-in

Both report types share Petpooja's standard notification-email layout:
    Row 1: "Date:"             , "<start> to <end>"
    Row 2: "Name:"             , "<report name>"
    Row 3: "Restaurant Name:"  , "<outlet>"
    Row 4: (blank)
    Row 5: column headers
    Row 6+: data

File format (real xlsx vs HTML-as-.xls) is auto-detected from content, not
from the file extension, since Petpooja doesn't use the extension
consistently. Until the Payment Wise Summary report is enabled for an
outlet, pass payment_wise_xlsx=None and every order is marked
"Pending (awaiting Payment Wise report)" in the workbook, ready to resolve
automatically once real data is supplied (see PaymentMap sheet).

PhonePe "Merchant Settlement Report" (from reports@phonepe.com, subject
"TUSKINFOOD Settlement Report") is a single .zip (one .csv inside, not
password-protected) shared across ALL outlets, rows differentiated by a
`StoreName` column ("Andheri TUSKIN COFFEE", "TUSKIN COFFEE BANDRA ", etc.
-- matched to an outlet by keyword, see `_outlet_keyword`/`match_phonepe_total`).
`TransactionDate` matches the report_date (settlement itself lags a day,
arriving alongside the next day's Petpooja emails). It covers every
non-cash payment instrument taken at the counter (UPI scan-and-pay AND
card swipes via the PhonePe EDC machine). Each outlet sheet's Dine-in
Total Sales minus this PhonePe total is highlighted as the implied cash
balance. Pass "phonepe_settlement": null (or omit it) on a day the report
hasn't arrived yet -- the balance shows as pending instead.

Usage:
    python3 build_dashboard.py <manifest.json> <output.xlsx>

Manifest schema:
{
  "report_date": "2026-09-22",
  "phonepe_settlement": "/path/Merchant_Settlement_Report_....zip" | null,
  "outlets": [
    {"name": "Tuskin Coffee Andheri West",
     "item_wise_xlsx": "/path/Item_bill_report_....xlsx",
     "payment_wise_xlsx": "/path/payment_wise_summary_....xls" | null}
  ]
}
"""
import sys
import json
import re
import csv
import io
import zipfile
from collections import defaultdict
from html.parser import HTMLParser

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
FILL_HIGH_ONLINE = PatternFill("solid", fgColor="FDBA74")  # >52% Swiggy/Zomato
FILL_TIER1 = PatternFill("solid", fgColor="FEF3C7")       # dine-in >15%
FILL_TIER2 = PatternFill("solid", fgColor="FDE68A")       # dine-in >30%
FILL_TIER3 = PatternFill("solid", fgColor="FCA5A5")       # dine-in >50%
FILL_CASH_RECON = PatternFill("solid", fgColor="A7F3D0")  # dine-in cash balance vs PhonePe
FILL_CASH_MISMATCH = PatternFill("solid", fgColor="FCA5A5")  # balance < 0: PhonePe exceeds dine-in total

PENDING_LABEL = "Pending (awaiting Payment Wise report)"
PENDING_PHONEPE_LABEL = "Pending (no PhonePe settlement report)"

ONLINE_HIGH_DISCOUNT_PCT = 52  # Swiggy/Zomato "high discount" highlight threshold


class _HTMLTableParser(HTMLParser):
    """Minimal stdlib parser for Petpooja's "Excel HTML" .xls exports
    (flat <table>/<tr>/<td|th> markup, no nesting or colspans)."""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._cell = []

    def handle_endtag(self, tag):
        if tag == "tr" and self._row is not None:
            self.rows.append(tuple(self._row))
            self._row = None
        elif tag in ("td", "th") and self._cell is not None:
            text = "".join(self._cell).strip()
            self._row.append(text if text else None)
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _load_rows(path):
    """Return report data as a list of row-tuples, regardless of whether the
    file is a real xlsx (OOXML zip) or an HTML table saved with a
    misleading .xls/.xlsx extension (Petpooja does both)."""
    with open(path, "rb") as f:
        head = f.read(4)
    if head[:2] == b"PK":
        wb = openpyxl.load_workbook(path, data_only=True)
        return list(wb.active.iter_rows(values_only=True))

    with open(path, "r", encoding="utf-8", errors="replace") as f:
        html = f.read()
    parser = _HTMLTableParser()
    parser.feed(html)
    width = max((len(r) for r in parser.rows), default=0)
    return [r + (None,) * (width - len(r)) for r in parser.rows]


def _find_header_row(rows, max_scan=10):
    """Petpooja report exports start with 3 metadata rows, a blank row, then headers."""
    for i, row in enumerate(rows[:max_scan]):
        non_null = [v for v in row if v not in (None, "")]
        if len(non_null) >= 4 and all(isinstance(v, str) for v in non_null):
            return i
    raise ValueError("Could not locate header row")


def _num(value):
    if value in (None, ""):
        return 0.0
    return float(str(value).replace(",", ""))


def _clean_invoice(value):
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def parse_item_wise(path):
    """Return list of {invoice, sub_total, discount, final_total} per line item."""
    rows = _load_rows(path)
    header_row = _find_header_row(rows)
    headers = [str(h).strip() if h else "" for h in rows[header_row]]
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

    out = []
    for r in rows[header_row + 1:]:
        if r[c_invoice] in (None, ""):
            continue
        invoice = _clean_invoice(r[c_invoice])
        if not invoice.isdigit():
            continue  # skip the report's trailing "Total" row
        out.append({
            "invoice": invoice,
            "sub_total": _num(r[c_subtotal]),
            "discount": _num(r[c_discount]),
            "final_total": _num(r[c_final]),
        })
    return out


def classify_platform(order_type: str, area: str, payment_type: str) -> str:
    o = (order_type or "").lower()
    a = (area or "").lower()
    p = (payment_type or "").lower()
    if "dine" in o:
        return "Dine-in"
    if "zomato" in a or "zomato" in p:
        return "Zomato"
    if "swiggy" in a or "swiggy" in p:
        return "Swiggy"
    return "Dine-in"


def parse_payment_wise(path):
    """Return {invoice: {"mode": <display string>, "platform": <Dine-in/Swiggy/Zomato>}}."""
    rows = _load_rows(path)
    header_row = _find_header_row(rows)
    headers = [str(h).strip() if h else "" for h in rows[header_row]]
    idx = {h.lower(): i for i, h in enumerate(headers)}

    def find(*keywords):
        for h, i in idx.items():
            if all(k in h for k in keywords):
                return i
        return None

    c_invoice = find("invoice")
    if c_invoice is None:
        c_invoice = find("bill")
    c_order_type = find("order", "type")
    c_area = find("area")
    c_payment_type = find("payment", "type")
    if c_payment_type is None:
        c_payment_type = find("payment", "mode")
    if c_invoice is None:
        raise KeyError(f"Could not find an invoice/bill column in headers: {headers}")

    mapping = {}
    for r in rows[header_row + 1:]:
        if r[c_invoice] in (None, ""):
            continue
        invoice = _clean_invoice(r[c_invoice])
        if not invoice.isdigit():
            continue  # skip the report's trailing "Total" row

        order_type = str(r[c_order_type]).strip() if c_order_type is not None and r[c_order_type] else ""
        area = str(r[c_area]).strip() if c_area is not None and r[c_area] else ""
        payment_type = str(r[c_payment_type]).strip() if c_payment_type is not None and r[c_payment_type] else ""

        platform = classify_platform(order_type, area, payment_type)
        mode = payment_type or order_type or "Unknown"
        if area:
            mode = f"{mode} ({area})"
        mapping[invoice] = {"mode": mode, "platform": platform}
    return mapping


def aggregate_orders(item_rows):
    """Group line items by invoice -> per-order sub_total/discount/final_total."""
    agg = defaultdict(lambda: {"sub_total": 0.0, "discount": 0.0, "final_total": 0.0})
    for row in item_rows:
        a = agg[row["invoice"]]
        a["sub_total"] += row["sub_total"]
        a["discount"] += row["discount"]
        a["final_total"] += row["final_total"]
    return agg


def parse_phonepe_settlement(path):
    """Parse a PhonePe "Merchant Settlement Report" export into per-row
    digital-payment amounts.

    One report is emailed daily (from reports@phonepe.com, subject
    "TUSKINFOOD Settlement Report") covering ALL outlets in a single .zip
    attachment containing one .csv, rows differentiated by a `StoreName`
    column. `TransactionDate` matches the report_date (the report itself
    arrives the next morning, T+1 settlement, alongside Petpooja's emails).
    Rows cover every non-cash payment instrument taken at the counter/table
    (UPI scan-and-pay AND card swipes via the PhonePe EDC machine) -- i.e.
    everything that isn't cash.
    """
    if str(path).lower().endswith(".zip"):
        with zipfile.ZipFile(path) as zf:
            csv_name = next(n for n in zf.namelist() if n.lower().endswith(".csv"))
            with zf.open(csv_name) as f:
                text = f.read().decode("utf-8-sig")
    else:
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()

    rows = []
    for r in csv.DictReader(io.StringIO(text)):
        store = (r.get("StoreName") or "").strip()
        if not store:
            continue
        amount = _num(r.get("Amount"))
        ptype = (r.get("PaymentType") or "").strip().upper()
        if ptype == "PAYMENT":
            signed = amount
        elif "REFUND" in ptype:
            signed = -amount
        else:
            continue
        rows.append({"store_name": store, "amount": signed})
    return rows


def _outlet_keyword(outlet_name: str) -> str:
    """Reduce e.g. "Tuskin Coffee Andheri West" to "andheri" for matching
    against PhonePe's differently-formatted StoreName strings (e.g. "Andheri
    TUSKIN COFFEE", "TUSKIN COFFEE BANDRA ")."""
    name = re.sub(r"(?i)tuskin\s*coffee", "", outlet_name).strip()
    return (name.split()[0] if name else outlet_name).lower()


def match_phonepe_total(outlet_name, phonepe_rows):
    """Sum PhonePe settlement amounts for the rows matching this outlet."""
    keyword = _outlet_keyword(outlet_name)
    matched = [r for r in phonepe_rows if keyword in r["store_name"].lower()]
    return sum(r["amount"] for r in matched), len(matched)


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


def build_outlet_sheet(wb, outlet_name, report_date, order_agg, payment_map, phonepe_total=None):
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
        "Orders >52% Discount — Swiggy", "Orders >52% Discount — Zomato",
        "100% Discount Dine-in Orders (Staff)",
        "Dine-in Orders >15% Discount", "Dine-in Orders >30% Discount", "Dine-in Orders >50% Discount",
        "Dine-in Total Sales (Rs.)", "PhonePe UPI+Card Total (Rs.)", "Balance — Cash (Rs.)",
    ]
    for i, lbl in enumerate(labels):
        ws.cell(row=5 + i, column=1, value=lbl).font = LABEL_FONT

    TABLE_HEADER_ROW = 5 + len(labels) + 2
    headers = ["Invoice No.", "Sub Total", "Discount", "Discount %", "Payment Mode",
               "Platform", "Flag: 100% Dine-in (Staff)", f"Flag: >{ONLINE_HIGH_DISCOUNT_PCT}% Online Discount",
               "Flag: Dine-in Discount Tier", "Final Total"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=TABLE_HEADER_ROW, column=c, value=h)
    style_header(ws, TABLE_HEADER_ROW, len(headers))

    invoices = sorted(order_agg.keys())
    first_data_row = TABLE_HEADER_ROW + 1
    for i, invoice in enumerate(invoices):
        r = first_data_row + i
        a = order_agg[invoice]
        entry = payment_map.get(invoice)
        mode = entry["mode"] if entry else ""
        platform = entry["platform"] if entry else PENDING_LABEL

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
                value=(f'=IF(AND(OR(F{r}="Swiggy",F{r}="Zomato"),'
                       f'D{r}>{ONLINE_HIGH_DISCOUNT_PCT / 100}),"HIGH DISCOUNT","")'))
        ws.cell(row=r, column=9,
                value=(f'=IF(F{r}<>"Dine-in","",'
                       f'IF(D{r}>0.5,">50%",IF(D{r}>0.3,">30%",IF(D{r}>0.15,">15%",""))))'))
        ws.cell(row=r, column=10, value=round(a["final_total"], 2)).number_format = "#,##0.00"
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
        ws["B12"] = f'=COUNTIFS({rng("F")},"Swiggy",{rng("D")},">{ONLINE_HIGH_DISCOUNT_PCT / 100}")'
        ws["B13"] = f'=COUNTIFS({rng("F")},"Zomato",{rng("D")},">{ONLINE_HIGH_DISCOUNT_PCT / 100}")'
        ws["B14"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">=0.999")'
        ws["B15"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.15")'
        ws["B16"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.3")'
        ws["B17"] = f'=COUNTIFS({rng("F")},"Dine-in",{rng("D")},">0.5")'
        ws["B18"] = f'=SUMIFS({rng("J")},{rng("F")},"Dine-in")'
    else:
        for r in range(5, 19):
            ws.cell(row=r, column=2, value=0)

    for r in (9, 10, 11):
        ws.cell(row=r, column=2).number_format = "0.0%"
    ws["B18"].number_format = "#,##0.00"

    # ---- PhonePe reconciliation: Dine-in total minus PhonePe (UPI+card)
    # settlement leaves the balance that should be cash. Flag it. ----
    if phonepe_total is None:
        ws["B19"] = PENDING_PHONEPE_LABEL
        ws["B20"] = ""
    else:
        ws["B19"] = round(phonepe_total, 2)
        ws["B19"].number_format = "#,##0.00"
        ws["B20"] = "=B18-B19"
        ws["B20"].number_format = "#,##0.00"
        for r in (18, 19):
            for c in (1, 2):
                ws.cell(row=r, column=c).fill = FILL_CASH_RECON
        # Balance should never be negative (PhonePe can't collect more than
        # was billed dine-in) -- red flags that as a real mismatch to
        # investigate rather than a normal cash figure.
        ws.conditional_formatting.add(
            "A20:B20",
            FormulaRule(formula=["AND(ISNUMBER($B20),$B20<0)"], fill=FILL_CASH_MISMATCH))
        ws.conditional_formatting.add(
            "A20:B20",
            FormulaRule(formula=["AND(ISNUMBER($B20),$B20>=0)"], fill=FILL_CASH_RECON))

    # Conditional formatting highlights on the order table
    if invoices:
        data_range = f"A{first_data_row}:J{last_data_row}"
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
    autosize(ws, [14, 12, 12, 12, 24, 12, 22, 22, 20, 14])
    return sheet_name, (first_data_row, last_data_row if invoices else None)


def build_payment_map_sheet(wb, all_rows):
    ws = wb.create_sheet("PaymentMap")
    ws["A1"] = "Raw Payment Wise Summary rows (used to resolve Platform on each outlet sheet)"
    ws["A1"].font = SUBTITLE_FONT
    headers = ["Outlet", "Invoice No.", "Payment Mode", "Platform"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, 3, len(headers))
    for i, (outlet, invoice, mode, platform) in enumerate(all_rows):
        r = 4 + i
        ws.cell(row=r, column=1, value=outlet)
        ws.cell(row=r, column=2, value=invoice)
        ws.cell(row=r, column=3, value=mode)
        ws.cell(row=r, column=4, value=platform)
    autosize(ws, [28, 14, 22, 12])
    if not all_rows:
        ws["A5"] = "No Payment Wise Summary data received yet for any outlet."
        ws["A5"].font = BODY_FONT


def build_dashboard_sheet(wb, outlet_sheets, report_date):
    ws = wb.create_sheet("Dashboard", 0)
    ws["A1"] = "Petpooja Daily Discount Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Report date: {report_date}"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["Outlet", "Avg Disc % Dine-in", "Avg Disc % Swiggy", "Avg Disc % Zomato",
               f">{ONLINE_HIGH_DISCOUNT_PCT}% Disc Swiggy", f">{ONLINE_HIGH_DISCOUNT_PCT}% Disc Zomato",
               "100% Disc Dine-in (Staff)",
               "Dine-in >15%", "Dine-in >30%", "Dine-in >50%", "Orders Pending Platform",
               "Dine-in Total Sales", "PhonePe UPI+Card Total", "Balance — Cash"]
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
        ws.cell(row=r, column=12, value=f"={q}!B18")
        ws.cell(row=r, column=13, value=f"=IF(ISNUMBER({q}!B19),{q}!B19,\"{PENDING_PHONEPE_LABEL}\")")
        ws.cell(row=r, column=14, value=f"=IF(ISNUMBER({q}!B19),{q}!B20,\"\")")
        for c in (2, 3, 4):
            ws.cell(row=r, column=c).number_format = "0.0%"
        for c in (12, 13, 14):
            ws.cell(row=r, column=c).number_format = "#,##0.00"
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    last_row = header_row + len(outlet_sheets)
    if outlet_sheets:
        balance_range = f"N{header_row + 1}:N{last_row}"
        ws.conditional_formatting.add(
            balance_range,
            FormulaRule(formula=["AND(ISNUMBER($N5),$N5<0)"], fill=FILL_CASH_MISMATCH))
        ws.conditional_formatting.add(
            balance_range,
            FormulaRule(formula=["AND(ISNUMBER($N5),$N5>=0)"], fill=FILL_CASH_RECON))
    autosize(ws, [28, 16, 16, 16, 14, 14, 18, 12, 12, 12, 18, 18, 20, 16])

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
            value=f"Legend: red = 100% discount dine-in (staff order) / dine-in >50% tier / "
                  f"Balance — Cash is negative (PhonePe exceeds dine-in total — investigate) · "
                  f"orange = >{ONLINE_HIGH_DISCOUNT_PCT}% discount on Swiggy or Zomato · "
                  f"amber/yellow = dine-in >15% / >30% discount tiers · "
                  f"green = implied cash balance (Dine-in total minus PhonePe UPI+card).").font = SUBTITLE_FONT
    ws.cell(row=notes_row + 1, column=1,
            value='Orders show as "Pending (awaiting Payment Wise report)" until that report '
                  "is enabled in Petpooja's Notification tab — see PaymentMap sheet. Balance — Cash "
                  "shows \"" + PENDING_PHONEPE_LABEL + "\" until that day's PhonePe settlement email "
                  "arrives.").font = SUBTITLE_FONT


def write_csv_snapshot(path, report_date, outlets_data):
    """Plain-text CSV mirror of the dashboard, one section per outlet.

    This exists because MCP tool attachments (Gmail) and binary Drive uploads
    (base64Content) require the agent to transcribe the file's raw bytes as a
    literal string in a tool call — verbatim reproduction of an opaque,
    high-entropy blob like that is not reliable at any practical size (this
    was verified the hard way: even small ~8KB chunks silently lost/altered
    characters on manual transcription). Plain text has none of that risk, so
    for delivery (as opposed to the full xlsx, which a human can download
    with SendUserFile and doesn't need retyping), always publish this CSV via
    Google Drive's `textContent` field instead of base64-encoding the xlsx.
    """
    import csv as _csv

    with open(path, "w", newline="") as f:
        w = _csv.writer(f)
        w.writerow(["Petpooja Daily Discount Dashboard"])
        w.writerow(["Report date", report_date])
        w.writerow([])

        for outlet in outlets_data:
            name = outlet["name"]
            order_agg = outlet["order_agg"]
            payment_map = outlet["payment_map"]
            phonepe_total = outlet.get("phonepe_total")

            totals = defaultdict(lambda: [0.0, 0.0, 0])
            for invoice, a in order_agg.items():
                entry = payment_map.get(invoice)
                platform = entry["platform"] if entry else PENDING_LABEL
                t = totals[platform]
                t[0] += a["sub_total"]
                t[1] += a["discount"]
                t[2] += 1

            def pct(invoice_data):
                sub, disc = invoice_data["sub_total"], invoice_data["discount"]
                return (disc / sub * 100) if sub else 0.0

            staff = sorted(
                (inv for inv, a in order_agg.items()
                 if a["sub_total"] > 0 and pct(a) >= 99.9
                 and payment_map.get(inv, {}).get("platform") == "Dine-in"),
                key=int)
            high_online = sorted(
                (inv for inv, a in order_agg.items()
                 if a["sub_total"] > 0 and pct(a) > ONLINE_HIGH_DISCOUNT_PCT
                 and payment_map.get(inv, {}).get("platform") in ("Swiggy", "Zomato")),
                key=int)
            dine15 = [inv for inv, a in order_agg.items()
                      if a["sub_total"] > 0 and pct(a) > 15
                      and payment_map.get(inv, {}).get("platform") == "Dine-in"]
            dine30 = [inv for inv, a in order_agg.items()
                      if a["sub_total"] > 0 and pct(a) > 30
                      and payment_map.get(inv, {}).get("platform") == "Dine-in"]
            dine50 = [inv for inv, a in order_agg.items()
                      if a["sub_total"] > 0 and pct(a) > 50
                      and payment_map.get(inv, {}).get("platform") == "Dine-in"]

            w.writerow(["Outlet", name])
            w.writerow([])
            w.writerow(["Platform", "Orders", "Avg Discount %"])
            for platform in ("Dine-in", "Swiggy", "Zomato", PENDING_LABEL):
                sub, disc, n = totals.get(platform, (0.0, 0.0, 0))
                if n == 0 and platform == PENDING_LABEL:
                    continue
                avg = round(disc / sub * 100, 1) if sub else 0
                w.writerow([platform, n, avg])
            w.writerow([])
            w.writerow(["Flag", "Count", "Invoices"])
            w.writerow(["100% Discount Dine-in (Staff)", len(staff), " ".join(staff)])
            w.writerow([f">{ONLINE_HIGH_DISCOUNT_PCT}% Discount Swiggy/Zomato", len(high_online), " ".join(high_online)])
            w.writerow(["Dine-in >15% Discount", len(dine15), ""])
            w.writerow(["Dine-in >30% Discount", len(dine30), ""])
            w.writerow(["Dine-in >50% Discount", len(dine50), ""])
            w.writerow([])

            dine_total = sum(a["final_total"] for inv, a in order_agg.items()
                              if payment_map.get(inv, {}).get("platform") == "Dine-in")
            w.writerow(["Payment Reconciliation (Dine-in)"])
            w.writerow(["Dine-in Total Sales", round(dine_total, 2)])
            if phonepe_total is None:
                w.writerow(["PhonePe UPI+Card Total", PENDING_PHONEPE_LABEL])
                w.writerow(["Balance — Cash", ""])
            else:
                w.writerow(["PhonePe UPI+Card Total", round(phonepe_total, 2)])
                w.writerow(["Balance — Cash", round(dine_total - phonepe_total, 2)])
            w.writerow([])

            w.writerow(["Invoice No.", "Sub Total", "Discount", "Discount %", "Payment Mode", "Platform", "Final Total"])
            for inv in sorted(order_agg.keys(), key=int):
                a = order_agg[inv]
                entry = payment_map.get(inv, {})
                w.writerow([inv, round(a["sub_total"], 2), round(a["discount"], 2),
                            round(pct(a), 1), entry.get("mode", ""), entry.get("platform", PENDING_LABEL),
                            round(a["final_total"], 2)])
            w.writerow([])


def main(manifest_path, output_path):
    with open(manifest_path) as f:
        manifest = json.load(f)

    report_date = manifest["report_date"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    outlet_sheets = []
    all_payment_rows = []
    outlets_data = []

    phonepe_rows = []
    phonepe_path = manifest.get("phonepe_settlement")
    if phonepe_path:
        phonepe_rows = parse_phonepe_settlement(phonepe_path)

    for outlet in manifest["outlets"]:
        name = outlet["name"]
        item_rows = parse_item_wise(outlet["item_wise_xlsx"])
        order_agg = aggregate_orders(item_rows)

        payment_map = {}
        pw_path = outlet.get("payment_wise_xlsx")
        if pw_path:
            payment_map = parse_payment_wise(pw_path)
            for invoice, entry in payment_map.items():
                all_payment_rows.append((name, invoice, entry["mode"], entry["platform"]))

        phonepe_total = None
        if phonepe_rows:
            phonepe_total, _ = match_phonepe_total(name, phonepe_rows)

        sheet_name, _ = build_outlet_sheet(wb, name, report_date, order_agg, payment_map, phonepe_total)
        outlet_sheets.append(sheet_name)
        outlets_data.append({"name": name, "order_agg": order_agg, "payment_map": payment_map,
                              "phonepe_total": phonepe_total})

    build_payment_map_sheet(wb, all_payment_rows)
    build_dashboard_sheet(wb, outlet_sheets, report_date)

    wb.save(output_path)

    csv_path = re.sub(r"\.xlsx$", "", output_path) + ".csv"
    write_csv_snapshot(csv_path, report_date, outlets_data)

    print(json.dumps({"output": output_path, "csv": csv_path, "outlets": outlet_sheets}))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: build_dashboard.py <manifest.json> <output.xlsx>", file=sys.stderr)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
