"""
GL Supplemental Analytics Workbook Builder

Builds a Finance Analytics workbook from a GL Supplemental report, providing
executive-friendly views of actual commission payments made to third-party
representatives.  Automates the monthly analysis that previously required
manual filtering and summarisation inside a large Excel workbook.

Usage:
    python build_gl_supplemental.py \\
        --gl-file GL_Supplemental.xlsx \\
        --fiscal-period P05 --fiscal-year 2026

Optional flags (see --help for the full list):
    --gl-sheet      Sheet name to read (default: first sheet).
    --output        Default: GL_Supplemental_Analytics_<Period>_<Year>.xlsx

Requires: openpyxl (see requirements.txt).
"""
import argparse
import sys
from collections import defaultdict, Counter
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef

CURRENCY = '$#,##0.00_);($#,##0.00)'
PCT = '0.0%'
NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
WHITE_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=16, color=NAVY)
LABEL_FONT = Font(bold=True, size=10)
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
KPI_LABEL_FONT = Font(size=10, color="404040")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
KPI_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

TAB_ORDER = [
    "Executive Summary",
    "Executive Dashboard",
    "Configuration",
    "Month-Year Summary",
    "Cost Center Summary",
    "Account Summary",
    "Cost Center + Account",
    "Filtered GL Detail",
    "Validation Report",
]

DATE_COLUMN_CANDIDATES = ["Accounting Period", "Accounting Date", "Date"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def month_year_key(dt):
    return (dt.year, dt.month)


def month_year_label(key):
    y, m = key
    return date(y, m, 1).strftime("%b-%Y")


def to_float(v):
    if v is None or v == '':
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def detect_data_sheet(wb, explicit_name):
    if explicit_name:
        return explicit_name
    candidates = [n for n in wb.sheetnames if n.strip().lower() != "screenshots"]
    if len(candidates) != 1:
        raise SystemExit(
            f"Could not auto-detect the data sheet (candidates: {candidates}). "
            f"Pass --gl-sheet explicitly."
        )
    return candidates[0]


def load_rows(path, sheetname):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheetname]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers) if h}
    rows = []
    total_rows = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        total_rows += 1
        if all(v is None or v == '' for v in r):
            continue
        rows.append(r)
    return idx, rows, total_rows


def style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def write_table(ws, start_row, headers, rows, currency_cols=(), pct_cols=(), start_col=1):
    for j, h in enumerate(headers):
        ws.cell(row=start_row, column=start_col + j, value=h)
    style_header(ws, start_row, len(headers), start_col)
    r = start_row + 1
    for row in rows:
        for j, val in enumerate(row):
            cell = ws.cell(row=r, column=start_col + j, value=val)
            cell.border = BORDER
            if j in currency_cols:
                cell.number_format = CURRENCY
            if j in pct_cols:
                cell.number_format = PCT
        r += 1
    return r


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def set_text_categories(chart, ref):
    for s in chart.ser:
        s.cat = AxDataSource(strRef=StrRef(f=str(ref)))


def truncate(label, n=20):
    label = str(label)
    return label if len(label) <= n else label[:n - 3] + "..."


# ---------------------------------------------------------------------------
# Data extraction
# ---------------------------------------------------------------------------

def detect_date_column(idx):
    for candidate in DATE_COLUMN_CANDIDATES:
        if candidate in idx:
            return candidate
    return None


def extract_gl(idx, rows, source_name):
    required = ["Cost Center", "Account", "Amount"]
    date_col = detect_date_column(idx)
    if date_col:
        required_check = required + [date_col]
    else:
        required_check = required

    missing_cols = [c for c in required_check if c not in idx]
    if missing_cols:
        raise SystemExit(f"GL Supplemental: missing required column(s): {missing_cols}")

    je_col = idx.get("JE Source")
    included = []
    missing_cc = missing_acct = blank_amt = invalid_amt = negative_amt = 0
    manual_excluded = 0
    je_sources = set()
    dup_counter = Counter()

    for r in rows:
        cc = r[idx["Cost Center"]]
        acct = r[idx["Account"]]
        amt_raw = r[idx["Amount"]]
        dt = to_date(r[idx[date_col]]) if date_col else None
        je_source = str(r[je_col]).strip() if je_col is not None and r[je_col] not in (None, '') else None
        amt = to_float(amt_raw)

        if cc in (None, ''):
            missing_cc += 1
        if acct in (None, ''):
            missing_acct += 1
        if amt_raw in (None, ''):
            blank_amt += 1
        elif amt is None:
            invalid_amt += 1
        elif amt < 0:
            negative_amt += 1

        dup_counter[r] += 1

        if je_source:
            je_sources.add(je_source)

        if je_source and je_source.lower() == "manual":
            manual_excluded += 1
            continue

        included.append({
            'cost_center': cc if cc not in (None, '') else 'Unassigned',
            'account': acct if acct not in (None, '') else 'Unassigned',
            'amount': amt or 0.0,
            'date': dt,
            'je_source': je_source or 'Unknown',
            'raw': r,
        })

    dup_count = sum(1 for v in dup_counter.values() if v > 1)

    validation = dict(
        source_name=source_name,
        imported_rows=len(rows),
        raw_file_rows=None,
        required_columns=required_check,
        date_column=date_col,
        missing_cost_center=missing_cc,
        missing_account=missing_acct,
        blank_amount=blank_amt,
        invalid_amount=invalid_amt,
        negative_amount=negative_amt,
        duplicate_transactions=dup_count,
        manual_excluded=manual_excluded,
        je_sources_found=sorted(je_sources),
    )
    return included, validation


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build(args):
    gl_wb = openpyxl.load_workbook(args.gl_file, data_only=True)
    gl_sheet = detect_data_sheet(gl_wb, args.gl_sheet)
    print(f"GL Supplemental: {args.gl_file} [{gl_sheet}]")

    idx, rows, raw_count = load_rows(args.gl_file, gl_sheet)
    included, val = extract_gl(idx, rows, args.gl_file.split('/')[-1])
    val['raw_file_rows'] = raw_count

    print(f"Fiscal Period: {args.fiscal_period} {args.fiscal_year}")
    print(f"Rows loaded: {val['imported_rows']}, Manual excluded: {val['manual_excluded']}, "
          f"Rows after filter: {len(included)}")

    # ---- KPIs ----
    TOTAL_PAID = sum(r['amount'] for r in included)
    NUM_TRANSACTIONS = len(included)
    cc_amounts = defaultdict(float)
    acct_amounts = defaultdict(float)
    for r in included:
        cc_amounts[r['cost_center']] += r['amount']
        acct_amounts[r['account']] += r['amount']
    NUM_COST_CENTERS = len(cc_amounts)
    NUM_ACCOUNTS = len(acct_amounts)

    # ---- Month-Year Summary ----
    my_amounts = defaultdict(float)
    for r in included:
        if r['date']:
            my_amounts[month_year_key(r['date'])] += r['amount']
    all_months = sorted(my_amounts.keys())
    month_year_summary = [{'label': month_year_label(k), 'amount': my_amounts[k]} for k in all_months]

    # ---- Cost Center Summary ----
    cost_center_summary = sorted(
        [{'cost_center': cc, 'amount': amt} for cc, amt in cc_amounts.items()],
        key=lambda x: x['amount'], reverse=True,
    )

    # ---- Account Summary ----
    account_summary = sorted(
        [{'account': acct, 'amount': amt} for acct, amt in acct_amounts.items()],
        key=lambda x: x['amount'], reverse=True,
    )

    # ---- Cost Center + Account Summary ----
    cca_amounts = defaultdict(float)
    for r in included:
        cca_amounts[(r['cost_center'], r['account'])] += r['amount']
    cc_account_summary = sorted(
        [{'cost_center': k[0], 'account': k[1], 'amount': amt} for k, amt in cca_amounts.items()],
        key=lambda x: x['amount'], reverse=True,
    )

    # ---- JE Source Distribution ----
    je_amounts = defaultdict(float)
    for r in included:
        je_amounts[r['je_source']] += r['amount']
    je_distribution = sorted(je_amounts.items(), key=lambda x: x[1], reverse=True)

    print(f"Total Amount Paid: {TOTAL_PAID:,.2f}  ({NUM_TRANSACTIONS} transactions)")
    print(f"Cost Centers: {NUM_COST_CENTERS}, Accounts: {NUM_ACCOUNTS}")

    top_cc = cost_center_summary[0] if cost_center_summary else {'cost_center': 'N/A', 'amount': 0}
    top_acct = account_summary[0] if account_summary else {'account': 'N/A', 'amount': 0}
    top_month = max(month_year_summary, key=lambda x: x['amount']) if month_year_summary else {'label': 'N/A', 'amount': 0}
    top10_cc_total = sum(c['amount'] for c in cost_center_summary[:10])
    top10_cc_pct = (top10_cc_total / TOTAL_PAID * 100) if TOTAL_PAID else 0

    commentary = (
        f"Total commission payments for {args.fiscal_period} {args.fiscal_year} amount to "
        f"${TOTAL_PAID:,.2f}, comprising {NUM_TRANSACTIONS} transactions across "
        f"{NUM_COST_CENTERS} cost centers and {NUM_ACCOUNTS} accounts. "
        f"The largest single month is {top_month['label']} (${top_month['amount']:,.2f}). "
        f"The largest Cost Center is {top_cc['cost_center']} (${top_cc['amount']:,.2f}), "
        f"and the largest Account is {top_acct['account']} (${top_acct['amount']:,.2f}). "
        f"{val['manual_excluded']} transactions with JE Source 'Manual' were excluded from analysis."
    )
    key_observations = (
        f"Top 10 Cost Centers represent ${top10_cc_total:,.2f} ({top10_cc_pct:.1f}% of total). "
        f"JE Source distribution includes: {', '.join(s for s, _ in je_distribution[:5])}. "
        f"Missing Cost Centers: {val['missing_cost_center']}, Missing Accounts: {val['missing_account']}, "
        f"Invalid Amounts: {val['invalid_amount']}, Negative Amounts: {val['negative_amount']}, "
        f"Duplicate Transactions: {val['duplicate_transactions']}. "
        f"See the Validation Report for full detail."
    )

    # ---- Assemble workbook ----
    wb = openpyxl.Workbook()

    # -- Executive Summary --
    ws1 = wb.active
    ws1.title = "Executive Summary"
    autosize(ws1, [28, 22, 60])
    ws1["A1"] = "GL Supplemental Analytics — Executive Summary"
    ws1["A1"].font = TITLE_FONT
    ws1.merge_cells("A1:C1")
    rows_es = [
        ("Reporting Period", f"{args.fiscal_period} {args.fiscal_year}"),
        ("Total Amount Paid", TOTAL_PAID),
        ("Number of Transactions", NUM_TRANSACTIONS),
        ("Number of Cost Centers", NUM_COST_CENTERS),
        ("Number of Accounts", NUM_ACCOUNTS),
    ]
    r = 3
    for label, value in rows_es:
        ws1.cell(row=r, column=1, value=label).font = LABEL_FONT
        cell = ws1.cell(row=r, column=2, value=value)
        if label == "Total Amount Paid":
            cell.number_format = CURRENCY
        r += 1
    ws1.cell(row=r + 1, column=1, value="Executive Commentary").font = LABEL_FONT
    ws1.cell(row=r + 1, column=2, value=commentary).alignment = WRAP
    ws1.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=3)
    ws1.row_dimensions[r + 1].height = 90
    ws1.cell(row=r + 3, column=1, value="Key Observations").font = LABEL_FONT
    ws1.cell(row=r + 3, column=2, value=key_observations).alignment = WRAP
    ws1.merge_cells(start_row=r + 3, start_column=2, end_row=r + 3, end_column=3)
    ws1.row_dimensions[r + 3].height = 90

    # -- Executive Dashboard --
    ws2 = wb.create_sheet("Executive Dashboard")
    for col in "ABCDEFGHIJK":
        ws2.column_dimensions[col].width = 14
    for col_idx in range(14, 30):
        ws2.column_dimensions[get_column_letter(col_idx)].hidden = True
    ws2["A1"] = f"Executive Dashboard — {args.fiscal_period} {args.fiscal_year}"
    ws2["A1"].font = TITLE_FONT
    ws2.merge_cells("A1:K1")

    kpis = [
        ("Total Amount Paid", TOTAL_PAID),
        ("Transactions", NUM_TRANSACTIONS),
        ("Cost Centers", NUM_COST_CENTERS),
        ("Accounts", NUM_ACCOUNTS),
    ]
    ws2.row_dimensions[2].height = 16
    ws2.row_dimensions[3].height = 28
    ws2.row_dimensions[4].height = 16
    for i, (label, value) in enumerate(kpis):
        col = 1 + i * 2
        ws2.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        c = ws2.cell(row=2, column=col, value=label)
        c.font, c.alignment, c.fill = KPI_LABEL_FONT, CENTER, KPI_FILL
        ws2.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col + 1)
        v = ws2.cell(row=3, column=col, value=value)
        v.font, v.alignment, v.fill = KPI_VALUE_FONT, CENTER, KPI_FILL
        if label == "Total Amount Paid":
            v.number_format = CURRENCY
        for rr in (2, 3, 4):
            for cc_ in (col, col + 1):
                ws2.cell(row=rr, column=cc_).border = BORDER

    HCOL = 14
    recent_months = month_year_summary[-12:]
    truncated_months = len(month_year_summary) > 12
    ws2.cell(row=1, column=HCOL, value="Month")
    ws2.cell(row=1, column=HCOL + 1, value="Amount Paid")
    for i, m in enumerate(recent_months, start=2):
        ws2.cell(row=i, column=HCOL, value=truncate(m['label']))
        ws2.cell(row=i, column=HCOL + 1, value=m['amount'])

    top10_cc = cost_center_summary[:10]
    ws2.cell(row=1, column=HCOL + 3, value="Cost Center")
    ws2.cell(row=1, column=HCOL + 4, value="Amount Paid")
    for i, c in enumerate(top10_cc, start=2):
        ws2.cell(row=i, column=HCOL + 3, value=truncate(c['cost_center']))
        ws2.cell(row=i, column=HCOL + 4, value=c['amount'])

    top10_acct = account_summary[:10]
    ws2.cell(row=1, column=HCOL + 6, value="Account")
    ws2.cell(row=1, column=HCOL + 7, value="Amount Paid")
    for i, a in enumerate(top10_acct, start=2):
        ws2.cell(row=i, column=HCOL + 6, value=truncate(a['account']))
        ws2.cell(row=i, column=HCOL + 7, value=a['amount'])

    ws2.cell(row=1, column=HCOL + 9, value="JE Source")
    ws2.cell(row=1, column=HCOL + 10, value="Amount Paid")
    for i, (je_src, je_amt) in enumerate(je_distribution, start=2):
        ws2.cell(row=i, column=HCOL + 9, value=truncate(je_src))
        ws2.cell(row=i, column=HCOL + 10, value=je_amt)

    # Charts
    month_chart = LineChart()
    month_chart.visible_cells_only = False
    month_chart.title = "Commission Payments by Month" + (" (last 12 months)" if truncated_months else "")
    month_chart.y_axis.title, month_chart.x_axis.title = "Amount Paid ($)", "Month-Year"
    month_chart.width, month_chart.height = 11, 8
    if recent_months:
        data_ref = Reference(ws2, min_col=HCOL + 1, min_row=1, max_row=1 + len(recent_months))
        cats_ref = Reference(ws2, min_col=HCOL, min_row=2, max_row=1 + len(recent_months))
        month_chart.add_data(data_ref, titles_from_data=True)
        month_chart.set_categories(cats_ref)
        set_text_categories(month_chart, cats_ref)
    month_chart.anchor = "A7"
    ws2.add_chart(month_chart)

    cc_chart = BarChart()
    cc_chart.visible_cells_only = False
    cc_chart.type = "bar"
    cc_chart.title = "Top 10 Cost Centers by Amount Paid"
    cc_chart.width, cc_chart.height = 11, 8
    if top10_cc:
        data_ref = Reference(ws2, min_col=HCOL + 4, min_row=1, max_row=1 + len(top10_cc))
        cats_ref = Reference(ws2, min_col=HCOL + 3, min_row=2, max_row=1 + len(top10_cc))
        cc_chart.add_data(data_ref, titles_from_data=True)
        cc_chart.set_categories(cats_ref)
        set_text_categories(cc_chart, cats_ref)
    cc_chart.anchor = "G7"
    ws2.add_chart(cc_chart)

    acct_chart = BarChart()
    acct_chart.visible_cells_only = False
    acct_chart.type = "bar"
    acct_chart.title = "Top 10 Accounts by Amount Paid"
    acct_chart.width, acct_chart.height = 11, 8
    if top10_acct:
        data_ref = Reference(ws2, min_col=HCOL + 7, min_row=1, max_row=1 + len(top10_acct))
        cats_ref = Reference(ws2, min_col=HCOL + 6, min_row=2, max_row=1 + len(top10_acct))
        acct_chart.add_data(data_ref, titles_from_data=True)
        acct_chart.set_categories(cats_ref)
        set_text_categories(acct_chart, cats_ref)
    acct_chart.anchor = "A25"
    ws2.add_chart(acct_chart)

    je_chart = PieChart()
    je_chart.visible_cells_only = False
    je_chart.title = "JE Source Distribution"
    je_chart.width, je_chart.height = 11, 8
    if je_distribution:
        data_ref = Reference(ws2, min_col=HCOL + 10, min_row=1, max_row=1 + len(je_distribution))
        cats_ref = Reference(ws2, min_col=HCOL + 9, min_row=2, max_row=1 + len(je_distribution))
        je_chart.add_data(data_ref, titles_from_data=True)
        je_chart.set_categories(cats_ref)
        set_text_categories(je_chart, cats_ref)
    je_chart.anchor = "G25"
    ws2.add_chart(je_chart)

    ws2.cell(row=43, column=1, value="Executive Commentary").font = LABEL_FONT
    ws2.merge_cells(start_row=44, start_column=1, end_row=52, end_column=10)
    c = ws2.cell(row=44, column=1, value=commentary + " " + key_observations)
    c.alignment, c.fill = WRAP, PatternFill("solid", fgColor="F2F2F2")

    ws2.cell(row=55, column=1, value="Data Quality Summary").font = LABEL_FONT
    dq_lines = (
        f"Required Columns: {', '.join(val['required_columns'])} — present in source workbook.\n"
        f"Missing Cost Centers: {val['missing_cost_center']}\n"
        f"Missing Accounts: {val['missing_account']}\n"
        f"Invalid Values: {val['invalid_amount']}\n"
        f"Negative Values: {val['negative_amount']}\n"
        f"Duplicate Transactions: {val['duplicate_transactions']}\n"
        f"Manual JE Source Excluded: {val['manual_excluded']}"
    )
    ws2.merge_cells(start_row=56, start_column=1, end_row=64, end_column=10)
    c = ws2.cell(row=56, column=1, value=dq_lines)
    c.alignment, c.fill = WRAP, PatternFill("solid", fgColor="F2F2F2")

    # -- Configuration --
    ws3 = wb.create_sheet("Configuration")
    autosize(ws3, [30, 50])
    ws3["A1"] = "Configuration"
    ws3["A1"].font = TITLE_FONT
    je_filter_info = (
        f"Excluded {val['manual_excluded']} 'Manual' entries. "
        f"Remaining JE Sources: {', '.join(val['je_sources_found'])}"
    )
    cfg_rows = [
        ("Reporting Period", f"{args.fiscal_period} {args.fiscal_year}"),
        ("Source Workbook Name", val['source_name']),
        ("Source Sheet Name", gl_sheet),
        ("Date Column Used", val['date_column'] or "Not detected"),
        ("Filters Applied", "JE Source = 'Manual' excluded"),
        ("JE Source Filter Info", je_filter_info),
        ("Total Rows in Source", val['raw_file_rows']),
        ("Rows After Filtering", len(included)),
    ]
    r = 3
    for label, value in cfg_rows:
        ws3.cell(row=r, column=1, value=label).font = LABEL_FONT
        ws3.cell(row=r, column=2, value=value).alignment = WRAP
        r += 1

    # -- Month-Year Summary --
    ws4 = wb.create_sheet("Month-Year Summary")
    autosize(ws4, [16, 18])
    ws4["A1"] = "Month-Year Summary"
    ws4["A1"].font = TITLE_FONT
    rows4 = [[m['label'], m['amount']] for m in month_year_summary]
    grand_total = sum(m['amount'] for m in month_year_summary)
    rows4.append(["Grand Total", grand_total])
    r = write_table(ws4, 3, ["Month-Year", "Amount Paid"], rows4, currency_cols=(1,))
    for c in range(1, 3):
        ws4.cell(row=r - 1, column=c).font = Font(bold=True)

    # -- Cost Center Summary --
    ws5 = wb.create_sheet("Cost Center Summary")
    autosize(ws5, [22, 18])
    ws5["A1"] = "Cost Center Summary"
    ws5["A1"].font = TITLE_FONT
    rows5 = [[c['cost_center'], c['amount']] for c in cost_center_summary]
    write_table(ws5, 3, ["Cost Center", "Amount Paid"], rows5, currency_cols=(1,))

    # -- Account Summary --
    ws6 = wb.create_sheet("Account Summary")
    autosize(ws6, [22, 18])
    ws6["A1"] = "Account Summary"
    ws6["A1"].font = TITLE_FONT
    rows6 = [[a['account'], a['amount']] for a in account_summary]
    write_table(ws6, 3, ["Account", "Amount Paid"], rows6, currency_cols=(1,))

    # -- Cost Center + Account Summary --
    ws7 = wb.create_sheet("Cost Center + Account")
    autosize(ws7, [22, 22, 18])
    ws7["A1"] = "Cost Center + Account Summary"
    ws7["A1"].font = TITLE_FONT
    rows7 = [[c['cost_center'], c['account'], c['amount']] for c in cc_account_summary]
    write_table(ws7, 3, ["Cost Center", "Account", "Amount Paid"], rows7, currency_cols=(2,))

    # -- Filtered GL Detail --
    ws8 = wb.create_sheet("Filtered GL Detail")
    headers_sorted = [None] * len(idx)
    for h, i in idx.items():
        headers_sorted[i] = h
    ws8.append(headers_sorted)
    style_header(ws8, 1, len(headers_sorted))
    for r_ in included:
        ws8.append(list(r_['raw']))
    for row in ws8.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = 'yyyy-mm-dd'

    # -- Validation Report --
    ws9 = wb.create_sheet("Validation Report")
    autosize(ws9, [35, 22])
    ws9["A1"] = "Validation Report"
    ws9["A1"].font = TITLE_FONT
    r = write_table(ws9, 3, ["Source Workbook", "Sheet", "Raw Rows", "Imported Rows"],
                    [[val['source_name'], gl_sheet, val['raw_file_rows'], val['imported_rows']]])
    r += 1
    ws9.cell(row=r, column=1, value="Required Columns").font = LABEL_FONT
    r += 1
    r = write_table(ws9, r, ["Required Column", "Present"],
                    [[col, "Yes"] for col in val['required_columns']])
    r += 1
    ws9.cell(row=r, column=1, value="Validation Metrics").font = LABEL_FONT
    r += 1
    write_table(ws9, r, ["Metric", "Count"],
                [
                    ["Missing Cost Centers", val['missing_cost_center']],
                    ["Missing Accounts", val['missing_account']],
                    ["Blank/Missing Amounts", val['blank_amount']],
                    ["Invalid (non-numeric) Amounts", val['invalid_amount']],
                    ["Negative Amounts", val['negative_amount']],
                    ["Duplicate Transactions", val['duplicate_transactions']],
                    ["Manual JE Source Excluded", val['manual_excluded']],
                ])

    # ---- Tab order & save ----
    assert set(TAB_ORDER) == set(wb.sheetnames), (set(TAB_ORDER) ^ set(wb.sheetnames))
    wb._sheets = [wb[name] for name in TAB_ORDER]
    wb.active = 0

    output = args.output or f"GL_Supplemental_Analytics_{args.fiscal_period}_{args.fiscal_year}.xlsx"
    wb.save(output)
    print(f"\nSaved: {output}")
    print(f"Sheets ({len(wb.sheetnames)}): {wb.sheetnames}")
    return output


def parse_args(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--gl-file", required=True, help="Path to the GL Supplemental workbook")
    p.add_argument("--gl-sheet", default=None, help="Sheet name (auto-detected if omitted)")
    p.add_argument("--fiscal-period", required=True, help="e.g. P05")
    p.add_argument("--fiscal-year", required=True, type=int)
    p.add_argument("--output", default=None, help="Output .xlsx path (default: GL_Supplemental_Analytics_<Period>_<Year>.xlsx)")
    return p.parse_args(argv)


if __name__ == "__main__":
    build(parse_args())
