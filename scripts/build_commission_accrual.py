"""
Commission Accrual Workbook Builder

Builds the audit-ready Commission Accrual Workbook (Executive Summary,
Configuration, Month-Year/Cost Center/Account/Cost Center+Account/Aging
summaries, raw detail sheets, Validation Report, Audit Evidence, and a
fixed-anchor Executive Dashboard) from a "Payment Made" and a "Payment
Remaining" source workbook.

Usage:
    python build_commission_accrual.py \\
        --made-file Comm_Payment_Report.xlsx \\
        --remaining-file InvoiceOrder_Line_Detail.xlsx \\
        --fiscal-period P05 --fiscal-year 2026

Optional flags (see --help for the full list):
    --made-sheet / --remaining-sheet   Sheet name to read (default: first
                                        sheet not literally named "Screenshots").
    --fiscal-divisor                   Override the P01-P12 lookup table.
    --remaining-business-days          Default 3.
    --aging-asof                       YYYY-MM-DD. Default: the source file's
                                        own "Selected Date" column if present,
                                        else today.
    --output                           Default: Commissions_Accrual_Output_<Period>_<Year>.xlsx

Requires: openpyxl, Pillow (see requirements.txt).
"""
import argparse
import io
import sys
import zipfile
from collections import defaultdict, Counter
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.drawing.image import Image as XLImage

FISCAL_CALENDAR = {
    "P01": 17, "P02": 17, "P03": 22, "P04": 17, "P05": 17, "P06": 22,
    "P07": 17, "P08": 17, "P09": 22, "P10": 17, "P11": 17, "P12": 22,
}
AGING_BUCKETS = ["Current (0-30 Days)", "31-60 Days", "61-90 Days",
                  "91-180 Days", "181-365 Days", "Over 365 Days"]

# Stakeholder-confirmed tab order. Edit this list if the order needs to
# change again in the future -- everything else derives from it.
TAB_ORDER = [
    "Executive Summary",
    "Executive Dashboard",
    "Cost Center Summary",
    "Payment Made Detail",
    "Payment Remaining Detail",
    "Audit Evidence",
    "Validation Report",
    "Configuration",
    "Month-Year Summary",
    "Cost Center + Account",
    "Account Summary",
    "Aging Summary",
]

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


# ---------------------------------------------------------------------------
# Data extraction & aggregation
# ---------------------------------------------------------------------------

def bucket_for_age(days):
    if days <= 30:
        return AGING_BUCKETS[0]
    if days <= 60:
        return AGING_BUCKETS[1]
    if days <= 90:
        return AGING_BUCKETS[2]
    if days <= 180:
        return AGING_BUCKETS[3]
    if days <= 365:
        return AGING_BUCKETS[4]
    return AGING_BUCKETS[5]


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
            f"Pass --made-sheet/--remaining-sheet explicitly."
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


def extract(idx, rows, label):
    required = ["Cost Center", "Bill To Account Number", "Commission Amount SUM", "Invoice Date"]
    missing_cols = [c for c in required if c not in idx]
    if missing_cols:
        raise SystemExit(f"{label}: missing required column(s): {missing_cols}")

    out = []
    missing_cc = missing_acct = blank_comm = invalid_comm = negative_comm = 0
    dup_counter = Counter()
    inv_i = idx.get('Invoice Number', idx.get('Invoice'))
    line_i = idx.get('Line Number')
    for r in rows:
        cc = r[idx['Cost Center']]
        acct = r[idx['Bill To Account Number']]
        comm_raw = r[idx['Commission Amount SUM']]
        inv_date = to_date(r[idx['Invoice Date']])
        comm = to_float(comm_raw)

        if cc in (None, ''):
            missing_cc += 1
        if acct in (None, ''):
            missing_acct += 1
        if comm_raw in (None, ''):
            blank_comm += 1
        elif comm is None:
            invalid_comm += 1
        elif comm < 0:
            negative_comm += 1

        if inv_i is not None and line_i is not None:
            dup_counter[(r[inv_i], r[line_i])] += 1

        out.append({
            'cost_center': cc if cc not in (None, '') else 'Unassigned',
            'account': acct if acct not in (None, '') else 'Unassigned',
            'commission': comm or 0.0,
            'invoice_date': inv_date,
            'raw': r,
        })

    dup_count = sum(1 for k, v in dup_counter.items() if v > 1 and k[0] is not None)
    validation = dict(
        label=label,
        imported_rows=len(rows),
        raw_file_rows=None,  # filled by caller
        missing_cost_center=missing_cc,
        missing_account=missing_acct,
        blank_commission=blank_comm,
        invalid_commission=invalid_comm,
        negative_commission=negative_comm,
        duplicate_transactions=dup_count,
    )
    return out, validation


def detect_asof_date(idx, rows):
    """If the source file has a 'Selected Date' column, use its value; else None."""
    col = idx.get('Selected Date')
    if col is None:
        return None
    for r in rows:
        v = r[col]
        if v in (None, ''):
            continue
        if isinstance(v, (datetime, date)):
            return to_date(v)
        try:
            return datetime.strptime(str(v).strip(), "%d-%b-%Y").date()
        except ValueError:
            continue
    return None


def extract_screenshot_images(path):
    """Read embedded images from a workbook's zip directly into memory."""
    images = []
    z = zipfile.ZipFile(path)
    for name in sorted(n for n in z.namelist() if n.startswith('xl/media/')):
        images.append((name.split('/')[-1], z.read(name)))
    return images


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

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
    # openpyxl's set_categories() always writes numRef, even for text labels;
    # force strRef so Excel renders the actual labels instead of blank/numeric axis ticks.
    for s in chart.ser:
        s.cat = AxDataSource(strRef=StrRef(f=str(ref)))


def truncate(label, n=20):
    label = str(label)
    return label if len(label) <= n else label[:n - 3] + "..."


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build(args):
    made_wb = openpyxl.load_workbook(args.made_file, data_only=True)
    made_sheet = detect_data_sheet(made_wb, args.made_sheet)
    remain_wb = openpyxl.load_workbook(args.remaining_file, data_only=True)
    remain_sheet = detect_data_sheet(remain_wb, args.remaining_sheet)

    print(f"Payment Made:      {args.made_file} [{made_sheet}]")
    print(f"Payment Remaining: {args.remaining_file} [{remain_sheet}]")

    made_idx, made_rows, made_raw_count = load_rows(args.made_file, made_sheet)
    remain_idx, remain_rows, remain_raw_count = load_rows(args.remaining_file, remain_sheet)

    made, made_val = extract(made_idx, made_rows, "Payment Made")
    made_val['raw_file_rows'] = made_raw_count
    remaining, remain_val = extract(remain_idx, remain_rows, "Payment Remaining")
    remain_val['raw_file_rows'] = remain_raw_count

    fiscal_divisor = args.fiscal_divisor or FISCAL_CALENDAR.get(args.fiscal_period)
    if fiscal_divisor is None:
        raise SystemExit(
            f"Unknown fiscal period '{args.fiscal_period}' and no --fiscal-divisor override given."
        )

    if args.aging_asof:
        aging_asof = datetime.strptime(args.aging_asof, "%Y-%m-%d").date()
    else:
        aging_asof = detect_asof_date(remain_idx, remain_rows) or date.today()
    print(f"Fiscal Period: {args.fiscal_period} {args.fiscal_year} (divisor {fiscal_divisor}), "
          f"Remaining Business Days: {args.remaining_business_days}, Aging as-of: {aging_asof}")

    PAYMENT_MADE_TOTAL = sum(r['commission'] for r in made)
    PAYMENT_REMAINING_TOTAL = sum(r['commission'] for r in remaining)
    ESTIMATED_COMMISSION = (PAYMENT_MADE_TOTAL / fiscal_divisor) * args.remaining_business_days
    FINAL_ACCRUAL = PAYMENT_MADE_TOTAL + PAYMENT_REMAINING_TOTAL + ESTIMATED_COMMISSION

    def allocate_estimate(dimension_made_total):
        if PAYMENT_MADE_TOTAL == 0:
            return 0.0
        return (dimension_made_total / PAYMENT_MADE_TOTAL) * ESTIMATED_COMMISSION

    # ---- Month-Year Summary ----
    my_made, my_remain = defaultdict(float), defaultdict(float)
    for r in made:
        if r['invoice_date']:
            my_made[month_year_key(r['invoice_date'])] += r['commission']
    for r in remaining:
        if r['invoice_date']:
            my_remain[month_year_key(r['invoice_date'])] += r['commission']
    all_months = sorted(set(my_made) | set(my_remain))
    month_year_summary = []
    for k in all_months:
        made_amt, rem_amt = my_made.get(k, 0.0), my_remain.get(k, 0.0)
        est_amt = allocate_estimate(made_amt)
        month_year_summary.append({'label': month_year_label(k), 'made': made_amt,
                                     'remaining': rem_amt, 'estimated': est_amt,
                                     'total': made_amt + rem_amt + est_amt})

    # ---- Cost Center Summary ----
    cc_made, cc_remain = defaultdict(float), defaultdict(float)
    for r in made:
        cc_made[r['cost_center']] += r['commission']
    for r in remaining:
        cc_remain[r['cost_center']] += r['commission']
    cost_center_summary = []
    for cc in set(cc_made) | set(cc_remain):
        made_amt, rem_amt = cc_made.get(cc, 0.0), cc_remain.get(cc, 0.0)
        est_amt = allocate_estimate(made_amt)
        cost_center_summary.append({'cost_center': cc, 'made': made_amt, 'remaining': rem_amt,
                                      'estimated': est_amt, 'total': made_amt + rem_amt + est_amt})
    cost_center_summary.sort(key=lambda x: x['total'], reverse=True)

    # ---- Cost Center + Account Summary ----
    cca_made, cca_remain = defaultdict(float), defaultdict(float)
    for r in made:
        cca_made[(r['cost_center'], r['account'])] += r['commission']
    for r in remaining:
        cca_remain[(r['cost_center'], r['account'])] += r['commission']
    cc_account_summary = []
    for key in set(cca_made) | set(cca_remain):
        made_amt, rem_amt = cca_made.get(key, 0.0), cca_remain.get(key, 0.0)
        est_amt = allocate_estimate(made_amt)
        cc_account_summary.append({'cost_center': key[0], 'account': key[1], 'made': made_amt,
                                     'remaining': rem_amt, 'estimated': est_amt,
                                     'total': made_amt + rem_amt + est_amt})
    cc_account_summary.sort(key=lambda x: x['total'], reverse=True)

    # ---- Account Summary ----
    acct_made, acct_remain = defaultdict(float), defaultdict(float)
    for r in made:
        acct_made[r['account']] += r['commission']
    for r in remaining:
        acct_remain[r['account']] += r['commission']
    account_summary = []
    for acct in set(acct_made) | set(acct_remain):
        made_amt, rem_amt = acct_made.get(acct, 0.0), acct_remain.get(acct, 0.0)
        est_amt = allocate_estimate(made_amt)
        account_summary.append({'account': acct, 'made': made_amt, 'remaining': rem_amt,
                                  'estimated': est_amt, 'total': made_amt + rem_amt + est_amt})
    account_summary.sort(key=lambda x: x['total'], reverse=True)

    # ---- Aging Summary ----
    aging_amounts, aging_counts = defaultdict(float), defaultdict(int)
    for r in remaining:
        if r['invoice_date']:
            b = bucket_for_age((aging_asof - r['invoice_date']).days)
            aging_amounts[b] += r['commission']
            aging_counts[b] += 1
    aging_summary = []
    for b in AGING_BUCKETS:
        amt = aging_amounts.get(b, 0.0)
        pct = (amt / PAYMENT_REMAINING_TOTAL) if PAYMENT_REMAINING_TOTAL else 0.0
        aging_summary.append({'bucket': b, 'amount': amt, 'pct': pct, 'count': aging_counts.get(b, 0)})

    print(f"Payment Made Total:      {PAYMENT_MADE_TOTAL:,.2f}  ({len(made)} rows)")
    print(f"Payment Remaining Total: {PAYMENT_REMAINING_TOTAL:,.2f}  ({len(remaining)} rows)")
    print(f"Estimated Commission:    {ESTIMATED_COMMISSION:,.2f}")
    print(f"Final Accrual:           {FINAL_ACCRUAL:,.2f}")

    top_cc = cost_center_summary[0]
    top_acct = account_summary[0]
    top_month = max(month_year_summary, key=lambda x: x['total'])
    over365 = next(a for a in aging_summary if a['bucket'] == "Over 365 Days")

    commentary = (
        f"Total Final Accrual for {args.fiscal_period} {args.fiscal_year} is ${FINAL_ACCRUAL:,.2f}, comprised of "
        f"${PAYMENT_MADE_TOTAL:,.2f} in payments made, ${PAYMENT_REMAINING_TOTAL:,.2f} in payments remaining, "
        f"and ${ESTIMATED_COMMISSION:,.2f} of estimated commission for the {args.remaining_business_days} remaining "
        f"business days (Payment Made Total / {fiscal_divisor} x {args.remaining_business_days}). "
        f"The largest single month by total accrual is {top_month['label']} (${top_month['total']:,.2f}). "
        f"The largest Cost Center is {top_cc['cost_center']} (${top_cc['total']:,.2f} total accrual), and the "
        f"largest Bill To Account Number is {top_acct['account']} (${top_acct['total']:,.2f})."
    )
    key_risks = (
        f"${over365['amount']:,.2f} of Payment Remaining ({over365['pct']*100:.1f}% of remaining total, "
        f"{over365['count']} transactions) is aged Over 365 Days as of {aging_asof.strftime('%d-%b-%Y')}. "
        f"{made_val['duplicate_transactions']} potential duplicate (Invoice Number + Line Number) records were "
        f"found in Payment Made and {remain_val['duplicate_transactions']} in Payment Remaining; "
        f"{remain_val['negative_commission']} negative Commission Amount SUM values were found in Payment Remaining. "
        f"See the Validation Report for full detail."
    )

    # ---- Assemble workbook ----
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Executive Summary"
    autosize(ws1, [28, 22, 60])
    ws1["A1"] = "Commission Accrual — Executive Summary"
    ws1["A1"].font = TITLE_FONT
    ws1.merge_cells("A1:C1")
    rows_es = [
        ("Fiscal Period", f"{args.fiscal_period} {args.fiscal_year}"),
        ("Fiscal Divisor", fiscal_divisor),
        ("Remaining Business Days", args.remaining_business_days),
        ("Payment Made Total", PAYMENT_MADE_TOTAL),
        ("Payment Remaining Total", PAYMENT_REMAINING_TOTAL),
        ("Estimated Commission", ESTIMATED_COMMISSION),
        ("Final Accrual", FINAL_ACCRUAL),
    ]
    r = 3
    for label, val in rows_es:
        ws1.cell(row=r, column=1, value=label).font = LABEL_FONT
        cell = ws1.cell(row=r, column=2, value=val)
        if label in ("Payment Made Total", "Payment Remaining Total", "Estimated Commission", "Final Accrual"):
            cell.number_format = CURRENCY
        r += 1
    ws1.cell(row=r + 1, column=1, value="Executive Commentary").font = LABEL_FONT
    ws1.cell(row=r + 1, column=2, value=commentary).alignment = WRAP
    ws1.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=3)
    ws1.row_dimensions[r + 1].height = 90
    ws1.cell(row=r + 3, column=1, value="Key Risks").font = LABEL_FONT
    ws1.cell(row=r + 3, column=2, value=key_risks).alignment = WRAP
    ws1.merge_cells(start_row=r + 3, start_column=2, end_row=r + 3, end_column=3)
    ws1.row_dimensions[r + 3].height = 90

    ws2 = wb.create_sheet("Configuration")
    autosize(ws2, [30, 45])
    ws2["A1"] = "Configuration"
    ws2["A1"].font = TITLE_FONT
    cfg_rows = [
        ("Fiscal Period", f"{args.fiscal_period} {args.fiscal_year}"),
        ("Fiscal Divisor", fiscal_divisor),
        ("Remaining Business Days", args.remaining_business_days),
        ("Estimation Formula", "(Payment Made Total / Fiscal Divisor) x Remaining Business Days"),
        ("Date Column Used", "Invoice Date"),
        ("Aging As-Of Date", f"{aging_asof.strftime('%d-%b-%Y')}"),
        ("Estimated Commission Allocation Method",
         "Estimated Commission is prorated across Month-Year, Cost Center, Account, and Cost Center+Account "
         "rows in proportion to each row's share of Payment Made Total, so every summary sheet reconciles "
         "exactly to the overall Final Accrual."),
    ]
    r = 3
    for label, val in cfg_rows:
        ws2.cell(row=r, column=1, value=label).font = LABEL_FONT
        ws2.cell(row=r, column=2, value=val).alignment = WRAP
        r += 1
    r += 1
    ws2.cell(row=r, column=1, value="Fiscal Calendar Lookup").font = LABEL_FONT
    r += 1
    r = write_table(ws2, r, ["Fiscal Period", "Divisor"], list(FISCAL_CALENDAR.items()))
    r += 1
    ws2.cell(row=r, column=1, value="Aging Bucket Definitions").font = LABEL_FONT
    r += 1
    write_table(ws2, r, ["Bucket"], [[b] for b in AGING_BUCKETS])

    ws3 = wb.create_sheet("Month-Year Summary")
    autosize(ws3, [16, 18, 18, 18, 18])
    ws3["A1"] = "Month-Year Summary"
    ws3["A1"].font = TITLE_FONT
    rows3 = [[m['label'], m['made'], m['remaining'], m['estimated'], m['total']] for m in month_year_summary]
    rows3.append(["Grand Total", sum(m['made'] for m in month_year_summary),
                  sum(m['remaining'] for m in month_year_summary),
                  sum(m['estimated'] for m in month_year_summary),
                  sum(m['total'] for m in month_year_summary)])
    r = write_table(ws3, 3, ["Month-Year", "Payment Made", "Payment Remaining", "Estimated Commission", "Total Accrual"],
                    rows3, currency_cols=(1, 2, 3, 4))
    for c in range(1, 6):
        ws3.cell(row=r - 1, column=c).font = Font(bold=True)

    ws4 = wb.create_sheet("Cost Center Summary")
    autosize(ws4, [18, 18, 18, 18, 18])
    ws4["A1"] = "Cost Center Summary"
    ws4["A1"].font = TITLE_FONT
    rows4 = [[c['cost_center'], c['made'], c['remaining'], c['estimated'], c['total']] for c in cost_center_summary]
    write_table(ws4, 3, ["Cost Center", "Payment Made", "Payment Remaining", "Estimated Commission", "Total Accrual"],
                rows4, currency_cols=(1, 2, 3, 4))

    ws5 = wb.create_sheet("Cost Center + Account")
    autosize(ws5, [18, 20, 18, 18, 18, 18])
    ws5["A1"] = "Cost Center + Account Summary"
    ws5["A1"].font = TITLE_FONT
    rows5 = [[c['cost_center'], c['account'], c['made'], c['remaining'], c['estimated'], c['total']]
             for c in cc_account_summary]
    write_table(ws5, 3, ["Cost Center", "Bill To Account Number", "Payment Made", "Payment Remaining",
                          "Estimated Commission", "Total Accrual"], rows5, currency_cols=(2, 3, 4, 5))

    ws6 = wb.create_sheet("Account Summary")
    autosize(ws6, [22, 18, 18, 18, 18])
    ws6["A1"] = "Account Summary"
    ws6["A1"].font = TITLE_FONT
    rows6 = [[a['account'], a['made'], a['remaining'], a['estimated'], a['total']] for a in account_summary]
    write_table(ws6, 3, ["Bill To Account Number", "Payment Made", "Payment Remaining", "Estimated Commission",
                          "Total Accrual"], rows6, currency_cols=(1, 2, 3, 4))

    ws7 = wb.create_sheet("Aging Summary")
    autosize(ws7, [22, 22, 16, 18])
    ws7["A1"] = "Aging Summary (Payment Remaining)"
    ws7["A1"].font = TITLE_FONT
    rows7 = [[a['bucket'], a['amount'], a['pct'], a['count']] for a in aging_summary]
    write_table(ws7, 3, ["Aging Bucket", "Payment Remaining Amount", "% of Remaining Total", "Transaction Count"],
                rows7, currency_cols=(1,), pct_cols=(2,))

    ws8 = wb.create_sheet("Payment Made Detail")
    made_headers_sorted = [None] * len(made_idx)
    for h, i in made_idx.items():
        made_headers_sorted[i] = h
    ws8.append(made_headers_sorted)
    style_header(ws8, 1, len(made_headers_sorted))
    for r_ in made:
        ws8.append(list(r_['raw']))
    for row in ws8.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = 'yyyy-mm-dd'

    ws9 = wb.create_sheet("Payment Remaining Detail")
    remain_headers_sorted = [None] * len(remain_idx)
    for h, i in remain_idx.items():
        remain_headers_sorted[i] = h
    ws9.append(remain_headers_sorted)
    style_header(ws9, 1, len(remain_headers_sorted))
    for r_ in remaining:
        ws9.append(list(r_['raw']))
    for row in ws9.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = 'yyyy-mm-dd'

    ws10 = wb.create_sheet("Validation Report")
    autosize(ws10, [30, 20, 20])
    ws10["A1"] = "Validation Report"
    ws10["A1"].font = TITLE_FONT
    r = write_table(ws10, 3, ["Report", "Imported Rows", "Screenshot Included"],
                     [["Payment Made", made_val['imported_rows'], "Yes"],
                      ["Payment Remaining", remain_val['imported_rows'], "Yes"]])
    r += 1
    ws10.cell(row=r, column=1, value="Workbook Names").font = LABEL_FONT
    r += 1
    r = write_table(ws10, r, ["Role", "Workbook Name", "Sheet Name"],
                     [["Payment Made", args.made_file.split('/')[-1], made_sheet],
                      ["Payment Remaining", args.remaining_file.split('/')[-1], remain_sheet]])
    r += 1
    write_table(ws10, r, ["Metric", "Payment Made", "Payment Remaining"],
                [
                    ["Raw file rows (incl. blank rows)", made_val['raw_file_rows'], remain_val['raw_file_rows']],
                    ["Imported (real) rows", made_val['imported_rows'], remain_val['imported_rows']],
                    ["Missing Cost Centers", made_val['missing_cost_center'], remain_val['missing_cost_center']],
                    ["Missing Bill To Account Number", made_val['missing_account'], remain_val['missing_account']],
                    ["Blank Commission Amount SUM", made_val['blank_commission'], remain_val['blank_commission']],
                    ["Invalid (non-numeric) values", made_val['invalid_commission'], remain_val['invalid_commission']],
                    ["Negative values", made_val['negative_commission'], remain_val['negative_commission']],
                    ["Duplicate transactions (Invoice + Line)", made_val['duplicate_transactions'], remain_val['duplicate_transactions']],
                ])

    ws11 = wb.create_sheet("Audit Evidence")
    autosize(ws11, [50])
    ws11["A1"] = "Audit Evidence — Screenshot sheets from source workbooks"
    ws11["A1"].font = TITLE_FONT
    row_cursor = 3
    for label, path in [("Payment Made", args.made_file), ("Payment Remaining", args.remaining_file)]:
        ws11.cell(row=row_cursor, column=1, value=f"From {label} workbook: {path.split('/')[-1]}").font = LABEL_FONT
        row_cursor += 1
        for name, data in extract_screenshot_images(path):
            img = XLImage(io.BytesIO(data))
            img.anchor = f"A{row_cursor}"
            ws11.add_image(img)
            row_cursor += 26  # clear the image before placing the next one/label

    ws12 = wb.create_sheet("Executive Dashboard")
    for col in "ABCDEFGHIJK":
        ws12.column_dimensions[col].width = 14
    for col_idx in range(14, 30):
        ws12.column_dimensions[get_column_letter(col_idx)].hidden = True
    ws12["A1"] = f"Executive Dashboard — {args.fiscal_period} {args.fiscal_year}"
    ws12["A1"].font = TITLE_FONT
    ws12.merge_cells("A1:K1")

    kpis = [("Payment Made Total", PAYMENT_MADE_TOTAL), ("Payment Remaining Total", PAYMENT_REMAINING_TOTAL),
            ("Estimated Commission", ESTIMATED_COMMISSION), ("Final Accrual", FINAL_ACCRUAL)]
    ws12.row_dimensions[2].height = 16
    ws12.row_dimensions[3].height = 28
    ws12.row_dimensions[4].height = 16
    for i, (label, val) in enumerate(kpis):
        col = 1 + i * 2
        ws12.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        c = ws12.cell(row=2, column=col, value=label)
        c.font, c.alignment, c.fill = KPI_LABEL_FONT, CENTER, KPI_FILL
        ws12.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col + 1)
        v = ws12.cell(row=3, column=col, value=val)
        v.font, v.number_format, v.alignment, v.fill = KPI_VALUE_FONT, CURRENCY, CENTER, KPI_FILL
        for rr in (2, 3, 4):
            for cc_ in (col, col + 1):
                ws12.cell(row=rr, column=cc_).border = BORDER

    HCOL = 14
    recent_months = month_year_summary[-12:]
    truncated_months = len(month_year_summary) > 12
    ws12.cell(row=1, column=HCOL, value="Month")
    ws12.cell(row=1, column=HCOL + 1, value="Total Accrual")
    for i, m in enumerate(recent_months, start=2):
        ws12.cell(row=i, column=HCOL, value=truncate(m['label']))
        ws12.cell(row=i, column=HCOL + 1, value=m['total'])

    ws12.cell(row=1, column=HCOL + 3, value="Component")
    ws12.cell(row=1, column=HCOL + 4, value="Amount")
    split_rows = [("Payment Made", PAYMENT_MADE_TOTAL), ("Payment Remaining", PAYMENT_REMAINING_TOTAL),
                  ("Estimated Commission", ESTIMATED_COMMISSION)]
    for i, (label, val) in enumerate(split_rows, start=2):
        ws12.cell(row=i, column=HCOL + 3, value=label)
        ws12.cell(row=i, column=HCOL + 4, value=val)

    ws12.cell(row=1, column=HCOL + 6, value="Cost Center")
    ws12.cell(row=1, column=HCOL + 7, value="Total Accrual")
    top10_cc = cost_center_summary[:10]
    for i, c in enumerate(top10_cc, start=2):
        ws12.cell(row=i, column=HCOL + 6, value=truncate(c['cost_center']))
        ws12.cell(row=i, column=HCOL + 7, value=c['total'])

    ws12.cell(row=1, column=HCOL + 9, value="Account")
    ws12.cell(row=1, column=HCOL + 10, value="Total Accrual")
    top10_acct = account_summary[:10]
    for i, a in enumerate(top10_acct, start=2):
        ws12.cell(row=i, column=HCOL + 9, value=truncate(a['account']))
        ws12.cell(row=i, column=HCOL + 10, value=a['total'])

    ws12.cell(row=1, column=HCOL + 12, value="Aging Bucket")
    ws12.cell(row=1, column=HCOL + 13, value="Remaining Amount")
    for i, a in enumerate(aging_summary, start=2):
        ws12.cell(row=i, column=HCOL + 12, value=a['bucket'])
        ws12.cell(row=i, column=HCOL + 13, value=a['amount'])

    month_chart = LineChart()
    month_chart.visible_cells_only = False
    month_chart.title = "Commission Accrual by Month" + (" (last 12 months)" if truncated_months else "")
    month_chart.y_axis.title, month_chart.x_axis.title = "Total Accrual ($)", "Month-Year"
    month_chart.width, month_chart.height = 11, 8
    data_ref = Reference(ws12, min_col=HCOL + 1, min_row=1, max_row=1 + len(recent_months))
    cats_ref = Reference(ws12, min_col=HCOL, min_row=2, max_row=1 + len(recent_months))
    month_chart.add_data(data_ref, titles_from_data=True)
    month_chart.set_categories(cats_ref)
    set_text_categories(month_chart, cats_ref)
    month_chart.anchor = "A7"
    ws12.add_chart(month_chart)

    pie_chart = PieChart()
    pie_chart.visible_cells_only = False
    pie_chart.title = "Payment Split"
    pie_chart.width, pie_chart.height = 11, 8
    data_ref = Reference(ws12, min_col=HCOL + 4, min_row=1, max_row=1 + len(split_rows))
    cats_ref = Reference(ws12, min_col=HCOL + 3, min_row=2, max_row=1 + len(split_rows))
    pie_chart.add_data(data_ref, titles_from_data=True)
    pie_chart.set_categories(cats_ref)
    set_text_categories(pie_chart, cats_ref)
    pie_chart.anchor = "G7"
    ws12.add_chart(pie_chart)

    cc_chart = BarChart()
    cc_chart.visible_cells_only = False
    cc_chart.type = "bar"
    cc_chart.title = "Top 10 Cost Centers by Total Accrual"
    cc_chart.width, cc_chart.height = 11, 8
    data_ref = Reference(ws12, min_col=HCOL + 7, min_row=1, max_row=1 + len(top10_cc))
    cats_ref = Reference(ws12, min_col=HCOL + 6, min_row=2, max_row=1 + len(top10_cc))
    cc_chart.add_data(data_ref, titles_from_data=True)
    cc_chart.set_categories(cats_ref)
    set_text_categories(cc_chart, cats_ref)
    cc_chart.anchor = "A25"
    ws12.add_chart(cc_chart)

    acct_chart = BarChart()
    acct_chart.visible_cells_only = False
    acct_chart.type = "bar"
    acct_chart.title = "Top 10 Accounts by Total Accrual"
    acct_chart.width, acct_chart.height = 11, 8
    data_ref = Reference(ws12, min_col=HCOL + 10, min_row=1, max_row=1 + len(top10_acct))
    cats_ref = Reference(ws12, min_col=HCOL + 9, min_row=2, max_row=1 + len(top10_acct))
    acct_chart.add_data(data_ref, titles_from_data=True)
    acct_chart.set_categories(cats_ref)
    set_text_categories(acct_chart, cats_ref)
    acct_chart.anchor = "G25"
    ws12.add_chart(acct_chart)

    aging_chart = BarChart()
    aging_chart.visible_cells_only = False
    aging_chart.type = "col"
    aging_chart.title = "Aging Distribution (Payment Remaining)"
    aging_chart.width, aging_chart.height = 23, 7
    data_ref = Reference(ws12, min_col=HCOL + 13, min_row=1, max_row=1 + len(aging_summary))
    cats_ref = Reference(ws12, min_col=HCOL + 12, min_row=2, max_row=1 + len(aging_summary))
    aging_chart.add_data(data_ref, titles_from_data=True)
    aging_chart.set_categories(cats_ref)
    set_text_categories(aging_chart, cats_ref)
    aging_chart.anchor = "A43"
    ws12.add_chart(aging_chart)

    ws12.cell(row=59, column=1, value="Executive Commentary").font = LABEL_FONT
    ws12.merge_cells(start_row=60, start_column=1, end_row=68, end_column=10)
    c = ws12.cell(row=60, column=1, value=commentary + " " + key_risks)
    c.alignment, c.fill = WRAP, PatternFill("solid", fgColor="F2F2F2")

    ws12.cell(row=71, column=1, value="Data Quality Summary").font = LABEL_FONT
    dq_lines = (
        f"Required Columns: Cost Center, Bill To Account Number, Commission Amount SUM, Invoice Date — present in both workbooks.\n"
        f"Missing Cost Centers: {made_val['missing_cost_center'] + remain_val['missing_cost_center']}\n"
        f"Missing Accounts: {made_val['missing_account'] + remain_val['missing_account']}\n"
        f"Invalid Values: {made_val['invalid_commission'] + remain_val['invalid_commission']}\n"
        f"Negative Values: {made_val['negative_commission'] + remain_val['negative_commission']}\n"
        f"Duplicate Transactions: {made_val['duplicate_transactions'] + remain_val['duplicate_transactions']}\n"
        f"Screenshots Preserved: Yes (both workbooks, Audit Evidence sheet)"
    )
    ws12.merge_cells(start_row=72, start_column=1, end_row=80, end_column=10)
    c = ws12.cell(row=72, column=1, value=dq_lines)
    c.alignment, c.fill = WRAP, PatternFill("solid", fgColor="F2F2F2")

    assert set(TAB_ORDER) == set(wb.sheetnames), (set(TAB_ORDER) ^ set(wb.sheetnames))
    wb._sheets = [wb[name] for name in TAB_ORDER]
    wb.active = 0

    output = args.output or f"Commissions_Accrual_Output_{args.fiscal_period}_{args.fiscal_year}.xlsx"
    wb.save(output)
    print(f"\nSaved: {output}")
    print(f"Sheets ({len(wb.sheetnames)}): {wb.sheetnames}")
    return output


def parse_args(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--made-file", required=True, help="Path to the Payment Made workbook")
    p.add_argument("--remaining-file", required=True, help="Path to the Payment Remaining workbook")
    p.add_argument("--made-sheet", default=None, help="Sheet name in the Payment Made workbook (auto-detected if omitted)")
    p.add_argument("--remaining-sheet", default=None, help="Sheet name in the Payment Remaining workbook (auto-detected if omitted)")
    p.add_argument("--fiscal-period", required=True, choices=sorted(FISCAL_CALENDAR), help="P01-P12")
    p.add_argument("--fiscal-year", required=True, type=int)
    p.add_argument("--fiscal-divisor", type=int, default=None, help="Override the P01-P12 lookup table")
    p.add_argument("--remaining-business-days", type=int, default=3)
    p.add_argument("--aging-asof", default=None, help="YYYY-MM-DD; default: source file's 'Selected Date' column, else today")
    p.add_argument("--output", default=None, help="Output .xlsx path (default: Commissions_Accrual_Output_<Period>_<Year>.xlsx)")
    return p.parse_args(argv)


if __name__ == "__main__":
    build(parse_args())
