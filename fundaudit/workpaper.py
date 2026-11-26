"""
Export the fund audit as an Excel workpaper with live formulas.

Every number on the Summary sheet is a formula that points back to a detail
sheet (Price Testing, Custodian Rec, Mgmt Fee) or to an input sheet
(NAV Statement, Assumptions, Evidence), so a reviewer can trace each figure.

Colour convention: blue = hard-coded input / evidence, green = link to another
sheet, black = calculation.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .procedures import AuditResult

FONT = "Arial"
NAVY = "1F3864"
NUM = '#,##0.00;(#,##0.00);"-"'
INT = '#,##0;(#,##0);"-"'
PCT = '0.00%;(0.00%);"-"'
PX = '#,##0.0000'
DATE = "dd-mmm-yyyy"

F = Font(name=FONT, size=10)
FB = Font(name=FONT, size=10, bold=True)
FT = Font(name=FONT, size=14, bold=True, color=NAVY)
FH = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FIN = Font(name=FONT, size=10, color="0000FF")
FLK = Font(name=FONT, size=10, color="008000")
FN = Font(name=FONT, size=9, italic=True, color="595959")
HEAD = PatternFill("solid", fgColor=NAVY)
TOTAL = PatternFill("solid", fgColor="D9E1F2")
SIGN = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", fgColor="F8CBAD")
AMBER = PatternFill("solid", fgColor="FFE699")
GREEN = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(style="thin", color="808080")
DOUBLE = Border(top=THIN, bottom=Side(style="double", color="000000"))
TOP = Border(top=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def _py(v):
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def _header(ws, row, headers, widths=None):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font, cell.fill = FH, HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        if widths:
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]


def _put(ws, ref, value, font=F, fmt=None, fill=None, border=None, align=None):
    c = ws[ref]
    c.value = _py(value)
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = align
    return c


def _title(ws, title, subtitle):
    _put(ws, "A1", title, FT)
    _put(ws, "A2", subtitle, FN)


def export_workpaper(r: AuditResult, path: Path, preparer: str = "[Your name]") -> Path:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    findings_ws = wb.create_sheet("Findings")
    pt = wb.create_sheet("Price Testing")
    cr = wb.create_sheet("Custodian Rec")
    fee = wb.create_sheet("Mgmt Fee")
    navs = wb.create_sheet("NAV Statement")
    asm = wb.create_sheet("Assumptions")
    ev = wb.create_sheet("Evidence")
    terms, st, s = r.terms, r.statement, r.settings

    # ------------------------------------------------------------------
    # Assumptions (inputs)
    # ------------------------------------------------------------------
    _title(asm, "Assumptions & parameters", "Blue cells are inputs. Change them and every sheet recalculates.")
    asm.column_dimensions["A"].width = 46
    asm.column_dimensions["B"].width = 16
    asm.column_dimensions["C"].width = 60
    _header(asm, 4, ["Parameter", "Value", "Source / rationale"])
    params = [
        ("Valuation date (year end)", pd.Timestamp(terms["valuation_date"]), DATE, "Fund financial year end"),
        ("Price tolerance", s["price_tolerance"], PCT, "Differences above this are exceptions (illustrative)"),
        ("Stale price threshold (days)", s["stale_days"], INT, "Fund price older than this is stale"),
        ("Overall materiality (% of NAV)", s["materiality_pct"], PCT, "Illustrative; NAV is the key measure for fund investors"),
        ("Performance materiality (% of overall)", s["performance_pct"], PCT, "Illustrative"),
        ("Clearly trivial (% of overall)", s["trivial_pct"], PCT, "Misstatements below this are not accumulated"),
        ("Management fee rate (p.a.)", terms["management_fee_rate"], PCT, "Fund prospectus (simulated)"),
        ("Day count", terms["day_count"], INT, "Actual/365"),
    ]
    A = {}
    for i, (label, val, fmt, src) in enumerate(params, start=5):
        _put(asm, f"A{i}", label)
        _put(asm, f"B{i}", val, FIN, fmt)
        _put(asm, f"C{i}", src, FN)
        A[label] = f"Assumptions!$B${i}"
    VDATE, TOL, STALE = A["Valuation date (year end)"], A["Price tolerance"], A["Stale price threshold (days)"]
    MATP, PMP, CTP = A["Overall materiality (% of NAV)"], A["Performance materiality (% of overall)"], A["Clearly trivial (% of overall)"]
    RATE, DAYS = A["Management fee rate (p.a.)"], A["Day count"]

    fx_start = 5 + len(params) + 2
    _header(asm, fx_start, ["FX rate (SGD per 1 unit)", "Rate", "Source"])
    FXREF = {}
    for j, (ccy, rate) in enumerate(terms["fx_sgd_per_unit"].items(), start=fx_start + 1):
        _put(asm, f"A{j}", ccy)
        _put(asm, f"B{j}", rate, FIN, PX)
        _put(asm, f"C{j}", "Closing rate at valuation date (simulated)", FN)
        FXREF[ccy] = j
    fx_lo, fx_hi = fx_start + 1, fx_start + len(FXREF)
    FXLOOK = lambda cell: f"INDEX(Assumptions!$B${fx_lo}:$B${fx_hi},MATCH({cell},Assumptions!$A${fx_lo}:$A${fx_hi},0))"

    # ------------------------------------------------------------------
    # NAV Statement (inputs, as reported by the fund administrator)
    # ------------------------------------------------------------------
    _title(navs, "NAV statement as reported", "Source: fund administrator NAV statement at the valuation date (simulated)")
    navs.column_dimensions["A"].width = 44
    navs.column_dimensions["B"].width = 18
    _header(navs, 4, ["Line item", "SGD"])
    lines = [("Investments at market value", "investments_at_market_value", 1),
             ("Cash at bank", "cash_at_bank", 1),
             ("Dividends receivable", "dividends_receivable", 1),
             ("Payable for investments purchased", "payable_for_investments_purchased", -1),
             ("Redemptions payable", "redemptions_payable", -1),
             ("Other payables", "other_payables", -1),
             ("Accrued management fee", "accrued_management_fee", -1)]
    N = {}
    for i, (label, key, sign) in enumerate(lines, start=5):
        _put(navs, f"A{i}", label)
        _put(navs, f"B{i}", st[key], FIN, NUM)
        N[key] = (f"'NAV Statement'!$B${i}", sign)
    k = 5 + len(lines)
    _put(navs, f"A{k}", "Reported NAV", FB)
    _put(navs, f"B{k}", st["reported_nav"], Font(name=FONT, size=10, bold=True, color="0000FF"), NUM, border=DOUBLE)
    _put(navs, f"A{k+1}", "Units in issue")
    _put(navs, f"B{k+1}", st["units_in_issue"], FIN, INT)
    _put(navs, f"A{k+2}", "Reported NAV per unit", FB)
    _put(navs, f"B{k+2}", st["reported_nav_per_unit"], FIN, PX)
    NAVREP, UNITS, NAVPU = f"'NAV Statement'!$B${k}", f"'NAV Statement'!$B${k+1}", f"'NAV Statement'!$B${k+2}"

    # ------------------------------------------------------------------
    # Evidence: pending trades and corporate actions
    # ------------------------------------------------------------------
    _title(ev, "Evidence obtained", "Broker confirmations and corporate action notices (simulated)")
    _header(ev, 4, ["Trade ID", "Security", "Side", "Quantity", "Price", "Ccy", "Trade date", "Settle date"],
            [14, 10, 8, 12, 10, 7, 13, 13])
    tr = r.raw["pending"]
    for i, t in enumerate(tr.itertuples(index=False), start=5):
        for c, v in enumerate(t, start=1):
            cell = ev.cell(row=i, column=c, value=_py(v))
            cell.font = FIN
            if c in (7, 8):
                cell.number_format = DATE
            if c == 4:
                cell.number_format = INT
    tr_lo, tr_hi = 5, 4 + max(len(tr), 1)
    ca_head = tr_hi + 3
    _put(ev, f"A{ca_head-1}", "Corporate actions", FB)
    _header(ev, ca_head, ["CA ID", "Security", "Type", "New shares", "Per old shares", "Ex date", "Pay date"])
    ca = r.raw["corp_actions"]
    for i, c_ in enumerate(ca.itertuples(index=False), start=ca_head + 1):
        for c, v in enumerate(c_, start=1):
            cell = ev.cell(row=i, column=c, value=_py(v))
            cell.font = FIN
            if c in (6, 7):
                cell.number_format = DATE
    ca_lo, ca_hi = ca_head + 1, ca_head + max(len(ca), 1)

    # ------------------------------------------------------------------
    # Price testing
    # ------------------------------------------------------------------
    _title(pt, "Price testing - valuation of investments",
           "Fund price vs independent price at the valuation date. Exceptions above tolerance are misstatements.")
    heads = ["Security", "Name", "Class", "Ccy", "Unit", "Quantity", "Fund price", "Fund price date",
             "Independent price", "Factor", "FX rate", "MV per fund (SGD)", "MV independent (SGD)",
             "Difference (SGD)", "Price diff %", "Price age (days)", "Stale?", "Exception?", "NAV misstatement (SGD)",
             "Status"]
    _header(pt, 4, heads, [9, 34, 8, 6, 8, 13, 11, 13, 11, 7, 9, 16, 16, 14, 10, 9, 8, 10, 15, 22])
    p = r.price_test
    PT_ROW = {}
    for i, row in enumerate(p.itertuples(index=False), start=5):
        PT_ROW[row.security_id] = i
        _put(pt, f"A{i}", row.security_id)
        _put(pt, f"B{i}", row.security_name)
        _put(pt, f"C{i}", row.asset_class)
        _put(pt, f"D{i}", row.currency)
        _put(pt, f"E{i}", row.unit)
        _put(pt, f"F{i}", row.quantity, FIN, INT)
        _put(pt, f"G{i}", row.fund_price, FIN, '#,##0.00')
        _put(pt, f"H{i}", row.fund_price_date, FIN, DATE)
        _put(pt, f"I{i}", row.independent_price, FIN, '#,##0.00')
        _put(pt, f"J{i}", f'=IF(E{i}="nominal",0.01,1)', fmt='0.00')
        _put(pt, f"K{i}", f"={FXLOOK(f'D{i}')}", FLK, PX)
        _put(pt, f"L{i}", f"=ROUND(F{i}*G{i}*J{i}*K{i},2)", fmt=NUM)
        _put(pt, f"M{i}", f"=ROUND(F{i}*I{i}*J{i}*K{i},2)", fmt=NUM)
        _put(pt, f"N{i}", f"=L{i}-M{i}", fmt=NUM)
        _put(pt, f"O{i}", f"=(G{i}-I{i})/I{i}", fmt=PCT)
        _put(pt, f"P{i}", f"={VDATE}-H{i}", fmt=INT)
        _put(pt, f"Q{i}", f'=IF(P{i}>{STALE},"Yes","No")')
        _put(pt, f"R{i}", f'=IF(ABS(O{i})>{TOL},"Yes","No")')
        _put(pt, f"S{i}", f'=IF(R{i}="Yes",N{i},0)', fmt=NUM)
        _put(pt, f"T{i}", row.status)
    pt_hi = 4 + len(p)
    t = pt_hi + 1
    _put(pt, f"B{t}", "Total", FB)
    for col in "LMNS":
        _put(pt, f"{col}{t}", f"=SUM({col}5:{col}{pt_hi})", FB, NUM, border=DOUBLE)
    _put(pt, f"B{t+1}", "Agrees to NAV statement investments?")
    _put(pt, f"L{t+1}", f'=IF(ABS(L{t}-{N["investments_at_market_value"][0]})<0.01,"OK","CHECK")', FB)
    for col in "QR":
        pt.conditional_formatting.add(f"{col}5:{col}{pt_hi}", CellIsRule(operator="equal", formula=['"Yes"'], fill=RED))
    pt.freeze_panes = "C5"
    INV_TOTAL = f"'Price Testing'!$L${t}"

    # ------------------------------------------------------------------
    # Custodian reconciliation
    # ------------------------------------------------------------------
    _title(cr, "Custodian reconciliation - existence & completeness",
           "Quantities per the fund's books vs the custodian statement, after expected differences.")
    heads = ["Security", "Name", "Ccy", "Qty per books", "Qty per custodian", "Difference (cust - books)",
             "Unsettled purchases", "Corporate action shares", "Unexplained difference", "Independent price",
             "Factor", "FX rate", "NAV misstatement (SGD)", "Status", "Assertion", "Explanation"]
    _header(cr, 4, heads, [9, 34, 6, 13, 13, 13, 12, 12, 12, 11, 7, 9, 15, 24, 13, 60])
    c = r.custodian_rec
    for i, row in enumerate(c.itertuples(index=False), start=5):
        _put(cr, f"A{i}", row.security_id)
        _put(cr, f"B{i}", row.security_name)
        _put(cr, f"C{i}", row.currency)
        _put(cr, f"D{i}", f"='Price Testing'!F{PT_ROW[row.security_id]}", FLK, INT)
        _put(cr, f"E{i}", row.quantity_held, FIN, INT)
        _put(cr, f"F{i}", f"=E{i}-D{i}", fmt=INT)
        _put(cr, f"G{i}", f'=SUMIFS(Evidence!$D${tr_lo}:$D${tr_hi},Evidence!$B${tr_lo}:$B${tr_hi},A{i},'
                          f'Evidence!$C${tr_lo}:$C${tr_hi},"BUY",Evidence!$H${tr_lo}:$H${tr_hi},">"&{VDATE})', FLK, INT)
        _put(cr, f"H{i}", f'=IFERROR(D{i}*INDEX(Evidence!$D${ca_lo}:$D${ca_hi},MATCH(A{i},Evidence!$B${ca_lo}:$B${ca_hi},0))'
                          f'/INDEX(Evidence!$E${ca_lo}:$E${ca_hi},MATCH(A{i},Evidence!$B${ca_lo}:$B${ca_hi},0)),0)', FLK, INT)
        _put(cr, f"I{i}", f"=F{i}+G{i}-H{i}", fmt=INT)
        pr = PT_ROW[row.security_id]
        _put(cr, f"J{i}", f"='Price Testing'!I{pr}", FLK, '#,##0.00')
        _put(cr, f"K{i}", f"='Price Testing'!J{pr}", FLK, '0.00')
        _put(cr, f"L{i}", f"='Price Testing'!K{pr}", FLK, PX)
        # Books should hold custodian qty adjusted for unsettled purchases; anything else misstates NAV
        _put(cr, f"M{i}", f"=-ROUND((F{i}+G{i})*J{i}*K{i}*L{i},2)", fmt=NUM)
        _put(cr, f"N{i}", row.status)
        _put(cr, f"O{i}", row.assertion)
        _put(cr, f"P{i}", row.explanation, align=WRAP)
    cr_hi = 4 + len(c)
    _put(cr, f"B{cr_hi+1}", "Total", FB)
    _put(cr, f"M{cr_hi+1}", f"=SUM(M5:M{cr_hi})", FB, NUM, border=DOUBLE)
    cr.conditional_formatting.add(f"N5:N{cr_hi}", CellIsRule(operator="equal", formula=['"Not held at custodian"'], fill=RED))
    cr.conditional_formatting.add(f"N5:N{cr_hi}", CellIsRule(operator="equal", formula=['"Unrecorded corporate action"'], fill=AMBER))
    _put(cr, f"B{cr_hi+3}", "Procedure: the misstatement is the custodian quantity, adjusted for unsettled purchases, less the "
                            "quantity in the books, valued at the independent price.", FN)
    cr.freeze_panes = "C5"
    CR_ROW = {sid: 5 + idx for idx, sid in enumerate(c["security_id"])}

    # ------------------------------------------------------------------
    # Management fee recalculation
    # ------------------------------------------------------------------
    _title(fee, "Management fee recalculation - accuracy",
           "Daily accrual = net assets x contractual rate / day count, recalculated for October 2026.")
    _header(fee, 4, ["Date", "Net assets (SGD)", "Daily fee (SGD)"], [14, 18, 16])
    d = r.fee_daily
    for i, row in enumerate(d.itertuples(index=False), start=5):
        _put(fee, f"A{i}", row.date, FIN, DATE)
        _put(fee, f"B{i}", row.net_assets_sgd, FIN, NUM)
        _put(fee, f"C{i}", f"=B{i}*{RATE}/{DAYS}", fmt=NUM)
    fh = 4 + len(d)
    fee.column_dimensions["E"].width = 44
    fee.column_dimensions["F"].width = 16
    _header(fee, 4, ["Date", "Net assets (SGD)", "Daily fee (SGD)", "", "Recalculation", "SGD"])
    fee["D4"].fill = PatternFill(fill_type=None)
    calc = [("Expected fee for October (recalculated)", f"=ROUND(SUM(C5:C{fh}),2)", NUM, F),
            ("Fee accrued by the manager (per NAV statement)", f"={N['accrued_management_fee'][0]}", NUM, FLK),
            ("Over / (under) accrual", "=F6-F5", NUM, FB),
            ("Implied annual rate charged", f"=F6*{DAYS}/SUM(B5:B{fh})", PCT, F),
            ("Contractual rate per prospectus", f"={RATE}", PCT, FLK),
            ("Effect on NAV: overstatement / (understatement)", "=-F7", NUM, FB)]
    for i, (label, f_, fmt, font) in enumerate(calc, start=5):
        _put(fee, f"E{i}", label)
        _put(fee, f"F{i}", f_, font, fmt)
    FEE_MIS = "'Mgmt Fee'!$F$10"

    # ------------------------------------------------------------------
    # Findings
    # ------------------------------------------------------------------
    _title(findings_ws, "Audit findings", "For discussion with management; responses to be obtained before sign-off.")
    heads = ["Finding", "Rating", "Procedure", "Assertion", "Reference", "Observation",
             "NAV overstatement / (understatement) SGD", "Recommendation", "Management response"]
    _header(findings_ws, 4, heads, [8, 9, 22, 13, 11, 60, 18, 62, 20])
    fd = r.findings
    for i, row in enumerate(fd.itertuples(index=False), start=5):
        vals = [row.finding, row.rating, row.procedure, row.assertion, row.reference, row.observation,
                row.nav_overstatement, row.recommendation, row.management_response]
        for col, v in enumerate(vals, start=1):
            cell = findings_ws.cell(row=i, column=col, value=_py(v))
            cell.font = F
            cell.alignment = WRAP
            if col == 7:
                cell.number_format = NUM
        findings_ws.row_dimensions[i].height = 64
    fd_hi = 4 + len(fd)
    findings_ws.conditional_formatting.add(f"B5:B{fd_hi}", CellIsRule(operator="equal", formula=['"High"'], fill=RED))
    findings_ws.conditional_formatting.add(f"B5:B{fd_hi}", CellIsRule(operator="equal", formula=['"Medium"'], fill=AMBER))
    findings_ws.freeze_panes = "C5"

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    ws = summary
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [3, 58, 18, 14, 4, 4]):
        ws.column_dimensions[col].width = w
    _put(ws, "B1", "Fund Audit Workpaper - Investments, Fees & NAV", FT)
    _put(ws, "B2", "SYNTHETIC DATA - fictional fund created for a portfolio project", FN)
    meta = [("Fund", terms["fund"]), ("Manager", terms["manager"]), ("Custodian", terms["custodian"]),
            ("Valuation date", f"={VDATE}"), ("Base currency", terms["base_currency"]),
            ("Prepared by / date", preparer), ("Reviewed by / date", "")]
    for i, (k_, v) in enumerate(meta, start=4):
        _put(ws, f"B{i}", k_, FB)
        cell = _put(ws, f"C{i}", v, FLK if str(v).startswith("=") else F, DATE if k_ == "Valuation date" else None,
                    align=Alignment(horizontal="left"))
        if k_.startswith(("Prepared", "Reviewed")):
            cell.fill = SIGN
    _put(ws, "B12", "Objective", FB)
    _put(ws, "B13", "Obtain evidence over the existence, completeness and valuation of investments, the accuracy of the "
                    "management fee, and the arithmetic accuracy of the reported NAV; evaluate misstatements against materiality.",
         align=WRAP)
    ws.merge_cells("B13:F13")
    ws.row_dimensions[13].height = 30

    row = 15

    def hdr(text, right="SGD"):
        nonlocal row
        for col in "BC":
            ws[f"{col}{row}"].fill = HEAD
        _put(ws, f"B{row}", text, FH)
        _put(ws, f"C{row}", right, FH, align=Alignment(horizontal="right"))
        row += 1

    def line(label, formula, font=F, fmt=NUM, border=None, fill=None, bold_label=False):
        nonlocal row
        _put(ws, f"B{row}", label, FB if bold_label else F, fill=fill)
        _put(ws, f"C{row}", formula, font, fmt, fill=fill, border=border)
        row += 1
        return row - 1

    # Materiality
    hdr("Materiality")
    line("Benchmark: reported NAV", f"={NAVREP}", FLK)
    mat_o = line("Overall materiality", f"=ROUND(C{row-1}*{MATP},2)")
    mat_p = line("Performance materiality", f"=ROUND(C{mat_o}*{PMP},2)")
    mat_t = line("Clearly trivial threshold", f"=ROUND(C{mat_o}*{CTP},2)")
    row += 1

    # NAV recomputation
    hdr("NAV recomputation")
    start = row
    for label, (ref, sign) in [("Investments at market value", N["investments_at_market_value"]),
                               ("Cash at bank", N["cash_at_bank"]),
                               ("Dividends receivable", N["dividends_receivable"]),
                               ("Payable for investments purchased", N["payable_for_investments_purchased"]),
                               ("Redemptions payable", N["redemptions_payable"]),
                               ("Other payables", N["other_payables"]),
                               ("Accrued management fee", N["accrued_management_fee"])]:
        line(("Less: " if sign < 0 else "") + label, f"={'-' if sign < 0 else ''}{ref}", FLK)
    nav_r = line("Recomputed NAV", f"=SUM(C{start}:C{row-1})", FB, border=DOUBLE, fill=TOTAL, bold_label=True)
    line("Reported NAV", f"={NAVREP}", FLK)
    chk1 = line("Recomputed NAV agrees to reported NAV?", f'=IF(ABS(C{nav_r}-C{row-1})<0.01,"OK","CHECK")', FB, None)
    chk2 = line("Investments agree to price testing population?", f'=IF(ABS({INV_TOTAL}-C{start})<0.01,"OK","CHECK")', FB, None)
    nav_pu = line("Recomputed NAV per unit", f"=ROUND(C{nav_r}/{UNITS},4)", F, PX)
    row += 1

    # Misstatement schedule
    for col, txt in zip("BCD", ["Summary of misstatements", "NAV over / (under)", "Above trivial?"]):
        ws[f"{col}{row}"].fill = HEAD
        _put(ws, f"{col}{row}", txt, FH, align=Alignment(horizontal="right" if col != "B" else "left"))
    row += 1
    mis_start = row
    m = r.misstatements
    m4_row = None
    for mr in m.itertuples(index=False):
        if mr.reference == "MGMT FEE":
            ref = FEE_MIS
        elif mr.area == "Valuation":
            ref = f"'Price Testing'!$S${PT_ROW[mr.reference]}"
        else:
            ref = f"'Custodian Rec'!$M${CR_ROW[mr.reference]}"
        _put(ws, f"B{row}", f"{mr.ref} {mr.description}", align=WRAP)
        _put(ws, f"C{row}", f"={ref}", FLK, NUM)
        _put(ws, f"D{row}", f'=IF(ABS(C{row})>=$C${mat_t},"Yes","No")', align=Alignment(horizontal="right"))
        if mr.area == "Existence":
            m4_row = row
        row += 1
    mis_end = row - 1
    acc = line("Accumulated misstatements (above clearly trivial)",
               f'=SUMIFS(C{mis_start}:C{mis_end},D{mis_start}:D{mis_end},"Yes")', FB, border=DOUBLE, fill=TOTAL, bold_label=True)
    pct = line("As % of reported NAV", f"=C{acc}/{NAVREP}", F, PCT)
    adj = line("Adjusted NAV if corrected", f"={NAVREP}-C{acc}")
    adj_pu = line("Adjusted NAV per unit", f"=ROUND(C{adj}/{UNITS},4)", F, PX)
    line("Reported NAV per unit", f"={NAVPU}", FLK, PX)
    concl = line("Evaluation",
                 f'=IF(ABS(C{acc})>=C{mat_o},"MATERIAL",IF(ABS(C{acc})>=C{mat_p},"Above PM","Below PM"))', FB, None)
    if m4_row:
        sens = line("Sensitivity: excluding the holding not at custodian (if evidenced)",
                    f"=C{acc}-IF(D{m4_row}=\"Yes\",C{m4_row},0)")
        line("  -> evaluation", f'=IF(ABS(C{sens})>=C{mat_o},"MATERIAL",IF(ABS(C{sens})>=C{mat_p},"Above PM","Below PM"))', FB, None)
    eval_end = row - 1
    for rng in (f"C{chk1}:C{chk2}", f"C{concl}:C{eval_end}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Below PM"'], fill=GREEN))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Above PM"'], fill=AMBER))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"MATERIAL"'], fill=RED))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"CHECK"'], fill=RED))
    for rr in range(chk1, eval_end + 1):
        ws[f"C{rr}"].alignment = Alignment(horizontal="right")
    row += 1

    # Conclusion
    _put(ws, f"B{row}", "Conclusion", FB)
    row += 1
    mat = r.materiality
    text = (f"The reported NAV recomputes without arithmetic error. Accumulated misstatements of SGD "
            f"{mat['accumulated_misstatement']:,.0f} ({mat['nav_impact_pct']:.2%} of NAV) exceed overall materiality of "
            f"SGD {mat['overall']:,.0f}, driven mainly by a holding that is not held at the custodian. Management is "
            f"requested to provide custodian evidence for that holding and to correct the valuation, corporate action and "
            f"fee errors. Until then the NAV cannot be relied on. The fee over-accrual is below the clearly trivial "
            f"threshold but is reported because it is a recurring breach of the prospectus.")
    _put(ws, f"B{row}", text, align=WRAP)
    ws.merge_cells(f"B{row}:F{row}")
    ws.row_dimensions[row].height = 80
    row += 2
    _put(ws, f"B{row}", "Legend: blue = input / evidence, green = link to another sheet, black = formula. Yellow = sign-off.", FN)

    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
