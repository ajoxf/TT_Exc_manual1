"""Dashboard: the MAIN view. Three spread families, each showing every leg's
Bid / Ask / Bid-Ask Gap, then the spread's own Bid / Ask / Gap, with the live
z-score, mean, standard deviation and take-profit alongside.

Everything else - costs, break-even, warm-up gates, feed rate, edge filter,
the fourth (listed) spread - moved to the Detail sheet. Nothing was removed.

NOTHING on this sheet is a formula. Every cell is written by VBA, and only when
the value differs from what is already there.
"""
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from wb_common import *

PX = "#,##0.0000"
USD = '$#,##0.00;($#,##0.00);-'
Z2 = "0.00"

# (title row, block title, [(symbol, description), ...], spread label, slot)
BLOCKS = [
    (6,  "BRENT / WTI",
     [("BZV6", "Brent Last Day Financial Oct26"),
      ("CLV6", "WTI Crude Oct26")],
     "BZ - CL   SPREAD", 1),
    (13, "ULSD / WTI   (HO/CL crack, $/bbl)",
     [("HOV6", "NY Harbor ULSD Oct26  (x42)"),
      ("CLV6", "WTI Crude Oct26")],
     "HO - CL   CRACK SPREAD", 2),
    (20, "3:2:1 CRACK   ((2 x RB + 1 x HO) x 42 / 3) - CL,  $/bbl",
     [("RBV6", "RBOB Gasoline Oct26  (x42, weight 2)"),
      ("HOV6", "NY Harbor ULSD Oct26  (x42, weight 1)"),
      ("CLV6", "WTI Crude Oct26  (weight 3)")],
     "3:2:1   CRACK SPREAD", 3),
]

STAT_LABELS = ["Live Z-score", "Mean", "Std Dev", "Window"]
TRADE_LABELS = ["Signal", "Entry (touch)", "TAKE PROFIT", "TP in sigma"]


def build(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 34, "C": 13, "D": 13, "E": 15, "F": 3,
                    "G": 19, "H": 15, "I": 3, "J": 20, "K": 18, "L": 3, "M": 46})
    ws.freeze_panes = "B5"

    # ---- header -----------------------------------------------------------
    title(ws, "B1", "TT RTD Z-SCORE MONITOR")
    ws["B2"] = ("z is for ENTRIES only.  Exits act on MONEY, never on z.   "
                "All execution is manual in TT - this workbook never places an order.")
    ws["B2"].font = f(10, bold=True, color=C_WARN)

    for cell, txt, fill in (("J1", "  >  START  ", "C6EFCE"), ("K1", "  #  STOP  ", "FFC7CE")):
        c = ws[cell]; c.value = txt
        c.font = f(10, bold=True); c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center"); c.border = BOX
    note(ws, "M1", "Assign macros to shapes over these two cells: StartMonitor / StopMonitor")

    label(ws, "B3", "Monitor", bold=True)
    value_cell(ws, "C3", "STOPPED", bold=True, fill=C_GREY)
    label(ws, "E3", "Rates", bold=True)
    ws.merge_cells("F3:K3")
    value_cell(ws, "F3", "-", align="left")
    label(ws, "B4", "Market", bold=True)
    input_cell(ws, "C4", "CME")
    label(ws, "E4", "Clock", bold=True)
    value_cell(ws, "F4", "--:--:--", bold=True)
    label(ws, "H4", "Session", bold=True)
    value_cell(ws, "J4", "-", align="left")
    note(ws, "M3", "Thresholds, costs and the take-profit rule all live on Config and "
                   "take effect without a restart.")
    note(ws, "M4", "Costs, break-even, warm-up gates, feed rate and the edge filter are on the "
                   "Detail sheet.")

    # ---- the three blocks -------------------------------------------------
    for trow, btitle, legs, spread_label, slot in BLOCKS:
        block_title(ws, f"B{trow}", btitle)
        hdr = trow + 1
        header_row(ws, hdr, {"B": "Instrument", "C": "Bid", "D": "Ask", "E": "Bid-Ask Gap",
                             "G": "Statistic", "H": "Value", "J": "Trade", "K": "Value"})
        top = hdr + 1

        for k, (sym, desc) in enumerate(legs):
            r = top + k
            value_cell(ws, f"B{r}", sym, bold=True, align="left")
            note(ws, f"M{r}", desc)
            for col in ("C", "D", "E"):
                value_cell(ws, f"{col}{r}", None, PX)

        sr = top + len(legs)                      # the spread row
        value_cell(ws, f"B{sr}", spread_label, bold=True, align="left")
        ws[f"B{sr}"].fill = PatternFill("solid", fgColor=C_BLOCK_FILL)
        for col in ("C", "D", "E"):
            c = value_cell(ws, f"{col}{sr}", None, PX, bold=True)
            c.fill = PatternFill("solid", fgColor=C_BLOCK_FILL)

        # statistics + trade panels, aligned to the four rows top..top+3
        for k in range(4):
            r = top + k
            label(ws, f"G{r}", STAT_LABELS[k], bold=(k == 0))
            ws[f"G{r}"].border = BOX
            label(ws, f"J{r}", TRADE_LABELS[k], bold=(k == 2))
            ws[f"J{r}"].border = BOX
            fmt = {0: Z2, 1: PX, 2: PX, 3: None}[k]
            value_cell(ws, f"H{r}", None, fmt, bold=(k == 0))
            tfmt = {0: None, 1: PX, 2: PX, 3: "0.00"}[k]
            value_cell(ws, f"K{r}", None, tfmt, bold=(k in (0, 2)))

        # conditional formatting, small ranges only
        sig = f"$K${top}"
        for rng in (f"H{top}", f"K{top}"):
            for fill, txt, formula in (
                (C_RED,   C_RED_T,   f'=OR({sig}="NO USABLE Z",{sig}="STALE FEED")'),
                (C_GREY,  C_GREY_T,  f'={sig}="WARMING UP"'),
                (C_AMBER, C_AMBER_T, f'={sig}="ABOVE CEILING"'),
                (C_GREEN, C_GREEN_T, f'={sig}="ENTRY BAND"'),
            ):
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[formula], stopIfTrue=False,
                    fill=PatternFill("solid", bgColor=fill),
                    font=Font(name=FONT, size=10, bold=True, color=txt)))
        # Take-profit reachability. The meaningful test is not a fixed number of
        # sigma - a trade entered at 2.5 sigma can legitimately travel 2.4 back
        # to the mean. It is whether the target sits PAST the mean, which needs
        # the spread to overshoot: a different bet from the one z measured.
        ws.conditional_formatting.add(f"K{top+3}", FormulaRule(
            formula=[f'=AND(ISNUMBER($K${top+3}),ISNUMBER($H${top}),$K${top+3}>ABS($H${top}))'],
            stopIfTrue=False,
            fill=PatternFill("solid", bgColor=C_AMBER),
            font=Font(name=FONT, size=10, bold=True, color=C_AMBER_T)))

    ws.conditional_formatting.add("C3", FormulaRule(
        formula=['=C3="RUNNING"'], stopIfTrue=False,
        fill=PatternFill("solid", bgColor=C_GREEN),
        font=Font(name=FONT, size=10, bold=True, color=C_GREEN_T)))

    # ---- take-profit rule, stated where it is used ------------------------
    block_title(ws, "B27", "TAKE-PROFIT RULE")
    label(ws, "B28", "Take Profit  =  Entry  +  Costs  +  (WIN_PCT x Notional)", bold=True)
    ws["B28"].font = f(10, bold=True, color=C_TITLE)
    for r, txt in (
        (29, "Signed by direction: a short-spread trade takes profit BELOW entry, a long-spread trade above it."),
        (30, "Costs = commission + crossing, from the Config rate card. Both are shown in full on the Detail sheet."),
        (31, "WIN_PCT and NOTIONAL_BASIS are Config cells. NOTIONAL_BASIS decides what the percentage is taken OF."),
        (32, "TP in sigma is the reachability check: how far the target sits from entry, in standard deviations. It "
             "turns amber when that exceeds the live |z| - meaning the target is PAST the mean, so it needs the "
             "spread to overshoot rather than merely revert. That is a different bet from the one z measured."),
    ):
        ws[f"B{r}"] = txt
        ws[f"B{r}"].font = f(9, italic=True, color=C_NOTE)

    NOTES = [
        (34, "READ THIS BEFORE TRADING FROM IT", True, C_WARN),
        (35, "z is for ENTRIES only. Exits act on MONEY, never on z. The TAKE PROFIT cell, not the z, tells you when to get out.", False, "000000"),
        (36, "z WILL STEP at each stats refresh. That is expected, not a fault - mean and sigma are recomputed every STATS_REFRESH_MIN minutes and held frozen in between.", False, "000000"),
        (37, "Freezing the reference is a feature. A continuously-updating mean chases the spread, so an open position's z drifts back toward zero without the price ever paying you.", False, "000000"),
        (38, "A blank z means the window is not usable yet - warming up, degenerate, or the feed is stale. A blank is honest; a number would not be.", False, "000000"),
        (39, "The chime fires on the CROSSING of ENTRY_Z, once, not continuously while the condition holds.", False, "000000"),
        (40, "SIGMA ON THESE SPREADS HAS NEVER BEEN MEASURED. Every cost and take-profit threshold is contingent on it. Measuring it is a main purpose of this tool.", True, C_WARN),
        (41, "Brent/WTI is a RELATED pair, not a basis pair - two different crudes, no arbitrage tying them, so no theoretical fair value. The only anchor is the spread's own empirical mean.", False, "000000"),
        (42, "NO CELL ON THIS SHEET IS A FORMULA. Every value is written by VBA only when it differs from what is already there - that is what stops the display flashing on every uptick.", True, C_TITLE),
        (43, "All RTD() formulas live on the hidden Feed sheet. There is no volatile function anywhere in this workbook.", False, C_NOTE),
    ]
    for row, txt, bold, color in NOTES:
        ws[f"B{row}"] = txt
        ws[f"B{row}"].font = f(9, bold=bold, italic=not bold, color=color)
    return ws
