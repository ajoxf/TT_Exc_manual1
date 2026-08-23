"""Dashboard: existing layout preserved, z-monitor block added below.

Cell coordinates fixed by the spec and honoured exactly:
  C4  market ("CME")          B6..B9  TT short names       C6..C9  instrument ids
  H7  HO bid, H8 CL bid       -> HO/CL crack = (H7*42)-H8
  P7  BZ mid, P8 CL mid       -> BZ - CL   = P7-P8
  H12 RB bid, I12 RB ask      H13 HO bid   H14 CL bid   I14 CL ask
  H15/I15 3:2:1 crack bid/ask -> H15 = ((H12*2*42)+(H13*1*42)-(I14*3))/3
                                 I15 = ((I12*2*42)+(I13*1*42)-(H14*3))/3   <- mirrored
  L3  clock  (was =TEXT(NOW(),"hh:mm:ss") - volatile, removed; VBA writes it)

NOTHING on this sheet is a formula. Every cell is written by VBA, compare-first.
"""
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from wb_common import *

PX = "#,##0.0000"
P2 = "#,##0.00"
USD = '$#,##0.00;($#,##0.00);-'
Z2 = "0.00"

LEGS = [
    (6,  "RBV6", "RBOB Gasoline Oct26",       "42,000 gal"),
    (7,  "HOV6", "NY Harbor ULSD Oct26",      "42,000 gal"),
    (8,  "CLV6", "WTI Crude Oct26",           "1,000 bbl"),
    (9,  "BZV6", "Brent Last Day Fin Oct26",  "1,000 bbl"),
]

CRACK = [
    (12, "RBV6", "RBOB Gasoline Oct26", 2, 42),
    (13, "HOV6", "NY Harbor ULSD Oct26", 1, 42),
    (14, "CLV6", "WTI Crude Oct26",      3, 1),
]

DERIVED = [
    (19, "BZ - CL", "legged",        "bid = BZ bid - CL ask ; ask = BZ ask - CL bid"),
    (20, "HO/CL crack", "legged",    "bid = HO bid x42 - CL ask ; ask = HO ask x42 - CL bid"),
    (21, "3:2:1 crack", "legged",    "bid sells products / buys crude ; ask mirrors it"),
    (22, "CL Oct26 - BZ Oct26 Inter-Product", "listed", "exchange-listed - one quote, no leg risk"),
    (23, "Oct26 HO-CL Crack", "listed",       "exchange-listed - one quote, no leg risk"),
]

# z-monitor metric rows: (row, label, unit/note, kind)
METRICS = [
    (27, "Spread",                      "",                    "text"),
    (28, "Definition",                  "LegB - beta x LegA",  "text"),
    (29, "Mode",                        "",                    "text"),
    (30, "HEDGE_RATIO (beta)",          "stamped pair ->",     "text"),
    (31, "USD per 1.00 of spread",      "per lot",             "usd0"),
    (32, "SECT:MARKET  (touch, shown separately from the statistic)", "", "sect"),
    (33, "Spread bid",                  "",                    "px"),
    (34, "Spread ask",                  "",                    "px"),
    (35, "Spread mid   <- feeds the statistic", "",            "px"),
    (36, "Market width",                "spread units",        "px"),
    (37, "SECT:WINDOW", "", "sect"),
    (38, "Stored samples in window",    "after dedupe + store gate", "int"),
    (39, "Window elapsed",              "min",                 "num2"),
    (40, "Warm-up gate",                "binding gate",        "text"),
    (41, "Feed rate",                   "quote CHANGES / min", "num1"),
    (42, "Feed status",                 "",                    "text"),
    (43, "SECT:STATISTICS  (FROZEN - republished every STATS_REFRESH_MIN)", "", "sect"),
    (44, "Mean (frozen)",               "",                    "px"),
    (45, "Sigma (frozen)",              "",                    "px"),
    (46, "Sigma in dollars",            "sigma x USD/pt x LOTS", "usd"),
    (47, "Stats age / next refresh",    "min",                 "text"),
    (48, "Z  (live, vs frozen mean & sigma)", "",              "z"),
    (49, "SIGNAL",                      "",                    "text"),
    (50, "SECT:COST OF THE ROUND TRIP", "", "sect"),
    (51, "Direction",                   "",                    "text"),
    (52, "Entry reference (touch)",     "",                    "px"),
    (53, "Commission",                  "contracts x rate x LOTS", "usd"),
    (54, "Crossing cost - LEGGING",     "",                    "usd"),
    (55, "Crossing cost - LISTED",      "",                    "usd"),
    (56, "TOTAL COST (this mode)",      "",                    "usd"),
    (57, "SECT:TAKE PROFIT  (exits act on MONEY, never on z)", "", "sect"),
    (58, "Break-even spread",           "absolute level",      "px"),
    (59, "TAKE-PROFIT spread",          "absolute level",      "px"),
    (60, "TP distance",                 "spread units",        "px"),
    (61, "TP in sigma",                 "x sigma",             "num2"),
    (62, "Target z",                    "z at the TP level",   "z"),
    (63, "TP beyond the mean?",         "",                    "text"),
    (64, "SECT:EDGE FILTER", "", "sect"),
    (65, "Expected capture",            "0.5 x |z| x sigma x USD/pt x LOTS", "usd"),
    (66, "Required",                    "EDGE_MULTIPLE x total cost", "usd"),
    (67, "EDGE VERDICT",                "",                    "text"),
    (68, "Min sigma to pass at live z", "spread units",        "px"),
]

SPREAD_COLS = ["D", "E", "F", "G"]


def build(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 32, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18,
                    "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12,
                    "N": 17, "O": 12, "P": 12, "Q": 12})
    ws.freeze_panes = "B5"

    # ---- header -----------------------------------------------------------
    title(ws, "B1", "TT RTD Z-SCORE MONITOR   -   Energy Futures Spreads")
    ws["B2"] = ("z is for ENTRIES only.  Exits act on MONEY, never on z.   "
                "All execution is manual in TT - this workbook never places an order.")
    ws["B2"].font = f(10, bold=True, color=C_WARN)

    for cell, txt, fill in (("N2", "  >  START  ", "C6EFCE"), ("P2", "  #  STOP  ", "FFC7CE")):
        c = ws[cell]; c.value = txt
        c.font = f(10, bold=True); c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center"); c.border = BOX
    note(ws, "N1", "Assign macros: StartMonitor / StopMonitor")

    label(ws, "B3", "Monitor", bold=True)
    value_cell(ws, "C3", "STOPPED", bold=True, fill=C_GREY)
    label(ws, "E3", "Rates", bold=True)
    ws.merge_cells("F3:J3")
    value_cell(ws, "F3", "-", align="left")
    label(ws, "K3", "Clock", bold=True)
    value_cell(ws, "L3", "--:--:--", bold=True)          # was =TEXT(NOW(),"hh:mm:ss")
    note(ws, "M3", "<- VBA writes this. NOW() was volatile and repainted the whole book on every RTD tick.")

    label(ws, "B4", "Market", bold=True)
    input_cell(ws, "C4", "CME")                          # $C$4 in the RTD Inst call
    label(ws, "E4", "Session", bold=True)
    value_cell(ws, "F4", "-", align="left")
    label(ws, "K4", "Last tick", bold=True)
    value_cell(ws, "L4", "-")

    # ---- outright legs ----------------------------------------------------
    header_row(ws, 5, {"B": "Symbol", "C": "Instrument ID", "D": "Description", "E": "Contract",
                       "F": "Tick", "G": "$/tick", "H": "Bid", "I": "Ask", "J": "High",
                       "K": "Low", "L": "Mid", "M": "Width"})
    for row, sym, desc, contract in LEGS:
        value_cell(ws, f"B{row}", sym, bold=True, align="left")
        value_cell(ws, f"C{row}", "", align="left")      # RTD("Inst") result, written by VBA
        ws[f"C{row}"].font = f(9, color=C_LINK)
        value_cell(ws, f"D{row}", desc, align="left"); ws[f"D{row}"].font = f(9)
        value_cell(ws, f"E{row}", contract); ws[f"E{row}"].font = f(9, color=C_NOTE)
        value_cell(ws, f"F{row}", 0.0001 if row in (6, 7) else 0.01, PX)
        value_cell(ws, f"G{row}", 4.20 if row in (6, 7) else 10.00, USD)
        for col, fmt in (("H", PX), ("I", PX), ("J", PX), ("K", PX), ("L", PX), ("M", PX)):
            value_cell(ws, f"{col}{row}", None, fmt)

    # ---- reference pair block (P7 / P8) -----------------------------------
    block_title(ws, "N5", "BRENT / WTI  (mid)")
    header_row(ws, 6, {"N": "Pair", "O": "Symbol", "P": "Mid", "Q": "Width"})
    value_cell(ws, "N7", "LegB", align="left"); ws["N7"].font = f(9, color=C_NOTE)
    value_cell(ws, "O7", "BZV6", bold=True); value_cell(ws, "P7", None, PX); value_cell(ws, "Q7", None, PX)
    value_cell(ws, "N8", "LegA", align="left"); ws["N8"].font = f(9, color=C_NOTE)
    value_cell(ws, "O8", "CLV6", bold=True); value_cell(ws, "P8", None, PX); value_cell(ws, "Q8", None, PX)
    value_cell(ws, "N9", "BZ - CL", bold=True, align="left")
    value_cell(ws, "O9", "P7 - P8"); ws["O9"].font = f(9, color=C_NOTE)
    value_cell(ws, "P9", None, PX, bold=True)

    # ---- 3:2:1 crack build ------------------------------------------------
    block_title(ws, "B10", "3:2:1 CRACK BUILD   ($/bbl)")
    header_row(ws, 11, {"B": "Leg", "C": "Instrument ID", "D": "Description", "E": "",
                        "F": "Weight", "G": "bbl conv", "H": "Bid", "I": "Ask", "J": "High",
                        "K": "Low", "L": "Mid"})
    for row, sym, desc, w, conv in CRACK:
        value_cell(ws, f"B{row}", sym, bold=True, align="left")
        value_cell(ws, f"C{row}", "", align="left"); ws[f"C{row}"].font = f(9, color=C_LINK)
        value_cell(ws, f"D{row}", desc, align="left"); ws[f"D{row}"].font = f(9)
        value_cell(ws, f"F{row}", w, "0")
        value_cell(ws, f"G{row}", conv, "0")
        for col in "HIJKL":
            value_cell(ws, f"{col}{row}", None, PX)
    value_cell(ws, "B15", "3:2:1 Crack", bold=True, align="left")
    value_cell(ws, "C15", "(2RB + 1HO)x42/3 - CL", align="left"); ws["C15"].font = f(9, color=C_NOTE)
    for col in ("H", "I", "L"):
        value_cell(ws, f"{col}15", None, PX, bold=True)
    note(ws, "N12", "CONVENTION (confirm):", color=C_WARN, bold=True, italic=False)
    note(ws, "N13", "H15 bid  = (2xH12x42 + 1xH13x42 - 3xI14)/3   sells products at the bid, buys crude at the ask")
    note(ws, "N14", "I15 ask  = (2xI12x42 + 1xI13x42 - 3xH14)/3   mirrors it - uses H14 (CL BID), not I14")
    note(ws, "N15", "The old sheet had I15 referencing I14. That double-counted the crude ask and understated the "
                    "quoted crack width.", color=C_WARN)

    # ---- derived + listed spreads ----------------------------------------
    block_title(ws, "B17", "DERIVED & LISTED SPREADS")
    header_row(ws, 18, {"B": "Spread", "C": "Basis", "D": "Crossing convention", "E": "", "F": "", "G": "",
                        "H": "Bid", "I": "Ask", "J": "", "K": "", "L": "Mid", "M": "Width",
                        "N": "Width $/lot"})
    for row, name, basis, conv in DERIVED:
        value_cell(ws, f"B{row}", name, bold=True, align="left")
        value_cell(ws, f"C{row}", basis); ws[f"C{row}"].font = f(9, color=C_NOTE)
        value_cell(ws, f"D{row}", conv, align="left"); ws[f"D{row}"].font = f(9, color=C_NOTE)
        for col in ("H", "I", "L", "M"):
            value_cell(ws, f"{col}{row}", None, PX)
        value_cell(ws, f"N{row}", None, USD)
    note(ws, "B24", "Legging crosses two markets; the listed inter-product spread crosses one, at roughly half the "
                    "cost and with no leg risk. Prefer it where liquidity allows.")

    # ---- z-score monitor block -------------------------------------------
    block_title(ws, "B26", "Z-SCORE MONITOR")
    hdr = {"B": "Metric", "C": "Unit / note"}
    for i, col in enumerate(SPREAD_COLS, start=1):
        hdr[col] = f"SPREAD {i}"
    header_row(ws, 26, hdr)

    fmt_map = {"px": PX, "usd": USD, "usd0": "#,##0", "num1": "0.0", "num2": "0.00",
               "int": "#,##0", "z": Z2, "text": None}

    for row, lab, unit, kind in METRICS:
        if kind == "sect":
            sect(ws, row, lab.replace("SECT:", ""), first="B", last="G")
            continue
        label(ws, f"B{row}", lab, bold=(lab in ("SIGNAL", "EDGE VERDICT", "TOTAL COST (this mode)",
                                                "TAKE-PROFIT spread")))
        ws[f"B{row}"].border = BOX
        value_cell(ws, f"C{row}", unit, align="left"); ws[f"C{row}"].font = f(9, color=C_NOTE)
        for col in SPREAD_COLS:
            c = value_cell(ws, f"{col}{row}", None, fmt_map[kind],
                           bold=(lab in ("SIGNAL", "EDGE VERDICT", "Z  (live, vs frozen mean & sigma)",
                                         "TAKE-PROFIT spread", "TOTAL COST (this mode)")))
            if kind == "text":
                c.font = f(10, bold=c.font.bold)
                c.alignment = Alignment(horizontal="center", wrap_text=False)

    # ---- conditional formatting: SMALL RANGES ONLY ------------------------
    def rule(fill, text_color, formula):
        return FormulaRule(formula=[formula], stopIfTrue=False,
                           fill=PatternFill("solid", bgColor=fill),
                           font=Font(name=FONT, size=10, bold=True, color=text_color))

    zr = "D48:G49"
    for r_ in (
        rule(C_RED,   C_RED_T,   '=OR(D$49="NO USABLE Z",D$49="STALE FEED")'),
        rule(C_GREY,  C_GREY_T,  '=D$49="WARMING UP"'),
        rule(C_AMBER, C_AMBER_T, '=D$49="ABOVE CEILING"'),
        rule(C_GREEN, C_GREEN_T, '=AND(D$49="ENTRY BAND",D$67="PASS")'),
        rule(C_AMBER, C_AMBER_T, '=AND(D$49="ENTRY BAND",D$67<>"PASS")'),
    ):
        ws.conditional_formatting.add(zr, r_)

    ws.conditional_formatting.add("D67:G67", rule(C_GREEN, C_GREEN_T, '=D67="PASS"'))
    ws.conditional_formatting.add("D67:G67", rule(C_AMBER, C_AMBER_T, '=D67="FAIL"'))
    ws.conditional_formatting.add("D42:G42", rule(C_RED,   C_RED_T,   '=D42="STALE"'))
    ws.conditional_formatting.add("D42:G42", rule(C_AMBER, C_AMBER_T, '=D42="THIN"'))
    ws.conditional_formatting.add("D42:G42", rule(C_GREEN, C_GREEN_T, '=D42="OK"'))
    ws.conditional_formatting.add("D40:G40", rule(C_GREY,  C_GREY_T,  '=D40<>"ready"'))
    ws.conditional_formatting.add("D63:G63", rule(C_AMBER, C_AMBER_T, '=LEFT(D63,3)="YES"'))
    ws.conditional_formatting.add("C3",      rule(C_GREEN, C_GREEN_T, '=C3="RUNNING"'))

    # ---- the things that must be stated plainly on the sheet --------------
    NOTES = [
        (70, "READ THIS BEFORE TRADING FROM IT", True, C_WARN),
        (71, "z is for ENTRIES only.  Exits act on MONEY, never on z.  The take-profit block, not the z, tells you when to get out.", False, "000000"),
        (72, "z WILL STEP at each stats refresh. That is expected, not a fault - mean and sigma are recomputed every STATS_REFRESH_MIN minutes and held frozen in between.", False, "000000"),
        (73, "Freezing the reference is a feature, not just a performance trick. A continuously-updating mean chases the spread, so an open position's z drifts back toward zero without the price ever paying you.", False, "000000"),
        (74, "Ticks are stored continuously and the window still slides continuously. Only the PUBLISHED mean and sigma are frozen.", False, "000000"),
        (75, "A highlighted z that fails the edge filter is NOT a trade. Read the EDGE VERDICT row before acting on the highlight.", False, C_WARN),
        (76, "The entry ceiling is its own rule: highlight only when ENTRY_Z <= |z| < MAX_ENTRY_Z. A z of 5+ is a momentum spike, not a reversion setup.", False, "000000"),
        (77, "A thin feed collapses sigma. If the feed rate is amber, the sigma you are trusting is thin.", False, "000000"),
        (78, "SIGMA ON THESE SPREADS HAS NEVER BEEN MEASURED. Every cost and edge threshold is contingent on it. Measuring it is a main purpose of this tool.", True, C_WARN),
        (79, "Brent/WTI is a RELATED pair, not a basis pair - two different crudes, no arbitrage tying them, so no theoretical fair value. The only anchor is the spread's own empirical mean. No carry or fair-value term is added anywhere.", False, "000000"),
        (79.5, "FOUR DECOUPLED RATES. CAPTURE_MS polls for a changed quote (buys latency to the chime, not a better sigma). STORE_MIN_INTERVAL_MS decides what reaches the window. PRICE_REFRESH_MS decides what repaints. STATS_REFRESH_MIN decides when mean and sigma are republished.", True, "000000"),
        (80, "Consequence of ENTRY_Z = 2.5 rather than 3.0: less expected capture per entry, so it RAISES the sigma the spread must have to be worth trading - by about 20%. See the Setup sheet for the table.", False, "000000"),
    ]
    NOTES2 = []
    for row, txt, bold, color in NOTES:
        NOTES2.append((int(row) if float(row).is_integer() else row, txt, bold, color))
    # renumber so the inserted line does not collide
    NOTES2 = [(70 + i, t, b, c) for i, (_, t, b, c) in enumerate(NOTES)]
    for row, txt, bold, color in NOTES2:
        ws[f"B{row}"] = txt
        ws[f"B{row}"].font = f(9, bold=bold, italic=not bold, color=color)

    ws["B83"] = ("NO CELL ON THIS SHEET IS A FORMULA. Every value is written by VBA only when it differs from what is "
                 "already there - that is what stops the display flashing on every uptick.")
    ws["B83"].font = f(9, bold=True, color=C_TITLE)
    ws["B84"] = "All RTD() formulas live on the hidden Feed sheet. There is no volatile function anywhere in this workbook."
    ws["B84"].font = f(9, italic=True, color=C_NOTE)
    return ws
