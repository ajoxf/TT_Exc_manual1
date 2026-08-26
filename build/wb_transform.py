"""Transform the ORIGINAL TTDashboard workbook in place.

The Dashboard's layout, fonts, fills, row heights, merges and column widths are
the user's own and are left exactly as they are. Only three things change:

  1. Every RTD() formula moves to the hidden Feed sheet. The visible cells keep
     their styling and become VBA-written values. A cell holding a live RTD()
     formula repaints whenever the feed moves, and that cannot be prevented
     while the formula is in the cell.
  2. =TEXT(NOW(),"hh:mm:ss") in L3 goes. NOW() is volatile, so it forces a full
     recalculation of the whole workbook on every RTD update.
  3. I15 is corrected to reference H14 (CL bid) rather than I14, so the 3:2:1
     ask mirrors the bid.

Then the z-score block is ADDED below, in the same visual style.

Original cell map, preserved:
  C2   title (merged C2:T2)          L3  clock
  B4/C4  TT MARKET / CME             B5/C5 headers
  B6..B12 TT short names, C6..C12 instrument ids
     B6 RBV6   B7 HOV6   B8 BZV6   B9 CLV6
     B10 CL Oct26 - BZ Oct26 Inter-Product   B12 Oct26 HO-CL Crack
  F..L  block: Symbol|Contract|Bid|Ask|Gap|High|Low
     rows 7 HO, 8 CL, 9 HO|CL crack       rows 12 RB, 13 HO, 14 CL, 15 3:2:1
  N..T  block: rows 7 BZ, 8 CL, 9 BZ - CL
"""
from copy import copy
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule

# every formula cell on the original Dashboard - cleared, styling kept
FORMULA_CELLS = (
    ["L3"]
    + [f"C{r}" for r in (6, 7, 8, 9, 10, 12)]
    + [f"{c}{r}" for r in (7, 8, 9) for c in "HIJKL"]
    + [f"{c}{r}" for r in (7, 8, 9) for c in "PQRST"]
    + [f"{c}{r}" for r in (12, 13, 14, 15) for c in "HIJKL"]
)

# the z-score block, added below the existing content
Z_TITLE_ROW, Z_HDR_ROW, Z_TOP = 17, 18, 19
CTRL_ROW = 23
# Columns A, B, C, G and O are HIDDEN in the original: A/B/C are the TT
# plumbing (short names and instrument ids) and G/O hold the contract name
# that also appears as the second wrapped line of F/N. So the added block uses
# the same visible 6 + 6 grid the existing blocks use.
Z_ROWS = [
    (19, "HO | CL Crack"),
    (20, "BZ - CL"),
    (21, "3:2:1"),
]
# Bid sits second, beside the spread name: it is the price you sell at, and
# the take-profit is measured from it. Direction is adjacent, in plain words.
Z_LEFT = [("F", "Spread"), ("H", "Bid"), ("I", "Direction"), ("J", "Z-score"),
          ("K", "Mean"), ("L", "Std Dev")]
Z_RIGHT = [("N", "Signal"), ("P", "Window"), ("Q", "Costs $"), ("R", "Win $"),
           ("S", "TAKE PROFIT"), ("T", "TP in sigma")]


def _restyle(ws, dst, src, size=None, bold=None, color=None, numfmt=None):
    """Copy a source cell's full style, then nudge the font/format."""
    d, s = ws[dst], ws[src]
    d._style = copy(s._style)
    fnt = copy(s.font)
    d.font = Font(name=fnt.name,
                  size=size if size is not None else fnt.size,
                  bold=bold if bold is not None else fnt.bold,
                  italic=fnt.italic,
                  color=color if color is not None else fnt.color)
    if numfmt is not None:
        d.number_format = numfmt
    return d


def transform(wb):
    ws = wb["Dashboard"]

    # --- 1 + 2: strip every formula, keep the styling -------------------
    for addr in FORMULA_CELLS:
        ws[addr].value = None

    # --- the Gap cell the original never had ----------------------------
    # J9 and R9 exist; J15 does not, so the 3:2:1 row had no Gap at all.
    _restyle(ws, "J15", "J14")

    # --- 3: the z-score block, in the same visual style -----------------
    ws.row_dimensions[Z_TITLE_ROW].height = 44.25
    ws.row_dimensions[Z_HDR_ROW].height = 23.25
    for r, _ in Z_ROWS:
        ws.row_dimensions[r].height = 50.25

    # One title band across the block, styled like the sheet's own C2 banner.
    # A separate note cell beside it would clip the title, because F is only
    # 16.1 wide and text cannot overflow into an occupied cell.
    ws.merge_cells(f"F{Z_TITLE_ROW}:T{Z_TITLE_ROW}")
    t = _restyle(ws, f"F{Z_TITLE_ROW}", "C2", size=13, bold=True)
    t.value = ("Z-SCORE MONITOR      z is for ENTRIES only - exits act on money, never on z."
               "      The chime fires on the CROSSING of ENTRY_Z, once.")
    t.alignment = Alignment(horizontal="center", vertical="center")

    for col, txt in Z_LEFT + Z_RIGHT:
        h = _restyle(ws, f"{col}{Z_HDR_ROW}", "H6", size=11, bold=True)
        h.value = txt

    for r, name in Z_ROWS:
        lab = _restyle(ws, f"F{r}", "F12", size=16, bold=True)
        lab.value = name
        # Bid, z and the take profit get the large treatment; the rest support them
        _restyle(ws, f"H{r}", "H12", size=22, numfmt="0.0000")     # Bid
        _restyle(ws, f"I{r}", "H12", size=11, numfmt="General")    # Direction
        _restyle(ws, f"J{r}", "H12", size=22, numfmt="0.00")       # Z-score
        for col in ("K", "L"):                                     # Mean, Std Dev
            _restyle(ws, f"{col}{r}", "H12", size=16, numfmt="0.0000")
        for col in ("N", "P"):                                     # Signal, Window
            _restyle(ws, f"{col}{r}", "H12", size=11, numfmt="General")
        for col in ("Q", "R"):                                     # Costs, Win
            _restyle(ws, f"{col}{r}", "H12", size=14, numfmt="$#,##0.00")
        _restyle(ws, f"S{r}", "H12", size=22, numfmt="0.0000")     # TAKE PROFIT
        _restyle(ws, f"T{r}", "H12", size=14, numfmt="0.00")       # TP in sigma

    # bottom border on the last row, matching the blocks above
    for col, _ in Z_LEFT + Z_RIGHT:
        c = ws[f"{col}{Z_ROWS[-1][0]}"]
        c.border = copy(ws["H15"].border)

    # --- conditional formatting, SMALL ranges only ----------------------
    def rule(fill, txt, formula):
        return FormulaRule(formula=[formula], stopIfTrue=False,
                           fill=PatternFill("solid", bgColor=fill),
                           font=Font(name="Arial", bold=True, color=txt))

    for r, _ in Z_ROWS:
        sig = f"$N${r}"
        for rng in (f"J{r}", f"N{r}"):
            for fill, txt, cond in (
                ("FFC7CE", "9C0006", f'=OR({sig}="NO USABLE Z",{sig}="STALE FEED")'),
                ("E7E6E6", "595959", f'={sig}="WARMING UP"'),
                ("FFEB9C", "9C5700", f'={sig}="ABOVE CEILING"'),
                ("C6EFCE", "006100", f'={sig}="ENTRY BAND"'),
            ):
                ws.conditional_formatting.add(rng, rule(fill, txt, cond))
        # TP beyond the mean needs an overshoot, not a reversion
        ws.conditional_formatting.add(f"T{r}", rule(
            "FFEB9C", "9C5700",
            f'=AND(ISNUMBER($T${r}),ISNUMBER($J${r}),$T${r}>ABS($J${r}))'))

    # --- control strip -------------------------------------------------
    # Rows 4 and 5 are 6.75 tall - spacer rows - so anything placed there would
    # be clipped. The strip goes below the added block instead, leaving rows
    # 1-15 completely untouched.
    ws.row_dimensions[CTRL_ROW].height = 36.0
    st = _restyle(ws, f"F{CTRL_ROW}", "B4", size=12, bold=True, color="FF111318")
    st.value = "STOPPED"
    st.fill = PatternFill("solid", fgColor="E7E6E6")
    rt = _restyle(ws, f"H{CTRL_ROW}", "G7", size=10, bold=False)
    rt.value = "-"
    rt.alignment = Alignment(horizontal="left", vertical="center")
    for addr, txt, fill in ((f"N{CTRL_ROW}", "  >  START  ", "C6EFCE"),
                            (f"P{CTRL_ROW}", "  #  STOP  ", "FFC7CE"),
                            (f"R{CTRL_ROW}", "  TEST CHIME  ", "D9E1F2")):
        c = _restyle(ws, addr, "B4", size=12, bold=True, color="FF111318")
        c.value = txt
        c.fill = PatternFill("solid", fgColor=fill)
    ws.conditional_formatting.add(f"F{CTRL_ROW}",
                                  rule("C6EFCE", "006100", f'=F{CTRL_ROW}="RUNNING"'))
    return ws
