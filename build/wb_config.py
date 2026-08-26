"""Config sheet: every tunable parameter in its own labelled cell."""
from openpyxl.styles import Font, PatternFill, Alignment
from wb_common import *

# name, value, default, units, meaning
ENGINE = [
    ("CAPTURE_MS",             100,   100,  "ms",    "RATE 1 of 4. How often the feed is POLLED for a changed quote. Drives latency to the chime, NOT sigma - after dedupe the window holds quote changes, not poll ticks. Sub-second needs the user32 timer; at 1000+ the module falls back to Application.OnTime."),
    ("STORE_MIN_INTERVAL_MS",    0,     0,  "ms",    "RATE 2 of 4. Minimum spacing between STORED samples. 0 = full fidelity. Leave it at 0 for the first week or two to characterise the microstructure, then set 250-1000: subsampling a mean-reverting level series is unbiased, so it is 3-10x less memory for the same mean and sigma."),
    ("PRICE_REFRESH_MS",       500,   500,  "ms",    "RATE 3 of 4. How often the display may repaint. Independent of capture and of storage - a number showing more digits than the market moves in will appear to change constantly."),
    ("CONFIG_REFRESH_MS",     1000,  1000,  "ms",    "How often this sheet is re-read. A value changed here takes effect within this long - no restart, no VBA edit."),
    ("RTD_THROTTLE_MS",        100,   100,  "ms",    "Application.RTD.ThrottleInterval set on start. Match it to CAPTURE_MS. Excel's default is 2000."),
    ("WATCHDOG_SEC",             5,     5,  "s",     "Application.OnTime watchdog. Re-arms the high-resolution timer if Excel killed it (a modal dialog will) and picks up a changed CAPTURE_MS."),
    ("STALE_SEC",               15,    15,  "s",     "No new quote id for this long => feed marked STALE and alerts go silent."),
    ("FLUSH_SEC",               60,    60,  "s",     "How often the recovery snapshot is written to the hidden Buffer sheet."),
    ("FLUSH_DECIMATE_MS",     1000,  1000,  "ms",    "The recovery snapshot is decimated to this spacing. Unbiased for mean and sigma, and it keeps a 120-minute window near 7,200 rows instead of tens of thousands."),
    ("BUFFER_CAPACITY",      80000, 80000,  "samples","Circular-buffer capacity per spread. 4 spreads x 80,000 x 16 bytes is about 5 MB of VBA array, held in memory, not in cells."),
    ("TICK_ARCHIVE_ENABLED",  True,  True,  "bool",  "Write every CHANGED quote to a CSV on disk. This is the one chance to characterise the spread's microstructure and to verify that coarser storage loses nothing real. Turn it off once STORE_MIN_INTERVAL_MS is set."),
    ("TICK_ARCHIVE_PATH",       "",    "",  "folder","Folder for the tick CSVs. Blank = a 'ticks' folder beside the workbook. One file per spread per day."),
    ("TICK_ARCHIVE_BATCH",     500,   500,  "lines", "Lines buffered in memory before a file write, so archiving never stalls the capture timer."),
]

WINDOW = [
    ("LOOKBACK_MIN",            120,   120,  "min",   "Rolling window width (90 or 120). Changing it re-evaluates stored history on the next stats refresh and re-applies the warm-up gates."),
    ("STATS_REFRESH_MIN",         5,     5,  "min",   "How often mean/sigma are recomputed and re-published. Held FROZEN in between; live z is measured against the frozen values."),
    ("MIN_SAMPLES",             300,   300,  "quotes","Warm-up gate 1. Quotes needed before any z is shown."),
    ("MIN_HISTORY_MIN",         120,   120,  "min",   "Warm-up gate 2. Window span needed before any z is shown. BOTH gates must pass. MUST BE LESS THAN LOOKBACK_MIN: eviction drops samples older than the lookback, so the span approaches it but never reaches it, and a value at or above LOOKBACK_MIN would never be satisfied. Anything above 98% of LOOKBACK_MIN is clamped to that, and the clamp is noted once in the Log."),
    ("MIN_SIGMA",                 0,     0,  "spread","Sigma floor. Below this the window is degenerate and NO z is shown. Set once sigma has actually been measured."),
    ("MAX_ABS_Z",                25,    25,  "z",     "Absurd-z guard. Above this, show 'no usable z' rather than a number."),
]

SIGNAL = [
    ("ENTRY_Z",                 2.5,   2.5,  "z",     "Highlight and chime at or above this |z|. 2.0 and 2.5 are both reasonable; 2.0 fires more often and needs a larger sigma to stay worth trading."),
    ("MAX_ENTRY_Z",             4.5,   4.5,  "z",     "Entry CEILING. At or above this, do NOT highlight - that is a momentum spike, not a reversion setup. Keep the band at least 1 sigma wide."),
    ("THIN_FEED_QPM",             6,     6,  "q/min", "Feed rate at or below which the quotes/min display turns amber."),
    ("SPREAD_QUOTE_CONVENTION", "CROSSING", "CROSSING", "text", "How the spread rows on Dashboard quote Bid and Ask. CROSSING: sell LegB at the bid and buy LegA at the ask, mirrored for the ask - the Gap is then the true cost of crossing and is always positive. SAME_SIDE: bid-minus-bid and ask-minus-ask, as the original sheet's rows 9 computed. Under SAME_SIDE the Gap is a DIFFERENCE of gaps, which understates the crossing cost and can go negative."),
]

COSTS = [
    ("LOTS",                      1,     1,  "lots",  "Contracts per leg, for the cost and take-profit maths."),
    ("TP_MODE",     "PCT_NOTIONAL", "PCT_NOTIONAL", "text", "PCT_NOTIONAL: Take Profit = Entry + Costs + (WIN_PCT x Notional). TARGET_USD: the older rule, Take Profit = Entry + Costs + TARGET_NET_USD."),
    ("WIN_PCT",                0.01,  0.01,  "fraction", "Percentage win, as a FRACTION - 0.01 is 1%, 0.02 is 2%. Used only when TP_MODE = PCT_NOTIONAL."),
    ("NOTIONAL_BASIS", "SPREAD_VALUE", "SPREAD_VALUE", "text", "What WIN_PCT is taken OF. SPREAD_VALUE = |spread| x USD/pt x LOTS. ONE_LEG = the crude leg's full contract value. BOTH_LEGS = both legs' contract value added. These differ by roughly 10x - read the note below the table before changing it."),
    ("TP_ENTRY_BASIS",        "BID", "BID",  "text",  "Which price the take-profit is measured from. BID: always the bid, the price you sell at - exact for a High-to-Low trade, optimistic by the spread's own gap on a Low-to-High one. TOUCH: prices each direction at the side it really trades on (bid when selling, ask when buying)."),
    ("TARGET_NET_USD",          150,   150,  "USD",   "Desired profit AFTER all costs. Used only when TP_MODE = TARGET_USD."),
    ("EDGE_MULTIPLE",           1.5,   1.5,  "x",     "Expected capture must clear the round trip by this multiple."),
    ("RATE_ORIENT_PER_SIDE",   0.35,  0.35,  "USD",   "Orient clearing + execution, per contract per side. Source: Orient rate card 2026-08-21, Ver.26.8.17."),
    ("RATE_FIX_PER_SIDE",      0.06,  0.06,  "USD",   "FIX / platform transaction fee, per contract per side. Same rate card."),
    ("RATE_CME_PER_SIDE",      1.50,  1.50,  "USD",   "CME exchange + clearing, per contract per side. Same rate card."),
]

ALERTS = [
    ("ALERT_ENABLED",          True,  True,  "bool",  "MASTER MUTE. FALSE silences every sound."),
    ("ALERT_SOUND_PATH",  r"C:\Windows\Media\chimes.wav", r"C:\Windows\Media\chimes.wav", "path", "Single short 'ting' played on an entry-band crossing."),
    ("ALERT_SOUND_PATH_CEILING", r"C:\Windows\Media\notify.wav", r"C:\Windows\Media\notify.wav", "path", "Distinct tone for crossing MAX_ENTRY_Z. That is 'stand down', not 'get in'."),
    ("ALERT_REARM_MARGIN",     0.25,  0.25,  "z",     "Hysteresis. Do not re-arm until |z| falls back below ENTRY_Z - this margin."),
    ("ALERT_COOLDOWN_SEC",       60,    60,  "s",     "Minimum seconds between alerts for the same spread."),
    ("ALERT_ONLY_IF_EDGE_PASSES", False, False, "bool", "FALSE = the chime follows ENTRY_Z alone. TRUE = also stay silent when the edge filter fails, on the argument that a sound for a trade you should not take trains you to ignore the sound. Ships FALSE by request; the edge verdict is still shown on Detail."),
    ("ALERT_SPEAK",           False, False,  "bool",  "TRUE also speaks the spread name (Application.Speech). Useful when watching four at once."),
]

PERSIST = [
    ("WARM_START_ENABLED",     True,  True,  "bool",  "On open, reload the flushed buffer so a restart does not cost another two hours of warm-up."),
    ("WARM_START_MAX_AGE_MIN",  180,   180,  "min",   "Only reload if the newest saved sample is younger than this."),
    ("LOG_ENABLED",            True,  True,  "bool",  "Write signal events to the Log sheet."),
    ("LOG_MAX_ROWS",          20000, 20000,  "rows",  "Log is trimmed to this many rows."),
]

SPREADS = [
    # slot, enabled, name, mode, hedge, beta stamp, usd/pt, contracts charged, tick, dp, listed equivalent
    (1, True,  "BZ - CL  (Brent LDF - WTI)",  "LEGGED", 1.0,
     "BZV6 / CLV6, derived 2026-08 (RELATED pair - no fair value, empirical mean only)",
     1000, 2, 0.01, 2, "CL Oct26 - BZ Oct26 Inter-Product"),
    (2, True,  "HO/CL crack  (ULSD - WTI)",   "LEGGED", 1.0,
     "HOV6 x42 / CLV6, 1:1 by design (crack, not a fitted beta)",
     1000, 2, 0.01, 2, "Oct26 HO-CL Crack"),
    (3, True,  "3:2:1 crack  (2RB+1HO-3CL)/3","LEGGED", 1.0,
     "RBV6/HOV6/CLV6, structural 3:2:1 - not a fitted beta",
     3000, 6, 0.01, 2, "(none listed)"),
    (4, True,  "CL-BZ Inter-Product (listed)","LISTED", 1.0,
     "Exchange-listed spread - beta not applicable",
     1000, 2, 0.01, 2, "n/a - this IS the listed spread"),
]


def build(wb):
    ws = wb.create_sheet("Config")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 30, "B": 34, "C": 14, "D": 10, "E": 96,
                    "F": 30, "G": 12, "H": 12, "I": 10, "J": 8, "K": 34})

    title(ws, "A1", "CONFIG  -  every parameter in its own labelled cell")
    note(ws, "A2", "LEGEND:  yellow / blue = EDIT THESE.  black = formula, do not edit.  "
                   "A changed value takes effect on the NEXT TIMER TICK - no restart, no VBA edit.", color="000000", bold=True)
    note(ws, "A3", "Exception: LOOKBACK_MIN redefines the window, so stored history is re-evaluated against the new "
                   "width on the next stats refresh and the warm-up gates re-apply.")
    note(ws, "A4", "VBA finds a parameter by matching column A exactly, so you may insert or reorder rows freely - "
                   "but never rename a parameter in column A.")
    note(ws, "A5", "FOUR DECOUPLED RATES:  CAPTURE_MS (poll for a changed quote)  >  STORE_MIN_INTERVAL_MS (what "
                   "reaches the window)  >  PRICE_REFRESH_MS (what repaints)  >  STATS_REFRESH_MIN (when mean and "
                   "sigma are republished).  Capturing fast buys alert latency, not a better sigma.",
         color="000000", bold=True)

    r = 7
    def section(name):
        nonlocal r
        ws[f"A{r}"] = name
        ws[f"A{r}"].font = f(11, bold=True, color=C_HDR_TEXT)
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=C_HDR_FILL)
            ws[f"{col}{r}"].font = f(11, bold=True, color=C_HDR_TEXT)
        r += 1
        header_row(ws, r, {"A": "Parameter", "B": "Value", "C": "Default", "D": "Units", "E": "Meaning"},
                   fill=C_SECT_FILL, text="1F3864")
        r += 1

    def rows(items):
        nonlocal r
        for name, val, dflt, unit, mean in items:
            label(ws, f"A{r}", name, bold=True)
            ws[f"A{r}"].border = BOX
            numfmt = "0.00" if isinstance(val, float) else None
            input_cell(ws, f"B{r}", val, numfmt)
            if isinstance(val, str):
                ws[f"B{r}"].alignment = Alignment(horizontal="left")
            value_cell(ws, f"C{r}", dflt, numfmt)
            ws[f"C{r}"].font = f(9, color=C_NOTE)
            value_cell(ws, f"D{r}", unit); ws[f"D{r}"].font = f(9, color=C_NOTE)
            label(ws, f"E{r}", mean, size=9)
            ws[f"E{r}"].border = BOX
            ws[f"E{r}"].alignment = Alignment(wrap_text=False)
            r += 1
        r += 1

    section("ENGINE  /  SAMPLING");            rows(ENGINE)
    section("WINDOW  /  STATISTICS");          rows(WINDOW)
    section("SIGNAL THRESHOLDS");              rows(SIGNAL)
    section("SIZING  /  COSTS");               rows(COSTS)
    section("AUDIBLE ALERT");                  rows(ALERTS)
    section("PERSISTENCE  /  LOGGING");        rows(PERSIST)

    # --- derived commission block (real formulas) -------------------------
    rate_rows = {}
    for row in range(1, r):
        v = ws.cell(row=row, column=1).value
        if isinstance(v, str) and v.startswith("RATE_"):
            rate_rows[v] = row
    ro, rf, rc = (rate_rows["RATE_ORIENT_PER_SIDE"], rate_rows["RATE_FIX_PER_SIDE"], rate_rows["RATE_CME_PER_SIDE"])

    ws[f"A{r}"] = "DERIVED COMMISSION  (formulas - do not edit)"
    ws[f"A{r}"].font = f(11, bold=True, color=C_HDR_TEXT)
    for col in "ABCDE":
        ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=C_HDR_FILL)
    r += 1
    header_row(ws, r, {"A": "Parameter", "B": "Value", "C": "Default", "D": "Units", "E": "Meaning"},
               fill=C_SECT_FILL, text="1F3864")
    r += 1

    label(ws, f"A{r}", "COMMISSION_CONSERVATIVE", bold=True); ws[f"A{r}"].border = BOX
    value_cell(ws, f"B{r}", f"=2*(B{ro}+B{rf}+B{rc})", "0.00", bold=True)
    label(ws, f"E{r}", "Per lot PER LEG, round turn, reading CME's 1.50 as per SIDE. This is the conservative "
                       "reading and the one to use.", size=9)
    r_cons = r; r += 1

    label(ws, f"A{r}", "COMMISSION_IF_CME_PER_RT", bold=True); ws[f"A{r}"].border = BOX
    value_cell(ws, f"B{r}", f"=2*(B{ro}+B{rf})+B{rc}", "0.00")
    label(ws, f"E{r}", "The same figure if CME's 1.50 turns out to be per ROUND TURN rather than per side.", size=9)
    r_alt = r; r += 1

    label(ws, f"A{r}", "COMMISSION_BASIS", bold=True); ws[f"A{r}"].border = BOX
    input_cell(ws, f"B{r}", "CONSERVATIVE")
    value_cell(ws, f"C{r}", "CONSERVATIVE"); ws[f"C{r}"].font = f(9, color=C_NOTE)
    label(ws, f"E{r}", 'CONSERVATIVE or CME_PER_RT. An understated cost model approves losing trades invisibly; '
                       'an overstated one only refuses trades, visibly. Leave it on CONSERVATIVE.', size=9)
    r_basis = r; r += 1

    label(ws, f"A{r}", "COMMISSION_PER_LOT_ROUND_TURN", bold=True); ws[f"A{r}"].border = BOX
    value_cell(ws, f"B{r}", f'=IF(B{r_basis}="CME_PER_RT",B{r_alt},B{r_cons})', "0.00", bold=True)
    label(ws, f"E{r}", "USD per lot per leg, round turn. This is what the cost model actually uses.", size=9)
    r += 2

    # --- spread definition table ------------------------------------------
    ws[f"A{r}"] = "SPREAD_TABLE_START   -   up to 4 monitored spreads, one buffer each"
    ws[f"A{r}"].font = f(11, bold=True, color=C_HDR_TEXT)
    for col in "ABCDEFGHIJK":
        ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=C_HDR_FILL)
    tbl_marker = r; r += 1
    note(ws, f"A{r}", "Leg COMPOSITION (which instruments, and their weights) is defined by the labelled formulas in "
                      "the MONITOR LEGS block on the hidden Feed sheet. Everything below is hot-reloaded each tick.")
    r += 1
    header_row(ws, r, {"A": "Slot", "B": "Enabled", "C": "Display name", "D": "Mode", "E": "HEDGE_RATIO (beta)",
                       "F": "Beta stamped for pair", "G": "USD per 1.00 of spread, per lot",
                       "H": "Contracts charged commission", "I": "Tick size", "J": "Decimals",
                       "K": "Listed equivalent (cost comparison)"},
               fill=C_SECT_FILL, text="1F3864")
    r += 1
    first_spread_row = r
    for slot, en, name, mode, hedge, stamp, usdpt, ncon, tick, dp, listed in SPREADS:
        value_cell(ws, f"A{r}", f"SPREAD_{slot}", bold=True)
        input_cell(ws, f"B{r}", en)
        input_cell(ws, f"C{r}", name); ws[f"C{r}"].alignment = Alignment(horizontal="left")
        input_cell(ws, f"D{r}", mode)
        input_cell(ws, f"E{r}", hedge, "0.0000")
        input_cell(ws, f"F{r}", stamp); ws[f"F{r}"].alignment = Alignment(horizontal="left")
        input_cell(ws, f"G{r}", usdpt, "#,##0")
        input_cell(ws, f"H{r}", ncon, "0")
        input_cell(ws, f"I{r}", tick, "0.0000")
        input_cell(ws, f"J{r}", dp, "0")
        value_cell(ws, f"K{r}", listed); ws[f"K{r}"].alignment = Alignment(horizontal="left")
        ws[f"K{r}"].font = f(9, color=C_NOTE)
        r += 1
    r += 1

    note(ws, f"A{r}", "HEDGE RATIO BELONGS TO THE PAIR. If the instruments change, beta must be re-derived - a stale "
                      "beta silently redefines the series. Column F stamps what each beta was computed for.",
         color=C_WARN, bold=True)
    r += 1
    note(ws, f"A{r}", "Contracts charged: BZ-CL and the listed spread are 2 contracts per lot (one per leg); the 3:2:1 "
                      "crack is 6 (2 RB + 1 HO + 3 CL).")
    r += 1
    note(ws, f"A{r}", "USD per 1.00 of spread: NYMEX energy contracts are 1,000 bbl, so one cent of differential is "
                      "$10 per leg per lot, i.e. $1,000 per 1.00.")
    r += 2
    note(ws, f"A{r}", "NOTIONAL_BASIS - read this before changing it:", color=C_WARN, bold=True)
    r += 1
    note(ws, f"A{r}", "A spread is not an outright, so 'notional' has no single meaning. At 1 lot with CL near $60 and "
                      "BZ-CL near $7.20:  SPREAD_VALUE ~ $7,200, ONE_LEG ~ $60,000, BOTH_LEGS ~ $127,000.")
    r += 1
    note(ws, f"A{r}", "1% of those is $72, $600 and $1,270. If sigma on this spread turns out to be ~5 cents ($50), "
                      "only the first is reachable - the others need the spread to travel 12 or 25 sigma, which it will "
                      "not do. Watch the TP-in-sigma cell on the Dashboard: above about 1 sigma it turns amber.",
         color=C_WARN)
    return ws, tbl_marker, first_spread_row
