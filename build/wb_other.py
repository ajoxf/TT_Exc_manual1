"""Buffer (hidden), Log, and Setup sheets."""
from openpyxl.styles import PatternFill, Alignment
from wb_common import *

# Buffer: per-slot column pairs, data from row 10
SLOT_COLS = {1: ("A", "B"), 2: ("D", "E"), 3: ("G", "H"), 4: ("J", "K")}
DATA_ROW = 10


def build_buffer(wb):
    ws = wb.create_sheet("Buffer")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {c: 15 for c in "ABDEGHJK"})
    set_widths(ws, {"C": 3, "F": 3, "I": 3, "M": 26, "N": 30})

    title(ws, "A1", "BUFFER  -  hidden.  Flushed circular buffers (crash / restart recovery).", 12)
    note(ws, "M1", "Written by VBA every FLUSH_SEC seconds, so a crash loses at most a minute.")
    note(ws, "M2", "On open, WarmStart reloads a slot only if its stamped LOOKBACK_MIN, HEDGE_RATIO and leg labels "
                   "still match Config, and the newest sample is younger than WARM_START_MAX_AGE_MIN.")
    note(ws, "M3", "Elapsed collection is credited from the OLDEST recovered sample - otherwise every restart costs "
                   "another two hours before a usable z.")
    note(ws, "M4", "Seeding drops CONSECUTIVE identical values only. A spread genuinely revisiting a level is a real "
                   "observation, not a duplicate.")

    for slot, (ct, cv) in SLOT_COLS.items():
        label(ws, f"{ct}3", f"SLOT {slot}", bold=True, color=C_TITLE)
        for r, lab in ((4, "Name"), (5, "Legs"), (6, "HEDGE_RATIO"), (7, "LOOKBACK_MIN"),
                       (8, "Count / last flush")):
            label(ws, f"{ct}{r}", lab, size=9, color=C_NOTE)
            ws[f"{cv}{r}"].border = BOX
        header_row(ws, 9, {ct: "Time (serial)", cv: "Spread value"})
    return ws


def build_log(wb):
    ws = wb.create_sheet("Log")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 20, "B": 30, "C": 20, "D": 9, "E": 12, "F": 12, "G": 12,
                    "H": 12, "I": 12, "J": 16, "K": 10, "L": 60})
    title(ws, "A1", "LOG  -  timestamped signal events", 12)
    note(ws, "A2", "One line when a state STARTS and one when it CLEARS - never one per poll. "
                   "Reviewable after the session.")
    header_row(ws, 3, {"A": "Time", "B": "Spread", "C": "Event", "D": "z", "E": "Mean", "F": "Sigma",
                       "G": "Spread", "H": "Bid", "I": "Ask", "J": "Signal", "K": "Edge", "L": "Note"})
    ws.freeze_panes = "A4"
    return ws


def build_setup(wb, cfg_rows):
    """cfg_rows: {PARAM_NAME: row on Config}"""
    ws = wb.create_sheet("Setup")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 4, "B": 46, "C": 16, "D": 16, "E": 16, "F": 16, "G": 60})

    def C(name):
        return f"Config!$B${cfg_rows[name]}"

    title(ws, "B1", "SETUP  -  turning this .xlsx into the working .xlsm")
    steps = [
        "openpyxl cannot author VBA, so this file ships as an .xlsx carrying all structure and formatting, "
        "plus a .bas module to import. Six steps, once.",
        "1.  Open TT_ZScore_Monitor.xlsx in Excel on the machine running TT and the TT RTD server.",
        "2.  File > Save As > Excel Macro-Enabled Workbook (.xlsm). Keep the same folder.",
        "3.  Alt+F11 to open the VBA editor.  File > Import File...  and choose TTZMonitor.bas.",
        "4.  In the editor, double-click ThisWorkbook and paste the two-line Workbook_Open / "
        "Workbook_BeforeClose stub printed at the bottom of TTZMonitor.bas.  (Optional - it just "
        "auto-starts the monitor and warm-starts the buffers.)",
        "5.  Back on Dashboard, insert two shapes over cells N2 and P2 (Insert > Shapes > Rectangle), "
        "right-click each > Assign Macro > StartMonitor and StopMonitor. Until you do, run them from Alt+F8.",
        "6.  Save. Click START. The Feed sheet stays hidden; you never need to look at it.",
        "Step 4 is NOT optional here. The capture timer is a Windows user32 timer, and one that outlives the "
        "workbook keeps firing into a module that no longer has its sheets - which crashes Excel. "
        "Workbook_BeforeClose calls StopMonitor and kills it.",
    ]
    r = 3
    for s in steps:
        ws[f"B{r}"] = s
        ws[f"B{r}"].font = f(10, bold=s[:2].strip().rstrip(".").isdigit())
        ws.merge_cells(f"B{r}:G{r}")
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if len(s) > 110 else 15
        r += 1
    r += 1

    block_title(ws, f"B{r}", "WHAT TO CHECK ON THE FIRST RUN")
    r += 1
    for s in [
        "Feed sheet column B (instrument ids) fills in. If it stays blank, the TT short name in Feed column A "
        "does not match TT, or the market in Dashboard!C4 is wrong.",
        "Dashboard prices move, but the sheet does NOT flash. If it flashes, something volatile has been "
        "reintroduced - search the workbook for NOW, TODAY, OFFSET, INDIRECT, RAND.",
        "The samples counter climbs at roughly the quotes/min rate, not at the timer rate. If it climbs at the "
        "timer rate, dedupe is not working and every z you see will be fiction.",
        "Warm-up takes MIN_HISTORY_MIN (default 120) minutes on a cold start. That is deliberate.",
        "A 'ticks' folder appears beside the workbook, one CSV per spread per day, holding every changed quote. "
        "That is the raw record - keep it for the first week or two.",
        "After a week, run Alt+F8 > CheckSubsamplingLoss. It compares mean and sigma over the live window at full "
        "fidelity against 250 / 500 / 1000 / 2000 / 5000 ms subsampling and writes the comparison to the Log. If "
        "sigma is materially unchanged, set STORE_MIN_INTERVAL_MS and reclaim the memory.",
    ]:
        ws[f"B{r}"] = "-  " + s
        ws[f"B{r}"].font = f(9)
        ws.merge_cells(f"B{r}:G{r}")
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1

    # ---- sigma threshold table (live formulas) ----------------------------
    block_title(ws, f"B{r}", "WHAT SIGMA THIS SPREAD MUST HAVE TO BE WORTH TRADING")
    r += 1
    note(ws, f"B{r}", "Live off Config. Edit the two assumed market widths below; everything else follows.")
    r += 1
    w_row = r
    label(ws, f"B{r}", "Assumed width, each outright leg (spread units)", size=9)
    input_cell(ws, f"C{r}", 0.02, "0.0000")
    label(ws, f"E{r}", "Assumed width, listed spread (spread units)", size=9)
    input_cell(ws, f"F{r}", 0.02, "0.0000")
    r += 2

    header_row(ws, r, {"B": "At LOTS from Config, per leg", "C": "Total cost", "D": "min sigma @ z = 3.0",
                       "E": "min sigma @ z = 2.5", "F": "min sigma @ live ENTRY_Z"})
    r += 1
    lots, comm, edge = C("LOTS"), C("COMMISSION_PER_LOT_ROUND_TURN"), C("EDGE_MULTIPLE")
    entry = C("ENTRY_Z")
    USD_PT = 1000

    for lab, ncon, cross in (
        ("Legging the outrights", 2, f"(2*$C${w_row})*{USD_PT}*{lots}"),
        ("Listed inter-product spread", 2, f"$F${w_row}*{USD_PT}*{lots}"),
    ):
        label(ws, f"B{r}", lab, bold=True)
        value_cell(ws, f"C{r}", f"={ncon}*{comm}*{lots}+{cross}", '$#,##0.00')
        for col, z in (("D", "3.0"), ("E", "2.5"), ("F", entry)):
            value_cell(ws, f"{col}{r}", f"={edge}*$C${r}/(0.5*{z}*{USD_PT}*{lots})", "0.0000")
        r += 1
    r += 1
    note(ws, f"B{r}", "capture = 0.5 x |z| x sigma x USD-per-point x LOTS.  An entry passes when "
                      "capture >= EDGE_MULTIPLE x total cost.")
    r += 1
    note(ws, f"B{r}", "A lower ENTRY_Z means less expected capture per entry, so it RAISES the sigma the spread must "
                      "have. 2.5 instead of 3.0 raises it by about 20%.", color=C_WARN, bold=True)
    r += 2

    block_title(ws, f"B{r}", "ASSUMPTIONS MADE WHILE BUILDING THIS - CHECK THEM")
    r += 1
    for s in [
        "The repository was empty, so the existing Dashboard layout was reconstructed from the cell references given "
        "in the spec (C4, B6, C7, H7/H8, P7/P8, H12/H13/H14/I14, L3). Those coordinates are honoured exactly; the "
        "surrounding rows are a best reconstruction. Compare against your live sheet before relying on it.",
        "The 3:2:1 ask row now mirrors the bid row: I15 uses H14 (CL BID), not I14. The bid row is unchanged. "
        "This is the convention the spec described; confirm it is the one you want.",
        "COMMISSION_BASIS defaults to CONSERVATIVE ($3.82 per lot per leg round turn, reading CME's $1.50 as per "
        "side). Switch it to CME_PER_RT on Config if the rate card proves otherwise.",
        "Contracts charged commission: 2 for BZ-CL and the listed spread, 6 for the 3:2:1 crack (2 RB + 1 HO + 3 CL). "
        "Change it in the Config spread table if your clearing differs.",
        "Slot 3 (3:2:1) has no exchange-listed equivalent, so its LISTED crossing-cost column reads n/a.",
        "MIN_SIGMA ships at 0 because sigma on these spreads has never been measured. Set it once you have a "
        "session of data - until then the degenerate-window guard rests on MAX_ABS_Z alone.",
        "Application.OnTime cannot go below about one second, so capture at 100 ms runs on a user32 SetTimer with "
        "an OnTime watchdog that re-arms it (Excel silently kills API timers when a modal dialog opens). If "
        "SetTimer is unavailable, CAPTURE_MS is clamped to 1000 and everything still works - only alert latency "
        "changes, not the statistics.",
        "Capturing at 100 ms does not mean storing at 100 ms, and after dedupe it does not change sigma either. "
        "The window holds quote CHANGES, not poll ticks; subsampling a mean-reverting level series is unbiased. "
        "100 ms buys latency to the chime, which for manual execution is not measurable in human terms - so the "
        "real reason to capture fast is simply that no genuine quote change is missed.",
    ]:
        ws[f"B{r}"] = "-  " + s
        ws[f"B{r}"].font = f(9)
        ws.merge_cells(f"B{r}:G{r}")
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 40
        r += 1
    return ws
