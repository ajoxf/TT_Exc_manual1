# TT RTD Z-Score Monitor

An Excel z-score monitor for **manually** trading energy futures spreads on Trading
Technologies Standard, fed by TT's RTD server. Excel stores the streaming tick data,
computes a rolling mean and standard deviation over a configurable lookback, shows a live
z-score, highlights when `|z|` crosses a configurable threshold, and works out a take-profit
level from the full cost of the round trip.

**All execution is manual in TT. This workbook never places an order.**

---

## Deliverables

| File | What it is |
|---|---|
| `TT_ZScore_Monitor.xlsx` | **Your own `TTDashboard_2108V1_1.xlsx`, transformed** — same layout, fonts, fills, row heights, merges and hidden columns, plus the added z-score block and supporting sheets |
| `TTZMonitor.bas` | The VBA engine — import into the workbook, then save as `.xlsm` |
| `reference/TTDashboard_2108V1_1.xlsx` | The untouched original, so the build is reproducible and any change is diffable |
| `build/` | The openpyxl scripts that transform the original, so it can be regenerated |

`openpyxl` cannot author VBA, so the workbook ships as `.xlsx` + `.bas`. The **Setup** sheet
inside the workbook has the six-step conversion; the short version:

1. Open the `.xlsx` on the machine running TT and the TT RTD server.
2. Save As → **Excel Macro-Enabled Workbook (.xlsm)**.
3. `Alt+F11` → File → Import File… → `TTZMonitor.bas`.
4. Paste the `Workbook_Open` / `Workbook_BeforeClose` stub (printed at the foot of the `.bas`)
   into `ThisWorkbook`. **Not optional** — see *Timers* below.
5. Draw two shapes over `N2` / `P2` on `Dashboard` and assign `StartMonitor` / `StopMonitor`.
6. Save, click **START**.

---

## Sheets

| Sheet | Visible | Purpose |
|---|---|---|
| `Dashboard` | yes | **Your sheet, unchanged.** Rows 1–15 keep every font, fill, height, merge and hidden column exactly. **No cell on it is a formula.** The z-score block is added at rows 17–21 in the same card style; the control strip at row 23. |
| `Detail` | yes | Costs, break-even, warm-up gates, feed rate and the edge filter, for all four slots. |
| `Config` | yes | Every parameter in its own labelled cell. Hot-reloaded — no restart, no VBA edit. |
| `Feed` | hidden | Every `RTD()` formula in the workbook. |
| `Buffer` | hidden | The flushed recovery snapshot of the circular buffers. |
| `Log` | yes | Timestamped signal events, edge-triggered. |
| `Setup` | yes | Conversion steps, first-run checklist, the sigma-threshold table, and the assumptions made while building this. |

---

## No flicker

Three causes, all removed:

1. **Volatile formulas.** `=TEXT(NOW(),"hh:mm:ss")` in `L3` is gone — VBA writes the same
   string into the same cell. There is no `NOW`, `TODAY`, `OFFSET`, `INDIRECT` or `RAND`
   anywhere in the workbook.
2. **Live formulas on the visible sheet.** Every `RTD()` lives on the hidden `Feed` sheet.
   A cell holding a live `RTD()` formula repaints whenever the feed moves, and that cannot be
   prevented while the formula is in the cell. `Dashboard` looks identical; only the mechanism
   changed.
3. **Blanket writes and whole-column conditional formatting.** Every write goes through one
   helper that **compares before it writes**, updates are wrapped in
   `ScreenUpdating = False … True`, and conditional formatting is applied to small ranges only
   (`D48:G49`, `D67:G67`, `D42:G42`, `D40:G40`, `D63:G63`, `C3`).

Prices display to tick precision only — a number showing more digits than the market moves in
will appear to change constantly.

---

## Four decoupled rates

Confusing these is how these tools go wrong.

| Rate | Config cell | Default | What it actually buys |
|---|---|---|---|
| 1. Capture | `CAPTURE_MS` | 100 | Polls the feed for a **changed** quote. Buys **latency to the chime**. Not a better sigma. |
| 2. Store | `STORE_MIN_INTERVAL_MS` | 0 | What reaches the window. `0` = full fidelity. |
| 3. Paint | `PRICE_REFRESH_MS` | 500 | What repaints. |
| 4. Publish | `STATS_REFRESH_MIN` | 5 | When mean and sigma are republished. |

After dedupe the window holds **quote changes, not poll ticks**, so capture rate can never
inflate the sample count or collapse the variance. Subsampling a mean-reverting level series
leaves its marginal distribution — hence mean and sigma — unchanged, and closely-spaced
samples are so autocorrelated that the extra rows buy little precision anyway.

**Recommended sequence:** leave `STORE_MIN_INTERVAL_MS` at `0` for the first week or two while
`TICK_ARCHIVE_ENABLED` writes every changed quote to `ticks/*.csv`. Then run
`Alt+F8 → CheckSubsamplingLoss`, which recomputes mean and sigma over the live window at full
fidelity against 250 / 500 / 1000 / 2000 / 5000 ms subsampling and writes the comparison to
`Log`. If sigma is materially unchanged, set `STORE_MIN_INTERVAL_MS` to 250–1000 and reclaim
the memory.

### Timers

`Application.OnTime` cannot go below about one second, so capture runs on a `user32`
`SetTimer` with an `OnTime` watchdog (default 5 s) that re-arms it — Excel silently kills API
timers when a modal dialog opens. The watchdog also picks up a changed `CAPTURE_MS`. If
`SetTimer` is unavailable, `CAPTURE_MS` falls back to a 1 s `OnTime` loop; only alert latency
changes, not the statistics.

A `user32` timer that outlives the workbook keeps firing into a module that no longer has its
sheets, and that crashes Excel — which is why the `Workbook_BeforeClose` stub is required.

---

## The statistics

```
spread = LegB − HEDGE_RATIO × LegA
z      = (spread − mean) / sigma
```

Nothing else — no carry, no swap, no fair-value term. Brent/WTI is a **RELATED** pair, not a
basis pair: two different crudes with no arbitrage tying them, so the only anchor is the
spread's own empirical mean.

Both legs' **mid** prices feed the statistic. The touch (bid/ask) is shown separately, because
the fix for a wide market and the fix for a bad fill are different fixes.

Mean and sigma are **frozen** between refreshes. Ticks are stored continuously and the window
still slides continuously; only the *published* values are held. Two consequences, both stated
on the sheet: **z will step at each refresh** (expected, not a fault), and freezing the
reference stops a continuously-updating mean chasing the spread — which otherwise drags an open
position's z back toward zero without the price ever paying you.

### The five traps, all coded in

1. **Dedupe.** Every sample is keyed on both legs' bid *and* ask. A quote id matching the last
   stored one is discarded. When seeding from saved history, only *consecutive* identical
   values are dropped — a spread genuinely revisiting a level is a real observation.
2. **Degenerate window.** `sigma < MIN_SIGMA` or `|z| > MAX_ABS_Z` shows **"NO USABLE Z"** and a
   blank z cell. A blank is honest; a huge number reads as opportunity.
3. **Two warm-up gates, both required.** `samples ≥ MIN_SAMPLES` **and**
   `elapsed ≥ MIN_HISTORY_MIN`. The binding gate is displayed.
4. **The entry ceiling is its own rule.** Highlight only when `ENTRY_Z ≤ |z| < MAX_ENTRY_Z`.
   Above the ceiling the state reads **ABOVE CEILING** and the ceiling tone fires — that is
   "stand down", not "get in".
5. **Feed-rate display.** Observed quote *changes* per minute, amber at or below
   `THIN_FEED_QPM`.

**z is for ENTRIES only. Exits act on money, never on z.**

---

## Costs and take-profit

Per contract per side, from the Orient rate card (2026-08-21, Ver.26.8.17): Orient
clearing + execution `$0.35`, FIX/platform `$0.06`, CME exchange + clearing `$1.50`. Per lot
per leg round turn that is **$3.82** (`COMMISSION_BASIS = CONSERVATIVE`), or `$2.32` if CME's
`$1.50` proves to be per round turn — a `Config` toggle. An understated cost model approves
losing trades invisibly; an overstated one only refuses trades, visibly.

```
commission_cost    = CONTRACTS_CHARGED × COMMISSION_PER_LOT_ROUND_TURN × LOTS
crossing_cost      = (spread_ask − spread_bid) × USD_PER_POINT × LOTS
total_cost         = commission_cost + crossing_cost
tp_distance_spread = (win_usd + total_cost) / (USD_PER_POINT × LOTS)
break_even_spread  = spread_mid ± total_cost / (USD_PER_POINT × LOTS)
take_profit_spread = spread_mid ± tp_distance_spread
exit_order_level   = level ± (spread_ask − spread_bid) / 2
capture            = 0.5 × |z| × sigma × USD_PER_POINT × LOTS
pass if capture ≥ EDGE_MULTIPLE × total_cost
```

Legging and listed crossing costs are shown side by side. The listed `CL-BZ` inter-product
spread crosses one market instead of two — roughly half the cost, and no leg risk.

`Detail` also carries the mean expressed on **each touch**, laid out per direction — four rows off
two numbers (`mean ∓ Gap/2`), using the live Gap:

```
SHORT (high to low)  ENTER            sell the bid    mean − Gap/2
SHORT (high to low)  EXIT at the mean buy  the ask    mean + Gap/2
LONG  (low to high)  ENTER            buy  the ask    mean + Gap/2
LONG  (low to high)  EXIT at the mean sell the bid    mean − Gap/2
```

A short **enters** on the bid and **exits** on the ask, so its two mean-relevant prices sit on
*opposite* sides of the mid. One row per direction would force the reader to flip sides in their
head — the same touch/mid confusion that put half a width into break-even. Entering at the mean's
touch and reverting exactly to the mean's other touch nets **minus the total cost**, which is the
identity that ties this block to the cost model. Shown from the moment statistics publish, warm-up
or not.

Break-even and take-profit are **mid levels**, on the same scale as the mean, sigma and z — a
mid move of `total_cost` is exactly what pays the round trip. The entry reference beside them is
the **touch** you actually get filled at, which already carries half the width; anchoring the
levels to it instead would charge 1.5 widths of crossing where only 1.0 is real. To work the
exit order, cross back: add half the Gap to close a short, subtract half to close a long.

Displayed live: break-even and take-profit as **absolute spread levels**, TP in sigma, target
z, sigma in dollars, a warning when the take-profit sits **beyond the mean** (that needs an
overshoot — a different bet from the one the z measured), and the edge verdict beside the
highlight. **A highlighted z that fails the edge filter is not a trade.**

`WIN_PCT` ships at **0.004** (0.4%), which targets about **1.25 sigma** — the same capture the edge
filter already assumes at `ENTRY_Z`. That keeps the take-profit *inside* the mean on all three legged
spreads. The older `0.01` put it at 2.8–3.0 sigma on two of the three, i.e. past the mean, which needs
an overshoot rather than a reversion. A flat `TARGET_NET_USD` cannot serve all four slots at once —
their sigma differs by 10x — which is why `PCT_NOTIONAL` is the default mode.

`ENTRY_Z = 2.5` rather than 3.0 means less expected capture per entry, so it *raises* the sigma
the spread must have by about 20%. The **Setup** sheet computes that table live off `Config`.

**Sigma on these spreads has never been measured.** Every number above is contingent on it, and
measuring it is a main purpose of this tool.

---

## The alert

One short chime **on the crossing, not on the state**:

- **Edge-triggered** — only on the transition from below the threshold to at-or-above it.
- **Hysteresis** — does not re-arm until `|z|` falls back below `ENTRY_Z − ALERT_REARM_MARGIN`.
- **Cooldown** — `ALERT_COOLDOWN_SEC` between alerts for the same spread.
- **Silent while not tradable** — warm-up, degenerate window, or stale feed.
- **Silent when the edge filter fails** (`ALERT_ONLY_IF_EDGE_PASSES`, default on). A sound for
  a trade you should not take trains you to ignore the sound.
- **Master mute** (`ALERT_ENABLED`) and a **distinct tone** for the `MAX_ENTRY_Z` ceiling.

Played through `winmm.dll PlaySound` asynchronously, so it never blocks the capture timer;
`Beep` is the fallback. `ALERT_SPEAK` optionally announces which spread fired.

---

## Volume and performance

- Fixed-size circular buffer per spread in **VBA arrays, not cells** (`BUFFER_CAPACITY`
  80,000 → ~5 MB across four spreads).
- Statistics recomputed in full only at each publish, which is cheap at that cadence and clears
  any accumulated float drift.
- **`MIN_HISTORY_MIN` must sit below `LOOKBACK_MIN`.** Eviction caps the window span at the
  lookback, so the elapsed-history gate can only ever approach `LOOKBACK_MIN` from below and a
  value at or above it is unsatisfiable — every slot reads `WARMING UP` forever, with no z, no
  signal and no chime. The engine now clamps such a value to 99% of `LOOKBACK_MIN`, logs it and
  shows it on the Dashboard status strip rather than warming up in silence.
- **Time-based eviction on every append**, so a dead feed drains the window and goes cold
  rather than freezing a stale z.
- Flushed to the hidden `Buffer` sheet every `FLUSH_SEC`, **decimated to `FLUSH_DECIMATE_MS`**
  so a 120-minute window stays near 7,200 rows and the flush never stalls capture.
- **Warm start** reloads that snapshot on demand, crediting elapsed collection from the
  *oldest* recovered sample, and only when the stamped lookback, hedge ratio and legs still
  match `Config`.
- Up to **four spreads**, each with its own buffer, head pointer and dedupe state.

---

## What changed in your Dashboard, and nothing else

Verified against `reference/TTDashboard_2108V1_1.xlsx`: every label, font, size, colour, number
format, row height, column width, hidden column (`A`, `B`, `C`, `G`, `O`), merge and the print
area are identical. Four deliberate changes:

- **Every `RTD()` formula moved to the hidden `Feed` sheet.** The visible cells keep their exact
  styling and become VBA-written values. A cell holding a live `RTD()` formula repaints whenever
  the feed moves, and that cannot be prevented while the formula is in the cell.
- **`=TEXT(NOW(),"hh:mm:ss")` in `L3` removed.** `NOW()` is volatile, so it forced a full
  recalculation of the workbook on every RTD update. VBA writes the same string into the same
  cell, in the same format.
- **`I15` corrected** to reference `H14` (CL **bid**) rather than `I14`, so the 3:2:1 ask mirrors
  the bid. The bid row `H15` was already right.
- **`J15` added** — the 3:2:1 row was the only spread row with no Gap cell.

Added: the z-score block at rows 17–21, and the control strip at row 23. Rows 1–15 are untouched.

### One arithmetic point to confirm

Your rows 9 quote the spread as bid-minus-bid and ask-minus-ask, so the **Gap** there is a
*difference of gaps* (`42×HO_gap − CL_gap`), which understates the cost of crossing and can go
negative. The crossing convention — sell LegB at the bid and buy LegA at the ask, mirrored for
the ask — makes the Gap the real round-trip cost and always positive. Your own 3:2:1 bid row
already used it.

`SPREAD_QUOTE_CONVENTION` on `Config` ships as `CROSSING`; set it to `SAME_SIDE` to get your
original rows-9 arithmetic back and compare the two side by side.
- **`USD_PER_POINT` for the 3:2:1 crack is 3,000, not 1,000.** The pack is three
  crude-equivalents, so one cent of crack is $30 per pack. `CONTRACTS_CHARGED` is 6
  (2 RB + 1 HO + 3 CL). Both are `Config` cells.
- **Slot 3 has no exchange-listed equivalent**, so its listed crossing-cost cell reads `n/a`.
- **`MIN_SIGMA` ships at 0** because sigma has never been measured. Until it is set, the
  degenerate-window guard rests on `MAX_ABS_Z` alone.
- **Leg composition** (which instruments, at what weights) is defined by labelled formulas in
  the *MONITOR LEGS* block on `Feed`, not on `Config`. Beta and every statistical, cost and
  alert parameter are on `Config` and hot-reload; changing which instruments make up a spread
  means editing those Feed formulas.
