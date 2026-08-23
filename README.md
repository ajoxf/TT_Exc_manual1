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
| `TT_ZScore_Monitor.xlsx` | All structure, layout, formatting, conditional formatting and the Feed formulas |
| `TTZMonitor.bas` | The VBA engine — import into the workbook, then save as `.xlsm` |
| `build/` | The openpyxl scripts that generate the `.xlsx`, so it can be regenerated |

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
| `Dashboard` | yes | The existing layout, unchanged. **No cell on it is a formula.** New z-monitor block below the existing one, same visual style. |
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
tp_distance_spread = (TARGET_NET_USD + total_cost) / (USD_PER_POINT × LOTS)
break_even_spread  = entry_spread ± total_cost / (USD_PER_POINT × LOTS)
take_profit_spread = entry_spread ± tp_distance_spread
capture            = 0.5 × |z| × sigma × USD_PER_POINT × LOTS
pass if capture ≥ EDGE_MULTIPLE × total_cost
```

Legging and listed crossing costs are shown side by side. The listed `CL-BZ` inter-product
spread crosses one market instead of two — roughly half the cost, and no leg risk.

Displayed live: break-even and take-profit as **absolute spread levels**, TP in sigma, target
z, sigma in dollars, a warning when the take-profit sits **beyond the mean** (that needs an
overshoot — a different bet from the one the z measured), and the edge verdict beside the
highlight. **A highlighted z that fails the edge filter is not a trade.**

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
- **Time-based eviction on every append**, so a dead feed drains the window and goes cold
  rather than freezing a stale z.
- Flushed to the hidden `Buffer` sheet every `FLUSH_SEC`, **decimated to `FLUSH_DECIMATE_MS`**
  so a 120-minute window stays near 7,200 rows and the flush never stalls capture.
- **Warm start** reloads that snapshot on demand, crediting elapsed collection from the
  *oldest* recovered sample, and only when the stamped lookback, hedge ratio and legs still
  match `Config`.
- Up to **four spreads**, each with its own buffer, head pointer and dedupe state.

---

## Assumptions made while building this

The repository was empty, so these are recorded on the **Setup** sheet as well:

- **The `Dashboard` layout was reconstructed** from the cell references in the spec — `C4`,
  `B6`, `C7`, `H7`/`H8`, `P7`/`P8`, `H12`/`H13`/`H14`/`I14`, `L3`. Those coordinates are
  honoured exactly; the rows around them are a best reconstruction. Compare against the live
  sheet before relying on it.
- **The 3:2:1 ask row now mirrors the bid row:** `I15` uses `H14` (CL **bid**), not `I14`. The
  bid row `H15` is unchanged and correct as written. This is the convention the spec described —
  please confirm it is the one you want.
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
