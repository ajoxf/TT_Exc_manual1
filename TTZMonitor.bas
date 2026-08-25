Attribute VB_Name = "TTZMonitor"
'==============================================================================
' TT RTD Z-SCORE MONITOR  -  manual trading of energy futures spreads on
' Trading Technologies Standard.  THIS MODULE NEVER PLACES AN ORDER.
'
' FOUR DECOUPLED RATES.  Confusing them is how these tools go wrong.
'
'   1  CAPTURE_MS (100)             poll the feed for a CHANGED quote.
'                                   Buys LATENCY TO THE CHIME.  Not sigma.
'   2  STORE_MIN_INTERVAL_MS (0)    what actually reaches the window.
'                                   0 = full fidelity.  Raising it to 250-1000
'                                   is unbiased subsampling: a mean-reverting
'                                   level series has the same marginal
'                                   distribution however finely you sample it,
'                                   so mean and sigma are unchanged - and
'                                   closely-spaced samples are so autocorrelated
'                                   that the extra rows buy little precision.
'   3  PRICE_REFRESH_MS (500)       what repaints.
'   4  STATS_REFRESH_MIN (5)        when mean and sigma are republished.
'
' Design rules this module exists to enforce:
'   * No cell on Dashboard is a formula, and no volatile function exists
'     anywhere in the workbook.  Every RTD() lives on the hidden Feed sheet.
'   * A Dashboard cell is written ONLY when the new value differs from what is
'     already there.  Compare first, write second.  That is what stops the
'     display flashing on every uptick.
'   * DEDUPE.  A sample is stored only when the quote actually changed.  A
'     timer sampling faster than the feed ticks would otherwise store the same
'     quote repeatedly, collapse the variance and manufacture enormous
'     z-scores.  The engine this is ported from once printed z = +53,026 on a
'     spread of 9.13 for exactly this reason.
'   * Mean and sigma are FROZEN between refreshes.  z steps at each refresh -
'     that is expected, not a fault.
'   * z is for ENTRIES only.  Exits act on money, never on z.
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
        ByVal lpszName As String, ByVal hModule As LongPtr, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private gTimerId As LongPtr
#Else
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
        ByVal lpszName As String, ByVal hModule As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private gTimerId As Long
#End If
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_NODEFAULT As Long = &H2

' MIN_HISTORY_MIN is capped at this fraction of LOOKBACK_MIN. The window
' span can only approach the lookback from below, so the gate has to sit
' strictly under it or it can never be met. Proportional, so it scales with
' whatever lookback is configured.
Private Const HIST_GATE_MAX_FRAC As Double = 0.99

'------------------------------------------------------------------ sheets ---
Private Const SH_DASH As String = "Dashboard"
Private Const SH_DET  As String = "Detail"
Private Const SH_CFG  As String = "Config"
Private Const SH_FEED As String = "Feed"
Private Const SH_BUF  As String = "Buffer"
Private Const SH_LOG  As String = "Log"

'--------------------------------------------------------- Feed row layout ---
Private Const F_INST_TOP As Long = 6     ' A6:H11  RB HO CL BZ CLBZ HOCL
Private Const F_INST_N   As Long = 6
Private Const F_DER_TOP  As Long = 16    ' A16:F20 derived + listed spreads
Private Const F_DER_N    As Long = 5
Private Const F_LEG_TOP  As Long = 25    ' C25:F32 monitor legs, 2 rows / slot
Private Const F_ALT_TOP  As Long = 36    ' A36:B39 alternative-execution ref row

Private Const I_RB As Long = 1
Private Const I_HO As Long = 2
Private Const I_CL As Long = 3
Private Const I_BZ As Long = 4
Private Const I_CLBZ As Long = 5
Private Const I_HOCL As Long = 6

'--------------------------------------------------- Detail row map ----------
Private Const R_NAME As Long = 21:   Private Const R_DEF As Long = 22
Private Const R_MODE As Long = 23:   Private Const R_BETA As Long = 24
Private Const R_USDPT As Long = 25
Private Const R_BID As Long = 27:    Private Const R_ASK As Long = 28
Private Const R_MID As Long = 29:    Private Const R_WIDTH As Long = 30
Private Const R_SAMP As Long = 32:   Private Const R_ELAP As Long = 33
Private Const R_GATE As Long = 34:   Private Const R_RATE As Long = 35
Private Const R_FSTAT As Long = 36
Private Const R_MEAN As Long = 38:   Private Const R_SIG As Long = 39
Private Const R_SIGUSD As Long = 40: Private Const R_AGE As Long = 41
Private Const R_Z As Long = 42:      Private Const R_SIGNAL As Long = 43
Private Const R_DIR As Long = 45:    Private Const R_ENTRY As Long = 46
Private Const R_COMM As Long = 47:   Private Const R_XLEG As Long = 48
Private Const R_XLST As Long = 49:   Private Const R_TOT As Long = 50
Private Const R_TPRULE As Long = 52: Private Const R_NOTIONAL As Long = 53
Private Const R_WINUSD As Long = 54: Private Const R_COSTU As Long = 55
Private Const R_WINU As Long = 56:   Private Const R_BE As Long = 57
Private Const R_TP As Long = 58:     Private Const R_TPD As Long = 59
Private Const R_TPS As Long = 60:    Private Const R_TZ As Long = 61
Private Const R_BEYOND As Long = 62
Private Const R_CAP As Long = 64:    Private Const R_REQ As Long = 65
' The mean on each touch, laid out per DIRECTION. Only two distinct numbers,
' but a short ENTERS on the bid and EXITS on the ask, so one row per
' direction would force the reader to flip sides in their head - which is
' exactly the touch/mid confusion that put half a width into break-even.
Private Const R_SH_ENT As Long = 70:  Private Const R_SH_EXIT As Long = 71
Private Const R_LG_ENT As Long = 72:  Private Const R_LG_EXIT As Long = 73
Private Const R_VERD As Long = 66:   Private Const R_MINSIG As Long = 67

'-------------------------------------------------- Dashboard row map --------
' The user's own layout, unchanged. VBA writes these cells; none is a formula.
'   C6 RB   C7 HO   C8 BZ   C9 CL   C10 CL-BZ listed   C12 HO-CL listed
'   F..L block  cols H bid, I ask, J gap, K high, L low
'        rows 7 HO, 8 CL, 9 HO|CL crack   /   12 RB, 13 HO, 14 CL, 15 3:2:1
'   N..T block  cols P bid, Q ask, R gap, S high, T low
'        rows 7 BZ, 8 CL, 9 BZ - CL
'   z block rows 19 HO|CL, 20 BZ-CL, 21 3:2:1
Private Const CB_H As Long = 8            ' column H - left block bid
Private Const CB_P As Long = 16           ' column P - right block bid
Private Const ZR1 As Long = 19            ' HO|CL crack   -> slot 2
Private Const ZR2 As Long = 20            ' BZ - CL       -> slot 1
Private Const ZR3 As Long = 21            ' 3:2:1         -> slot 3

Private Const MAX_SPREADS As Long = 4
Private Const SLOT_COL As String = "DEFG"      ' Detail column per slot
Private Const CHG_RING As Long = 2048     ' quote-change times, for the feed rate

'------------------------------------------------------------------- state ---
Private Type SpreadState
    ' --- definition (hot-reloaded from Config) ---------------------------
    Enabled          As Boolean
    Name             As String
    Mode             As String            ' LEGGED | LISTED
    Hedge            As Double
    BetaStamp        As String
    UsdPerPoint      As Double
    Contracts        As Double
    TickSize         As Double
    Decimals         As Long
    AltRow           As Long
    LegLabelA        As String
    LegLabelB        As String

    ' --- window ----------------------------------------------------------
    Cap              As Long
    Head             As Long
    Count            As Long
    BufT()           As Double
    BufV()           As Double

    ' --- capture / store separation --------------------------------------
    SeenQid          As String            ' last quote id OBSERVED
    LastQuoteTime    As Double
    StoredQid        As String            ' last quote id STORED
    LastStoredT      As Double
    PendQid          As String            ' newest observation not yet stored
    PendV            As Double
    PendT            As Double
    HavePend         As Boolean
    ChgT()           As Double            ' ring of observed-change times
    ChgHead          As Long
    ChgCount         As Long

    ' --- published statistics --------------------------------------------
    Mean             As Double
    Sigma            As Double
    ' The mean is a MID. These are the same mean expressed on each touch, so
    ' each side reads as the price that side would actually deal at.
    MeanBid          As Double
    MeanAsk          As Double
    HaveMeanTouch    As Boolean
    StatsValid       As Boolean
    StatsN           As Long
    LastStatsTime    As Double

    ' --- derived, recomputed each capture, painted on the slow cadence ----
    CurBid           As Double
    CurAsk           As Double
    CurMid           As Double
    LegAMid          As Double
    LegBMid          As Double
    HaveTouch        As Boolean
    Z                As Double
    HaveZ            As Boolean
    Elapsed          As Double
    Qpm              As Double
    Signal           As String
    Gate             As String
    FeedStatus       As String
    Commission       As Double
    XLeg             As Variant
    XLst             As Variant
    Total            As Double
    HaveTrade        As Boolean
    DirTxt           As String
    Entry            As Double
    BE               As Double
    TP               As Double
    TPDist           As Double
    TPSig            As Double
    TargetZ          As Double
    Beyond           As String
    TpRule           As String
    Notional         As Double
    WinUsd           As Double
    CostUnits        As Double
    WinUnits         As Double
    Capture          As Double
    Required         As Double
    Verdict          As String
    MinSigPass       As Double

    ' --- alert / logging edge state --------------------------------------
    Armed            As Boolean
    LastAlert        As Double
    CeilArmed        As Boolean
    LastCeilAlert    As Double
    PrevSignal       As String
    PrevGate         As String
    PrevFeedStatus   As String

    ' --- tick archive -----------------------------------------------------
    Arc()            As String
    ArcN             As Long
End Type

Private S(1 To MAX_SPREADS) As SpreadState
Private gCfg            As Object
Private gRunning        As Boolean
Private gHiRes          As Boolean
Private gTimerMs        As Long
Private gInCapture      As Boolean
Private gNextWatch      As Date
Private gNextOnTime     As Date
Private gLastCapture    As Double
Private gLastConfig     As Double
Private gLastPaint      As Double
Private gLastFlush      As Double
Private gCaptureCount   As Long          ' captures since the last watchdog
Private gMeasuredHz     As Double        ' measured captures per second
Private gLastRateCalc   As Double
Private gTTState        As String        ' TT's own status text, when it is not an id
Private gHistClamp      As String        ' set when MIN_HISTORY_MIN had to be clamped
Private gFlushRows(1 To MAX_SPREADS) As Long

'==============================================================================
' PUBLIC ENTRY POINTS
'==============================================================================
Public Sub StartMonitor()
    On Error GoTo Fail
    If gRunning Then Exit Sub

    LoadConfig
    Application.RTD.ThrottleInterval = CfgL("RTD_THROTTLE_MS", 100)

    LoadSpreadDefs
    InitBuffers
    If CfgB("WARM_START_ENABLED", True) Then WarmStart

    ApplyFormats
    gRunning = True
    gLastCapture = TNow()
    gLastRateCalc = TNow()
    gCaptureCount = 0
    gMeasuredHz = 0
    gLastConfig = TNow()
    gLastFlush = TNow()
    gLastPaint = 0

    ArmCaptureTimer
    ScheduleWatchdog

    W Dash, "F23", "RUNNING"
    W Dash, "H23", RateSummary()
    
    LogEvent "", "MONITOR STARTED", "", RateSummary() & _
             "  lookback=" & CfgD("LOOKBACK_MIN", 120) & "min" & _
             "  entry_z=" & CfgD("ENTRY_Z", 2.5) & "  ceiling=" & CfgD("MAX_ENTRY_Z", 4.5) & _
             IIf(gHiRes, "  [high-resolution timer]", "  [OnTime fallback, 1s floor]")
    Exit Sub
Fail:
    gRunning = False
    MsgBox "StartMonitor failed: " & Err.Description, vbExclamation, "TT Z-Monitor"
End Sub

Public Sub StopMonitor()
    On Error Resume Next
    gRunning = False
    DisarmCaptureTimer
    Application.OnTime gNextWatch, "WatchdogTick", , False
    FlushBuffers
    FlushArchives True
    W Dash, "F23", "STOPPED"
    W Dash, "H23", "-"
    LogEvent "", "MONITOR STOPPED", "", ""
    Err.Clear
End Sub

Public Sub Auto_Open()
    ' Deliberately does NOT auto-start: a cold start begins a fresh warm-up and
    ' the operator should decide when that clock begins.
    W Dash, "F23", "STOPPED"
    W Dash, "L3", Format$(Now, "hh:mm:ss")
End Sub

'==============================================================================
' TIMERS
'   High-resolution capture runs on the user32 timer, which is the only way to
'   get below Excel's one-second Application.OnTime floor.  An OnTime watchdog
'   re-arms it, because Excel silently kills API timers when a modal dialog
'   opens.  If SetTimer is unavailable, everything still works at 1 s.
'==============================================================================
Private Sub ArmCaptureTimer()
    Dim ms As Long
    ms = CfgL("CAPTURE_MS", 100)
    If ms < 20 Then ms = 20

    DisarmCaptureTimer
    gTimerMs = ms
    If ms < 1000 Then
        gTimerId = SetTimer(0, 0, ms, AddressOf CaptureProc)
        gHiRes = (gTimerId <> 0)
    Else
        gHiRes = False
    End If
    If Not gHiRes Then ScheduleOnTime
End Sub

Private Sub DisarmCaptureTimer()
    On Error Resume Next
    If gTimerId <> 0 Then KillTimer 0, gTimerId
    gTimerId = 0
    gHiRes = False
    If gNextOnTime <> 0 Then Application.OnTime gNextOnTime, "OnTimeTick", , False
    gNextOnTime = 0
    Err.Clear
End Sub

Private Sub ScheduleOnTime()
    Dim ms As Long
    ms = gTimerMs
    If ms < 1000 Then ms = 1000
    gNextOnTime = Now + CDbl(ms) / 86400000#
    Application.OnTime gNextOnTime, "OnTimeTick"
End Sub

Private Sub ScheduleWatchdog()
    Dim s As Double
    s = CfgD("WATCHDOG_SEC", 5)
    If s < 1 Then s = 1
    gNextWatch = Now + s / 86400#
    Application.OnTime gNextWatch, "WatchdogTick"
End Sub

' user32 timer callback.  It must NEVER raise - an unhandled error here takes
' Excel down with it.
#If VBA7 Then
Public Sub CaptureProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                       ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub CaptureProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                       ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next
    Tick
    Err.Clear
End Sub

Public Sub OnTimeTick()
    If Not gRunning Then Exit Sub
    On Error Resume Next
    Tick
    If Not gHiRes Then ScheduleOnTime
    Err.Clear
End Sub

Public Sub WatchdogTick()
    Dim ms As Long
    If Not gRunning Then Exit Sub
    On Error Resume Next

    ms = CfgL("CAPTURE_MS", 100)
    ' Re-arm if the timer died, or if CAPTURE_MS was changed on Config.
    If ms <> gTimerMs Then
        ArmCaptureTimer
    ElseIf gHiRes And (TNow() - gLastCapture) * 86400# > WorksheetFunction.Max(2, gTimerMs * 10# / 1000#) Then
        ArmCaptureTimer
        LogEvent "", "TIMER RE-ARMED", "", "capture timer had stopped"
    End If

    ' Measure what the capture timer is ACTUALLY doing, rather than assume it
    ' matches CAPTURE_MS. This is the number that shows whether the
    ' high-resolution timer is running or the OnTime fallback took over.
    If gLastRateCalc > 0 Then
        Dim secs As Double
        secs = (TNow() - gLastRateCalc) * 86400#
        If secs > 0.2 Then gMeasuredHz = gCaptureCount / secs
    End If
    gLastRateCalc = TNow()
    gCaptureCount = 0

    W Dash, "L3", Format$(Now, "hh:mm:ss")
    W Dash, "H23", RateSummary()

    If (TNow() - gLastFlush) * 86400# >= CfgL("FLUSH_SEC", 60) Then
        FlushBuffers
        FlushArchives False
        gLastFlush = TNow()
    End If

    ScheduleWatchdog
    Err.Clear
End Sub

Private Function RateSummary() As String
    Dim st As Long
    st = CfgL("STORE_MIN_INTERVAL_MS", 0)
    If Len(gTTState) > 0 Then
        RateSummary = gTTState & "   -   log in to TT; the RTD server is answering"
        Exit Function
    End If
    ' A clamped history gate goes FIRST: it is a config error the operator has
    ' to see, and the rate detail behind it is only ever informational.
    If Len(gHistClamp) > 0 Then
        RateSummary = "CONFIG: " & gHistClamp & "  |  capture " & _
                      CfgL("CAPTURE_MS", 100) & "ms  |  paint " & _
                      CfgL("PRICE_REFRESH_MS", 500) & "ms"
        Exit Function
    End If
    RateSummary = "capture " & CfgL("CAPTURE_MS", 100) & "ms (" & _
                  Format$(gMeasuredHz, "0.0") & "/s measured, " & _
                  IIf(gHiRes, "hi-res", "OnTime fallback") & ")  |  store " & _
                  IIf(st <= 0, "full fidelity", st & "ms") & "  |  paint " & _
                  CfgL("PRICE_REFRESH_MS", 500) & "ms  |  stats " & _
                  CfgD("STATS_REFRESH_MIN", 5) & "min"
End Function

'==============================================================================
' THE TICK
'   Capture + dedupe + store-gate + z + alerts run at CAPTURE_MS (cheap: one
'   range read, no cell writes).  Painting runs at PRICE_REFRESH_MS.
'==============================================================================
Private Sub Tick()
    Dim t As Double, su As Boolean
    Dim vLeg As Variant, vDer As Variant, vInst As Variant, fs As Worksheet
    Dim d0 As Worksheet

    If Not gRunning Then Exit Sub
    If gInCapture Then Exit Sub                  ' no re-entry
    On Error GoTo Done
    gInCapture = True
    ' Excel busy, in cell-edit mode, or showing a dialog: skip this capture
    ' rather than reach into it.
    If Not Application.Ready Then GoTo Done

    t = TNow()
    gLastCapture = t
    gCaptureCount = gCaptureCount + 1

    If (t - gLastConfig) * 86400000# >= CfgL("CONFIG_REFRESH_MS", 1000) Then
        LoadConfig
        LoadSpreadDefs
        gLastConfig = t
    End If

    Set fs = ThisWorkbook.Sheets(SH_FEED)
    Set d0 = ThisWorkbook.Sheets(SH_DASH)
    vLeg = fs.Range(fs.Cells(F_LEG_TOP, 4), fs.Cells(F_LEG_TOP + 2 * MAX_SPREADS - 1, 6)).Value2
    vDer = fs.Range(fs.Cells(F_DER_TOP, 1), fs.Cells(F_DER_TOP + F_DER_N - 1, 6)).Value2

    Dim i As Long
    For i = 1 To MAX_SPREADS
        If S(i).Enabled Then
            CaptureSlot i, vLeg, t
            EvalSlot i, vDer, t
            HandleAlert i, t
        End If
    Next i

    If (t - gLastPaint) * 86400000# >= CfgL("PRICE_REFRESH_MS", 500) Then
        vInst = fs.Range(fs.Cells(F_INST_TOP, 1), fs.Cells(F_INST_TOP + F_INST_N - 1, 8)).Value2
        su = Application.ScreenUpdating
        Application.ScreenUpdating = False
        PaintPrices vInst, vDer
        PaintMonitor
        ' The clock rides the paint cadence, so it ticks once a second. It used
        ' to be written only by the watchdog, which made it a watchdog
        ' heartbeat rather than proof the capture timer was alive.
        W d0, "L3", Format$(Now, "hh:mm:ss")
        Application.ScreenUpdating = su
        gLastPaint = t
    End If

    ' Logging is edge-triggered: one line when a state starts, one when it
    ' clears - never one per poll.
    For i = 1 To MAX_SPREADS
        If S(i).Enabled Then LogEdges i
    Next i

Done:
    gInCapture = False
    If Err.Number <> 0 Then Err.Clear
End Sub

'------------------------------------------------------------------ capture --
Private Sub CaptureSlot(ByVal i As Long, vLeg As Variant, ByVal t As Double)
    Dim ra As Long, rb As Long
    Dim aB As Variant, aA As Variant, aM As Variant
    Dim bB As Variant, bA As Variant, bM As Variant
    Dim mid As Double, qid As String, storeMs As Double, listed As Boolean

    ra = (i - 1) * 2 + 1: rb = ra + 1
    aB = vLeg(ra, 1): aA = vLeg(ra, 2): aM = vLeg(ra, 3)
    bB = vLeg(rb, 1): bA = vLeg(rb, 2): bM = vLeg(rb, 3)
    listed = (S(i).Mode = "LISTED")

    S(i).HaveTouch = IsNum(bB) And IsNum(bA)
    If Not listed Then S(i).HaveTouch = S(i).HaveTouch And IsNum(aB) And IsNum(aA)

    If S(i).HaveTouch Then
        ' sell LegB at the bid and buy LegA at the ask; the ask mirrors it
        S(i).CurBid = Nz(bB) - S(i).Hedge * Nz(aA)
        S(i).CurAsk = Nz(bA) - S(i).Hedge * Nz(aB)
        S(i).CurMid = (S(i).CurBid + S(i).CurAsk) / 2
    End If

    If Not IsNum(bM) Then Exit Sub
    If Not listed And Not IsNum(aM) Then Exit Sub

    ' spread = LegB - HEDGE_RATIO x LegA.  Nothing else: no carry, no swap, no
    ' fair-value term.  Brent/WTI is a RELATED pair, not a basis pair.
    mid = Nz(bM) - S(i).Hedge * Nz(aM)
    S(i).LegAMid = Nz(aM)
    S(i).LegBMid = Nz(bM)

    ' DEDUPE, keyed on BOTH legs' bid AND ask.
    qid = QKey(aB) & "|" & QKey(aA) & "|" & QKey(bB) & "|" & QKey(bA)
    If qid <> S(i).SeenQid Then
        S(i).SeenQid = qid
        S(i).LastQuoteTime = t
        NoteChange i, t
        ArchiveTick i, t, aB, aA, bB, bA, mid
        S(i).PendQid = qid
        S(i).PendV = mid
        S(i).PendT = t
        S(i).HavePend = True
    End If

    ' STORE GATE.  Rate 2: the window holds quote changes, subsampled to at
    ' most one per STORE_MIN_INTERVAL_MS.  Nothing is stored twice, so a poll
    ' faster than the feed can never collapse the variance.
    If S(i).HavePend Then
        If S(i).PendQid <> S(i).StoredQid Then
            storeMs = CfgD("STORE_MIN_INTERVAL_MS", 0)
            If storeMs <= 0 Or S(i).LastStoredT = 0 Or _
               (t - S(i).LastStoredT) * 86400000# >= storeMs Then
                Append i, S(i).PendT, S(i).PendV
                S(i).StoredQid = S(i).PendQid
                S(i).LastStoredT = S(i).PendT
                S(i).HavePend = False
            End If
        Else
            S(i).HavePend = False        ' it changed and changed back
        End If
    End If
End Sub

Private Sub NoteChange(ByVal i As Long, ByVal t As Double)
    If S(i).ChgHead < 0 Then S(i).ChgHead = 0
    S(i).ChgT(S(i).ChgHead) = t
    S(i).ChgHead = (S(i).ChgHead + 1) Mod CHG_RING
    If S(i).ChgCount < CHG_RING Then S(i).ChgCount = S(i).ChgCount + 1
End Sub

'==============================================================================
' EVALUATE  -  z, gates, costs, take-profit, edge filter.  No cell writes.
'==============================================================================
Private Sub EvalSlot(ByVal i As Long, vDer As Variant, ByVal t As Double)
    Dim lookback As Double, refreshMin As Double
    Dim entryZ As Double, maxZ As Double, minSig As Double, maxAbsZ As Double
    Dim minSamp As Long, minHist As Double, thin As Double, staleSec As Double
    Dim lots As Double, comm As Double, edgeMult As Double, targetUsd As Double
    Dim degenerate As Boolean, warming As Boolean, stale As Boolean
    Dim xOwn As Double, xAlt As Double, k As Long, dirSign As Double

    lookback = CfgD("LOOKBACK_MIN", 120)
    refreshMin = CfgD("STATS_REFRESH_MIN", 5)
    entryZ = CfgD("ENTRY_Z", 2.5)
    maxZ = CfgD("MAX_ENTRY_Z", 4.5)
    minSig = CfgD("MIN_SIGMA", 0)
    maxAbsZ = CfgD("MAX_ABS_Z", 25)
    minSamp = CfgL("MIN_SAMPLES", 300)
    minHist = CfgD("MIN_HISTORY_MIN", 120)
    thin = CfgD("THIN_FEED_QPM", 6)
    staleSec = CfgD("STALE_SEC", 15)
    lots = CfgD("LOTS", 1)
    comm = CfgD("COMMISSION_PER_LOT_ROUND_TURN", 3.82)
    edgeMult = CfgD("EDGE_MULTIPLE", 1.5)
    targetUsd = CfgD("TARGET_NET_USD", 0)

    ' Time-based eviction runs every capture, not only when a quote arrives:
    ' a dead feed must drain its window and go cold, not freeze a stale z.
    EvictOld i, t, lookback

    ' ---- statistics refresh (frozen in between) --------------------------
    If S(i).Count >= 2 Then
        If Not S(i).StatsValid Or (t - S(i).LastStatsTime) * 1440# >= refreshMin Then
            RecomputeStats i, t, lookback
        End If
    Else
        ' The window drained below two samples - a frozen mean from before that
        ' is not a reference, it is a fossil.
        S(i).StatsValid = False
    End If

    S(i).Elapsed = WindowElapsedMin(i, t)
    S(i).Qpm = QuotesPerMin(i, t)

    stale = (S(i).LastQuoteTime = 0) Or ((t - S(i).LastQuoteTime) * 86400# > staleSec)
    If stale Then
        S(i).FeedStatus = "STALE"
    ElseIf S(i).Qpm <= thin Then
        S(i).FeedStatus = "THIN"
    Else
        S(i).FeedStatus = "OK"
    End If

    ' ---- two warm-up gates, BOTH required --------------------------------
    warming = (S(i).Count < minSamp) Or (S(i).Elapsed < minHist)
    ' Report the SHORTFALL, not two rounded numbers. Formatting both sides at
    ' "0" printed a satisfied-looking "120/120 min" while 119.99 < 120 still
    ' held, which hid an unreachable gate behind a display artefact.
    If S(i).Count < minSamp And S(i).Elapsed < minHist Then
        S(i).Gate = "both: " & S(i).Count & "/" & minSamp & " samples, " & _
                    HistShort(S(i).Elapsed, minHist)
    ElseIf S(i).Count < minSamp Then
        S(i).Gate = "samples: " & S(i).Count & "/" & minSamp
    ElseIf S(i).Elapsed < minHist Then
        S(i).Gate = "history: " & HistShort(S(i).Elapsed, minHist)
    Else
        S(i).Gate = "ready"
    End If

    ' ---- z ---------------------------------------------------------------
    S(i).HaveZ = False: S(i).Z = 0: degenerate = False
    If S(i).StatsValid And S(i).HaveTouch Then
        If S(i).Sigma <= 0 Or S(i).Sigma < minSig Then
            degenerate = True
        Else
            S(i).Z = (S(i).CurMid - S(i).Mean) / S(i).Sigma
            If Abs(S(i).Z) > maxAbsZ Then
                degenerate = True
            Else
                S(i).HaveZ = True
            End If
        End If
    End If

    If stale Then
        S(i).Signal = "STALE FEED"
    ElseIf warming Then
        S(i).Signal = "WARMING UP"
    ElseIf degenerate Or Not S(i).HaveZ Then
        S(i).Signal = "NO USABLE Z"          ' a blank is honest; a huge number
    ElseIf Abs(S(i).Z) >= maxZ Then          ' reads as opportunity
        S(i).Signal = "ABOVE CEILING"
    ElseIf Abs(S(i).Z) >= entryZ Then
        S(i).Signal = "ENTRY BAND"
    Else
        S(i).Signal = "FLAT"
    End If
    S(i).HaveTrade = S(i).HaveZ And Not warming And Not stale And Not degenerate

    ' ---- costs -----------------------------------------------------------
    S(i).Commission = S(i).Contracts * comm * lots
    xOwn = 0
    If S(i).HaveTouch Then xOwn = (S(i).CurAsk - S(i).CurBid) * S(i).UsdPerPoint * lots

    xAlt = -1
    If S(i).AltRow > 0 Then
        k = S(i).AltRow - F_DER_TOP + 1
        If k >= 1 And k <= F_DER_N Then
            If IsNum(vDer(k, 6)) Then xAlt = Nz(vDer(k, 6)) * S(i).UsdPerPoint * lots
        End If
    End If

    If S(i).Mode = "LISTED" Then
        S(i).XLst = IIf(S(i).HaveTouch, xOwn, "")
        S(i).XLeg = IIf(xAlt >= 0, xAlt, "n/a")
    Else
        S(i).XLeg = IIf(S(i).HaveTouch, xOwn, "")
        S(i).XLst = IIf(xAlt >= 0, xAlt, "n/a")
    End If
    S(i).Total = S(i).Commission + xOwn

    ' ---- the mean, expressed on each touch --------------------------------
    ' Mean and Sigma are computed on MIDS, but you never deal at the mid: you
    ' sell the bid to get short and buy the ask to get long. Offset the mean by
    ' half the LIVE gap so each side reads as the price that side would deal at
    ' if the spread were sitting exactly on its mean. Deliberately outside the
    ' HaveTrade gate below - it is worth seeing while the window is still
    ' warming, which is when the operator is deciding whether to bother.
    If S(i).StatsValid And S(i).HaveTouch Then
        S(i).MeanBid = S(i).Mean - (S(i).CurAsk - S(i).CurBid) / 2
        S(i).MeanAsk = S(i).Mean + (S(i).CurAsk - S(i).CurBid) / 2
        S(i).HaveMeanTouch = True
    Else
        S(i).MeanBid = 0: S(i).MeanAsk = 0: S(i).HaveMeanTouch = False
    End If

    ' ---- direction, break-even, take-profit, edge filter -----------------
    If Not S(i).HaveTrade Or Not S(i).HaveTouch Then
        S(i).DirTxt = "-": S(i).Beyond = "-": S(i).Verdict = "n/a"
        S(i).TpRule = "-": S(i).Notional = 0: S(i).WinUsd = 0
        S(i).CostUnits = 0: S(i).WinUnits = 0
        Exit Sub
    End If

    If S(i).Z > 0 Then
        dirSign = -1: S(i).DirTxt = "SHORT the spread": S(i).Entry = S(i).CurBid
    Else
        dirSign = 1: S(i).DirTxt = "LONG the spread": S(i).Entry = S(i).CurAsk
    End If

    ' TAKE PROFIT = Mid + Costs + (WIN_PCT x Notional), signed by direction.
    ' A spread is not an outright, so "notional" has no single meaning - what
    ' the percentage is taken OF is a Config choice, and the three answers
    ' differ by roughly 10x. TP in sigma is the reachability check.
    '
    ' ANCHOR THE LEVELS TO THE MID, NOT TO THE ENTRY TOUCH. Entry is a touch,
    ' and a touch already carries half the width (CurBid = CurMid - width/2).
    ' Adding the FULL round-trip cost to it charged 1.5 widths of crossing
    ' where only 1.0 is real, pushing break-even and take-profit half a width
    ' too far - and TargetZ then compared a touch-scale TP against Mean, which
    ' is mid-scale. Mean, Sigma, z and Capture all live on the mid, so the
    ' levels do too: a mid move of CostUnits is exactly what pays the round
    ' trip. Entry stays the touch, because that is the price you actually get.
    S(i).CostUnits = S(i).Total / (S(i).UsdPerPoint * lots)
    S(i).Notional = NotionalUsd(i, lots)

    If UCase$(CfgS("TP_MODE", "PCT_NOTIONAL")) = "TARGET_USD" Then
        S(i).WinUsd = targetUsd
        S(i).TpRule = "Mid + Costs + TARGET_NET_USD"
    Else
        S(i).WinUsd = CfgD("WIN_PCT", 0.01) * S(i).Notional
        S(i).TpRule = "Mid + Costs + " & Format$(CfgD("WIN_PCT", 0.01) * 100, "0.##") & _
                      "% x " & UCase$(CfgS("NOTIONAL_BASIS", "SPREAD_VALUE"))
    End If
    S(i).WinUnits = S(i).WinUsd / (S(i).UsdPerPoint * lots)

    ' TPDist is the required move OF THE MID, which is what Sigma measures.
    S(i).TPDist = S(i).CostUnits + S(i).WinUnits
    S(i).BE = S(i).CurMid + dirSign * S(i).CostUnits
    S(i).TP = S(i).CurMid + dirSign * S(i).TPDist

    ' To work the exit order, cross back: buy at the ask to close a short, sell
    ' at the bid to close a long. That touch is TP - dirSign x (width / 2).
    S(i).TPSig = S(i).TPDist / S(i).Sigma
    S(i).TargetZ = (S(i).TP - S(i).Mean) / S(i).Sigma

    ' A take-profit beyond the mean needs the spread to OVERSHOOT, which is a
    ' different bet from the one the z-score measured.
    If (dirSign < 0 And S(i).TP < S(i).Mean) Or (dirSign > 0 And S(i).TP > S(i).Mean) Then
        S(i).Beyond = "YES - needs overshoot past the mean"
    Else
        S(i).Beyond = "no - inside the mean"
    End If

    S(i).Capture = 0.5 * Abs(S(i).Z) * S(i).Sigma * S(i).UsdPerPoint * lots
    S(i).Required = edgeMult * S(i).Total
    S(i).Verdict = IIf(S(i).Capture >= S(i).Required, "PASS", "FAIL")
    S(i).MinSigPass = S(i).Required / (0.5 * Abs(S(i).Z) * S(i).UsdPerPoint * lots)
End Sub

' What WIN_PCT is taken OF. The three answers are not close to each other:
' at 1 lot with CL near $60 and BZ-CL near $7.20 they are roughly $7,200,
' $60,000 and $127,000.
Private Function NotionalUsd(ByVal i As Long, ByVal lots As Double) As Double
    Dim basis As String, legA As Double, legB As Double
    basis = UCase$(CfgS("NOTIONAL_BASIS", "SPREAD_VALUE"))
    legA = Abs(S(i).LegAMid): legB = Abs(S(i).LegBMid)

    Select Case basis
        Case "ONE_LEG"
            ' The crude leg carries the contract value the spread is quoted against.
            If legA > 0 Then
                NotionalUsd = legA * S(i).UsdPerPoint * lots
            Else
                NotionalUsd = legB * S(i).UsdPerPoint * lots
            End If
        Case "BOTH_LEGS"
            NotionalUsd = (legA + legB) * S(i).UsdPerPoint * lots
        Case Else                        ' SPREAD_VALUE
            NotionalUsd = Abs(S(i).CurMid) * S(i).UsdPerPoint * lots
    End Select
End Function

'==============================================================================
' CIRCULAR BUFFER
'==============================================================================
Private Sub InitBuffers()
    Dim i As Long, cap As Long
    cap = CfgL("BUFFER_CAPACITY", 80000)
    If cap < 1000 Then cap = 1000
    For i = 1 To MAX_SPREADS
        If S(i).Cap <> cap Then
            S(i).Cap = cap
            ReDim S(i).BufT(0 To cap - 1)
            ReDim S(i).BufV(0 To cap - 1)
            S(i).Head = 0
            S(i).Count = 0
        End If
        ReDim S(i).ChgT(0 To CHG_RING - 1)
        S(i).ChgHead = 0: S(i).ChgCount = 0
        ReDim S(i).Arc(0 To CfgL("TICK_ARCHIVE_BATCH", 500))
        S(i).ArcN = 0
        S(i).Armed = True
        S(i).CeilArmed = True
        S(i).StatsValid = False
        S(i).StoredQid = "": S(i).SeenQid = "": S(i).HavePend = False
        S(i).LastStoredT = 0: S(i).LastQuoteTime = 0
    Next i
End Sub

Private Sub Append(ByVal i As Long, ByVal t As Double, ByVal v As Double)
    S(i).BufT(S(i).Head) = t
    S(i).BufV(S(i).Head) = v
    S(i).Head = (S(i).Head + 1) Mod S(i).Cap
    If S(i).Count < S(i).Cap Then S(i).Count = S(i).Count + 1
End Sub

' Time-based eviction: a dead feed drains the window and goes cold, rather
' than freezing a stale z in place.
Private Sub EvictOld(ByVal i As Long, ByVal t As Double, ByVal lookback As Double)
    Dim cutoff As Double
    cutoff = t - lookback / 1440#
    Do While S(i).Count > 0
        If S(i).BufT(OldestIdx(i)) >= cutoff Then Exit Do
        S(i).Count = S(i).Count - 1
    Loop
End Sub

Private Function OldestIdx(ByVal i As Long) As Long
    OldestIdx = (S(i).Head - S(i).Count + S(i).Cap) Mod S(i).Cap
End Function

' Full recompute over the live window.  Cheap, because it runs only every
' STATS_REFRESH_MIN minutes - and it clears any accumulated float drift.
Private Sub RecomputeStats(ByVal i As Long, ByVal t As Double, ByVal lookback As Double)
    Dim k As Long, idx As Long, n As Long
    Dim sm As Double, ss As Double, v As Double, mu As Double, va As Double
    Dim cutoff As Double
    cutoff = t - lookback / 1440#

    For k = 0 To S(i).Count - 1
        idx = (OldestIdx(i) + k) Mod S(i).Cap
        If S(i).BufT(idx) >= cutoff Then
            v = S(i).BufV(idx)
            n = n + 1: sm = sm + v: ss = ss + v * v
        End If
    Next k

    S(i).StatsN = n
    S(i).LastStatsTime = t
    If n >= 2 Then
        mu = sm / n
        va = (ss - n * mu * mu) / (n - 1)
        If va < 0 Then va = 0
        S(i).Mean = mu
        S(i).Sigma = Sqr(va)
        S(i).StatsValid = True
    Else
        S(i).Mean = 0: S(i).Sigma = 0: S(i).StatsValid = False
    End If
End Sub

' "119.99/120.00 min, 0.007 short" - the shortfall carries enough precision
' that it can never itself round away to zero while the gate is still failing.
Private Function HistShort(ByVal elapsed As Double, ByVal minHist As Double) As String
    HistShort = Format$(elapsed, "0.00") & "/" & Format$(minHist, "0.00") & " min, " & _
                Format$(minHist - elapsed, "0.000") & " short"
End Function

Private Function WindowElapsedMin(ByVal i As Long, ByVal t As Double) As Double
    If S(i).Count = 0 Then Exit Function
    WindowElapsedMin = (t - S(i).BufT(OldestIdx(i))) * 1440#
End Function

' Observed quote CHANGES per minute - the true feed rate, independent of
' whatever STORE_MIN_INTERVAL_MS is set to.
Private Function QuotesPerMin(ByVal i As Long, ByVal t As Double) As Double
    Dim k As Long, idx As Long, n As Long, cutoff As Double
    cutoff = t - 1# / 1440#
    For k = S(i).ChgCount - 1 To 0 Step -1
        idx = (S(i).ChgHead - S(i).ChgCount + k + 2 * CHG_RING) Mod CHG_RING
        If S(i).ChgT(idx) < cutoff Then Exit For
        n = n + 1
    Next k
    QuotesPerMin = n
End Function

'==============================================================================
' SUBSAMPLING CHECK  -  run this after a week of full-fidelity storage to see,
' on your own data, what STORE_MIN_INTERVAL_MS would actually cost you.
' Alt+F8 > CheckSubsamplingLoss.  Results go to the Log sheet.
'==============================================================================
Public Sub CheckSubsamplingLoss()
    Dim i As Long, j As Long, ivl As Variant
    Dim n As Long, sm As Double, ss As Double, mu As Double, sg As Double
    Dim k As Long, idx As Long, lastT As Double, v As Double, msg As String

    ivl = Array(0#, 250#, 500#, 1000#, 2000#, 5000#)
    For i = 1 To MAX_SPREADS
        If Not S(i).Enabled Or S(i).Count < 2 Then GoTo NextSlot
        msg = ""
        For j = LBound(ivl) To UBound(ivl)
            n = 0: sm = 0: ss = 0: lastT = 0
            For k = 0 To S(i).Count - 1
                idx = (OldestIdx(i) + k) Mod S(i).Cap
                If ivl(j) <= 0 Or lastT = 0 Or _
                   (S(i).BufT(idx) - lastT) * 86400000# >= ivl(j) Then
                    v = S(i).BufV(idx)
                    n = n + 1: sm = sm + v: ss = ss + v * v
                    lastT = S(i).BufT(idx)
                End If
            Next k
            If n >= 2 Then
                mu = sm / n
                sg = Sqr(WorksheetFunction.Max(0, (ss - n * mu * mu) / (n - 1)))
                msg = msg & IIf(ivl(j) = 0, "full", Format$(ivl(j), "0") & "ms") & _
                      ": n=" & n & " mean=" & Format$(mu, "0.000000") & _
                      " sigma=" & Format$(sg, "0.000000") & "   "
            End If
        Next j
        LogEvent S(i).Name, "SUBSAMPLING CHECK", "", msg
NextSlot:
    Next i
    MsgBox "Subsampling comparison written to the Log sheet." & vbCrLf & vbCrLf & _
           "If sigma is materially unchanged at 250-1000 ms, set " & _
           "STORE_MIN_INTERVAL_MS on Config and reclaim the memory.", _
           vbInformation, "TT Z-Monitor"
End Sub

'==============================================================================
' PAINT  -  every write goes through W(), which compares before it writes
'==============================================================================
Private Sub PaintPrices(vInst As Variant, vDer As Variant)
    ' ---- the user's Dashboard, cell for cell --------------------------
    Dim d As Worksheet
    Set d = ThisWorkbook.Sheets(SH_DASH)

    gTTState = TTStateFrom(vInst)

    W d, "C6", TxtOf(vInst(I_RB, 2))
    W d, "C7", TxtOf(vInst(I_HO, 2))
    W d, "C8", TxtOf(vInst(I_BZ, 2))
    W d, "C9", TxtOf(vInst(I_CL, 2))
    W d, "C10", TxtOf(vInst(I_CLBZ, 2))
    W d, "C12", TxtOf(vInst(I_HOCL, 2))

    ' left block: HO, CL, then the HO|CL crack
    PaintLeg d, 7, CB_H, vInst, I_HO
    PaintLeg d, 8, CB_H, vInst, I_CL
    PaintSpreadRow d, 9, CB_H, 2, vInst, I_HO, I_CL, 42#

    ' right block: BZ, CL, then BZ - CL
    PaintLeg d, 7, CB_P, vInst, I_BZ
    PaintLeg d, 8, CB_P, vInst, I_CL
    PaintSpreadRow d, 9, CB_P, 1, vInst, I_BZ, I_CL, 1#

    ' second block: RB, HO, CL, then the 3:2:1 pack
    PaintLeg d, 12, CB_H, vInst, I_RB
    PaintLeg d, 13, CB_H, vInst, I_HO
    PaintLeg d, 14, CB_H, vInst, I_CL
    Paint321 d, 15, CB_H, 3, vInst

    PaintDetailPrices vInst, vDer
End Sub

' TT returns a status string where an instrument id belongs when it cannot
' resolve one - "Not_Connected" when the platform is not logged in. Surface
' TT's own words rather than leaving a screen of blanks, which looks identical
' to a wrong symbol or a quiet market.
Private Function TTStateFrom(vInst As Variant) As String
    Dim k As Long, id As String, good As Long, bad As Long, firstBad As String
    For k = 1 To F_INST_N
        id = TxtOf(vInst(k, 2))
        If Len(id) > 0 Then
            If UCase$(Left$(id, 4)) = "NOT_" Then
                bad = bad + 1
                If Len(firstBad) = 0 Then firstBad = id
            Else
                good = good + 1
            End If
        End If
    Next k

    If bad > 0 And good = 0 Then
        TTStateFrom = "TT: " & Replace(firstBad, "_", " ")
    ElseIf bad > 0 Then
        TTStateFrom = "TT: " & bad & " of " & (good + bad) & " unresolved"
    ElseIf good = 0 Then
        TTStateFrom = "WAITING FOR TT"
    Else
        TTStateFrom = ""
    End If
End Function

Private Sub PaintLeg(d As Worksheet, ByVal r As Long, ByVal c0 As Long, _
                     vInst As Variant, ByVal idx As Long)
    WC d, r, c0, vInst(idx, 3)                            ' Bid
    WC d, r, c0 + 1, vInst(idx, 4)                        ' Ask
    WC d, r, c0 + 2, Sub2(vInst(idx, 4), vInst(idx, 3))   ' Gap
    WC d, r, c0 + 3, vInst(idx, 5)                        ' High
    WC d, r, c0 + 4, vInst(idx, 6)                        ' Low
End Sub

' The spread row. Its Bid and Ask follow SPREAD_QUOTE_CONVENTION:
'   CROSSING  (default) - sell LegB at the bid and buy LegA at the ask, and
'                         mirror it for the ask. Gap is then the real cost of
'                         crossing, and it is always positive.
'   SAME_SIDE           - bid minus bid and ask minus ask, as the original
'                         sheet computed rows 9. Gap is a DIFFERENCE of gaps
'                         there, which understates the crossing cost and can
'                         even go negative.
Private Sub PaintSpreadRow(d As Worksheet, ByVal r As Long, ByVal c0 As Long, _
                           ByVal slot As Long, vInst As Variant, _
                           ByVal idxB As Long, ByVal idxA As Long, ByVal scaleB As Double)
    Dim bid As Variant, ask As Variant

    If SameSideQuotes() Then
        bid = Comb2(vInst(idxB, 3), scaleB, vInst(idxA, 3))
        ask = Comb2(vInst(idxB, 4), scaleB, vInst(idxA, 4))
    ElseIf S(slot).HaveTouch Then
        bid = S(slot).CurBid
        ask = S(slot).CurAsk
    Else
        bid = "": ask = ""
    End If

    WC d, r, c0, bid
    WC d, r, c0 + 1, ask
    WC d, r, c0 + 2, Sub2(ask, bid)
    WC d, r, c0 + 3, Comb2(vInst(idxB, 5), scaleB, vInst(idxA, 5))   ' High
    WC d, r, c0 + 4, Comb2(vInst(idxB, 6), scaleB, vInst(idxA, 6))   ' Low
End Sub

' The 3:2:1 pack: (2 x RB x 42 + 1 x HO x 42 - 3 x CL) / 3
Private Sub Paint321(d As Worksheet, ByVal r As Long, ByVal c0 As Long, _
                     ByVal slot As Long, vInst As Variant)
    Dim bid As Variant, ask As Variant

    If SameSideQuotes() Then
        bid = Pack321(vInst(I_RB, 3), vInst(I_HO, 3), vInst(I_CL, 3))
        ask = Pack321(vInst(I_RB, 4), vInst(I_HO, 4), vInst(I_CL, 4))
    ElseIf S(slot).HaveTouch Then
        bid = S(slot).CurBid
        ask = S(slot).CurAsk
    Else
        bid = "": ask = ""
    End If

    WC d, r, c0, bid
    WC d, r, c0 + 1, ask
    WC d, r, c0 + 2, Sub2(ask, bid)
    WC d, r, c0 + 3, Pack321(vInst(I_RB, 5), vInst(I_HO, 5), vInst(I_CL, 5))
    WC d, r, c0 + 4, Pack321(vInst(I_RB, 6), vInst(I_HO, 6), vInst(I_CL, 6))
End Sub

Private Function SameSideQuotes() As Boolean
    SameSideQuotes = (UCase$(CfgS("SPREAD_QUOTE_CONVENTION", "CROSSING")) = "SAME_SIDE")
End Function

Private Function Pack321(a As Variant, b As Variant, c As Variant) As Variant
    If IsNum(a) And IsNum(b) And IsNum(c) Then
        Pack321 = ((Nz(a) * 2 * 42) + (Nz(b) * 1 * 42) - (Nz(c) * 3)) / 3
    Else
        Pack321 = ""
    End If
End Function

Private Function Comb2(b As Variant, ByVal scaleB As Double, a As Variant) As Variant
    If IsNum(a) And IsNum(b) Then
        Comb2 = Nz(b) * scaleB - Nz(a)
    Else
        Comb2 = ""
    End If
End Function

Private Function Sub2(x As Variant, y As Variant) As Variant
    If IsNum(x) And IsNum(y) Then Sub2 = Nz(x) - Nz(y) Else Sub2 = ""
End Function

Private Sub PaintDetailPrices(vInst As Variant, vDer As Variant)
    Dim d As Worksheet, r As Long, k As Long
    Dim legRow As Variant, usdpt As Variant, lots As Double
    On Error Resume Next
    Set d = ThisWorkbook.Sheets(SH_DET)
    If d Is Nothing Then Exit Sub
    lots = CfgD("LOTS", 1)

    legRow = Array(I_RB, I_HO, I_CL, I_BZ)
    For k = 0 To 3
        r = 6 + k
        W d, "C" & r, TxtOf(vInst(legRow(k), 2))
        W d, "H" & r, vInst(legRow(k), 3)
        W d, "I" & r, vInst(legRow(k), 4)
        W d, "J" & r, vInst(legRow(k), 5)
        W d, "K" & r, vInst(legRow(k), 6)
        W d, "L" & r, vInst(legRow(k), 7)
        W d, "M" & r, vInst(legRow(k), 8)
    Next k

    ' USD per 1.00 of spread is 1,000 per lot, except the 3:2:1 pack, which is
    ' three crude-equivalents: 3,000.
    usdpt = Array(1000#, 1000#, 3000#, 1000#, 1000#)
    For k = 1 To F_DER_N
        r = 12 + k
        W d, "H" & r, vDer(k, 3)
        W d, "I" & r, vDer(k, 4)
        W d, "L" & r, vDer(k, 5)
        W d, "M" & r, vDer(k, 6)
    Next k
    Err.Clear
End Sub

Private Sub PaintMonitor()
    ' ---- the z-score block added below the user's own layout -----------
    Dim d As Worksheet, k As Long, r As Long, slot As Long
    Dim rows_ As Variant, slots_ As Variant
    Set d = ThisWorkbook.Sheets(SH_DASH)

    rows_ = Array(ZR1, ZR2, ZR3)
    slots_ = Array(2, 1, 3)          ' display order: HO|CL, BZ-CL, 3:2:1

    For k = 0 To 2
        r = rows_(k): slot = slots_(k)

        ' Live Z-score, Mean, Std Dev. Blank whenever the window is not
        ' usable - a blank is honest; a number would not be.
        If S(slot).HaveTrade Then WC d, r, 8, S(slot).Z Else WC d, r, 8, ""
        If S(slot).StatsValid Then
            WC d, r, 9, S(slot).Mean
            WC d, r, 10, S(slot).Sigma
        Else
            WC d, r, 9, "": WC d, r, 10, ""
        End If
        WC d, r, 11, S(slot).Signal
        WC d, r, 12, WindowText(slot)

        ' Columns G and O are hidden in this sheet, so the block uses
        ' F, H, I, J, K, L and N, P, Q, R, S - the same visible grid the
        ' blocks above use.
        If S(slot).HaveTrade And S(slot).HaveTouch Then
            WC d, r, 14, S(slot).Entry
            WC d, r, 16, S(slot).Total
            WC d, r, 17, S(slot).WinUsd
            WC d, r, 18, S(slot).TP
            WC d, r, 19, S(slot).TPSig
        Else
            WC d, r, 14, "": WC d, r, 16, IIf(S(slot).HaveTouch, S(slot).Total, "")
            WC d, r, 17, "": WC d, r, 18, "": WC d, r, 19, ""
        End If
    Next k

    W d, "F23", IIf(gRunning, "RUNNING", "STOPPED")
    W d, "H23", RateSummary()

    PaintDetail
End Sub

' The compact status shown in the Window column: whichever of stale feed,
' warm-up gate, or sample count and feed rate the operator most needs.
Private Function WindowText(ByVal i As Long) As String
    If Len(gTTState) > 0 Then
        WindowText = gTTState
    ElseIf S(i).FeedStatus = "STALE" Then
        WindowText = "STALE FEED"
    ElseIf S(i).Gate <> "ready" Then
        WindowText = S(i).Gate
    Else
        WindowText = S(i).Count & " samples / " & Format$(S(i).Qpm, "0") & " q-min"
    End If
End Function

Private Sub PaintDetail()
    Dim d As Worksheet, i As Long, c As String
    Dim lots As Double, refreshMin As Double, ageMin As Double, t As Double
    On Error Resume Next
    Set d = ThisWorkbook.Sheets(SH_DET)
    If d Is Nothing Then Exit Sub
    lots = CfgD("LOTS", 1)
    refreshMin = CfgD("STATS_REFRESH_MIN", 5)
    t = TNow()

    For i = 1 To MAX_SPREADS
        c = Mid$(SLOT_COL, i, 1)

        If Not S(i).Enabled Then
            BlankSlot d, c
            W d, c & R_SIGNAL, "OFF"
            GoTo NextSlot
        End If

        W d, c & R_NAME, S(i).Name
        W d, c & R_DEF, S(i).LegLabelB & IIf(S(i).Mode = "LISTED", "", _
              "  -  " & Format$(S(i).Hedge, "0.####") & " x " & S(i).LegLabelA)
        W d, c & R_MODE, S(i).Mode
        W d, c & R_BETA, Format$(S(i).Hedge, "0.0000") & "   [" & S(i).BetaStamp & "]"
        W d, c & R_USDPT, S(i).UsdPerPoint

        If S(i).HaveTouch Then
            W d, c & R_BID, S(i).CurBid
            W d, c & R_ASK, S(i).CurAsk
            W d, c & R_MID, S(i).CurMid
            W d, c & R_WIDTH, S(i).CurAsk - S(i).CurBid
        Else
            W d, c & R_BID, "": W d, c & R_ASK, "": W d, c & R_MID, "": W d, c & R_WIDTH, ""
        End If

        W d, c & R_SAMP, S(i).Count
        W d, c & R_ELAP, S(i).Elapsed
        W d, c & R_RATE, S(i).Qpm
        W d, c & R_FSTAT, S(i).FeedStatus
        W d, c & R_GATE, S(i).Gate

        If S(i).StatsValid Then
            ageMin = (t - S(i).LastStatsTime) * 1440#
            W d, c & R_MEAN, S(i).Mean
            W d, c & R_SIG, S(i).Sigma
            W d, c & R_SIGUSD, S(i).Sigma * S(i).UsdPerPoint * lots
            W d, c & R_AGE, Format$(ageMin, "0.0") & " old / next in " & _
                            Format$(WorksheetFunction.Max(0, refreshMin - ageMin), "0.0")
        Else
            W d, c & R_MEAN, "": W d, c & R_SIG, "": W d, c & R_SIGUSD, "": W d, c & R_AGE, "-"
        End If

        If S(i).HaveMeanTouch Then
            ' SHORT is high -> low: sell the bid now, buy the ask back at the mean.
            W d, c & R_SH_ENT, S(i).MeanBid
            W d, c & R_SH_EXIT, S(i).MeanAsk
            ' LONG is low -> high: buy the ask now, sell the bid back at the mean.
            W d, c & R_LG_ENT, S(i).MeanAsk
            W d, c & R_LG_EXIT, S(i).MeanBid
        Else
            W d, c & R_SH_ENT, "": W d, c & R_SH_EXIT, ""
            W d, c & R_LG_ENT, "": W d, c & R_LG_EXIT, ""
        End If

        If S(i).HaveTrade Then W d, c & R_Z, S(i).Z Else W d, c & R_Z, ""
        W d, c & R_SIGNAL, S(i).Signal

        W d, c & R_COMM, S(i).Commission
        W d, c & R_XLEG, S(i).XLeg
        W d, c & R_XLST, S(i).XLst
        W d, c & R_TOT, IIf(S(i).HaveTouch, S(i).Total, "")

        If S(i).HaveTrade And S(i).HaveTouch Then
            W d, c & R_DIR, S(i).DirTxt
            W d, c & R_ENTRY, S(i).Entry
            W d, c & R_TPRULE, S(i).TpRule
            W d, c & R_NOTIONAL, S(i).Notional
            W d, c & R_WINUSD, S(i).WinUsd
            W d, c & R_COSTU, S(i).CostUnits
            W d, c & R_WINU, S(i).WinUnits
            W d, c & R_BE, S(i).BE
            W d, c & R_TP, S(i).TP
            W d, c & R_TPD, S(i).TPDist
            W d, c & R_TPS, S(i).TPSig
            W d, c & R_TZ, S(i).TargetZ
            W d, c & R_BEYOND, S(i).Beyond
            W d, c & R_CAP, S(i).Capture
            W d, c & R_REQ, S(i).Required
            W d, c & R_VERD, S(i).Verdict
            W d, c & R_MINSIG, S(i).MinSigPass
        Else
            W d, c & R_DIR, "-": W d, c & R_ENTRY, ""
            W d, c & R_TPRULE, "-": W d, c & R_NOTIONAL, "": W d, c & R_WINUSD, ""
            W d, c & R_COSTU, "": W d, c & R_WINU, ""
            W d, c & R_BE, "": W d, c & R_TP, "": W d, c & R_TPD, ""
            W d, c & R_TPS, "": W d, c & R_TZ, "": W d, c & R_BEYOND, "-"
            W d, c & R_CAP, "": W d, c & R_REQ, "": W d, c & R_VERD, "n/a"
            W d, c & R_MINSIG, ""
        End If
NextSlot:
    Next i
    Err.Clear
End Sub

Private Sub BlankSlot(d As Worksheet, ByVal c As String)
    Dim r As Variant, k As Long
    r = Array(R_NAME, R_DEF, R_MODE, R_BETA, R_USDPT, R_BID, R_ASK, R_MID, R_WIDTH, _
              R_SAMP, R_ELAP, R_GATE, R_RATE, R_FSTAT, R_MEAN, R_SIG, R_SIGUSD, R_AGE, _
              R_Z, R_DIR, R_ENTRY, R_COMM, R_XLEG, R_XLST, R_TOT, R_TPRULE, R_NOTIONAL, _
              R_WINUSD, R_COSTU, R_WINU, R_BE, R_TP, R_TPD, R_TPS, R_TZ, R_BEYOND, _
              R_CAP, R_REQ, R_VERD, R_MINSIG, _
              R_SH_ENT, R_SH_EXIT, R_LG_ENT, R_LG_EXIT)
    For k = LBound(r) To UBound(r)
        W d, c & r(k), ""
    Next k
End Sub

'==============================================================================
' AUDIBLE ALERT  -  fire on the CROSSING, not on the state.  Runs at capture
' rate, which is the whole point of capturing at 100 ms.
'==============================================================================
Private Sub HandleAlert(ByVal i As Long, ByVal t As Double)
    Dim az As Double, margin As Double, cool As Double
    Dim entryZ As Double, maxZ As Double

    entryZ = CfgD("ENTRY_Z", 2.5)
    maxZ = CfgD("MAX_ENTRY_Z", 4.5)
    margin = CfgD("ALERT_REARM_MARGIN", 0.25)
    cool = CfgD("ALERT_COOLDOWN_SEC", 60)
    az = Abs(S(i).Z)

    ' Re-arm on hysteresis regardless of mute, so unmuting does not immediately
    ' fire on a condition that has been true for an hour.
    If Not S(i).HaveZ Or az < entryZ - margin Then S(i).Armed = True
    If Not S(i).HaveZ Or az < maxZ - margin Then S(i).CeilArmed = True

    If Not S(i).HaveTrade Then Exit Sub            ' silent while not tradable
    If Not CfgB("ALERT_ENABLED", True) Then Exit Sub

    ' --- ceiling crossing: "stand down", not "get in" ---------------------
    If az >= maxZ Then
        If S(i).CeilArmed And (t - S(i).LastCeilAlert) * 86400# >= cool Then
            Ting CfgS("ALERT_SOUND_PATH_CEILING", "C:\Windows\Media\notify.wav")
            Speak "Ceiling " & S(i).Name
            S(i).CeilArmed = False
            S(i).LastCeilAlert = t
            S(i).Armed = False          ' an entry ting must not follow it down
            LogEvent S(i).Name, "ALERT ceiling", Format$(S(i).Z, "0.00"), "stand down"
        End If
        Exit Sub
    End If

    ' --- entry-band crossing ----------------------------------------------
    If az < entryZ Then Exit Sub
    If Not S(i).Armed Then Exit Sub
    If (t - S(i).LastAlert) * 86400# < cool Then Exit Sub
    If CfgB("ALERT_ONLY_IF_EDGE_PASSES", True) And S(i).Verdict <> "PASS" Then Exit Sub

    Ting CfgS("ALERT_SOUND_PATH", "C:\Windows\Media\chimes.wav")
    Speak S(i).Name
    S(i).Armed = False
    S(i).LastAlert = t
    LogEvent S(i).Name, "ALERT entry", Format$(S(i).Z, "0.00"), "edge=" & S(i).Verdict
End Sub

Private Sub Ting(ByVal path As String)
    On Error Resume Next
    If Len(path) > 0 Then
        PlaySound path, 0, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT
        If Err.Number <> 0 Then Err.Clear: Beep
    Else
        Beep
    End If
End Sub

Private Sub Speak(ByVal what As String)
    On Error Resume Next
    If CfgB("ALERT_SPEAK", False) Then Application.Speech.Speak what, True
    Err.Clear
End Sub

' Alt+F8 > TestChime. Plays the entry tone, then the ceiling tone, so the
' sound path can be checked without waiting for warm-up and a real crossing.
' Alerts are silent during warm-up by design, which would otherwise make this
' the last thing verified rather than the first.
Public Sub TestChime()
    Dim p1 As String, p2 As String
    LoadConfig
    p1 = CfgS("ALERT_SOUND_PATH", "C:\Windows\Media\chimes.wav")
    p2 = CfgS("ALERT_SOUND_PATH_CEILING", "C:\Windows\Media\notify.wav")

    If Len(Dir(p1)) = 0 Then
        MsgBox "ALERT_SOUND_PATH not found on disk:" & vbCrLf & p1 & vbCrLf & vbCrLf & _
               "Fix the path on Config, or clear the cell to fall back to Beep.", _
               vbExclamation, "TT Z-Monitor"
        Exit Sub
    End If

    Ting p1
    MsgBox "That was the ENTRY chime." & vbCrLf & p1 & vbCrLf & vbCrLf & _
           "Click OK to hear the MAX_ENTRY_Z ceiling tone - the one that means " & _
           "stand down, not get in.", vbInformation, "TT Z-Monitor"

    If Len(Dir(p2)) = 0 Then
        MsgBox "ALERT_SOUND_PATH_CEILING not found on disk:" & vbCrLf & p2, _
               vbExclamation, "TT Z-Monitor"
        Exit Sub
    End If
    Ting p2
    MsgBox "That was the CEILING tone. The two must be distinguishable by ear - " & _
           "if they are not, change one on Config.", vbInformation, "TT Z-Monitor"
End Sub

'==============================================================================
' TICK ARCHIVE  -  every CHANGED quote, to CSV.  This is the record that lets
' you verify, on your own data, that coarser storage loses nothing real.
'==============================================================================
Private Sub ArchiveTick(ByVal i As Long, ByVal t As Double, aB As Variant, aA As Variant, _
                        bB As Variant, bA As Variant, ByVal mid As Double)
    On Error Resume Next
    If Not CfgB("TICK_ARCHIVE_ENABLED", True) Then Exit Sub
    If S(i).ArcN > UBound(S(i).Arc) - 1 Then FlushArchives False
    S(i).Arc(S(i).ArcN) = Format$(t, "yyyy-mm-dd hh:nn:ss") & "." & _
        Format$(Int((t - Int(t)) * 86400000#) Mod 1000, "000") & "," & _
        QKey(aB) & "," & QKey(aA) & "," & QKey(bB) & "," & QKey(bA) & "," & _
        Format$(mid, "0.##########")
    S(i).ArcN = S(i).ArcN + 1
    Err.Clear
End Sub

Private Sub FlushArchives(ByVal force As Boolean)
    Dim i As Long, k As Long, fnum As Integer, folder As String, fp As String
    On Error Resume Next
    If Not CfgB("TICK_ARCHIVE_ENABLED", True) And Not force Then Exit Sub

    folder = CfgS("TICK_ARCHIVE_PATH", "")
    If Len(folder) = 0 Then folder = ThisWorkbook.path & "\ticks"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    For i = 1 To MAX_SPREADS
        If S(i).ArcN > 0 Then
            fp = folder & "\ticks_slot" & i & "_" & Format$(Date, "yyyymmdd") & ".csv"
            fnum = FreeFile
            If Len(Dir(fp)) = 0 Then
                Open fp For Append As #fnum
                Print #fnum, "timestamp,legA_bid,legA_ask,legB_bid,legB_ask,spread_mid"
            Else
                Open fp For Append As #fnum
            End If
            For k = 0 To S(i).ArcN - 1
                Print #fnum, S(i).Arc(k)
            Next k
            Close #fnum
            S(i).ArcN = 0
        End If
    Next i
    Err.Clear
End Sub

'==============================================================================
' PERSISTENCE  -  the recovery snapshot is DECIMATED to FLUSH_DECIMATE_MS.
' Unbiased for mean and sigma, and it keeps a 120-minute window near 7,200 rows
' instead of tens of thousands, so the flush never stalls the capture timer.
'==============================================================================
Private Sub FlushBuffers()
    Dim b As Worksheet, i As Long, k As Long, idx As Long, n As Long
    Dim ct As Long, cv As Long, arr() As Double, dec As Double, lastT As Double
    On Error Resume Next
    Set b = ThisWorkbook.Sheets(SH_BUF)
    If b Is Nothing Then Exit Sub
    dec = CfgD("FLUSH_DECIMATE_MS", 1000)
    Application.EnableEvents = False

    For i = 1 To MAX_SPREADS
        ct = 1 + (i - 1) * 3: cv = ct + 1
        b.Cells(4, cv).Value2 = S(i).Name
        b.Cells(5, cv).Value2 = S(i).LegLabelB & " - b x " & S(i).LegLabelA
        b.Cells(6, cv).Value2 = S(i).Hedge
        b.Cells(7, cv).Value2 = CfgD("LOOKBACK_MIN", 120)

        If gFlushRows(i) > 0 Then
            b.Range(b.Cells(10, ct), b.Cells(9 + gFlushRows(i), cv)).ClearContents
        End If

        n = 0: lastT = 0
        If S(i).Count > 0 Then
            ReDim arr(1 To S(i).Count, 1 To 2)
            For k = 0 To S(i).Count - 1
                idx = (OldestIdx(i) + k) Mod S(i).Cap
                If dec <= 0 Or lastT = 0 Or (S(i).BufT(idx) - lastT) * 86400000# >= dec Then
                    n = n + 1
                    arr(n, 1) = S(i).BufT(idx)
                    arr(n, 2) = S(i).BufV(idx)
                    lastT = S(i).BufT(idx)
                End If
            Next k
            If n > 0 Then b.Range(b.Cells(10, ct), b.Cells(9 + n, cv)).Value2 = arr
        End If
        gFlushRows(i) = n
        b.Cells(8, cv).Value2 = S(i).Count & " held / " & n & " saved @ " & _
                                Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Next i

    Application.EnableEvents = True
    Err.Clear
End Sub

' Reload the flushed buffer so a restart does not cost another two hours of
' warm-up.  Elapsed collection is credited from the OLDEST recovered sample.
' Only reuse rows whose lookback, hedge ratio and legs still match Config.
Public Sub WarmStart()
    Dim b As Worksheet, i As Long, ct As Long, cv As Long
    Dim lastRow As Long, v As Variant, k As Long, n As Long
    Dim t As Double, cutoff As Double, maxAge As Double, lookback As Double
    Dim prev As Double, havePrev As Boolean, tt As Double, vv As Double

    On Error GoTo Done
    Set b = ThisWorkbook.Sheets(SH_BUF)
    t = TNow()
    maxAge = CfgD("WARM_START_MAX_AGE_MIN", 180)
    lookback = CfgD("LOOKBACK_MIN", 120)

    For i = 1 To MAX_SPREADS
        ct = 1 + (i - 1) * 3: cv = ct + 1
        If Not S(i).Enabled Then GoTo NextSlot
        If CStr(b.Cells(4, cv).Value2) <> S(i).Name Then GoTo NextSlot
        If CStr(b.Cells(5, cv).Value2) <> S(i).LegLabelB & " - b x " & S(i).LegLabelA Then GoTo NextSlot
        If Nz(b.Cells(6, cv).Value2) <> S(i).Hedge Then GoTo NextSlot
        If Nz(b.Cells(7, cv).Value2) <> lookback Then GoTo NextSlot

        lastRow = b.Cells(b.Rows.Count, ct).End(xlUp).Row
        If lastRow < 10 Then GoTo NextSlot
        If (t - Nz(b.Cells(lastRow, ct).Value2)) * 1440# > maxAge Then GoTo NextSlot

        v = b.Range(b.Cells(10, ct), b.Cells(lastRow, cv)).Value2
        cutoff = t - lookback / 1440#
        havePrev = False: n = 0
        For k = 1 To UBound(v, 1)
            tt = Nz(v(k, 1)): vv = Nz(v(k, 2))
            If tt >= cutoff Then
                ' Drop CONSECUTIVE identical values only.  A spread genuinely
                ' revisiting a level is a real observation, not a duplicate.
                If Not havePrev Or vv <> prev Then
                    Append i, tt, vv
                    prev = vv: havePrev = True: n = n + 1
                End If
            End If
        Next k
        If n > 0 Then
            S(i).LastQuoteTime = 0          ' the live feed is still unproven
            RecomputeStats i, t, lookback
            LogEvent S(i).Name, "WARM START", "", n & " samples recovered, oldest " & _
                     Format$(WindowElapsedMin(i, t), "0") & " min ago"
        End If
NextSlot:
    Next i
Done:
    Err.Clear
End Sub

Private Sub LogEdges(ByVal i As Long)
    If S(i).Signal <> S(i).PrevSignal Then
        LogEvent S(i).Name, "SIGNAL -> " & S(i).Signal, Format$(S(i).Z, "0.00"), _
                 "mean=" & Fmt(S(i).Mean) & " sigma=" & Fmt(S(i).Sigma) & _
                 " spread=" & Fmt(S(i).CurMid) & " edge=" & S(i).Verdict
        S(i).PrevSignal = S(i).Signal
    End If
    If S(i).Gate <> S(i).PrevGate Then
        LogEvent S(i).Name, "GATE -> " & S(i).Gate, "", ""
        S(i).PrevGate = S(i).Gate
    End If
    If S(i).FeedStatus <> S(i).PrevFeedStatus Then
        LogEvent S(i).Name, "FEED -> " & S(i).FeedStatus, "", Format$(S(i).Qpm, "0") & " q/min"
        S(i).PrevFeedStatus = S(i).FeedStatus
    End If
End Sub

Private Sub LogEvent(ByVal spreadName As String, ByVal ev As String, _
                     ByVal zTxt As String, ByVal noteTxt As String)
    Dim L As Worksheet, r As Long, maxRows As Long
    On Error Resume Next
    If Not CfgB("LOG_ENABLED", True) Then Exit Sub
    Set L = ThisWorkbook.Sheets(SH_LOG)
    If L Is Nothing Then Exit Sub
    r = L.Cells(L.Rows.Count, 1).End(xlUp).Row + 1
    If r < 4 Then r = 4
    L.Cells(r, 1).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    L.Cells(r, 2).Value2 = spreadName
    L.Cells(r, 3).Value2 = ev
    L.Cells(r, 4).Value2 = zTxt
    L.Cells(r, 12).Value2 = noteTxt
    maxRows = CfgL("LOG_MAX_ROWS", 20000)
    If r > maxRows + 3 Then L.Rows("4:1003").Delete
    Err.Clear
End Sub

'==============================================================================
' CONFIG  -  re-read every CONFIG_REFRESH_MS, so a changed cell takes effect
' without a restart and without editing VBA.
'==============================================================================
Private Sub LoadConfig()
    Dim ws As Worksheet, v As Variant, r As Long, nm As String
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets(SH_CFG)
    If gCfg Is Nothing Then Set gCfg = CreateObject("Scripting.Dictionary")
    gCfg.RemoveAll
    v = ws.Range("A1:B400").Value2
    For r = 1 To UBound(v, 1)
        If VarType(v(r, 1)) = vbString Then
            nm = Trim$(CStr(v(r, 1)))
            If Len(nm) > 0 Then
                If Not gCfg.Exists(nm) Then gCfg(nm) = v(r, 2)
            End If
        End If
    Next r
    ClampHistoryGate
    Exit Sub
Fail:
    ' Every Cfg* getter falls back to its default, so a failure here degrades
    ' to the shipped settings rather than stopping the monitor. Those defaults
    ' are themselves MIN_HISTORY_MIN 120 / LOOKBACK_MIN 120, so the guard has
    ' to run on this path too or a failed read reinstates the unreachable gate.
    Err.Clear
    ClampHistoryGate
End Sub

' MIN_HISTORY_MIN >= LOOKBACK_MIN is unsatisfiable BY CONSTRUCTION, and it
' fails silently. EvictOld drops every sample older than LOOKBACK_MIN, so the
' span WindowElapsedMin measures is bounded ABOVE by LOOKBACK_MIN and only ever
' approaches it from below. Asking for the whole lookback leaves the history
' gate permanently unmet: no z, no signal, no chime, on every slot - while the
' gate cell rounds to a satisfied-looking "120/120". Clamp it under the horizon
' and say so on the status strip, rather than warming up forever in silence.
'
' The clamp is recomputed from the sheet on every load, so it never drifts and
' it releases the moment the operator sets a workable value.
Private Sub ClampHistoryGate()
    Dim lookback As Double, minHist As Double, capped As Double, msg As String
    On Error Resume Next
    ' Reachable from the Fail path, where the dictionary may not exist yet.
    If gCfg Is Nothing Then Set gCfg = CreateObject("Scripting.Dictionary")
    lookback = CfgD("LOOKBACK_MIN", 120)
    minHist = CfgD("MIN_HISTORY_MIN", 120)

    If lookback > 0 And minHist >= lookback Then
        capped = lookback * HIST_GATE_MAX_FRAC
        gCfg("MIN_HISTORY_MIN") = capped
        msg = "MIN_HISTORY_MIN " & Format$(minHist, "0.##") & " >= LOOKBACK_MIN " & _
              Format$(lookback, "0.##") & " - unreachable, using " & Format$(capped, "0.##")
    Else
        msg = ""
    End If

    ' LoadConfig runs every CONFIG_REFRESH_MS, so log the transition only.
    If msg <> gHistClamp Then
        If Len(msg) > 0 Then
            LogEvent "", "CONFIG CLAMPED", Format$(minHist, "0.##"), msg
        Else
            LogEvent "", "CONFIG OK", Format$(minHist, "0.##"), _
                     "MIN_HISTORY_MIN is below LOOKBACK_MIN - history gate reachable"
        End If
        gHistClamp = msg
    End If
    Err.Clear
End Sub

Private Sub LoadSpreadDefs()
    Dim ws As Worksheet, fs As Worksheet, v As Variant, r As Long, i As Long
    Dim tag As String, legs As Variant, alts As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_CFG)
    Set fs = ThisWorkbook.Sheets(SH_FEED)
    v = ws.Range("A1:K400").Value2
    legs = fs.Range(fs.Cells(F_LEG_TOP, 3), fs.Cells(F_LEG_TOP + 2 * MAX_SPREADS - 1, 3)).Value2
    alts = fs.Range(fs.Cells(F_ALT_TOP, 1), fs.Cells(F_ALT_TOP + MAX_SPREADS - 1, 2)).Value2

    For r = 1 To UBound(v, 1)
        If VarType(v(r, 1)) = vbString Then
            tag = Trim$(CStr(v(r, 1)))
            If Left$(tag, 7) = "SPREAD_" And Len(tag) = 8 Then
                i = CLng(Mid$(tag, 8, 1))
                If i >= 1 And i <= MAX_SPREADS Then
                    S(i).Enabled = TruthyV(v(r, 2))
                    S(i).Name = TxtOf(v(r, 3))
                    S(i).Mode = UCase$(TxtOf(v(r, 4)))
                    S(i).Hedge = Nz(v(r, 5))
                    S(i).BetaStamp = TxtOf(v(r, 6))
                    S(i).UsdPerPoint = Nz(v(r, 7))
                    S(i).Contracts = Nz(v(r, 8))
                    S(i).TickSize = Nz(v(r, 9))
                    S(i).Decimals = CLng(Nz(v(r, 10)))
                    If S(i).UsdPerPoint <= 0 Then S(i).UsdPerPoint = 1000
                    If S(i).Contracts <= 0 Then S(i).Contracts = 2
                    If S(i).Decimals <= 0 Then S(i).Decimals = 2
                    S(i).LegLabelA = TxtOf(legs((i - 1) * 2 + 1, 1))
                    S(i).LegLabelB = TxtOf(legs((i - 1) * 2 + 2, 1))
                    S(i).AltRow = CLng(Nz(alts(i, 2)))
                End If
            End If
        End If
    Next r
End Sub

Private Function CfgD(ByVal nm As String, ByVal dflt As Double) As Double
    On Error GoTo Fb
    If gCfg Is Nothing Then CfgD = dflt: Exit Function
    If Not gCfg.Exists(nm) Then CfgD = dflt: Exit Function
    If Not IsNum(gCfg(nm)) Then CfgD = dflt: Exit Function
    CfgD = CDbl(gCfg(nm))
    Exit Function
Fb: CfgD = dflt: Err.Clear
End Function

Private Function CfgL(ByVal nm As String, ByVal dflt As Long) As Long
    CfgL = CLng(CfgD(nm, CDbl(dflt)))
End Function

Private Function CfgS(ByVal nm As String, ByVal dflt As String) As String
    On Error GoTo Fb
    If gCfg Is Nothing Then CfgS = dflt: Exit Function
    If Not gCfg.Exists(nm) Then CfgS = dflt: Exit Function
    If Len(TxtOf(gCfg(nm))) = 0 Then CfgS = dflt: Exit Function
    CfgS = TxtOf(gCfg(nm))
    Exit Function
Fb: CfgS = dflt: Err.Clear
End Function

Private Function CfgB(ByVal nm As String, ByVal dflt As Boolean) As Boolean
    On Error GoTo Fb
    If gCfg Is Nothing Then CfgB = dflt: Exit Function
    If Not gCfg.Exists(nm) Then CfgB = dflt: Exit Function
    CfgB = TruthyV(gCfg(nm))
    Exit Function
Fb: CfgB = dflt: Err.Clear
End Function

'==============================================================================
' FORMATS  -  set ONCE.  Prices show tick precision only: a number showing more
' digits than the market moves in will appear to change constantly.
'==============================================================================
Public Sub ApplyFormats()
    ' The user's own cells keep the number formats they already have - the
    ' Dashboard's appearance is not ours to change. Only the added z-block and
    ' the Detail sheet are formatted here.
    Dim d As Worksheet, det As Worksheet, i As Long, c As String, fmt As String
    Dim k As Long, r As Long, rows_ As Variant
    On Error Resume Next
    Set d = ThisWorkbook.Sheets(SH_DASH)
    rows_ = Array(ZR1, ZR2, ZR3)
    For k = 0 To 2
        r = rows_(k)
        d.Cells(r, 8).NumberFormat = "0.00"                          ' z
        d.Range(d.Cells(r, 9), d.Cells(r, 10)).NumberFormat = "0.0000"
        d.Cells(r, 14).NumberFormat = "0.0000"                       ' entry
        d.Range(d.Cells(r, 16), d.Cells(r, 17)).NumberFormat = "$#,##0.00"
        d.Cells(r, 18).NumberFormat = "0.0000"                       ' take profit
        d.Cells(r, 19).NumberFormat = "0.00"                         ' TP in sigma
    Next k

    ' The clock must stay TEXT. VBA writes Format$(Now, "hh:mm:ss"), and against
    ' a General-formatted cell Excel silently coerces that string into a time
    ' VALUE. The cell then holds a serial number while W() keeps comparing it
    ' with a string, the comparison can never match, and L3 is rewritten on
    ' every single paint - defeating the compare-first rule on the one cell
    ' that updates most often. The original =TEXT(NOW(),...) produced text, so
    ' text is also what preserves its appearance.
    With d.Range("L3")
        If .NumberFormat <> "@" Then
            .ClearContents
            .NumberFormat = "@"
        End If
    End With

    Set det = ThisWorkbook.Sheets(SH_DET)
    If Not det Is Nothing Then
        For i = 1 To MAX_SPREADS
            c = Mid$(SLOT_COL, i, 1)
            fmt = PriceFmt(S(i).Decimals)
            det.Range(c & R_BID & "," & c & R_ASK & "," & c & R_MID & "," & c & R_WIDTH & "," & _
                      c & R_MEAN & "," & c & R_SIG & "," & c & R_ENTRY & "," & c & R_BE & "," & _
                      c & R_TP & "," & c & R_TPD & "," & c & R_COSTU & "," & c & R_WINU & "," & _
                      c & R_SH_ENT & "," & c & R_SH_EXIT & "," & _
                      c & R_LG_ENT & "," & c & R_LG_EXIT & "," & _
                      c & R_MINSIG).NumberFormat = fmt
        Next i
        det.Range("H6:M9,H13:M17").NumberFormat = "#,##0.0000"
    End If
    Err.Clear
End Sub

Private Function PriceFmt(ByVal dp As Long) As String
    If dp <= 0 Then
        PriceFmt = "#,##0"
    Else
        PriceFmt = "#,##0." & String$(dp, "0")
    End If
End Function

'==============================================================================
' HELPERS
'==============================================================================
Private Function Dash() As Worksheet
    Set Dash = ThisWorkbook.Sheets(SH_DASH)
End Function

' THE CORE RULE: compare first, write second.
Private Sub W(ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    Dim c As Range, cur As Variant
    On Error GoTo Fail
    Set c = ws.Range(addr)
    cur = c.Value2

    If IsError(cur) Then c.Value2 = v: Exit Sub

    If VarType(v) = vbString Then
        If IsEmpty(cur) Then
            If Len(CStr(v)) > 0 Then c.Value2 = v
        ElseIf CStr(cur) <> CStr(v) Then
            c.Value2 = v
        End If
        Exit Sub
    End If

    If Not IsNum(v) Then Exit Sub
    If IsEmpty(cur) Or VarType(cur) = vbString Then
        c.Value2 = CDbl(v)
    ElseIf Abs(CDbl(cur) - CDbl(v)) > 0.0000000001 Then
        c.Value2 = CDbl(v)
    End If
    Exit Sub
Fail:
    Err.Clear
End Sub

' Same compare-first rule as W(), addressed by row and column.
Private Sub WC(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    On Error GoTo Fail
    W ws, ws.Cells(r, c).Address(False, False), v
    Exit Sub
Fail:
    Err.Clear
End Sub

Private Function TNow() As Double
    ' ~10 ms resolution, and immune to the Now() one-second boundary.
    TNow = CDbl(Date) + Timer / 86400#
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
        IsNum = IsNumeric(v)
        Exit Function
    End If
    IsNum = IsNumeric(v)
End Function

Private Function Nz(ByVal v As Variant) As Double
    If IsNum(v) Then Nz = CDbl(v)
End Function

Private Function TxtOf(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Then
        TxtOf = ""
    Else
        TxtOf = Trim$(CStr(v))
    End If
    Err.Clear
End Function

' Dedupe key: exact text of the quote, so 7.20 and 7.2000 are the same quote.
Private Function QKey(ByVal v As Variant) As String
    If IsNum(v) Then
        QKey = Format$(CDbl(v), "0.##########")
    Else
        QKey = "-"
    End If
End Function

Private Function TruthyV(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then TruthyV = CBool(v): Exit Function
    If IsNumeric(v) Then TruthyV = (CDbl(v) <> 0): Exit Function
    s = UCase$(Trim$(CStr(v)))
    TruthyV = (s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Or s = "ON")
End Function

Private Function Fmt(ByVal d As Double) As String
    Fmt = Format$(d, "0.0000")
End Function

'==============================================================================
' PASTE THIS INTO THE "ThisWorkbook" MODULE:
'
'   Private Sub Workbook_Open()
'       TTZMonitor.Auto_Open
'   End Sub
'
'   Private Sub Workbook_BeforeClose(Cancel As Boolean)
'       TTZMonitor.StopMonitor
'   End Sub
'
' Workbook_BeforeClose is NOT optional when the high-resolution timer is in
' use: a user32 timer that outlives the workbook will keep firing into a
' module that no longer has its sheets, and that crashes Excel.
'==============================================================================
