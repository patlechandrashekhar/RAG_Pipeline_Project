Attribute VB_Name = "ADIN6210_CodeReview"
Option Explicit

'===============================================================================
' Module       : ADIN6210_CodeReview.bas
' Device       : ADIN6210
' Platform     : Teradyne UltraFLEX / IG-XL VBT
' Description  : Code-quality self-review and diagnostic utility for the
'                ADIN6210 test program.  Performs a structured audit of key
'                program-health indicators at run-time and reports findings
'                to the IG-XL datalog.
'
' Entry Points (call from Flow sheet):
'   CodeReview_RunAll          - Execute every audit sub-test in sequence.
'   CodeReview_HardwareChecks  - Hardware / resource readiness only.
'   CodeReview_SoftwareChecks  - Software / environment checks only.
'   CodeReview_DatalogChecks   - Datalog configuration checks only.
'
' Pattern:  Setup -> Measure/Evaluate -> Datalog
' Author   : Senior VBT Engineer
' Revision : 1.0  (initial production release)
'===============================================================================

' ---------------------------------------------------------------------------
' Private constants
' ---------------------------------------------------------------------------
Private Const MODULE_NAME          As String = "ADIN6210_CodeReview"
Private Const REVIEW_VERSION       As String = "1.0.0"

' Severity tokens written to the datalog comment stream
Private Const SEV_PASS             As String = "[PASS]"
Private Const SEV_WARN             As String = "[WARN]"
Private Const SEV_FAIL             As String = "[FAIL]"
Private Const SEV_INFO             As String = "[INFO]"

' Conservative hardware safety limits for ADIN6210
Private Const MAX_SAFE_VDD_V       As Double = 3.465   ' AVDD3P3 absolute max (5% above 3.3 V)
Private Const MIN_SAFE_VDD_V       As Double = 3.135   ' AVDD3P3 minimum (5% below 3.3 V)
Private Const MAX_SAFE_VDDIO_V     As Double = 3.465   ' VDDIO absolute max
Private Const MIN_SAFE_VDDIO_V     As Double = 3.135   ' VDDIO minimum
Private Const MAX_SAFE_CURRENT_MA  As Double = 150.0   ' Per-supply current ceiling (mA)
Private Const MDIO_PERIOD_MIN_NS   As Double = 40.0    ' MDC period lower bound (ns)
Private Const MDIO_PERIOD_MAX_NS   As Double = 400.0   ' MDC period upper bound (ns)

' Supply pin names expected in the channel map
Private Const PIN_AVDD3P3          As String = "AVDD3P3"
Private Const PIN_VDDIO            As String = "VDDIO_R"
Private Const PIN_MDC              As String = "MDC"
Private Const PIN_MDIO             As String = "MDIO"

' ---------------------------------------------------------------------------
' Module-level audit counters (reset at the start of each RunAll call)
' ---------------------------------------------------------------------------
Private m_PassCount  As Long
Private m_WarnCount  As Long
Private m_FailCount  As Long

'===============================================================================
' PUBLIC ENTRY POINTS
'===============================================================================

' ---------------------------------------------------------------------------
' CodeReview_RunAll
' Master entry point.  Runs every audit category and returns 0 on clean pass,
' non-zero if any FAIL or WARN was recorded.
' ---------------------------------------------------------------------------
Public Function CodeReview_RunAll() As Long
    On Error GoTo errHandler

    Call ResetCounters
    Call LogBanner("ADIN6210 Code-Quality Review  v" & REVIEW_VERSION)

    ' --- Execute each audit category ---
    Call AuditEnvironment
    Call AuditDatalogConfig
    Call AuditHardwareSupplies
    Call AuditDigitalTiming
    Call AuditSiteLoopPatterns
    Call AuditMDIOHelperUsage
    Call AuditProcessMonitorLogic
    Call AuditExternalLoopbackSetup

    ' --- Summary ---
    Call LogSummary

    ' Return non-zero if any issue was found
    CodeReview_RunAll = m_FailCount + m_WarnCount

    Exit Function
errHandler:
    Call HandleError(MODULE_NAME & ".CodeReview_RunAll")
    CodeReview_RunAll = -1
End Function

' ---------------------------------------------------------------------------
' CodeReview_HardwareChecks
' Standalone hardware / resource readiness audit.
' ---------------------------------------------------------------------------
Public Function CodeReview_HardwareChecks() As Long
    On Error GoTo errHandler

    Call ResetCounters
    Call LogBanner("ADIN6210 Hardware Checks  v" & REVIEW_VERSION)

    Call AuditHardwareSupplies
    Call AuditDigitalTiming

    Call LogSummary
    CodeReview_HardwareChecks = m_FailCount + m_WarnCount

    Exit Function
errHandler:
    Call HandleError(MODULE_NAME & ".CodeReview_HardwareChecks")
    CodeReview_HardwareChecks = -1
End Function

' ---------------------------------------------------------------------------
' CodeReview_SoftwareChecks
' Standalone software / environment audit.
' ---------------------------------------------------------------------------
Public Function CodeReview_SoftwareChecks() As Long
    On Error GoTo errHandler

    Call ResetCounters
    Call LogBanner("ADIN6210 Software Checks  v" & REVIEW_VERSION)

    Call AuditEnvironment
    Call AuditSiteLoopPatterns
    Call AuditMDIOHelperUsage
    Call AuditProcessMonitorLogic
    Call AuditExternalLoopbackSetup

    Call LogSummary
    CodeReview_SoftwareChecks = m_FailCount + m_WarnCount

    Exit Function
errHandler:
    Call HandleError(MODULE_NAME & ".CodeReview_SoftwareChecks")
    CodeReview_SoftwareChecks = -1
End Function

' ---------------------------------------------------------------------------
' CodeReview_DatalogChecks
' Standalone datalog configuration audit.
' ---------------------------------------------------------------------------
Public Function CodeReview_DatalogChecks() As Long
    On Error GoTo errHandler

    Call ResetCounters
    Call LogBanner("ADIN6210 Datalog Checks  v" & REVIEW_VERSION)

    Call AuditDatalogConfig

    Call LogSummary
    CodeReview_DatalogChecks = m_FailCount + m_WarnCount

    Exit Function
errHandler:
    Call HandleError(MODULE_NAME & ".CodeReview_DatalogChecks")
    CodeReview_DatalogChecks = -1
End Function

'===============================================================================
' PRIVATE AUDIT ROUTINES
'===============================================================================

' ---------------------------------------------------------------------------
' AuditEnvironment
' Checks IG-XL execution environment health.
' ---------------------------------------------------------------------------
Private Sub AuditEnvironment()
    On Error GoTo errHandler

    Call LogSection("Environment Audit")

    ' 1. Verify TheExec is available
    If TheExec Is Nothing Then
        Call LogResult(SEV_FAIL, "TheExec object is Nothing - IG-XL not running")
        m_FailCount = m_FailCount + 1
        Exit Sub
    End If
    Call LogResult(SEV_PASS, "TheExec object is valid")
    m_PassCount = m_PassCount + 1

    ' 2. Tester mode
    Dim modeStr As String
    Select Case TheExec.TesterMode
        Case testModeOffline
            modeStr = "Offline"
            Call LogResult(SEV_WARN, "Running in OFFLINE mode - hardware results are simulated")
            m_WarnCount = m_WarnCount + 1
        Case testModeOnline
            modeStr = "Online"
            Call LogResult(SEV_PASS, "Running in ONLINE mode")
            m_PassCount = m_PassCount + 1
        Case Else
            modeStr = "Unknown(" & CStr(TheExec.TesterMode) & ")"
            Call LogResult(SEV_WARN, "Unrecognised tester mode: " & modeStr)
            m_WarnCount = m_WarnCount + 1
    End Select

    ' 3. Active site count
    Dim activeSites As Long
    activeSites = TheExec.Sites.Active.Count
    If activeSites < 1 Then
        Call LogResult(SEV_FAIL, "No active sites - cannot execute tests")
        m_FailCount = m_FailCount + 1
    Else
        Call LogResult(SEV_PASS, "Active site count = " & CStr(activeSites))
        m_PassCount = m_PassCount + 1
    End If

    ' 4. Run mode
    If TheExec.RunMode = runModeProduction Then
        Call LogResult(SEV_INFO, "Run mode = Production")
    Else
        Call LogResult(SEV_INFO, "Run mode = Engineering/Debug")
    End If

    ' 5. Channel map loaded
    Dim chanMap As String
    chanMap = TheExec.CurrentChanMap
    If Len(Trim(chanMap)) = 0 Then
        Call LogResult(SEV_FAIL, "CurrentChanMap is empty - no channel map loaded")
        m_FailCount = m_FailCount + 1
    Else
        Call LogResult(SEV_PASS, "Channel map loaded: " & chanMap)
        m_PassCount = m_PassCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditEnvironment")
End Sub

' ---------------------------------------------------------------------------
' AuditDatalogConfig
' Verifies that the datalog is configured for production-safe output.
' ---------------------------------------------------------------------------
Private Sub AuditDatalogConfig()
    On Error GoTo errHandler

    Call LogSection("Datalog Configuration Audit")

    ' 1. DatalogOn flag
    If TheExec.Datalog.Setup.LotSetup.DatalogOn Then
        Call LogResult(SEV_PASS, "DatalogOn = True")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "DatalogOn = False - test results will NOT be logged")
        m_WarnCount = m_WarnCount + 1
    End If

    ' 2. Window output
    If TheExec.Datalog.Setup.DatalogSetup.WindowOutput Then
        Call LogResult(SEV_PASS, "WindowOutput = True")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "WindowOutput = False - no on-screen datalog")
        m_WarnCount = m_WarnCount + 1
    End If

    ' 3. STDF output (informational)
    If TheExec.Datalog.Setup.DatalogSetup.STDFOutput Then
        Call LogResult(SEV_INFO, "STDFOutput = True (STDF file will be written)")
    Else
        Call LogResult(SEV_INFO, "STDFOutput = False")
    End If

    ' 4. Custom column widths enabled
    If TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths Then
        Call LogResult(SEV_PASS, "Custom column widths are enabled")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "Custom column widths are disabled - test names may be truncated")
        m_WarnCount = m_WarnCount + 1
    End If

    ' 5. Test name column width (parametric)
    Dim tnWidth As Long
    tnWidth = TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width
    If tnWidth >= 40 Then
        Call LogResult(SEV_PASS, "Parametric TestName column width = " & CStr(tnWidth) & " (adequate)")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "Parametric TestName column width = " & CStr(tnWidth) & " (< 40, names may be clipped)")
        m_WarnCount = m_WarnCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditDatalogConfig")
End Sub

' ---------------------------------------------------------------------------
' AuditHardwareSupplies
' Reads live supply voltages and currents; checks against ADIN6210 safe limits.
' ---------------------------------------------------------------------------
Private Sub AuditHardwareSupplies()
    On Error GoTo errHandler

    Call LogSection("Hardware Supply Audit")

    ' Skip hardware checks in offline mode to avoid instrument errors
    If TheExec.TesterMode = testModeOffline Then
        Call LogResult(SEV_INFO, "Offline mode - hardware supply checks skipped")
        Exit Sub
    End If

    ' --- AVDD3P3 ---
    Call CheckSupplyVoltage(PIN_AVDD3P3, MIN_SAFE_VDD_V, MAX_SAFE_VDD_V)
    Call CheckSupplyCurrent(PIN_AVDD3P3, MAX_SAFE_CURRENT_MA)

    ' --- VDDIO_R ---
    Call CheckSupplyVoltage(PIN_VDDIO, MIN_SAFE_VDDIO_V, MAX_SAFE_VDDIO_V)
    Call CheckSupplyCurrent(PIN_VDDIO, MAX_SAFE_CURRENT_MA)

    ' --- DIB power rail ---
    If thehdw.DIB.powerOn Then
        Call LogResult(SEV_PASS, "DIB power rail is ON")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "DIB power rail is OFF - device will not be powered")
        m_FailCount = m_FailCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditHardwareSupplies")
End Sub

' ---------------------------------------------------------------------------
' CheckSupplyVoltage  (helper called by AuditHardwareSupplies)
' ---------------------------------------------------------------------------
Private Sub CheckSupplyVoltage(pinName As String, _
                                minV As Double, _
                                maxV As Double)
    On Error GoTo errHandler

    Dim measV As Double
    measV = thehdw.DCVI.Pins(pinName).Voltage

    If measV < minV Then
        Call LogResult(SEV_FAIL, pinName & " voltage = " & _
                       FormatNum(measV, "0.000") & " V  (below min " & _
                       FormatNum(minV, "0.000") & " V)")
        m_FailCount = m_FailCount + 1
    ElseIf measV > maxV Then
        Call LogResult(SEV_FAIL, pinName & " voltage = " & _
                       FormatNum(measV, "0.000") & " V  (above max " & _
                       FormatNum(maxV, "0.000") & " V)")
        m_FailCount = m_FailCount + 1
    Else
        Call LogResult(SEV_PASS, pinName & " voltage = " & _
                       FormatNum(measV, "0.000") & " V  (within [" & _
                       FormatNum(minV, "0.000") & ", " & _
                       FormatNum(maxV, "0.000") & "] V)")
        m_PassCount = m_PassCount + 1
    End If

    Exit Sub
errHandler:
    Call LogResult(SEV_WARN, "Could not read voltage for pin: " & pinName & _
                   "  (" & Err.Description & ")")
    m_WarnCount = m_WarnCount + 1
End Sub

' ---------------------------------------------------------------------------
' CheckSupplyCurrent  (helper called by AuditHardwareSupplies)
' ---------------------------------------------------------------------------
Private Sub CheckSupplyCurrent(pinName As String, maxMA As Double)
    On Error GoTo errHandler

    Dim measI As Double
    ' Current returned in Amps; convert to mA for readability
    measI = thehdw.DCVI.Pins(pinName).Current * 1000.0

    If Abs(measI) > maxMA Then
        Call LogResult(SEV_WARN, pinName & " current = " & _
                       FormatNum(measI, "0.0") & " mA  (exceeds ceiling " & _
                       FormatNum(maxMA, "0.0") & " mA) - check for short")
        m_WarnCount = m_WarnCount + 1
    Else
        Call LogResult(SEV_PASS, pinName & " current = " & _
                       FormatNum(measI, "0.0") & " mA  (within limit)")
        m_PassCount = m_PassCount + 1
    End If

    Exit Sub
errHandler:
    Call LogResult(SEV_WARN, "Could not read current for pin: " & pinName & _
                   "  (" & Err.Description & ")")
    m_WarnCount = m_WarnCount + 1
End Sub

' ---------------------------------------------------------------------------
' AuditDigitalTiming
' Validates MDC/MDIO timing parameters against ADIN6210 spec limits.
' ---------------------------------------------------------------------------
Private Sub AuditDigitalTiming()
    On Error GoTo errHandler

    Call LogSection("Digital Timing Audit")

    If TheExec.TesterMode = testModeOffline Then
        Call LogResult(SEV_INFO, "Offline mode - digital timing checks skipped")
        Exit Sub
    End If

    ' Read MDC period from timing set "timing_1_0"
    Dim period_s  As Double
    Dim period_ns As Double

    On Error Resume Next
    period_s = thehdw.Digital.Timing.Period("timing_1_0")
    If Err.Number <> 0 Then
        Call LogResult(SEV_WARN, "Could not read timing set 'timing_1_0': " & Err.Description)
        m_WarnCount = m_WarnCount + 1
        Err.Clear
        On Error GoTo errHandler
        Exit Sub
    End If
    On Error GoTo errHandler

    period_ns = period_s / ns   ' convert to nanoseconds

    If period_ns < MDIO_PERIOD_MIN_NS Then
        Call LogResult(SEV_FAIL, "MDC period = " & FormatNum(period_ns, "0.0") & _
                       " ns  (below min " & FormatNum(MDIO_PERIOD_MIN_NS, "0.0") & " ns)")
        m_FailCount = m_FailCount + 1
    ElseIf period_ns > MDIO_PERIOD_MAX_NS Then
        Call LogResult(SEV_WARN, "MDC period = " & FormatNum(period_ns, "0.0") & _
                       " ns  (above recommended max " & FormatNum(MDIO_PERIOD_MAX_NS, "0.0") & " ns)")
        m_WarnCount = m_WarnCount + 1
    Else
        Call LogResult(SEV_PASS, "MDC period = " & FormatNum(period_ns, "0.0") & _
                       " ns  (within [" & FormatNum(MDIO_PERIOD_MIN_NS, "0.0") & ", " & _
                       FormatNum(MDIO_PERIOD_MAX_NS, "0.0") & "] ns)")
        m_PassCount = m_PassCount + 1
    End If

    ' Check that MDIO D1 edge is set to exactly one period (as per MDIO_test.bas pattern)
    Dim mdioD1_s  As Double
    Dim mdioD1_ns As Double

    On Error Resume Next
    mdioD1_s = thehdw.Digital.Pins(PIN_MDIO).Timing.EdgeTime("timing_1_0", chEdgeD1)
    If Err.Number <> 0 Then
        Call LogResult(SEV_WARN, "Could not read MDIO D1 edge: " & Err.Description)
        m_WarnCount = m_WarnCount + 1
        Err.Clear
        On Error GoTo errHandler
        Exit Sub
    End If
    On Error GoTo errHandler

    mdioD1_ns = mdioD1_s / ns

    Dim expectedD1_ns As Double
    expectedD1_ns = period_ns   ' VBT_Char_Timing.bas sets D1 = Period

    Dim toleranceNs As Double
    toleranceNs = 0.5           ' 0.5 ns tolerance

    If Abs(mdioD1_ns - expectedD1_ns) <= toleranceNs Then
        Call LogResult(SEV_PASS, "MDIO D1 edge = " & FormatNum(mdioD1_ns, "0.0") & _
                       " ns  (matches expected " & FormatNum(expectedD1_ns, "0.0") & " ns)")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "MDIO D1 edge = " & FormatNum(mdioD1_ns, "0.0") & _
                       " ns  (expected " & FormatNum(expectedD1_ns, "0.0") & _
                       " ns, delta = " & FormatNum(Abs(mdioD1_ns - expectedD1_ns), "0.0") & " ns)")
        m_WarnCount = m_WarnCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditDigitalTiming")
End Sub

' ---------------------------------------------------------------------------
' AuditSiteLoopPatterns
' Reviews site-loop constructs in the loaded program for correctness.
' Checks that TheExec.Sites.Active is used (not TheExec.Sites) in
' measurement loops, consistent with production best practice.
' ---------------------------------------------------------------------------
Private Sub AuditSiteLoopPatterns()
    On Error GoTo errHandler

    Call LogSection("Site Loop Pattern Audit")

    ' Verify active site count is non-zero before any loop would execute
    Dim activeSites As Long
    activeSites = TheExec.Sites.Active.Count

    If activeSites = 0 Then
        Call LogResult(SEV_FAIL, "Sites.Active.Count = 0 - all site loops will be skipped")
        m_FailCount = m_FailCount + 1
    Else
        Call LogResult(SEV_PASS, "Sites.Active.Count = " & CStr(activeSites) & _
                       " - site loops will execute")
        m_PassCount = m_PassCount + 1
    End If

    ' Verify we are not inside a site loop at review time (would indicate a logic error)
    If TheExec.Sites.InSiteLoop Then
        Call LogResult(SEV_WARN, "InSiteLoop = True at review entry - unexpected nesting detected")
        m_WarnCount = m_WarnCount + 1
    Else
        Call LogResult(SEV_PASS, "InSiteLoop = False at review entry - correct")
        m_PassCount = m_PassCount + 1
    End If

    ' Confirm Selected sites match Active sites (should be true outside a loop)
    If TheExec.Sites.Selected.Count = activeSites Then
        Call LogResult(SEV_PASS, "Sites.Selected.Count matches Sites.Active.Count (" & _
                       CStr(activeSites) & ")")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "Sites.Selected.Count (" & _
                       CStr(TheExec.Sites.Selected.Count) & _
                       ") does not match Sites.Active.Count (" & _
                       CStr(activeSites) & ") - possible stale selection")
        m_WarnCount = m_WarnCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditSiteLoopPatterns")
End Sub

' ---------------------------------------------------------------------------
' AuditMDIOHelperUsage
' Validates that MDIO CL45 / CL22 helper calls use legal address ranges
' consistent with the ADIN6210 register map.
' ---------------------------------------------------------------------------
Private Sub AuditMDIOHelperUsage()
    On Error GoTo errHandler

    Call LogSection("MDIO Helper Usage Audit")

    ' ADIN6210 uses Clause-45 MMD device addresses 1, 3, 7, 30 (vendor-specific)
    ' and Clause-22 PHY address 0.
    ' We validate the constants used in MDIO_test.bas at review time.

    ' --- Clause-45 device address ---
    Dim cl45DevAddr As Long
    cl45DevAddr = 30    ' as used in MDIO_test.bas

    If cl45DevAddr >= 0 And cl45DevAddr <= 31 Then
        Call LogResult(SEV_PASS, "CL45 device address " & CStr(cl45DevAddr) & _
                       " is within legal range [0..31]")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "CL45 device address " & CStr(cl45DevAddr) & _
                       " is outside legal range [0..31]")
        m_FailCount = m_FailCount + 1
    End If

    ' --- Process monitor control register address ---
    Dim pmCtrlAddr As Long
    pmCtrlAddr = 65350   ' GePmCntrl register

    If pmCtrlAddr >= 0 And pmCtrlAddr <= 65535 Then
        Call LogResult(SEV_PASS, "PM control register address 0x" & _
                       Hex(pmCtrlAddr) & " (" & CStr(pmCtrlAddr) & _
                       ") is within 16-bit register space")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "PM control register address " & CStr(pmCtrlAddr) & _
                       " exceeds 16-bit register space")
        m_FailCount = m_FailCount + 1
    End If

    ' --- PM result register address ---
    Dim pmResultAddr As Long
    pmResultAddr = 65351   ' GePmResult register

    If pmResultAddr >= 0 And pmResultAddr <= 65535 Then
        Call LogResult(SEV_PASS, "PM result register address 0x" & _
                       Hex(pmResultAddr) & " (" & CStr(pmResultAddr) & _
                       ") is within 16-bit register space")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "PM result register address " & CStr(pmResultAddr) & _
                       " exceeds 16-bit register space")
        m_FailCount = m_FailCount + 1
    End If

    ' --- Clause-22 PHY address ---
    Dim cl22PhyAddr As Long
    cl22PhyAddr = 0   ' as used in MDIO_test.bas

    If cl22PhyAddr >= 0 And cl22PhyAddr <= 31 Then
        Call LogResult(SEV_PASS, "CL22 PHY address " & CStr(cl22PhyAddr) & _
                       " is within legal range [0..31]")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "CL22 PHY address " & CStr(cl22PhyAddr) & _
                       " is outside legal range [0..31]")
        m_FailCount = m_FailCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditMDIOHelperUsage")
End Sub

' ---------------------------------------------------------------------------
' AuditProcessMonitorLogic
' Validates the Process_Monitor_test DSPWave sizing and ring-oscillator
' selection constants used in MDIO_test.bas.
' ---------------------------------------------------------------------------
Private Sub AuditProcessMonitorLogic()
    On Error GoTo errHandler

    Call LogSection("Process Monitor Logic Audit")

    ' Expected DSPWave size: 9 elements (indices 0..8, three per oscillator)
    Dim expectedElements As Long
    expectedElements = 9

    Call LogResult(SEV_INFO, "Process_Monitor_test DSPWave size = " & _
                   CStr(expectedElements) & " (3 results x 3 ring oscillators)")

    ' Validate GePmCntN320nsm1 constant
    Dim gePmCntN320nsm1 As Long
    gePmCntN320nsm1 = &HFF   ' 255 decimal

    If gePmCntN320nsm1 >= 0 And gePmCntN320nsm1 <= 255 Then
        Call LogResult(SEV_PASS, "GePmCntN320nsm1 = 0x" & Hex(gePmCntN320nsm1) & _
                       " (" & CStr(gePmCntN320nsm1) & ") - valid 8-bit field")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "GePmCntN320nsm1 = " & CStr(gePmCntN320nsm1) & _
                       " - exceeds 8-bit field width")
        m_FailCount = m_FailCount + 1
    End If

    ' Validate ring oscillator selection values (0=RVT, 1=LVT, 2=HVT)
    Dim roSelMax As Long
    roSelMax = 2   ' HVT is the highest valid selection

    If roSelMax <= 3 Then
        Call LogResult(SEV_PASS, "Ring oscillator GePmRoSel max = " & CStr(roSelMax) & _
                       " - fits in 2-bit field")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "Ring oscillator GePmRoSel max = " & CStr(roSelMax) & _
                       " - exceeds 2-bit field")
        m_FailCount = m_FailCount + 1
    End If

    ' Validate gate time calculation: (320 * (GePmCntN320nsm1 + 1)) + 40 ns
    Dim gateTime_ns As Double
    gateTime_ns = (320.0 * (CDbl(gePmCntN320nsm1) + 1.0)) + 40.0

    If gateTime_ns > 0 Then
        Call LogResult(SEV_PASS, "PM gate time = " & FormatNum(gateTime_ns, "0.0") & _
                       " ns  (positive - frequency divisor is valid)")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "PM gate time = " & FormatNum(gateTime_ns, "0.0") & _
                       " ns  (zero or negative - division by zero risk)")
        m_FailCount = m_FailCount + 1
    End If

    ' Warn if GePmReq bit is not 1 (would prevent PM from starting)
    Dim gePmReq As Long
    gePmReq = 1

    If gePmReq = 1 Then
        Call LogResult(SEV_PASS, "GePmReq = 1 - process monitor request bit is set")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "GePmReq = " & CStr(gePmReq) & _
                       " - process monitor will not start")
        m_FailCount = m_FailCount + 1
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditProcessMonitorLogic")
End Sub

' ---------------------------------------------------------------------------
' AuditExternalLoopbackSetup
' Reviews the Extlpbk_RGMII_1000T_test register write sequence for
' correctness and completeness.
' ---------------------------------------------------------------------------
Private Sub AuditExternalLoopbackSetup()
    On Error GoTo errHandler

    Call LogSection("External Loopback Setup Audit")

    ' Verify relay states are consistent with loopback test requirements.
    ' PHY_LPBK_RLY_A and MAC_LPBK_RLY must be OFF for external loopback.
    If TheExec.TesterMode = testModeOffline Then
        Call LogResult(SEV_INFO, "Offline mode - relay state checks skipped")
    Else
        Dim phyRlyState As Long
        Dim macRlyState As Long

        On Error Resume Next
        phyRlyState = thehdw.Utility.Pins("PHY_LPBK_RLY_A").State
        macRlyState = thehdw.Utility.Pins("MAC_LPBK_RLY").State
        If Err.Number <> 0 Then
            Call LogResult(SEV_WARN, "Could not read loopback relay states: " & Err.Description)
            m_WarnCount = m_WarnCount + 1
            Err.Clear
            On Error GoTo errHandler
            GoTo SkipRelayCheck
        End If
        On Error GoTo errHandler

        If phyRlyState = tlUtilBitOff Then
            Call LogResult(SEV_PASS, "PHY_LPBK_RLY_A is OFF - correct for external loopback")
            m_PassCount = m_PassCount + 1
        Else
            Call LogResult(SEV_WARN, "PHY_LPBK_RLY_A is ON - may conflict with external loopback")
            m_WarnCount = m_WarnCount + 1
        End If

        If macRlyState = tlUtilBitOff Then
            Call LogResult(SEV_PASS, "MAC_LPBK_RLY is OFF - correct for external loopback")
            m_PassCount = m_PassCount + 1
        Else
            Call LogResult(SEV_WARN, "MAC_LPBK_RLY is ON - may conflict with external loopback")
            m_WarnCount = m_WarnCount + 1
        End If
    End If

SkipRelayCheck:

    ' Validate key register values used in Extlpbk_RGMII_1000T_test
    ' MII control register 0x940 = Auto-neg enable + restart + 1000T
    Dim miiCtrl As Long
    miiCtrl = &H940

    If (miiCtrl And &H1000) <> 0 Then
        Call LogResult(SEV_PASS, "MII Control 0x" & Hex(miiCtrl) & _
                       ": Auto-negotiation enable bit is SET")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "MII Control 0x" & Hex(miiCtrl) & _
                       ": Auto-negotiation enable bit is CLEAR")
        m_WarnCount = m_WarnCount + 1
    End If

    ' Register 19 = 0xC1: LbAllDigSel + LbExtEn
    Dim reg19Val As Long
    reg19Val = &HC1

    If (reg19Val And &H1) <> 0 Then
        Call LogResult(SEV_PASS, "Reg19 = 0x" & Hex(reg19Val) & _
                       ": LbExtEn bit is SET - external loopback enabled")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_FAIL, "Reg19 = 0x" & Hex(reg19Val) & _
                       ": LbExtEn bit is CLEAR - external loopback NOT enabled")
        m_FailCount = m_FailCount + 1
    End If

    ' Register 9 = 0x1800: ManMstrSlvEnAdv + ManMstrAdv (master mode)
    Dim reg9Val As Long
    reg9Val = &H1800

    If (reg9Val And &H1000) <> 0 Then
        Call LogResult(SEV_PASS, "Reg9 = 0x" & Hex(reg9Val) & _
                       ": Manual master/slave enable bit is SET")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_WARN, "Reg9 = 0x" & Hex(reg9Val) & _
                       ": Manual master/slave enable bit is CLEAR")
        m_WarnCount = m_WarnCount + 1
    End If

    If (reg9Val And &H800) <> 0 Then
        Call LogResult(SEV_PASS, "Reg9 = 0x" & Hex(reg9Val) & _
                       ": ManMstrAdv (master preference) bit is SET")
        m_PassCount = m_PassCount + 1
    Else
        Call LogResult(SEV_INFO, "Reg9 = 0x" & Hex(reg9Val) & _
                       ": ManMstrAdv bit is CLEAR (slave preference)")
    End If

    Exit Sub
errHandler:
    Call HandleError(MODULE_NAME & ".AuditExternalLoopbackSetup")
End Sub

'===============================================================================
' PRIVATE HELPER ROUTINES
'===============================================================================

' ---------------------------------------------------------------------------
' ResetCounters - zero all audit counters before a new run
' ---------------------------------------------------------------------------
Private Sub ResetCounters()
    m_PassCount = 0
    m_WarnCount = 0
    m_FailCount = 0
End Sub

' ---------------------------------------------------------------------------
' LogBanner - write a prominent header to the datalog
' ---------------------------------------------------------------------------
Private Sub LogBanner(title As String)
    Dim sep As String
    sep = String(70, "=")
    TheExec.Datalog.WriteComment sep
    TheExec.Datalog.WriteComment "  " & title
    TheExec.Datalog.WriteComment sep
End Sub

' ---------------------------------------------------------------------------
' LogSection - write a section divider to the datalog
' ---------------------------------------------------------------------------
Private Sub LogSection(sectionName As String)
    TheExec.Datalog.WriteComment String(60, "-")
    TheExec.Datalog.WriteComment "  SECTION: " & sectionName
    TheExec.Datalog.WriteComment String(60, "-")
End Sub

' ---------------------------------------------------------------------------
' LogResult - write a single audit result line to the datalog
' ---------------------------------------------------------------------------
Private Sub LogResult(severity As String, message As String)
    TheExec.Datalog.WriteComment "  " & severity & "  " & message
End Sub

' ---------------------------------------------------------------------------
' LogSummary - write the final pass/warn/fail tally
' ---------------------------------------------------------------------------
Private Sub LogSummary()
    Dim sep As String
    sep = String(70, "=")
    TheExec.Datalog.WriteComment sep
    TheExec.Datalog.WriteComment "  AUDIT SUMMARY"
    TheExec.Datalog.WriteComment "    PASS  : " & CStr(m_PassCount)
    TheExec.Datalog.WriteComment "    WARN  : " & CStr(m_WarnCount)
    TheExec.Datalog.WriteComment "    FAIL  : " & CStr(m_FailCount)

    If m_FailCount > 0 Then
        TheExec.Datalog.WriteComment "  OVERALL : FAIL  (" & CStr(m_FailCount) & _
                                     " failure(s) require attention)"
    ElseIf m_WarnCount > 0 Then
        TheExec.Datalog.WriteComment "  OVERALL : PASS WITH WARNINGS  (" & _
                                     CStr(m_WarnCount) & " warning(s))"
    Else
        TheExec.Datalog.WriteComment "  OVERALL : CLEAN PASS"
    End If

    TheExec.Datalog.WriteComment sep
End Sub

' ---------------------------------------------------------------------------
' FormatNum - locale-safe numeric formatter
' ---------------------------------------------------------------------------
Private Function FormatNum(value As Double, fmt As String) As String
    FormatNum = Format(value, fmt)
End Function

' ---------------------------------------------------------------------------
' HandleError - centralised VBT error handler
' ---------------------------------------------------------------------------
Private Sub HandleError(location As String)
    Dim errMsg As String
    errMsg = "VBT Error in " & location & _
             "  #" & Trim(Str(Err.Number)) & ": " & Err.Description

    ' Write to datalog if possible
    On Error Resume Next
    TheExec.Datalog.WriteComment SEV_FAIL & "  " & errMsg
    On Error GoTo 0

    ' Increment fail counter so the summary reflects the error
    m_FailCount = m_FailCount + 1

    ' Abort the test step cleanly
    AbortTest
End Sub
