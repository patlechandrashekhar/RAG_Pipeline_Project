# UltraFlex VBT Function Templates

Generic VBT function skeletons for the Teradyne UltraFlex / IG-XL platform.
Each template shows the minimal correct structure (setup, measure, cleanup,
datalog). Replace placeholder names with device-specific pin names and groups.

UltraFlex programs use `DT` sheet headers with `platform=Jaguar`, support
UltraVI80 / UltraPin800 / HDVS / DCIO instruments, concurrent test,
background DSP, and optional vblite-generated register map objects.

---

## Conventions

### Function Naming

- Top-level test functions: `TF_` prefix, return `Long`
  (e.g. `Public Function TF_Continuity() As Long`)
- Helper functions: no prefix, can return any type
- Module names: `VBT_` prefix for test modules, no prefix for infrastructure

### Variable Naming

- `g_` prefix for global variables
- `c_` prefix for constants
- `m_` prefix for private object members
- Site-aware variables: always `Dim x As New SiteDouble` /
  `As New SiteLong` / `As New SiteBoolean`

### Standard Error Handler

```vba
On Error GoTo errHandler
' ... test code ...
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
```

### Standard Profiling

```vba
theTimer.VBTFunctionStart
' ... test code ...
theTimer.VBTFunctionFinish
```

### TestLimit

- Always use `ForceResults:=tlForceFlow`
- For dimensionless values: `unit:=unitNone`
- For custom units: `unit:=unitCustom, customUnit:="mA"`
- UltraFlex supports deferred evaluation with background DSP

---

## Table of Contents

1. [Globals Module Template](#1-globals-module-template)
2. [DUTconditions Class Template](#2-dutconditions-class-template)
3. [Instruments Module Template](#3-instruments-module-template)
4. [Exec Interpose Module Template](#4-exec-interpose-module-template)
5. [Continuity Test (PPMU-based)](#5-continuity-test-ppmu-based)
6. [Continuity Test (DCVI-based)](#6-continuity-test-dcvi-based)
7. [Pin Leakage Test](#7-pin-leakage-test)
8. [Supply Current Measurement](#8-supply-current-measurement)
9. [IDDQ Test (Supply Switching Method)](#9-iddq-test-supply-switching-method)
10. [Scan Chain Test (DSSC-based)](#10-scan-chain-test-dssc-based)
11. [Logic Level Verification (Production)](#11-logic-level-verification-production)
12. [Logic Level Characterisation (Threshold Search)](#12-logic-level-characterisation-threshold-search)
13. [GPIO Functional Test (Register-based)](#13-gpio-functional-test-register-based)
14. [Trim Template (Binary Search)](#14-trim-template-binary-search)
15. [Power-On Reset / Power Cycle](#15-power-on-reset--power-cycle)
16. [Reference Voltage Measurement](#16-reference-voltage-measurement)
17. [eFuse / OTP Programming](#17-efuse--otp-programming)
18. [DUTconditions Power Sequence Utilities](#18-dutconditions-power-sequence-utilities)
19. [Instrument Initialization Utilities](#19-instrument-initialization-utilities)

---

## 1. Globals Module Template

```vba
Attribute VB_Name = "Globals"
Option Explicit
''' Global variables and constants.
''' Owner: <name>

' --- DUT Conditions Class Instance ---
Public DUTconds As New DUTconditions

' --- Site-aware State Variables ---
Public g_DeviceID As New SiteLong
Public g_SiliconRev As String
Public g_DIE_Temp As New SiteDouble
Public g_CalibratedPart As New SiteBoolean

' --- Trim Code Storage (site-aware) ---
Public TC_OSCTRIM As New SiteLong
Public TC_REFTRIM As New SiteLong
Public TC_LDOTRIM As New SiteLong
' <add trim codes as needed>

' --- Register Map Objects (requires vblite-generated code) ---
' If using vblite, declare RM_ and BF_ objects here for each
' peripheral block. Otherwise, use pattern-based register access.

' --- Program Control ---
Public g_PROGRAM As String
Public g_GlobalVarReset As Boolean

' --- Measurement Constants ---
Public Const c_DCVISampSize As Double = 5
Public Const c_DCVISampRate As Double = 1000
Public Const c_pinSettPPMU As Double = 5    ' ms
Public Const c_pinSettDCVI As Double = 5    ' ms
Public Const c_nearZero As Double = 0.000000000001

' --- Supply Definitions ---
Public Const c_nomVlogSupp As Double = 3.3
Public Const c_nomAVDD As Double = 3.3
Public Const c_nomDVDD As Double = 1.8

' --- Pin List Definitions ---
Public Const c_supPinList As String = "SUPPLY_PINS"
Public Const c_allRelays As String = "ALL_RELAYS"
Public Const c_allDigPins As String = "DIG_PINS"

' --- Leakage Pre/Post Storage ---
Public g_PreLeakEvenPins0V As New PinListData
Public g_PreLeakEvenPinsVDD As New PinListData
Public g_PreLeakOddPins0V As New PinListData
Public g_PreLeakOddPinsVDD As New PinListData

' --- Test Condition Type ---
Public Type TestCondition
    vdd As Double
    vddIO As Double
    temperature As Double
    vih As Double
    vil As Double
    voh As Double
    vol As Double
    ioh As Double
    iol As Double
    vt As Double
End Type
Public g_tc(0 To 10) As TestCondition
Public g_CurrentTestCondition As Integer
Public Const c_NomTC As Long = 1
Public Const c_MinTC As Long = 2
Public Const c_MaxTC As Long = 3

' --- Conditional Compilation ---
' Set in Tools -> Project Properties -> Conditional Compile Arguments:
' PGM_MASK=32:FT=240:PROBE=15
' Use:
' #If (PGM_MASK And PROBE) Then
'     ' probe-specific code
' #ElseIf (PGM_MASK And FT) Then
'     ' final test code
' #End If

' --- MUX Selection Enum ---
Public Enum MuxSelectEnum
    MUX_GND = 0
    MUX_VREF = 1
    MUX_AIN0 = 2
    ' <add channels as needed>
End Enum
Public g_MuxOffset As New SiteDouble
```

---

## 2. DUTconditions Class Template

```vba
Attribute VB_Name = "DUTconditions"
Option Explicit
''' DUTconditions class -- manages supply voltages, logic levels,
''' and power sequencing for the DUT.

' --- Private State ---
Private m_AVdd As Double
Private m_DVdd As Double
Private m_IOVdd0 As Double
Private m_IOVdd1 As Double
Private m_Vlogic As Double
Private m_Vih As Double
Private m_Vil As Double
Private m_Voh As Double
Private m_Vol As Double
Private m_Vt As Double

' --- Saved State (for save/restore) ---
Private m_savedAVdd As Double
Private m_savedDVdd As Double
Private m_savedIOVdd0 As Double
Private m_savedIOVdd1 As Double
Private m_savedVlogic As Double

' --- Absolute Max Ratings ---
Private Const c_maxAVDD As Double = 3.6
Private Const c_maxDVDD As Double = 2.0
Private Const c_maxIOVDD As Double = 3.6
Private Const c_maxVlogic As Double = 3.6

' --- Ramping Constants ---
Private Const c_supSettle As Double = 5      ' ms settling time
Private Const c_stepWait As Double = 0.5     ' ms per ramp step
Private Const c_maxLastStep As Double = 0.1  ' V max final step size
Private Const c_maxNumSteps As Long = 10     ' maximum ramp steps

Private Sub Class_Initialize()
    Call Class_Reset
End Sub

Public Sub Class_Reset()
    m_AVdd = 0
    m_DVdd = 0
    m_IOVdd0 = 0
    m_IOVdd1 = 0
    m_Vlogic = c_nomVlogSupp
    m_Vih = m_Vlogic
    m_Vil = 0
    m_Voh = m_Vlogic - 0.5
    m_Vol = 0.6
    m_Vt = m_Vlogic * 0.5
End Sub

' ---------------------------------------------------------------
' Save / Restore State
' ---------------------------------------------------------------
Public Sub saveState()
    m_savedAVdd = m_AVdd
    m_savedDVdd = m_DVdd
    m_savedIOVdd0 = m_IOVdd0
    m_savedIOVdd1 = m_IOVdd1
    m_savedVlogic = m_Vlogic
End Sub

Public Sub restoreState()
    Call setAVdd(m_savedAVdd)
    Call setDVdd(m_savedDVdd)
    Call setIOVdd(m_savedIOVdd0, "IOVDD0")
    Call setIOVdd(m_savedIOVdd1, "IOVDD1")
    Call setLevelsSafe(m_savedVlogic)
End Sub

' ---------------------------------------------------------------
' EOS-Safe Exponential Ramp (common helper)
' ---------------------------------------------------------------
Private Sub rampDCVI(pinset As String, newVal As Double, _
                     currentVal As Double)
    Dim deltaVal As Double
    Dim lastStepSize As Double
    Dim numSteps As Long
    Dim i As Long

    deltaVal = newVal - currentVal
    If Abs(deltaVal) < c_nearZero Then Exit Sub

    ' Calculate number of exponentially decreasing steps
    lastStepSize = Abs(deltaVal)
    i = 0
    While lastStepSize > c_maxLastStep And i <= c_maxNumSteps
        i = i + 1
        lastStepSize = Abs(deltaVal) / (2 ^ i)
    Wend
    numSteps = i

    ' Ramp using exponentially decreasing step sizes
    For i = 1 To numSteps
        TheHdw.DCVI.Pins(pinset).Voltage = newVal - (deltaVal / (2 ^ i))
        TheHdw.Wait c_stepWait
    Next i

    ' Final value
    TheHdw.DCVI.Pins(pinset).Voltage = newVal
    TheHdw.SetSettlingTimer c_supSettle
End Sub

' ---------------------------------------------------------------
' Supply Setters (EOS-safe)
' ---------------------------------------------------------------
Public Sub setAVdd(newVal As Double)
    If newVal > c_maxAVDD Then
        TheExec.Datalog.WriteComment "AVDD exceeds absolute max!"
        Exit Sub
    End If
    Call rampDCVI("AVDD", newVal, m_AVdd)
    m_AVdd = newVal
End Sub

Public Sub setDVdd(newVal As Double)
    If newVal > c_maxDVDD Then
        TheExec.Datalog.WriteComment "DVDD exceeds absolute max!"
        Exit Sub
    End If
    Call rampDCVI("DVDD", newVal, m_DVdd)
    m_DVdd = newVal
End Sub

Public Sub setIOVdd(newVal As Double, _
                    Optional pinset As String = "IOVDD0")
    If newVal > c_maxIOVDD Then
        TheExec.Datalog.WriteComment "IOVDD exceeds absolute max!"
        Exit Sub
    End If
    If pinset = "IOVDD0" Then
        Call rampDCVI(pinset, newVal, m_IOVdd0)
        m_IOVdd0 = newVal
    Else
        Call rampDCVI(pinset, newVal, m_IOVdd1)
        m_IOVdd1 = newVal
    End If
End Sub

' ---------------------------------------------------------------
' Logic Level Setters (with safety checks)
' ---------------------------------------------------------------
Public Sub setVih(pins As String, val As Double)
    If val > c_maxVlogic Then Exit Sub
    TheHdw.PinLevels.Pins(pins).ModifyLevel chVih, val
    m_Vih = val
End Sub

Public Sub setVil(pins As String, val As Double)
    If val < -0.5 Then Exit Sub
    TheHdw.PinLevels.Pins(pins).ModifyLevel chVil, val
    m_Vil = val
End Sub

Public Sub setVoh(pins As String, val As Double)
    If val > c_maxVlogic Then Exit Sub
    TheHdw.PinLevels.Pins(pins).ModifyLevel chVoh, val
    m_Voh = val
End Sub

Public Sub setVol(pins As String, val As Double)
    If val > m_Voh Then Exit Sub
    TheHdw.PinLevels.Pins(pins).ModifyLevel chVol, val
    m_Vol = val
End Sub

' ---------------------------------------------------------------
' Getters
' ---------------------------------------------------------------
Public Function getVlogic() As Double
    getVlogic = m_Vlogic
End Function

Public Function getIOVDD(Optional domain As String = "IOVDD0") As Double
    If domain = "IOVDD0" Then
        getIOVDD = m_IOVdd0
    Else
        getIOVDD = m_IOVdd1
    End If
End Function

Public Function getVih() As Double
    getVih = m_Vih
End Function

' ---------------------------------------------------------------
' Power On -- full power-up sequence
' ---------------------------------------------------------------
Public Sub powerOn()
    ' 1. Initialize all DCVI properties
    Call InitDCVI

    ' 2. Ramp core supply first
    Call setDVdd(c_nomDVDD)

    ' 3. Ramp analog supply
    Call setAVdd(c_nomAVDD)

    ' 4. Ramp IO supplies
    Call setIOVdd(c_nomVlogSupp, "IOVDD0")
    Call setIOVdd(c_nomVlogSupp, "IOVDD1")

    ' 5. Connect digital pins and apply levels/timing
    TheHdw.Digital.ApplyLevelsTiming
    TheHdw.Wait c_supSettle

    ' 6. Hold reset high via PPMU
    With TheHdw.PPMU.Pins("RESETN")
        .ForceV c_nomVlogSupp, 2 * mA
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait c_supSettle

    m_Vlogic = c_nomVlogSupp
End Sub

' ---------------------------------------------------------------
' Power Off -- safe power-down sequence
' ---------------------------------------------------------------
Public Sub powerOff()
    ' 1. Disconnect digital pins
    TheHdw.Digital.DisconnectAllPins

    ' 2. Set all supplies to 0V (reverse order)
    Call setIOVdd(0, "IOVDD1")
    Call setIOVdd(0, "IOVDD0")
    Call setAVdd(0)
    Call setDVdd(0)

    ' 3. Disconnect all DCVI
    TheHdw.DCVI.Pins(c_supPinList).Disconnect

    ' 4. Disconnect PPMU
    TheHdw.PPMU.Pins("RESETN").Disconnect

    ' 5. Open all relays
    TheHdw.Utility.Pins(c_allRelays).RelayMode = tlRelayModeOff
    TheHdw.Wait c_supSettle
End Sub

' ---------------------------------------------------------------
' Power Cycle
' ---------------------------------------------------------------
Public Sub powerCycle()
    Call powerOff
    TheHdw.Wait 50   ' ms dead time
    Call powerOn
End Sub
```

---

## 3. Instruments Module Template

```vba
Attribute VB_Name = "Instruments"
Option Explicit
''' Instrument configuration utilities for UltraFlex.
''' Covers DCVI (UltraVI80), PPMU, PLMeter, and DCDiffMeter.

' ---------------------------------------------------------------
' DCVI Voltage Range Enum
' ---------------------------------------------------------------
Public Enum DCVIVoltageRangeEnum
    DCVI_VRANGE_5V = 0
    DCVI_VRANGE_8V = 1
    DCVI_VRANGE_20V = 2
End Enum

' ---------------------------------------------------------------
' DCVI Current Range Enum
' ---------------------------------------------------------------
Public Enum DCVICurrentRangeEnum
    DCVI_IRANGE_2uA = 0
    DCVI_IRANGE_20uA = 1
    DCVI_IRANGE_200uA = 2
    DCVI_IRANGE_2mA = 3
    DCVI_IRANGE_20mA = 4
    DCVI_IRANGE_200mA = 5
    DCVI_IRANGE_2A = 6
End Enum

' ---------------------------------------------------------------
' PPMU Current Range Enum
' ---------------------------------------------------------------
Public Enum PPMUCurrentRangeEnum
    PPMU_IRANGE_10nA = 0
    PPMU_IRANGE_200nA = 1
    PPMU_IRANGE_2uA = 2
    PPMU_IRANGE_20uA = 3
    PPMU_IRANGE_200uA = 4
    PPMU_IRANGE_2mA = 5
End Enum

' ---------------------------------------------------------------
' PLMeter Mode Enum
' ---------------------------------------------------------------
Public Enum PLMeterModeEnum
    PLMETER_DIRECT = 0
    PLMETER_DIFFERENTIAL = 1
End Enum

' ---------------------------------------------------------------
' DCVI Current Range Conversion
' ---------------------------------------------------------------
Public Function GetDCVICurrentRange(rangeEnum As DCVICurrentRangeEnum) _
        As Double
    Select Case rangeEnum
        Case DCVI_IRANGE_2uA:   GetDCVICurrentRange = 2 * uA
        Case DCVI_IRANGE_20uA:  GetDCVICurrentRange = 20 * uA
        Case DCVI_IRANGE_200uA: GetDCVICurrentRange = 200 * uA
        Case DCVI_IRANGE_2mA:   GetDCVICurrentRange = 2 * mA
        Case DCVI_IRANGE_20mA:  GetDCVICurrentRange = 20 * mA
        Case DCVI_IRANGE_200mA: GetDCVICurrentRange = 200 * mA
        Case DCVI_IRANGE_2A:    GetDCVICurrentRange = 2
        Case Else:              GetDCVICurrentRange = 200 * mA
    End Select
End Function

' ---------------------------------------------------------------
' Init DCVI -- configure all supply channels
' ---------------------------------------------------------------
Public Sub InitDCVI()
    With TheHdw.DCVI.Pins(c_supPinList)
        .Gate = False
        .Mode = tlDCVIModeVoltage
        .Current = 50 * mA
        .CurrentRange = 200 * mA
        .NominalBandwidth = 50000
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 200 * mA
        .Meter.Filter.Value = 500
        .Meter.Filter.Bypass = False
        .ComplianceRange(tlDCVICompliancePositive) = 10
        .ComplianceRange(tlDCVIComplianceNegative) = -10
        .BleederResistor = tlDCVIBleederEnabled
        .Connect
        .Gate = True
    End With
End Sub

' ---------------------------------------------------------------
' Init PPMU -- configure parametric pins
' ---------------------------------------------------------------
Public Sub InitPPMU(pins As String)
    With TheHdw.PPMU.Pins(pins)
        .ForceV 0, 2 * mA
        .Connect
        .Gate = tlOn
    End With
End Sub

' ---------------------------------------------------------------
' Init PLMeter -- configure precision level meter
' ---------------------------------------------------------------
Public Sub InitPLMeter(meterPin As String)
    With TheHdw.PLMeter.Pins(meterPin)
        .Mode = tlPLMeterModeDirect
        .Measurement.SampleRate = 100000
        .Measurement.SampleSize = 200
        .Measurement.Filter = tlPLMeterFilter200K
        .Measurement.FilterDelay = 0
        .Connect
        .AsynchronousTrigger.Action = tlPLMeterActionIgnore
        .AlarmLatching = True
        .VoltageRange = 10
    End With
End Sub

' ---------------------------------------------------------------
' Init DCDiffMeter -- configure differential voltage meter
' ---------------------------------------------------------------
Public Sub InitDiffMeter(meterPin As String)
    With TheHdw.DCDiffMeter.Pins(meterPin)
        .HardwareAveraging = 16
        .VoltageRange = 7
        .Measurement.SampleRate = 100000
        .Measurement.SampleSize = 1
        .Connect
    End With
End Sub

' ---------------------------------------------------------------
' MUX Select -- route analog channel through DIB MUX
' ---------------------------------------------------------------
Public Sub MuxSelect(channel As MuxSelectEnum)
    ' <write MUX control via utility pins or digital pattern>
    ' Example using utility relay pins:
    ' TheHdw.Utility.Pins("MUX_A0").RelayMode = ...
    ' TheHdw.Utility.Pins("MUX_A1").RelayMode = ...
    TheHdw.Wait 2   ' ms relay settling
End Sub
```

---

## 4. Exec Interpose Module Template

```vba
Attribute VB_Name = "Interpose"
Option Explicit
''' Exec Interpose hooks -- lifecycle management for UltraFlex programs.

' ---------------------------------------------------------------
' OnProgramLoaded
' ---------------------------------------------------------------
Public Sub OnProgramLoaded()
    ' Set working directory to program location
    ChDrive TheExec.ProgramPath
    ChDir TheExec.ProgramPath

    ' If using vblite, add register map library reference here.
    ' Otherwise, no action needed.

    ' Sort pin list data for consistent ordering
    tl_PinListDataSort False
End Sub

' ---------------------------------------------------------------
' OnProgramValidated
' ---------------------------------------------------------------
Public Sub OnProgramValidated()
    g_GlobalVarReset = True

    ' Enable all sites initially
    Dim site As Variant
    For Each site In TheExec.Sites
        TheExec.Sites.Item(site).Active = True
    Next site

    ' SetOnceSetup: identify part revision, set compile flags
    ' <read device ID register via JTAG/SWD>
    ' <set conditional compilation flags based on revision>

    ' Configure TDR exclusion pins
    ' TheHdw.Digital.TDR.ExcludePins "<TDR_EXCLUDE_PINS>"

    ' Pin respecification for simulation / offline mode
    If TheExec.TesterMode = testModeOffline Then
        ' <respec pins for offline simulation>
    End If
End Sub

' ---------------------------------------------------------------
' OnProgramStarted
' ---------------------------------------------------------------
Public Sub OnProgramStarted()
    ' Reinitialize if globals were reset
    If g_GlobalVarReset Then
        g_GlobalVarReset = False
        ' <re-init any COM objects, DLL handles, etc.>
    End If

    ' Start test timer
    theTimer.DieStart

    ' Set flow flag to latch mode
    TheExec.Flow.DefaultBranch = tlFlowBranchLatch

    ' Init pattern directories from program path
    ' TheHdw.Digital.PatternDirectory = TheExec.ProgramPath & "\patterns"

    ' Init DIB EEPROM data (board traceability)
    ' <read DIB EEPROM via utility instrument>

    ' Init differential meter calibration
    ' Call InitDiffMeter("<DIFF_METER_PIN>")

    ' If using vblite, init register map objects here.

    ' Power on DIB (5V then 15V relays in sequence)
    TheHdw.Utility.Pins("DIB_5V_RELAY").RelayMode = tlRelayModeOn
    TheHdw.Wait 10
    TheHdw.Utility.Pins("DIB_15V_RELAY").RelayMode = tlRelayModeOn
    TheHdw.Wait 10

    ' Configure MW receiver if applicable
    ' <init MW receiver block>

    ' Init JTAG/SWD communication class
    ' Set g_SWD = New SWDClass
    ' g_SWD.Init

    ' Configure datalog setup
    TheExec.Datalog.Setup.SortType = tlDatalogSortByDUT

    ' Set DUT conditions to nominal
    DUTconds.Class_Reset
    DUTconds.powerOn
End Sub

' ---------------------------------------------------------------
' OnProgramEnded
' ---------------------------------------------------------------
Public Sub OnProgramEnded()
    ' Disconnect differential meter
    ' TheHdw.DCDiffMeter.Pins("<DIFF_METER_PIN>").Disconnect

    ' Power down DUT safely
    DUTconds.powerOff

    ' Open all relays
    TheHdw.Utility.Pins(c_allRelays).RelayMode = tlRelayModeOff

    ' Handle characterisation data output
    ' <flush CSV files, close file handles>

    ' Evaluate board checker pass/fail
    ' <check DIB health counters>
End Sub

' ---------------------------------------------------------------
' OnGlobalVariableReset (UltraFlex-specific)
' ---------------------------------------------------------------
Public Sub OnGlobalVariableReset()
    g_GlobalVarReset = False

    ' Terminate any COM objects
    ' Set g_SWD = Nothing
    ' <release DLL handles>
End Sub

' ---------------------------------------------------------------
' OnAlarmOccurred (UltraFlex-specific)
' ---------------------------------------------------------------
Public Sub OnAlarmOccurred(alarmSource As String, _
                           alarmType As Long, _
                           alarmDescription As String)
    ' Log alarm details
    TheExec.Datalog.WriteComment _
        "ALARM: Source=" & alarmSource & _
        " Type=" & CStr(alarmType) & _
        " Desc=" & alarmDescription

    ' Parse alarm source and take corrective action
    Select Case alarmSource
        Case "DCVI"
            ' Reduce current range or disable gate
            TheExec.Datalog.WriteComment _
                "DCVI alarm -- check for short on supply pins."
        Case "Digital"
            ' Log digital pattern error
            TheExec.Datalog.WriteComment _
                "Digital alarm -- check pattern execution."
        Case Else
            TheExec.Datalog.WriteComment _
                "Unhandled alarm source: " & alarmSource
    End Select
End Sub
```

---

## 5. Continuity Test (PPMU-based)

Uses checkerboard pin grouping (`grp_OS_1` and `grp_OS_2`) to detect
adjacent shorts. Each group is forced while the other is grounded.

```vba
Attribute VB_Name = "VBT_Continuity"
Option Explicit

Public Function TF_Continuity() As Long
    On Error GoTo errHandler
    Dim site As Variant
    Dim tempRes As New PinListData

    theTimer.VBTFunctionStart

    ' --- SETUP: Ground all supplies ---
    With TheHdw.DCVI.Pins(c_supPinList)
        .Mode = tlDCVIModeVoltage
        .Voltage = 0
        .Connect
        .Gate = True
    End With
    TheHdw.Wait 5 * ms

    ' --- SETUP: Disconnect digital, connect PPMU ---
    TheHdw.Digital.DisconnectPins ("OS_PPMU_PINS")
    With TheHdw.PPMU.Pins("OS_PPMU_PINS")
        .ForceV 0, 2 * mA
        .Connect
        .Gate = tlOn
    End With

    ' ---------------------------------------------------------------
    ' MEASURE: Diode to VDD -- group A forced positive, group B at 0V
    ' ---------------------------------------------------------------
    TheHdw.PPMU.Pins("grp_OS_1").ForceI 0.002    ' +2 mA
    TheHdw.Wait 2 * ms
    tempRes = TheHdw.PPMU.Pins("grp_OS_1").Read( _
        tlPPMUReadMeasurements, 1)
    ' <log per-pin results via TestLimit>
    TheHdw.PPMU.Pins("grp_OS_1").ForceV 0, 2 * mA

    ' ---------------------------------------------------------------
    ' MEASURE: Diode to GND -- group A forced negative, group B at 0V
    ' ---------------------------------------------------------------
    TheHdw.PPMU.Pins("grp_OS_1").ForceI -0.002   ' -2 mA
    TheHdw.Wait 2 * ms
    tempRes = TheHdw.PPMU.Pins("grp_OS_1").Read( _
        tlPPMUReadMeasurements, 1)
    ' <log per-pin results via TestLimit>
    TheHdw.PPMU.Pins("grp_OS_1").ForceV 0, 2 * mA

    ' ---------------------------------------------------------------
    ' Repeat for grp_OS_2 (groups swapped)
    ' ---------------------------------------------------------------
    TheHdw.PPMU.Pins("grp_OS_2").ForceI 0.002
    TheHdw.Wait 2 * ms
    tempRes = TheHdw.PPMU.Pins("grp_OS_2").Read( _
        tlPPMUReadMeasurements, 1)
    TheHdw.PPMU.Pins("grp_OS_2").ForceV 0, 2 * mA

    TheHdw.PPMU.Pins("grp_OS_2").ForceI -0.002
    TheHdw.Wait 2 * ms
    tempRes = TheHdw.PPMU.Pins("grp_OS_2").Read( _
        tlPPMUReadMeasurements, 1)
    TheHdw.PPMU.Pins("grp_OS_2").ForceV 0, 2 * mA

    ' --- CLEANUP ---
    TheHdw.PPMU.Pins("OS_PPMU_PINS").ForceV 0
    TheHdw.PPMU.Pins("OS_PPMU_PINS").Disconnect

    ' --- DATALOG ---
    ' <TestLimit calls for each pin's supply and ground diode voltage>
    ' TheExec.Flow.TestLimit resultVal:=tempRes, _
    '     ScaleType:=scaleMilli, unit:=unitVolt, _
    '     ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 6. Continuity Test (DCVI-based)

For supply and high-voltage pins measured through UltraVI80 DCVI channels.
Uses `.BleederResistor` and `.Gate = tlDCVIGateOffHiZ` for safe disconnect.

```vba
Public Function TF_ContinuityDCVI() As Long
    On Error GoTo errHandler
    Dim dcviRes As New PinListData
    Dim site As Variant

    theTimer.VBTFunctionStart

    ' --- SETUP: Configure DCVI for current forcing ---
    With TheHdw.DCVI.Pins("OS_DCVI_PINS")
        .Gate = False
        .Mode = tlDCVIModeCurrent
        .BleederResistor = tlDCVIBleederDisabled
        .SetVoltageAndRange 0, 5
        .SetCurrentAndRange 0.001, 0.002   ' 1 mA force, 2 mA range
        .Connect
        .Gate = True
    End With
    TheHdw.Wait c_pinSettDCVI

    ' --- MEASURE: Diode to VDD (+1 mA) ---
    TheHdw.DCVI.Pins("OS_DCVI_PINS").Current = 0.001
    TheHdw.Wait c_pinSettDCVI
    dcviRes = TheHdw.DCVI.Pins("OS_DCVI_PINS").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)

    ' Check per-site per-channel results
    For Each site In TheExec.Sites
        ' TheHdw.Pins("<pin>").FailCountEx(site)
        ' <compare measured voltage against expected diode drop>
    Next site

    ' --- MEASURE: Diode to GND (-1 mA) ---
    TheHdw.DCVI.Pins("OS_DCVI_PINS").Current = -0.001
    TheHdw.Wait c_pinSettDCVI
    dcviRes = TheHdw.DCVI.Pins("OS_DCVI_PINS").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)

    ' --- CLEANUP ---
    TheHdw.DCVI.Pins("OS_DCVI_PINS").Current = 0
    TheHdw.DCVI.Pins("OS_DCVI_PINS").Gate = tlDCVIGateOffHiZ
    TheHdw.DCVI.Pins("OS_DCVI_PINS").Disconnect

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=dcviRes, _
        ScaleType:=scaleMilli, unit:=unitVolt, _
        ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 7. Pin Leakage Test

Checkerboard forcing pattern: even pins at 0V while odd pins at VDD,
then swap. Detects adjacent pin-to-pin leakage.

```vba
Attribute VB_Name = "VBT_Leakage"
Option Explicit

Public Function TF_Leakage() As Long
    On Error GoTo errHandler
    Dim ppmuLeak_0V As New PinListData
    Dim ppmuLeak_VDD As New PinListData
    Dim dcviLeak_0V As New PinListData
    Dim dcviLeak_VDD As New PinListData

    theTimer.VBTFunctionStart

    ' ---------------------------------------------------------------
    ' PPMU Pins: Phase 1 -- even at 0V, odd at VDD
    ' ---------------------------------------------------------------
    TheHdw.Digital.DisconnectPins ("LKG_PPMU_PINS")
    With TheHdw.PPMU.Pins("LKG_PPMU_EVEN")
        .ForceV 0, 200 * nA
        .Connect
        .Gate = tlOn
    End With
    With TheHdw.PPMU.Pins("LKG_PPMU_ODD")
        .ForceV g_tc(g_CurrentTestCondition).vddIO, 200 * nA
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait c_pinSettPPMU

    ppmuLeak_0V = TheHdw.PPMU.Pins("LKG_PPMU_EVEN").Read( _
        tlPPMUReadMeasurements, 1)
    ppmuLeak_VDD = TheHdw.PPMU.Pins("LKG_PPMU_ODD").Read( _
        tlPPMUReadMeasurements, 1)

    ' ---------------------------------------------------------------
    ' PPMU Pins: Phase 2 -- swap (even at VDD, odd at 0V)
    ' ---------------------------------------------------------------
    TheHdw.PPMU.Pins("LKG_PPMU_EVEN").ForceV _
        g_tc(g_CurrentTestCondition).vddIO, 200 * nA
    TheHdw.PPMU.Pins("LKG_PPMU_ODD").ForceV 0, 200 * nA
    TheHdw.Wait c_pinSettPPMU

    g_PreLeakEvenPinsVDD = TheHdw.PPMU.Pins("LKG_PPMU_EVEN").Read( _
        tlPPMUReadMeasurements, 1)
    g_PreLeakOddPins0V = TheHdw.PPMU.Pins("LKG_PPMU_ODD").Read( _
        tlPPMUReadMeasurements, 1)

    ' ---------------------------------------------------------------
    ' DCVI Pins: Leakage at 0V
    ' ---------------------------------------------------------------
    With TheHdw.DCVI.Pins("LKG_DCVI_PINS")
        .Mode = tlDCVIModeVoltage
        .Meter.Mode = tlDCVIMeterCurrent
        .SetCurrentAndRange 20 * uA, 20 * uA
        .SetVoltageAndRange 0, 5
        .Connect
        .Gate = True
    End With
    TheHdw.Wait c_pinSettDCVI
    dcviLeak_0V = TheHdw.DCVI.Pins("LKG_DCVI_PINS").Meter.Read( _
        tlStrobe, 10, 1000)

    ' ---------------------------------------------------------------
    ' DCVI Pins: Leakage at VDD
    ' ---------------------------------------------------------------
    TheHdw.DCVI.Pins("LKG_DCVI_PINS").Voltage = _
        g_tc(g_CurrentTestCondition).vdd
    TheHdw.Wait c_pinSettDCVI
    dcviLeak_VDD = TheHdw.DCVI.Pins("LKG_DCVI_PINS").Meter.Read( _
        tlStrobe, 10, 1000)

    ' --- CLEANUP ---
    TheHdw.PPMU.Pins("LKG_PPMU_PINS").Disconnect
    TheHdw.DCVI.Pins("LKG_DCVI_PINS").Disconnect

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=ppmuLeak_0V, _
        ScaleType:=scaleNano, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=ppmuLeak_VDD, _
        ScaleType:=scaleNano, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=dcviLeak_0V, _
        ScaleType:=scaleNano, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=dcviLeak_VDD, _
        ScaleType:=scaleNano, unit:=unitAmp, _
        ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 8. Supply Current Measurement

Multi-supply current measurement supporting active, standby, powerdown, and
hibernate modes. Each supply is measured independently with explicit current
ranges.

```vba
Attribute VB_Name = "VBT_SupplyCurrent"
Option Explicit

Public Function TF_SupplyCurrent(mode As String) As Long
    On Error GoTo errHandler
    Dim avddCurr As New PinListData
    Dim dvddCurr As New PinListData
    Dim iovdd0Curr As New PinListData
    Dim iovdd1Curr As New PinListData

    theTimer.VBTFunctionStart

    ' --- Guard: skip in offline mode ---
    If TheExec.TesterMode = testModeOffline Then
        TheExec.Datalog.WriteComment _
            "Offline mode -- skipping supply current."
        theTimer.VBTFunctionFinish
        Exit Function
    End If

    ' --- SETUP: Set DUT to requested power mode ---
    Select Case mode
        Case "ACTIVE"
            ' <run pattern to set active mode>
        Case "STANDBY"
            ' <write power mode register to standby value>
        Case "POWERDOWN"
            ' <write power mode register to powerdown value>
        Case "HIBERNATE"
            ' <write power mode register to hibernate value>
        Case Else
            TheExec.Datalog.WriteComment _
                "Unknown power mode: " & mode
    End Select
    TheHdw.Wait 10   ' ms mode settling

    ' --- SETUP: Configure meter ranges per supply ---
    TheHdw.DCVI.Pins("AVDD").Meter.CurrentRange = 200 * mA
    TheHdw.DCVI.Pins("DVDD").Meter.CurrentRange = 200 * mA
    TheHdw.DCVI.Pins("IOVDD0").Meter.CurrentRange = 20 * mA
    TheHdw.DCVI.Pins("IOVDD1").Meter.CurrentRange = 20 * mA

    ' --- MEASURE: Read all supplies ---
    avddCurr = TheHdw.DCVI.Pins("AVDD").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)
    dvddCurr = TheHdw.DCVI.Pins("DVDD").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)
    iovdd0Curr = TheHdw.DCVI.Pins("IOVDD0").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)
    iovdd1Curr = TheHdw.DCVI.Pins("IOVDD1").Meter.Read( _
        tlStrobe, c_DCVISampSize, c_DCVISampRate)

    ' --- DATALOG: Log each supply separately ---
    TheExec.Flow.TestLimit resultVal:=avddCurr, _
        ScaleType:=scaleMilli, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=dvddCurr, _
        ScaleType:=scaleMilli, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=iovdd0Curr, _
        ScaleType:=scaleMilli, unit:=unitAmp, _
        ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=iovdd1Curr, _
        ScaleType:=scaleMilli, unit:=unitAmp, _
        ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 9. IDDQ Test (Supply Switching Method)

UltraFlex IDDQ uses BBAC capture of voltage decay on an external capacitor.
The pattern switches the supply from DCVI to the capacitor, and the leakage
current is calculated from the voltage drop.

```vba
Attribute VB_Name = "VBT_IDDQ"
Option Explicit

Public Function TF_IDDQ() As Long
    On Error GoTo errHandler
    Dim site As Variant
    Dim capVoltage() As Double
    Dim iddqCurrent As New SiteDouble
    Dim deltaV As Double
    Dim deltaT As Double

    ' External cap value in Farads (on DIB)
    Dim Const c_iddqCapValue As Double = 0.0000001   ' 100 nF

    ' Number of IDDQ measurement points in the pattern
    Dim Const c_iddqMeasPoints As Long = 10

    theTimer.VBTFunctionStart

    ' --- Guard: skip in offline mode ---
    If TheExec.TesterMode = testModeOffline Then
        iddqCurrent = 0
        GoTo DatalogSection
    End If

    ' --- SETUP: Configure HRAM for fail capture ---
    TheHdw.Digital.HRAM.Enable = True
    TheHdw.Digital.Patgen.NoHaltMode = noHaltAlways

    ' --- SETUP: Connect BBAC capture to IDDQ circuit ---
    ' <configure BBAC source/capture for voltage measurement>
    ' TheHdw.Digital.DSSC.Source("<BBAC_PIN>").LoadWave ...
    ' TheHdw.Digital.DSSC.Capture("<BBAC_PIN>").Enable = True

    ' --- MEASURE: Run IDDQ pattern ---
    ' Pattern switches supply from DCVI to cap at each vector
    TheHdw.Digital.Patterns.Pat("iddq_measure.pat").Start
    TheHdw.Digital.Patgen.HaltWait

    ' --- CALCULATE: I = deltaV * C / deltaT ---
    ' <extract captured voltage samples from BBAC>
    ' deltaV = capVoltage(0) - capVoltage(c_iddqMeasPoints - 1)
    ' deltaT = <time between first and last sample>
    ' For Each site In TheExec.Sites
    '     iddqCurrent = deltaV * c_iddqCapValue / deltaT
    ' Next site

    ' --- Find min/max/mean across measure points ---
    ' <calculate statistics from per-point measurements>

DatalogSection:
    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=iddqCurrent, _
        ScaleType:=scaleNano, unit:=unitAmp, _
        ForceResults:=tlForceFlow

    ' --- CLEANUP ---
    TheHdw.Digital.HRAM.Enable = False
    TheHdw.Digital.Patgen.NoHaltMode = noHaltDisable

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 10. Scan Chain Test (DSSC-based)

UltraFlex scan uses DSSC source/capture with external text files for scan
chain data. Chain errors are counted per scan chain.

```vba
Attribute VB_Name = "VBT_ScanChain"
Option Explicit

Public Function TF_ScanChain() As Long
    On Error GoTo errHandler
    Dim site As Variant
    Dim chainErrors As New SiteLong
    Dim capturedData As New DSPWave
    Dim expectedData As New DSPWave
    Dim maskData As New DSPWave

    theTimer.VBTFunctionStart

    ' --- SETUP: Load scan source/capture/mask from text files ---
    ' Dim srcFile As String
    ' srcFile = TheExec.ProgramPath & "\scan_data\chain_src.txt"
    ' Dim expFile As String
    ' expFile = TheExec.ProgramPath & "\scan_data\chain_exp.txt"
    ' Dim mskFile As String
    ' mskFile = TheExec.ProgramPath & "\scan_data\chain_mask.txt"

    ' expectedData.LoadFromFile expFile
    ' maskData.LoadFromFile mskFile

    ' --- SETUP: Configure DSSC source with scan-in data ---
    ' TheHdw.Digital.DSSC.Source("SCAN_IN").LoadWaveFromFile srcFile
    ' TheHdw.Digital.DSSC.Source("SCAN_IN").Enable = True

    ' --- SETUP: Configure DSSC capture for scan-out ---
    ' TheHdw.Digital.DSSC.Capture("SCAN_OUT").Enable = True

    ' --- MEASURE: Run scan pattern ---
    TheHdw.Digital.Patterns.Pat("scan_shift.pat").Start
    TheHdw.Digital.Patgen.HaltWait

    ' --- COMPARE: Captured vs expected using mask ---
    ' capturedData = TheHdw.Digital.DSSC.Capture("SCAN_OUT").ReadWave
    ' For Each site In TheExec.Sites
    '     chainErrors = 0
    '     Dim i As Long
    '     For i = 0 To capturedData.Length - 1
    '         If maskData.Data(i) = 1 Then
    '             If capturedData.Data(i) <> expectedData.Data(i) Then
    '                 chainErrors = chainErrors + 1
    '             End If
    '         End If
    '     Next i
    ' Next site

    ' --- CLEANUP ---
    ' TheHdw.Digital.DSSC.Source("SCAN_IN").Enable = False
    ' TheHdw.Digital.DSSC.Capture("SCAN_OUT").Enable = False

    ' --- DATALOG: Per-chain error counts ---
    TheExec.Flow.TestLimit resultVal:=chainErrors, _
        unit:=unitNone, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 11. Logic Level Verification (Production)

Verifies input and output logic levels meet datasheet specifications at
stressed voltage values.

```vba
Attribute VB_Name = "VBT_LogicLevels"
Option Explicit

Public Function TF_CheckLogicLevels() As Long
    On Error GoTo errHandler
    Dim vihResult As New SiteLong
    Dim vilResult As New SiteLong
    Dim vohResult As New SiteLong
    Dim volResult As New SiteLong
    Dim site As Variant

    ' Stressed voltage levels (datasheet limits with margin)
    Dim Const c_vihStressed As Double = 2.4    ' worst-case VIH
    Dim Const c_vilStressed As Double = 0.8    ' worst-case VIL
    Dim Const c_vohStressed As Double = 2.4    ' VOH threshold
    Dim Const c_volStressed As Double = 0.6    ' VOL threshold

    theTimer.VBTFunctionStart

    ' --- SETUP: Set levels to nominal ---
    Call setLevelsSafe(c_nomVlogSupp)

    ' ---------------------------------------------------------------
    ' VIH check: set VIH to stressed value, run pattern, check pass
    ' ---------------------------------------------------------------
    TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel chVih, c_vihStressed
    TheHdw.Digital.Patterns.Pat("comms_check.pat").Start
    TheHdw.Digital.Patgen.HaltWait
    For Each site In TheExec.Sites
        vihResult = TheHdw.Pins("CHECK_PIN").FailCountEx(site)
    Next site
    Call setLevelsSafe(c_nomVlogSupp)

    ' ---------------------------------------------------------------
    ' VIL check
    ' ---------------------------------------------------------------
    TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel chVil, c_vilStressed
    TheHdw.Digital.Patterns.Pat("comms_check.pat").Start
    TheHdw.Digital.Patgen.HaltWait
    For Each site In TheExec.Sites
        vilResult = TheHdw.Pins("CHECK_PIN").FailCountEx(site)
    Next site
    Call setLevelsSafe(c_nomVlogSupp)

    ' ---------------------------------------------------------------
    ' VOH check
    ' ---------------------------------------------------------------
    TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel chVoh, c_vohStressed
    TheHdw.Digital.Patterns.Pat("comms_check.pat").Start
    TheHdw.Digital.Patgen.HaltWait
    For Each site In TheExec.Sites
        vohResult = TheHdw.Pins("CHECK_PIN").FailCountEx(site)
    Next site
    Call setLevelsSafe(c_nomVlogSupp)

    ' ---------------------------------------------------------------
    ' VOL check
    ' ---------------------------------------------------------------
    TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel chVol, c_volStressed
    TheHdw.Digital.Patterns.Pat("comms_check.pat").Start
    TheHdw.Digital.Patgen.HaltWait
    For Each site In TheExec.Sites
        volResult = TheHdw.Pins("CHECK_PIN").FailCountEx(site)
    Next site
    Call setLevelsSafe(c_nomVlogSupp)

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=vihResult, _
        unit:=unitNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=vilResult, _
        unit:=unitNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=vohResult, _
        unit:=unitNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=volResult, _
        unit:=unitNone, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 12. Logic Level Characterisation (Threshold Search)

Helper functions that sweep a logic level from known-pass to known-fail (P2F)
or fail-to-pass (F2P), returning the threshold voltage.

```vba
''' Sweep from known-pass to fail. Returns last passing voltage.
''' Returns -99 if immediate fail, 99 if never failed.
Public Function levelThresP2F(levelType As Long, _
                               startVal As Double, _
                               stopVal As Double, _
                               stepSize As Double, _
                               patternName As String, _
                               checkPin As String) As SiteDouble
    On Error GoTo errHandler
    Dim result As New SiteDouble
    Dim currentVal As Double
    Dim failCount As New SiteLong
    Dim lastPass As New SiteDouble
    Dim foundFail As New SiteBoolean
    Dim site As Variant

    Dim Const c_errorImmedFail As Double = -99
    Dim Const c_errorNeverFail As Double = 99

    result = c_errorNeverFail
    lastPass = startVal
    foundFail = False

    currentVal = startVal
    Do While currentVal <= stopVal
        ' Set the level under test
        TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel _
            levelType, currentVal

        ' Run check pattern
        TheHdw.Digital.Patterns.Pat(patternName).Start
        TheHdw.Digital.Patgen.HaltWait

        ' Check per-site pass/fail
        For Each site In TheExec.Sites
            failCount = TheHdw.Pins(checkPin).FailCountEx(site)
            If failCount = 0 Then
                lastPass = currentVal
            Else
                If Not foundFail Then
                    foundFail = True
                    result = lastPass
                End If
            End If
        Next site

        currentVal = currentVal + stepSize
    Loop

    ' Check for immediate fail (first value failed)
    For Each site In TheExec.Sites
        If foundFail And lastPass = startVal Then
            result = c_errorImmedFail
        End If
    Next site

    ' Restore nominal levels
    Call setLevelsSafe(c_nomVlogSupp)

    Set levelThresP2F = result
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function

''' Sweep from known-fail to pass. Returns first passing voltage.
''' Returns -99 if immediate pass, 99 if never passed.
Public Function levelThresF2P(levelType As Long, _
                               startVal As Double, _
                               stopVal As Double, _
                               stepSize As Double, _
                               patternName As String, _
                               checkPin As String) As SiteDouble
    On Error GoTo errHandler
    Dim result As New SiteDouble
    Dim currentVal As Double
    Dim failCount As New SiteLong
    Dim foundPass As New SiteBoolean
    Dim site As Variant

    Dim Const c_errorImmedPass As Double = -99
    Dim Const c_errorNeverPass As Double = 99

    result = c_errorNeverPass
    foundPass = False

    currentVal = startVal
    Do While currentVal <= stopVal
        TheHdw.PinLevels.Pins(c_allDigPins).ModifyLevel _
            levelType, currentVal

        TheHdw.Digital.Patterns.Pat(patternName).Start
        TheHdw.Digital.Patgen.HaltWait

        For Each site In TheExec.Sites
            failCount = TheHdw.Pins(checkPin).FailCountEx(site)
            If failCount = 0 And Not foundPass Then
                foundPass = True
                result = currentVal
            End If
        Next site

        currentVal = currentVal + stepSize
    Loop

    Call setLevelsSafe(c_nomVlogSupp)

    Set levelThresF2P = result
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 13. GPIO Functional Test (Register-based)

Tests GPIO input and output functionality using register access via
SWD/JTAG at multiple IO voltage levels. Register read/write can use
vblite-generated APIs or pattern-based SPI/JTAG transactions.

```vba
Attribute VB_Name = "VBT_GPIO"
Option Explicit

Public Function TF_GPIOTest() As Long
    On Error GoTo errHandler
    Dim site As Variant
    Dim readVal As New SiteLong
    Dim expectedVal As Long
    Dim gpioResult As New SiteBoolean
    Dim measVoltage As New PinListData

    ' IO voltage levels to test
    Dim Const c_numVoltLevels As Long = 3
    Dim voltLevels(0 To 2) As Double
    voltLevels(0) = 3.3
    voltLevels(1) = 1.8
    voltLevels(2) = 1.2

    theTimer.VBTFunctionStart
    gpioResult = True

    Dim vIdx As Long
    For vIdx = 0 To c_numVoltLevels - 1
        ' --- Set IO voltage ---
        DUTconds.setIOVdd voltLevels(vIdx), "IOVDD0"
        TheHdw.Wait 5

        ' ---------------------------------------------------------------
        ' INPUT TEST: Configure GPIO as inputs, force via PPMU, read reg
        ' ---------------------------------------------------------------
        ' bfWrite BF_gpio.GP0OEN, &H0, True     ' disable outputs
        ' bfWrite BF_gpio.GP0IE, &HFF, True      ' enable inputs

        ' Force logic high on all GPIO pins
        TheHdw.Digital.DisconnectPins ("GPIO_PINS")
        TheHdw.PPMU.Pins("GPIO_PINS").ForceV voltLevels(vIdx), 2 * mA
        TheHdw.PPMU.Pins("GPIO_PINS").Connect
        TheHdw.PPMU.Pins("GPIO_PINS").Gate = tlOn
        TheHdw.Wait 1

        ' <read GPIO input register via your device communication method>
        expectedVal = &HFF
        For Each site In TheExec.Sites
            If readVal <> expectedVal Then gpioResult = False
        Next site

        ' Force logic low on all GPIO pins
        TheHdw.PPMU.Pins("GPIO_PINS").ForceV 0, 2 * mA
        TheHdw.Wait 1
        ' readVal = bfRead(BF_gpio.GP0IN_Y)
        expectedVal = &H0
        For Each site In TheExec.Sites
            If readVal <> expectedVal Then gpioResult = False
        Next site

        TheHdw.PPMU.Pins("GPIO_PINS").Disconnect

        ' ---------------------------------------------------------------
        ' OUTPUT TEST: Configure as outputs, load with PPMU, measure V
        ' ---------------------------------------------------------------
        ' bfWrite BF_gpio.GP0OEN, &HFF, True    ' enable outputs
        ' bfWrite BF_gpio.GP0IE, &H0, True       ' disable inputs
        ' bfWrite BF_gpio.GP0OUT, &HFF, True     ' drive high

        TheHdw.PPMU.Pins("GPIO_PINS").ForceI -0.0002  ' -200 uA load
        TheHdw.PPMU.Pins("GPIO_PINS").Connect
        TheHdw.PPMU.Pins("GPIO_PINS").Gate = tlOn
        TheHdw.Wait 1

        measVoltage = TheHdw.PPMU.Pins("GPIO_PINS").Read( _
            tlPPMUReadMeasurements, 1)
        ' <verify VOH: measVoltage > voltLevels(vIdx) - 0.5>

        TheHdw.PPMU.Pins("GPIO_PINS").Disconnect
    Next vIdx

    ' --- Restore nominal ---
    DUTconds.setIOVdd c_nomVlogSupp, "IOVDD0"

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=gpioResult, _
        unit:=unitNone, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 14. Trim Template (Binary Search)

Generic binary-search trim loop. Write trim code to DUT register, measure
the result, and converge on the target value.

```vba
Attribute VB_Name = "VBT_Trim"
Option Explicit

Public Function TF_BinaryTrim(targetVal As Double, _
                               maxDiff As Double, _
                               minCode As Long, _
                               maxCode As Long) As Long
    On Error GoTo errHandler
    Dim trimCode As New SiteLong
    Dim measVal As New SiteDouble
    Dim lo As New SiteLong
    Dim hi As New SiteLong
    Dim mid As New SiteLong
    Dim site As Variant

    Dim Const c_maxIterations As Long = 20

    theTimer.VBTFunctionStart

    lo = minCode
    hi = maxCode

    ' --- Binary search loop ---
    Dim iteration As Long
    For iteration = 0 To c_maxIterations
        For Each site In TheExec.Sites
            mid = (lo + hi) / 2
        Next site

        ' <write trim code to DUT register via your device communication method>
        TheHdw.Wait 1   ' ms settling after trim write

        ' Measure result
        ' measVal = <call measurement function>

        ' Update search bounds per site
        For Each site In TheExec.Sites
            If measVal < targetVal Then
                lo = mid
            Else
                hi = mid
            End If

            ' Check convergence
            If Abs(measVal - targetVal) < maxDiff Then
                trimCode = mid
            End If
        Next site
    Next iteration

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=trimCode, _
        unit:=unitNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=measVal, _
        unit:=unitVolt, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 15. Power-On Reset / Power Cycle

Verifies that the DUT boots correctly after a power cycle and that
registers contain expected default values.

```vba
Public Function TF_PowerOnReset() As Long
    On Error GoTo errHandler
    Dim porResult As New SiteBoolean
    Dim regDefault As New SiteLong

    theTimer.VBTFunctionStart

    ' --- SETUP: Power cycle the DUT ---
    DUTconds.powerCycle
    TheHdw.Wait 50   ' ms boot time

    ' --- MEASURE: Read a known default register ---
    ' <read chip ID register via your device communication method>
    ' porResult = (regDefault = <expected_default>)

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=porResult, _
        unit:=unitNone, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 16. Reference Voltage Measurement

Measures an on-chip reference voltage using the PLMeter or DCDiffMeter
instrument.

```vba
Public Function TF_RefVoltage() As Long
    On Error GoTo errHandler
    Dim refMeas As New PinListData

    theTimer.VBTFunctionStart

    ' --- SETUP: Route VREF to measurement pin ---
    ' <configure internal MUX via register write>
    ' MuxSelect MUX_VREF
    TheHdw.Wait 5   ' ms settling

    ' --- MEASURE: Read voltage via PLMeter ---
    Call InitPLMeter("VREF_METER")
    refMeas = TheHdw.PLMeter.Pins("VREF_METER").Measurement.Read

    ' --- CLEANUP ---
    TheHdw.PLMeter.Pins("VREF_METER").Disconnect

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=refMeas, _
        ScaleType:=scaleNone, unit:=unitVolt, _
        ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 17. eFuse / OTP Programming

Programs one-time-programmable (OTP) memory cells. Requires key-unlock
sequences to access protected registers. Register access can use
vblite-generated APIs or pattern-based transactions.

```vba
Public Function TF_eFuseProgram() As Long
    On Error GoTo errHandler
    Dim fuseResult As New SiteBoolean

    theTimer.VBTFunctionStart

    ' --- SETUP: Unlock OTP registers via key sequence ---
    ' <write unlock key 0 to OTP key register>
    ' <write unlock key 1 to OTP key register>
    ' TheHdw.Wait 1

    ' --- PROGRAM: Write fuse data and trigger blow ---
    ' <write fuse data value to OTP data register>
    ' <write blow command to OTP command register>
    ' TheHdw.Wait 10   ' ms programming pulse

    ' --- VERIFY: Read back and compare ---
    ' <read back OTP data register>
    ' fuseResult = (readback = <fuse_value>)

    ' --- CLEANUP: Re-lock OTP registers ---
    ' <write 0 to OTP key registers to re-lock>

    ' --- DATALOG ---
    TheExec.Flow.TestLimit resultVal:=fuseResult, _
        unit:=unitNone, ForceResults:=tlForceFlow

    theTimer.VBTFunctionFinish
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment _
        "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next
End Function
```

---

## 18. DUTconditions Power Sequence Utilities

Standard power-up and power-down sequences used by the `DUTconditions` class.
These are reference patterns showing the correct ordering and EOS-safe
ramping for UltraFlex DCVI channels.

### Power-Up Sequence

```vba
''' Standard power-up sequence:
''' 1. Set DCVI properties (mode, current range, bandwidth, compliance)
''' 2. Connect and gate supplies (all initially at 0V)
''' 3. Ramp DVDD with EOS-safe exponential steps
''' 4. Ramp AVDD with EOS-safe exponential steps
''' 5. Ramp IOVDD0 with EOS-safe exponential steps
''' 6. Ramp IOVDD1 with EOS-safe exponential steps
''' 7. ApplyLevelsTiming (connect digital pins at correct levels)
''' 8. Hold reset pin high via PPMU
''' 9. Wait for full settling

Public Sub powerUpSequence()
    ' Step 1: Init all DCVI channels
    With TheHdw.DCVI.Pins(c_supPinList)
        .Gate = False
        .Mode = tlDCVIModeVoltage
        .Current = 50 * mA
        .CurrentRange = 200 * mA
        .NominalBandwidth = 50000
        .ComplianceRange(tlDCVICompliancePositive) = 10
        .ComplianceRange(tlDCVIComplianceNegative) = -10
        .Connect
        .Gate = True
    End With

    ' Steps 3-6: Ramp supplies in order (core first, IO last)
    Call rampDCVI("DVDD", c_nomDVDD, 0)
    Call rampDCVI("AVDD", c_nomAVDD, 0)
    Call rampDCVI("IOVDD0", c_nomVlogSupp, 0)
    Call rampDCVI("IOVDD1", c_nomVlogSupp, 0)

    ' Step 7: Connect digital pins
    TheHdw.Digital.ApplyLevelsTiming

    ' Step 8: Assert reset via PPMU
    With TheHdw.PPMU.Pins("RESETN")
        .ForceV c_nomVlogSupp, 2 * mA
        .Connect
        .Gate = tlOn
    End With

    ' Step 9: Final settling
    TheHdw.Wait 10   ' ms
End Sub
```

### Power-Down Sequence

```vba
''' Standard power-down sequence:
''' 1. Disconnect digital pins
''' 2. Set all supplies to 0V (reverse order of power-up)
''' 3. Open all relays
''' 4. Disconnect all DCVI channels
''' 5. Disconnect PPMU

Public Sub powerDownSequence()
    ' Step 1: Disconnect digital
    TheHdw.Digital.DisconnectAllPins

    ' Step 2: Ramp supplies down (reverse order)
    Call rampDCVI("IOVDD1", 0, c_nomVlogSupp)
    Call rampDCVI("IOVDD0", 0, c_nomVlogSupp)
    Call rampDCVI("AVDD", 0, c_nomAVDD)
    Call rampDCVI("DVDD", 0, c_nomDVDD)

    ' Step 3: Open all relays
    TheHdw.Utility.Pins(c_allRelays).RelayMode = tlRelayModeOff

    ' Step 4: Disconnect DCVI
    TheHdw.DCVI.Pins(c_supPinList).Gate = False
    TheHdw.DCVI.Pins(c_supPinList).Disconnect

    ' Step 5: Disconnect PPMU
    TheHdw.PPMU.Pins("RESETN").Gate = tlOff
    TheHdw.PPMU.Pins("RESETN").Disconnect

    TheHdw.Wait 5   ' ms final wait
End Sub
```

---

## 19. Instrument Initialization Utilities

UltraFlex-specific instrument initialization patterns and the
`setLevelsSafe` utility for establishing safe digital pin levels.

### DCVI Initialization

```vba
Public Sub InitDCVI()
    With TheHdw.DCVI.Pins(c_supPinList)
        .Gate = False
        .Mode = tlDCVIModeVoltage
        .Current = 50 * mA
        .CurrentRange = 200 * mA
        .NominalBandwidth = 50000
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 200 * mA
        .Meter.Filter.Value = 500
        .Meter.Filter.Bypass = False
        .ComplianceRange(tlDCVICompliancePositive) = 10
        .ComplianceRange(tlDCVIComplianceNegative) = -10
        .Connect
        .Gate = True
    End With
End Sub
```

### PLMeter Initialization

```vba
Public Sub InitPLMeter(meterPin As String)
    With TheHdw.PLMeter.Pins(meterPin)
        .Mode = tlPLMeterModeDirect
        .Measurement.SampleRate = 100000
        .Measurement.SampleSize = 200
        .Measurement.Filter = tlPLMeterFilter200K
        .Measurement.FilterDelay = 0
        .Connect
        .AsynchronousTrigger.Action = tlPLMeterActionIgnore
        .AlarmLatching = True
        .VoltageRange = 10
    End With
End Sub
```

### DCDiffMeter Initialization

```vba
Public Sub InitDiffMeter(meterPin As String)
    With TheHdw.DCDiffMeter.Pins(meterPin)
        .HardwareAveraging = 16
        .VoltageRange = 7
        .Measurement.SampleRate = 100000
        .Measurement.SampleSize = 1
        .Connect
    End With
End Sub
```

### Safe Level Setting

Sets all digital pin levels to safe values based on the logic supply voltage.
Call this before and after any level stress tests.

```vba
Public Sub setLevelsSafe(Vlog As Double)
    With TheHdw.PinLevels.Pins(c_allDigPins)
        .ModifyLevel chVch, Vlog + 0.5      ' clamp high
        .ModifyLevel chVcl, -0.5             ' clamp low
        .ModifyLevel chVil, 0                ' input low
        .ModifyLevel chVih, Vlog             ' input high
        .ModifyLevel chVoh, Vlog - 0.5       ' output high threshold
        .ModifyLevel chVol, 0.6              ' output low threshold
        .ModifyLevel chVt, Vlog * 0.5        ' termination voltage
        .ModifyLevel chIoh, -200 * uA        ' output high current
        .ModifyLevel chIol, 200 * uA         ' output low current
    End With
End Sub
```

### PPMU Quick Connect/Disconnect

```vba
Public Sub PPMUConnect(pins As String, forceV As Double, _
                       iRange As Double)
    With TheHdw.PPMU.Pins(pins)
        .ForceV forceV, iRange
        .Connect
        .Gate = tlOn
    End With
End Sub

Public Sub PPMUDisconnect(pins As String)
    With TheHdw.PPMU.Pins(pins)
        .ForceV 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
End Sub
```

### DCVI Quick Voltage Read

```vba
Public Function ReadDCVIVoltage(pin As String, _
                                 sampleSize As Long, _
                                 sampleRate As Double) As PinListData
    TheHdw.DCVI.Pins(pin).Meter.Mode = tlDCVIMeterVoltage
    Set ReadDCVIVoltage = TheHdw.DCVI.Pins(pin).Meter.Read( _
        tlStrobe, sampleSize, sampleRate)
End Function
```

### DCVI Quick Current Read

```vba
Public Function ReadDCVICurrent(pin As String, _
                                 sampleSize As Long, _
                                 sampleRate As Double) As PinListData
    TheHdw.DCVI.Pins(pin).Meter.Mode = tlDCVIMeterCurrent
    Set ReadDCVICurrent = TheHdw.DCVI.Pins(pin).Meter.Read( _
        tlStrobe, sampleSize, sampleRate)
End Function
```
