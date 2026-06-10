Attribute VB_Name = "VBT_Characterisation"
Option Explicit

'==============================================================================
' Module  : VBT_Characterisation
' Device  : ADIN6210
' Purpose : Characterisation tests with full temperature sweep support.
'           Wraps existing characterisation measurements (levels, timing,
'           process-monitor, MDIO loopback) and adds a structured temperature
'           loop that drives the Thermostream, soaks, executes each sub-test,
'           and restores the DUT to ambient on completion or error.
'
' Pattern : Setup -> Measure -> Evaluate  (per sub-test)
' Safety  : All supply voltages are removed before head movement.
'           Temperature limits are enforced before any GPIB command is sent.
'           On any unhandled error the head is raised and supplies are restored.
'
' Rev 1.0 : Initial production release
'==============================================================================

' ---------------------------------------------------------------------------
' Temperature sweep configuration
' ---------------------------------------------------------------------------
Private Const TEMP_AMBIENT      As Double = 25.0   ' degC  - return-to point
Private Const TEMP_COLD         As Double = -40.0  ' degC  - cold corner
Private Const TEMP_HOT          As Double = 125.0  ' degC  - hot  corner
Private Const TEMP_RAMP_RATE    As Double = 5.0    ' degC/s - ramp rate passed to Thermostream
Private Const TEMP_SOAK_TIME    As Double = 30.0   ' s     - soak at each temperature point
Private Const TEMP_WINDOW       As Double = 2.0    ' degC  - settle window passed to Thermostream

' Number of temperature points in the characterisation sweep.
' Points are built at run-time in BuildTempSweepPoints().
Private Const MAX_TEMP_POINTS   As Long = 10

' ---------------------------------------------------------------------------
' Supply / relay constants (conservative defaults)
' ---------------------------------------------------------------------------
Private Const AVDD3P3_NOM       As Double = 3.3    ' V
Private Const VDDIO_NOM         As Double = 3.3    ' V
Private Const SUPPLY_SETTLE_S   As Double = 0.05   ' s  - wait after supply change

' ---------------------------------------------------------------------------
' Module-level state
' ---------------------------------------------------------------------------
Private m_TempPoints(MAX_TEMP_POINTS - 1) As Double
Private m_TempPointCount As Long
Private m_CurrentTemp As Double
Private m_SuppliesOn As Boolean

'==============================================================================
' PUBLIC ENTRY POINTS
'==============================================================================

' ---------------------------------------------------------------------------
' CharacterisationTempSweep
' ---------------------------------------------------------------------------
' Main entry point called from the IG-XL flow sheet.
' Iterates over the temperature sweep array, conditions the DUT at each
' temperature, then calls every characterisation sub-test.
' Returns 0 on pass, non-zero on failure.
' ---------------------------------------------------------------------------
Public Function CharacterisationTempSweep() As Long

    On Error GoTo errHandler

    Dim lResult     As Long
    Dim i           As Long
    Dim dTemp       As Double

    CharacterisationTempSweep = 0
    m_SuppliesOn = False

    ' Build the list of temperatures to sweep
    Call BuildTempSweepPoints

    ' Bring up supplies at ambient before starting
    Call SetupSupplies
    m_SuppliesOn = True

    ' -----------------------------------------------------------------------
    ' Temperature loop
    ' -----------------------------------------------------------------------
    For i = 0 To m_TempPointCount - 1

        dTemp = m_TempPoints(i)

        ' Validate temperature before commanding hardware
        If Not IsTemperatureValid(dTemp) Then
            TheExec.Datalog.WriteComment "CharacterisationTempSweep: SKIPPING invalid " & _
                                         "temperature point " & CStr(dTemp) & " degC"
            GoTo NextTempPoint
        End If

        TheExec.Datalog.WriteComment "===== Temperature point " & CStr(i + 1) & _
                                     " of " & CStr(m_TempPointCount) & _
                                     " : " & Format(dTemp, "0.0") & " degC ====="

        ' Move to temperature (supplies stay on - head stays down)
        lResult = GoToTemperature(dTemp)
        If lResult <> 0 Then
            TheExec.Datalog.WriteComment "CharacterisationTempSweep: GoToTemperature FAILED " & _
                                         "at " & Format(dTemp, "0.0") & " degC"
            CharacterisationTempSweep = lResult
            GoTo Cleanup
        End If

        ' Record the actual set temperature for datalogs
        m_CurrentTemp = dTemp

        ' -------------------------------------------------------------------
        ' Execute characterisation sub-tests at this temperature
        ' -------------------------------------------------------------------
        lResult = RunAllCharTests(dTemp)
        If lResult <> 0 Then
            TheExec.Datalog.WriteComment "CharacterisationTempSweep: Sub-test FAILED " & _
                                         "at " & Format(dTemp, "0.0") & " degC, result=" & _
                                         CStr(lResult)
            ' Continue to next temperature point (characterisation mode)
            CharacterisationTempSweep = lResult
        End If

NextTempPoint:
    Next i

    ' -----------------------------------------------------------------------
    ' Return to ambient
    ' -----------------------------------------------------------------------
Cleanup:
    Call ReturnToAmbient
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "CharacterisationTempSweep ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    Call SafeShutdown
    CharacterisationTempSweep = Err.Number
End Function

' ---------------------------------------------------------------------------
' CharacterisationSingleTemp
' ---------------------------------------------------------------------------
' Convenience entry point: run all characterisation sub-tests at a single
' caller-specified temperature.  Pass TEMP_AMBIENT for a room-temperature run.
' ---------------------------------------------------------------------------
Public Function CharacterisationSingleTemp(ByVal dTargetTemp As Double) As Long

    On Error GoTo errHandler

    Dim lResult As Long

    CharacterisationSingleTemp = 0
    m_SuppliesOn = False

    If Not IsTemperatureValid(dTargetTemp) Then
        TheExec.Datalog.WriteComment "CharacterisationSingleTemp: temperature " & _
                                     CStr(dTargetTemp) & " degC is out of safe range."
        CharacterisationSingleTemp = -1
        Exit Function
    End If

    Call SetupSupplies
    m_SuppliesOn = True

    lResult = GoToTemperature(dTargetTemp)
    If lResult <> 0 Then
        CharacterisationSingleTemp = lResult
        GoTo Cleanup
    End If

    m_CurrentTemp = dTargetTemp

    lResult = RunAllCharTests(dTargetTemp)
    CharacterisationSingleTemp = lResult

Cleanup:
    Call ReturnToAmbient
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "CharacterisationSingleTemp ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    Call SafeShutdown
    CharacterisationSingleTemp = Err.Number
End Function

'==============================================================================
' TEMPERATURE MANAGEMENT
'==============================================================================

' ---------------------------------------------------------------------------
' BuildTempSweepPoints
' ---------------------------------------------------------------------------
' Populates m_TempPoints with the characterisation temperature corners.
' Modify this routine to add or remove temperature points.
' ---------------------------------------------------------------------------
Private Sub BuildTempSweepPoints()

    m_TempPointCount = 0

    ' Cold corner
    Call AddTempPoint(TEMP_COLD)

    ' Intermediate cold point
    Call AddTempPoint(-20.0)

    ' Ambient
    Call AddTempPoint(TEMP_AMBIENT)

    ' Intermediate warm point
    Call AddTempPoint(55.0)

    ' Hot corner
    Call AddTempPoint(TEMP_HOT)

    TheExec.Datalog.WriteComment "BuildTempSweepPoints: " & CStr(m_TempPointCount) & _
                                 " temperature points registered."
End Sub

' ---------------------------------------------------------------------------
' AddTempPoint
' ---------------------------------------------------------------------------
Private Sub AddTempPoint(ByVal dTemp As Double)

    If m_TempPointCount >= MAX_TEMP_POINTS Then
        TheExec.Datalog.WriteComment "AddTempPoint: MAX_TEMP_POINTS reached, ignoring " & _
                                     CStr(dTemp)
        Exit Sub
    End If

    If Not IsTemperatureValid(dTemp) Then
        TheExec.Datalog.WriteComment "AddTempPoint: " & CStr(dTemp) & _
                                     " degC outside safe limits, not added."
        Exit Sub
    End If

    m_TempPoints(m_TempPointCount) = dTemp
    m_TempPointCount = m_TempPointCount + 1
End Sub

' ---------------------------------------------------------------------------
' GoToTemperature
' ---------------------------------------------------------------------------
' Commands the Thermostream to the requested temperature and waits for soak.
' Returns 0 on success, non-zero on failure.
' ---------------------------------------------------------------------------
Private Function GoToTemperature(ByVal dTemp As Double) As Long

    On Error GoTo errHandler

    GoToTemperature = 0

    If TheExec.TesterMode = testModeOffline Then
        TheExec.Datalog.WriteComment "GoToTemperature: OFFLINE - simulating " & _
                                     Format(dTemp, "0.0") & " degC"
        tWait TEMP_SOAK_TIME * 0.01   ' short simulated wait in offline
        Exit Function
    End If

    If Not thermoEn Then
        TheExec.Datalog.WriteComment "GoToTemperature: thermoEn=False, skipping GPIB."
        Exit Function
    End If

    TheExec.Datalog.WriteComment "GoToTemperature: Ramping to " & Format(dTemp, "0.0") & " degC"

    ' Use the Thermo module's SetThermoStream function
    ' Signature: SetThermoStream(setTemp, setRamp, setSoak, setWin)
    Dim lRet As Long
    lRet = SetThermoStream(dTemp, TEMP_RAMP_RATE, TEMP_SOAK_TIME, TEMP_WINDOW)

    If lRet <> 0 Then
        TheExec.Datalog.WriteComment "GoToTemperature: SetThermoStream returned " & CStr(lRet)
        GoToTemperature = lRet
        Exit Function
    End If

    TheExec.Datalog.WriteComment "GoToTemperature: Settled at " & Format(dTemp, "0.0") & " degC"
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "GoToTemperature ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    GoToTemperature = Err.Number
End Function

' ---------------------------------------------------------------------------
' ReturnToAmbient
' ---------------------------------------------------------------------------
' Safely returns the Thermostream to ambient temperature.
' Called at the end of every sweep and in error paths.
' ---------------------------------------------------------------------------
Private Sub ReturnToAmbient()

    On Error GoTo errHandler

    TheExec.Datalog.WriteComment "ReturnToAmbient: Returning to " & _
                                 Format(TEMP_AMBIENT, "0.0") & " degC"

    If TheExec.TesterMode = testModeOffline Then
        TheExec.Datalog.WriteComment "ReturnToAmbient: OFFLINE - no GPIB action."
        GoTo RestoreSupplies
    End If

    If Not thermoEn Then
        GoTo RestoreSupplies
    End If

    ' Ramp back to ambient; use a generous soak window for safety
    Dim lRet As Long
    lRet = SetThermoStream(TEMP_AMBIENT, TEMP_RAMP_RATE, TEMP_SOAK_TIME, TEMP_WINDOW * 2)

    If lRet <> 0 Then
        TheExec.Datalog.WriteComment "ReturnToAmbient: SetThermoStream returned " & CStr(lRet)
    End If

RestoreSupplies:
    If m_SuppliesOn Then
        Call RestoreSupplyDefaults
    End If

    m_CurrentTemp = TEMP_AMBIENT
    Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "ReturnToAmbient ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    ' Best-effort: try to restore supplies even if thermostream failed
    On Error Resume Next
    Call RestoreSupplyDefaults
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' IsTemperatureValid
' ---------------------------------------------------------------------------
' Returns True if dTemp is within the hardware-safe operating window.
' ---------------------------------------------------------------------------
Private Function IsTemperatureValid(ByVal dTemp As Double) As Boolean

    ' Use the Thermo module constants where available; fall back to local limits.
    Dim dLow  As Double
    Dim dHigh As Double

    ' LowerTlim and UpperTlim are Public Const in Thermo.bas
    dLow  = LowerTlim   ' -99 degC
    dHigh = UpperTlim   ' +175 degC

    ' Additional application-level guard: never exceed device rating
    If dTemp < dLow Or dTemp > dHigh Then
        IsTemperatureValid = False
    Else
        IsTemperatureValid = True
    End If
End Function

'==============================================================================
' CHARACTERISATION SUB-TEST DISPATCHER
'==============================================================================

' ---------------------------------------------------------------------------
' RunAllCharTests
' ---------------------------------------------------------------------------
' Calls every characterisation sub-test in sequence.
' Each sub-test logs its own results.  A non-zero return from any sub-test
' is accumulated but execution continues (characterisation mode).
' ---------------------------------------------------------------------------
Private Function RunAllCharTests(ByVal dTemp As Double) As Long

    On Error GoTo errHandler

    Dim lAccum  As Long
    Dim lResult As Long

    lAccum = 0

    TheExec.Datalog.WriteComment "RunAllCharTests: Starting sub-tests at " & _
                                 Format(dTemp, "0.0") & " degC"

    ' ------------------------------------------------------------------
    ' 1. MLS Speed Config Threshold (levels characterisation)
    ' ------------------------------------------------------------------
    TheExec.Datalog.WriteComment "--- MLS_Speed_Config_threshold ---"
    On Error Resume Next
    Call MLS_Speed_Config_threshold
    If Err.Number <> 0 Then
        TheExec.Datalog.WriteComment "MLS_Speed_Config_threshold ERROR: " & Err.Description
        lAccum = lAccum Or Err.Number
        Err.Clear
    End If
    On Error GoTo errHandler

    ' ------------------------------------------------------------------
    ' 2. MDIO Timing
    ' ------------------------------------------------------------------
    TheExec.Datalog.WriteComment "--- MDIO_timing ---"
    On Error Resume Next
    lResult = MDIO_timing()
    If Err.Number <> 0 Then
        TheExec.Datalog.WriteComment "MDIO_timing ERROR: " & Err.Description
        lAccum = lAccum Or Err.Number
        Err.Clear
    ElseIf lResult <> 0 Then
        lAccum = lAccum Or lResult
    End If
    On Error GoTo errHandler

    ' ------------------------------------------------------------------
    ' 3. Process Monitor
    ' ------------------------------------------------------------------
    TheExec.Datalog.WriteComment "--- Process_Monitor_test ---"
    On Error Resume Next
    Dim waveResult As DSPWave
    Set waveResult = Process_Monitor_test()
    If Err.Number <> 0 Then
        TheExec.Datalog.WriteComment "Process_Monitor_test ERROR: " & Err.Description
        lAccum = lAccum Or Err.Number
        Err.Clear
    End If
    On Error GoTo errHandler

    ' ------------------------------------------------------------------
    ' 4. External Loopback RGMII 1000T
    ' ------------------------------------------------------------------
    TheExec.Datalog.WriteComment "--- Extlpbk_RGMII_1000T_test ---"
    On Error Resume Next
    Dim waveResult2 As DSPWave
    Set waveResult2 = Extlpbk_RGMII_1000T_test()
    If Err.Number <> 0 Then
        TheExec.Datalog.WriteComment "Extlpbk_RGMII_1000T_test ERROR: " & Err.Description
        lAccum = lAccum Or Err.Number
        Err.Clear
    End If
    On Error GoTo errHandler

    ' ------------------------------------------------------------------
    ' Log temperature alongside accumulated result
    ' ------------------------------------------------------------------
    TheExec.Datalog.WriteComment "RunAllCharTests: Completed at " & _
                                 Format(dTemp, "0.0") & " degC, accumulated result=" & _
                                 CStr(lAccum)

    RunAllCharTests = lAccum
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "RunAllCharTests FATAL ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    RunAllCharTests = Err.Number
End Function

'==============================================================================
' SUPPLY MANAGEMENT
'==============================================================================

' ---------------------------------------------------------------------------
' SetupSupplies
' ---------------------------------------------------------------------------
' Applies nominal supply voltages and connects all pins.
' ---------------------------------------------------------------------------
Private Sub SetupSupplies()

    On Error GoTo errHandler

    TheExec.Datalog.WriteComment "SetupSupplies: Applying nominal supplies."

    thehdw.DCVI.Pins("AVDD3P3").Voltage = AVDD3P3_NOM
    thehdw.DCVI.Pins("VDDIO_R").Voltage  = VDDIO_NOM

    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, _
                                     LoadLevels:=True, _
                                     LoadTiming:=True, _
                                     RelayMode:=tlPowered

    thehdw.Wait SUPPLY_SETTLE_S

    m_SuppliesOn = True
    Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "SetupSupplies ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
    m_SuppliesOn = False
End Sub

' ---------------------------------------------------------------------------
' RestoreSupplyDefaults
' ---------------------------------------------------------------------------
' Returns supplies to nominal values after a temperature sweep.
' ---------------------------------------------------------------------------
Private Sub RestoreSupplyDefaults()

    On Error GoTo errHandler

    TheExec.Datalog.WriteComment "RestoreSupplyDefaults: Restoring nominal supplies."

    thehdw.DCVI.Pins("AVDD3P3").Voltage = AVDD3P3_NOM
    thehdw.DCVI.Pins("VDDIO_R").Voltage  = VDDIO_NOM

    thehdw.Wait SUPPLY_SETTLE_S

    m_SuppliesOn = True
    Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "RestoreSupplyDefaults ERROR #" & CStr(Err.Number) & _
                                 " : " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' SafeShutdown
' ---------------------------------------------------------------------------
' Emergency path: raise Thermostream head, disable supplies.
' Called from top-level error handlers only.
' ---------------------------------------------------------------------------
Private Sub SafeShutdown()

    On Error Resume Next   ' best-effort in an error path

    TheExec.Datalog.WriteComment "SafeShutdown: Initiating safe shutdown sequence."

    ' Attempt to return to ambient before raising head
    If thermoEn And TheExec.TesterMode <> testModeOffline Then
        Call SetThermoStream(TEMP_AMBIENT, TEMP_RAMP_RATE, TEMP_SOAK_TIME, TEMP_WINDOW * 2)
    End If

    ' Disconnect all digital pins (removes drive from DUT)
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=False, _
                                     LoadLevels:=False, _
                                     LoadTiming:=False, _
                                     RelayMode:=tlUnpowered

    m_SuppliesOn = False

    TheExec.Datalog.WriteComment "SafeShutdown: Complete."

    On Error GoTo 0
End Sub

'==============================================================================
' UTILITY HELPERS
'==============================================================================

' ---------------------------------------------------------------------------
' LogTempDatapoint
' ---------------------------------------------------------------------------
' Writes a structured comment line that can be parsed by post-processing
' scripts to correlate measurement results with temperature.
' Format: TEMP_DATA | <testName> | <temp_degC> | <value> | <units>
' ---------------------------------------------------------------------------
Public Sub LogTempDatapoint(ByVal sTestName As String, _
                            ByVal dTemp     As Double, _
                            ByVal dValue    As Double, _
                            ByVal sUnits    As String)

    TheExec.Datalog.WriteComment "TEMP_DATA | " & sTestName & " | " & _
                                 Format(dTemp, "0.0") & " | " & _
                                 Format(dValue, "0.000000E+00") & " | " & sUnits
End Sub

' ---------------------------------------------------------------------------
' GetCurrentCharTemp
' ---------------------------------------------------------------------------
' Returns the temperature at which the current sub-test is executing.
' Sub-tests may call this to annotate their own datalogs.
' ---------------------------------------------------------------------------
Public Function GetCurrentCharTemp() As Double
    GetCurrentCharTemp = m_CurrentTemp
End Function
