Attribute VB_Name = "Thermo"
Option Explicit

'This module contains the low level Thermostream control functions
'Tested to work with models TP04000, TP04310 and ATS710.

Public Const thermoEn As Boolean = True    'Defines whether to enable GPIB commands to the Thermostream.
Public Const ThermoGPIBAddr As Long = 30    'GPIB address of Thermostream
Public Const HeadDown As String = "Down"
Public Const HeadUp As String = "Up"
Public Const OfflineString As String = "TestingOffline"
Public Const Autohead As Boolean = True

Private Const WaitGPIB As Double = 0.4    'Wait time between commands. GPIB is _very_ slow.
Private Const FlexBoardNum As Long = 0    'Flex GPIB board number
Private Const Invalid As String = "32766"    'Chosen to be wrong result to several queries and too large to be valid.
Private Const winTitle As String = "Thermostream error"

'DUT sensor type DSNS n: 0 - no DUT sensor, 1 - type T thermocouple, 2 - type K thermocouple, 3 - RTD, 4 - diode
Private Const SensorType As Long = 0
Private Const Air2DUTMax As Double = 300
Private Const AirFlow As Double = 14
Private Const LowerTlim As Double = -99    'Lowest air temperature setting for TP04310 = -99C.
Private Const UpperTlim As Double = 175    'Do not exceed max temperature rating for Johnstech plastic lids of 175C.
Private Const DutType As Long = 1
Private Const DutTherConst As Long = 100
Private Const EmailNotify As Boolean = False


Public Sub InitThermoStream()
'''Initialise the Thermostream and set up general parameters.
'''
    On Error GoTo errHandler
    Dim cmdString As String
    Dim readBuf As String

    If Autohead Then
        'A reset here normally causes the head to lift.
        TheExec.Datalog.WriteComment "****** Resetting Thermostream, head may lift. Wait A few seconds."
        cmdString = "*RST"    'reset Thermostream
        Call SendCommand(cmdString)
        tWait 4    'manual recommends 4s wait after reset
    End If

    cmdString = "*IDN?"    'Identify model
    Call SendCommand(cmdString)
    readBuf = Invalid
    Call ReadData(readBuf)
    TheExec.Datalog.WriteComment "****** Connected to " & readBuf
    cmdString = "*CLS"    'clear status registers
    Call SendCommand(cmdString)

    'DUT Type (DTYP) Thermal Mass Examples
    '0 Smallest DUT mass 28 pin, 350 mil, ceramic or plastic device
    '1 Larger DUT mass 32 pin, 400 mil ceramic or plastic device
    '2 Even Larger DUT mass 68 pin PLCC plastic device
    '3 The Largest DUT mass Larger hybrid chips, Thermal box
    '4 System Derived (Autotune) Thermal fixture, thermal box
    'Device specific thermal constant (DUTC)
    cmdString = fmt("DTYP %x;DUTC %x", Array(DutType, DutTherConst))    'Set as required for DUT package.
    Call SendCommand(cmdString)

    'Set up Thermostream general conditions.
    'ULIM= lower air temperature limit. Don't freeze the board.
    'ULIM= upper air temperature limit. Don't burn the board.
    'ADMD= air-to-DUT maximum temperature difference.
    'FLWM= airflow rate, between 5 and 18 scfm.
    'DSNS= DUT sensor type, 0= none  1= type T thermocouple  2= type k thermocouple  3= RTD  4= diode
    cmdString = fmt("LLIM %x;ULIM %x;ADMD %x;FLWM %x;DSNS %x", Array(LowerTlim, UpperTlim, Air2DUTMax, AirFlow, SensorType))
    Call SendCommand(cmdString)


    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function SetThermoStream(setTemp As Double, setRamp As Double, setSoak As Double, setWin As Double) As Long
'''Set the Thermostream to the specified temperature, with the specified soak time.
'''Wait for it to get to temperature.
'''
    On Error GoTo errHandler

    Dim timeLimit As Double
    Dim cmdString As String
    Dim wtime As Double
    Dim setPtNum As Long
    Dim readBuf As String
    Dim headPos As String
    Dim winBodyPrompt As String

    readBuf = Invalid
    SetThermoStream = 0
    wtime = 0

    'For SETN, must use 0=hot,1=ambient, 2=cold as defined in manual.
    If setTemp <= 0 Then    'Cold
        setPtNum = 2
    ElseIf setTemp > 30 Then    'Hot
        setPtNum = 0
    Else    'Room temp
        setPtNum = 1
    End If

    'timeLimit determines when decision is made to abort test. Must be longer than soak time.
    'This is A safeguard when Thermostream cannot get to temperature - due to poor sealing, Thermocouple problem etc.
    timeLimit = 1800 + cPar(colSoak)

    If thermoEn And TheExec.TesterMode = testModeOnline Then

        'Check that the Thermostream head is down on the board.
        headPos = ReadHeadPosition()
        If thermoEn And headPos <> HeadDown And headPos <> OfflineString Then
            TheExec.Datalog.WriteComment "!!!!!! Head down check before setting temperature has failed."
            SetThermoStream = 1001
            winBodyPrompt = "Thermostream head reads UP. Please call A technician."
            BoxTop_ADIUI winTitle
            Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)
            Exit Function
        End If

        'Configure A setpoint on the Thermostream
        cmdString = fmt("SETN %x;SETP %x;RAMP %x;SOAK %x;WNDW %x", Array(setPtNum, setTemp, setRamp, setSoak, setWin))
        Call SendCommand(cmdString)
        'Check if Thermostream accepted the setpoint.
        cmdString = "SETP?"
        Call SendCommand(cmdString)
        'Read back the buffer contents

        Call ReadData(readBuf)
        If Abs(CDbl(readBuf) - setTemp) > 0.1 Then
            TheExec.Datalog.WriteComment ("!!!!!! =====================================================================")
            TheExec.Datalog.WriteComment ("!!!!!!   The Thermostream has failed to configure setpoint " & setPtNum & " for: " & setTemp & "C.")
            TheExec.Datalog.WriteComment ("!!!!!!   Please call A technician. ")
            TheExec.Datalog.WriteComment ("!!!!!! =====================================================================")
            SetThermoStream = 2002
            Call AbortThermoStream(setTemp)
            Exit Function
        End If

        Call SendCommand("FLOW 1")    'Turn on airflow, if not already on.

        wtime = waitOut(TheExec.Timer, timeLimit)
        If wtime < 0 Then
            TheExec.Datalog.WriteComment ("!!!!!! ===========================================================")
            TheExec.Datalog.WriteComment ("!!!!!!   The Thermostream failed to reach " & setTemp & "C")
            TheExec.Datalog.WriteComment ("!!!!!!   after " & Format(timeLimit, "0.0") & " seconds.")
            TheExec.Datalog.WriteComment ("!!!!!!   Ensure good air seal of head to DIB. Place tape over thermocouple.")
            TheExec.Datalog.WriteComment ("!!!!!!   Please call A technician if these measures do not work.                               ")
            TheExec.Datalog.WriteComment ("!!!!!! ===========================================================")
            SetThermoStream = 3003
            Call AbortThermoStream(setTemp)
            Exit Function
        End If

    End If

    TheExec.Datalog.WriteComment "****** Soak at " & setTemp & "C complete after " & Format(wtime, "0.0") & " seconds."

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub SetThermoStreamNoWait(setTemp As Double, setSoak As Double, setWin As Double)
'''Set the Thermostream to the specified temperature, soak time and window.
'''Do not wait for it to get to temperature, but return control immediately.
'''

    On Error GoTo errHandler
    Const setRamp As Double = 9999
    Dim cmdString As String
    Dim setPtNum As Long
    Dim readBuf As String
    Dim headPos As String
    Dim winBodyPrompt As String

    readBuf = Invalid

    'For SETN, must use 0=hot,1=ambient, 2=cold as defined in manual.

    If setTemp <= 0 Then    'Cold
        setPtNum = 2
    ElseIf setTemp > 30 Then    'Hot
        setPtNum = 0
    Else    'Room temperature
        setPtNum = 1
    End If

    If thermoEn And TheExec.TesterMode = testModeOnline Then

        'Check that the Thermostream head is down on the board.
        headPos = ReadHeadPosition()
        If thermoEn And headPos <> HeadDown And headPos <> OfflineString Then
            TheExec.Datalog.WriteComment "!!!!!! Head down check before setting temperature has failed."
            winBodyPrompt = "Thermostream head reads UP. Please call A technician."
            BoxTop_ADIUI winTitle
            Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)
            Exit Sub
        End If

        'Configure A setpoint on the Thermostream
        cmdString = fmt("SETN %x;SETP %x;RAMP %x;SOAK %x;WNDW %x", Array(setPtNum, setTemp, setRamp, setSoak, setWin))
        Call SendCommand(cmdString)
        'Check if Thermostream accepted the setpoint.
        cmdString = "SETP?"
        Call SendCommand(cmdString)
        'Read back the buffer contents

        Call ReadData(readBuf)
        If Abs(CDbl(readBuf) - setTemp) > 0.1 Then
            TheExec.Datalog.WriteComment ("!!!!!! =====================================================================")
            TheExec.Datalog.WriteComment ("!!!!!!   The Thermostream has failed to configure setpoint " & setPtNum & " for: " & setTemp & "C.")
            TheExec.Datalog.WriteComment ("!!!!!!   Please call A technician. ")
            TheExec.Datalog.WriteComment ("!!!!!! =====================================================================")
            Call AbortThermoStream(setTemp)
            Exit Sub
        End If

        Call SendCommand("FLOW 1")    'Turn on airflow, if not already on.

    End If

    'TheExec.Datalog.WriteComment "****** Set temperature (no_wait) to " & setTemp & "C"

    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Function ReadThermoStreamTemperature() As Double
'''Return the Current temperature from the Thermostream.
'''
    On Error GoTo errHandler

    Dim cmdString As String
    Dim readBuf As String
    readBuf = Invalid

    'Request temperature

    cmdString = "TEMP?"    'Check reading from thermocouple.
    Call SendCommand(cmdString)
    'read back the buffer contents
    Call ReadData(readBuf)

    ReadThermoStreamTemperature = CDbl(readBuf)

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub SetHead(HeadPosition As String)
'''Set the head position (up | down) of the Thermostream.
'''
    On Error GoTo errHandler

    Dim cmdString As String

    If Autohead Then
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "****** Setting Thermostream head " & HeadPosition

        'Build the command string
        Select Case HeadPosition
        Case "Up"
            Call SendCommand("FLOW 0")    'Turn off airflow before raising head.
            tWait 1.5
            Call SendCommand("HEAD 0")
            Interaction.Beep    'Warn user: head lifted.

        Case "Down"
            Call SendCommand("HEAD 1;FLOW 0")    'Lower head, airflow off.
            tWait 1.5
        Case Else
            'error handler
        End Select
    End If

    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Function ReadHeadPosition() As String
'''Read the head position (up | down) of the Thermostream.
'''
    On Error GoTo errHandler

    Dim cmdString As String
    Dim numPos As Long
    Dim readBuf As String
    readBuf = Invalid
    numPos = 4    'Default to read Up

    'Read back temp
    cmdString = "AUXC?"

    If thermoEn And TheExec.TesterMode = testModeOnline Then
        'send new CMD to ThermoStream
        Call SendCommand(cmdString)

        'Read back buffer
        Call ReadData(readBuf)
        numPos = CLng(readBuf)
        If readBuf = Invalid Then
            ReadHeadPosition = "UNKNOWN"
        Else
            If (numPos And 4) = 0 Then
                ReadHeadPosition = HeadDown
            Else
                ReadHeadPosition = HeadUp
            End If
        End If
    Else
        ReadHeadPosition = OfflineString
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IsInManualMode() As Boolean
'''Returns Thermostream manual | program control ststus.
'''
    On Error GoTo errHandler

    'This function will return True when system is in manual Mode
    Dim cmdString As String
    Dim auxcRegister As Long
    Dim readBuf As String

    readBuf = Invalid
    'Read back temp
    cmdString = "AUXC?"
    'send new CMD to ThermoStream
    Call SendCommand(cmdString)

    'Read back buffer
    Call ReadData(readBuf)
    auxcRegister = CLng(readBuf)    'Convert ASCII buffer to number
    IsInManualMode = auxcRegister And 256    '0=Engineering|Program|Cycle Mode,  1=Operator|Manual Mode

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function CheckTemperatureDrift(expecTemp As Double) As Long
'''Check Thermostream temperature is still at the setpoint.
'''
    On Error GoTo errHandler

    Const window As Double = 2.7
    Dim cmdString As String
    Dim tempActual As Double
    Dim headPos As String
    Dim readBuf As String
    Dim winBodyPrompt As String

    CheckTemperatureDrift = 0
    readBuf = Invalid

    If thermoEn And TheExec.TesterMode = testModeOnline Then

        headPos = ReadHeadPosition()
        If thermoEn And headPos <> HeadDown And headPos <> OfflineString Then    'Head problem
            TheExec.Datalog.WriteComment "!!!!!! Head down check has failed."
            CheckTemperatureDrift = 4004
            winBodyPrompt = "Thermostream head position fail. Please call A technician."
            BoxTop_ADIUI winTitle
            Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)
            Exit Function
        End If

        cmdString = "TEMP?"    'Check if Thermostream temperature now matches what was requested.
        Call SendCommand(cmdString)
        'read back the buffer contents
        Call ReadData(readBuf)
        tempActual = CDbl(readBuf)
        If (Abs(tempActual - expecTemp) > window) Then
            TheExec.Datalog.WriteComment ("!!!!!! ======================================================")
            TheExec.Datalog.WriteComment ("!!!!!!    The Thermostream failed to maintain temperature at " & expecTemp & "C")
            TheExec.Datalog.WriteComment ("!!!!!!    It now reads: " & CDbl(readBuf) & "C")
            TheExec.Datalog.WriteComment ("!!!!!!    Please call A technician.          ")
            TheExec.Datalog.WriteComment ("!!!!!! ======================================================")
            CheckTemperatureDrift = 5005
            Call AbortThermoStream(expecTemp)
            Exit Function
        End If
        'If this point reached temperature is correct.
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub AbortThermoStream(curTemp As Double)
'''Abort test, bring DIB to safe temperature, lift Thermostream head.
'''
    On Error GoTo errHandler
    Dim winBodyPrompt As String

    TheExec.Datalog.WriteComment ("!!!!!! Aborting Thermostream operation.")

    If thermoEn Then
        If ReadHeadPosition <> HeadDown And ReadHeadPosition <> OfflineString Then
            TheExec.Datalog.WriteComment "!!!!!! Head down check fail at Abort Thermostream."
            winBodyPrompt = "Thermostream head position check fail. Please call A technician."
        Else
            'For SETN, must use 0=hot,1=ambient, 2=cold as defined in manual.
            If curTemp < 15 Then
                TheExec.Datalog.WriteComment "****** Heating up board to remove moisture."
                Call SendCommand("SETN 0;SETP 85;RAMP 9999;SOAK 20;WNDW 4")
                tWait 75
                TheExec.Datalog.WriteComment "****** Cooling down board."
                Call SendCommand("SETN 0;SETP 40;RAMP 9999;SOAK 20;WNDW 4")
                tWait 55
            End If

            If curTemp > 40 Then
                TheExec.Datalog.WriteComment "****** Cooling the board."
                Call SendCommand("SETN 0;SETP 40;RAMP 9999;SOAK 20;WNDW 4")
                tWait 55
            End If

            Call SetHead(HeadUp)
            winBodyPrompt = "Thermostream issue. Please call A technician."
        End If
    End If

    TheExec.Datalog.WriteComment ("****** Thermostream operations complete.")
    'Use enable words to stop program looping.
    TheExec.EnableWord("CharEN") = False
    TheExec.EnableWord("Calok") = False
    If EmailNotify Then sendEmail    'Notify that program has ended.
    BoxTop_ADIUI winTitle
    Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)

    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub EndThermoTest(curTemp As Double)
'''After testing finished, bring DIB to safe temperature, lift Thermostream head.
'''
    On Error GoTo errHandler
    Dim wtime As Double
    wtime = 0

    If thermoEn Then
        If ReadHeadPosition <> HeadDown And ReadHeadPosition <> OfflineString Then
            TheExec.Datalog.WriteComment "!!!!!! Head down check fail at End Thermostream testing."
        Else
            'For SETN, must use 0=hot,1=ambient, 2=cold as defined in manual.
            If curTemp < 15 Then
                TheExec.Datalog.WriteComment "****** Heating up board to remove moisture."
                Call SendCommand("SETN 0;SETP 85;RAMP 9999;SOAK 20;WNDW 4")
                wtime = waitOut(TheExec.Timer, 75)

                TheExec.Datalog.WriteComment "****** Cooling down board."
                Call SendCommand("SETN 0;SETP 40;RAMP 9999;SOAK 20;WNDW 4")
                wtime = wtime + waitOut(TheExec.Timer, 55)
            ElseIf curTemp > 40 Then
                TheExec.Datalog.WriteComment "****** Cooling the board."
                Call SendCommand("SETN 0;SETP 40;RAMP 9999;SOAK 20;WNDW 4")
                wtime = waitOut(TheExec.Timer, 55)
            Else
                'Already at safe temperature, do nothing.
            End If

        End If

        Call SetHead(HeadUp)
    End If
    TheExec.Datalog.WriteComment "****** Thermostream operations complete. "    'Cleanup took " & wtime & " secs."
    If EmailNotify Then sendEmail    'Notify that program has ended.

    Exit Sub
errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GetController(ByVal FlexBoardNum As Long) As DriverGPIBController
'''Access GPIB Service properties and methods through TheHdw.GPIB.
'''The count property indicates how many different controllers have been
'''created. There will be one controller per GPIB board or per board entry
'''in the GPIB section of the tester configuration file.
'''

    On Error GoTo errHandler

    Set GetController = Nothing    'Assume that no controllers exist.
    If (thehdw.GPIB.Count > 0) Then    'If no controllers exist, exit function.
        Set GetController = thehdw.GPIB(FlexBoardNum)
    End If

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "!!!!!! Flex GPIB controller board not found."
    Resume Next
End Function

Public Function ConnectDeviceAndPlaceOnline() As Boolean
'''Check if the GPIB address of the Thermostream is in use.
'''Obtain A controller for the board number. If A controller exists,
'''attempt to connect to the Thermostream and place it online.
'''
    On Error GoTo errHandler

    Const DevName As String = "TStream"
    Const CommandTimeout As Long = 5
    Const RTerm As String = ""
    Const WTerm As String = ""
    Const EOIVal As Boolean = True
    Dim isAddrUsed As Long
    Dim ctrl As DriverGPIBController

    If thermoEn And TheExec.TesterMode = testModeOnline Then

        Call thehdw.GPIB(FlexBoardNum).AddressIsUsed(ThermoGPIBAddr, isAddrUsed)
        If isAddrUsed = 0 Then
            ''The AddressIsUsed check result cannot be relied upon.
            'TheExec.Datalog.WriteComment "!!!!!! Thermostream not detected at address " & ThermoGPIBaddr
        End If

        Set ctrl = GetController(FlexBoardNum)
        If (Not (ctrl Is Nothing)) Then
            'Below line gives A GPIB connect error if Thermostream is already configured.
            Call ctrl.Connect(ThermoGPIBAddr, RTerm, WTerm, EOIVal, DevName, CommandTimeout)
            tWait WaitGPIB
            Call ctrl.PlaceDeviceOnLine(ThermoGPIBAddr, 1)
            tWait WaitGPIB
            Call InitThermoStream
        Else
            TheExec.Datalog.WriteComment "!!!!!! No Flex GPIB controller board available."
        End If

    End If

    Exit Function

errHandler:
    'TheExec.Datalog.WriteComment "!!!!!! Thermostream Connect & Place Online error"
    tWait WaitGPIB
    Resume Next

End Function

Public Sub ReleaseThermoStreamGPIB()
'''If Thermostream connected, call disconnect function to release the GPIB bus.
'''
    On Error GoTo errHandler

    Dim isAddrUsed As Long

    'find out if address is used
    Call thehdw.GPIB(FlexBoardNum).AddressIsUsed(ThermoGPIBAddr, isAddrUsed)

    'Disconnect ThermoStream
    If isAddrUsed = 1 Then Call DisconnectDevice(FlexBoardNum, ThermoGPIBAddr)


    Exit Sub
errHandler:
    'TheExec.Datalog.WriteComment "!!!!!! Thermostream release error"
    tWait WaitGPIB
    Resume Next

End Sub

Public Function DisconnectDevice(FlexBoardNum As Long, ThermoGPIBAddr As Long) As Boolean
'''Disconnect the Thermostream from the GPIB bus.
'''

    On Error GoTo errHandler

    Dim ctrl As DriverGPIBController

    DisconnectDevice = False

    If thermoEn And TheExec.TesterMode = testModeOnline Then
        Set ctrl = GetController(FlexBoardNum)
        If (Not (ctrl Is Nothing)) Then
            Call ctrl.Disconnect(ThermoGPIBAddr)
            tWait WaitGPIB
            DisconnectDevice = True
        End If
    End If

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "!!!!!! Thermostream DisconnectDevice error"
    Resume Next

End Function

Public Function SendCommand(ByVal cmd As String) As Boolean
'''Send the specified command to the Thermostream via GPIB.
'''

    On Error GoTo errHandler
    'Write A command string to the Thermostream.
    Dim ctrl As DriverGPIBController
    Dim winBodyPrompt As String
    SendCommand = False

    'First obtain A controller for the board number. If A controller exists
    'write the string to the device. The controller automatically handles
    'termination according to the specifications of the
    'device in the Tester Configuration file.

GPIBwrite:
    If thermoEn And TheExec.TesterMode = testModeOnline Then
        Set ctrl = GetController(FlexBoardNum)
        If (Not (ctrl Is Nothing)) Then
            Call ctrl.Write(ThermoGPIBAddr, cmd)    'Need error handler for GPIB comms fail.
            tWait WaitGPIB
            SendCommand = True
            DoEvents    'passes control to OS while waiting.
        End If
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "!!!!!! Thermostream GPIB fail on SendCommand: " & cmd
    winBodyPrompt = "Thermostream failed to acknowledge. Check GPIB cable and address, then hit OK."
    BoxTop_ADIUI winTitle
    Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)
    'Retry when cable is connected
    Resume GPIBwrite
End Function

Public Function ReadData(ByRef ReplyBuf As String) As Boolean
'''Reads A response | status string from the Thermostream.
'''

    On Error GoTo errHandler

    Dim ctrl As DriverGPIBController
    Dim winBodyPrompt As String

    ReadData = False
    ReplyBuf = Invalid

GPIBread:
    If thermoEn And TheExec.TesterMode = testModeOnline Then
        Set ctrl = GetController(FlexBoardNum)    'Get the Flex GPIB controller.
        If (Not (ctrl Is Nothing)) Then
            Call ctrl.Read(ThermoGPIBAddr, ReplyBuf)    'Need error handler for GPIB comms fail
            tWait WaitGPIB
            ReadData = True
        End If
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "!!!!!! Thermostream GPIB fail on ReadData"
    winBodyPrompt = "Thermostream: failed to respond. Check GPIB cable and address, then hit OK."
    BoxTop_ADIUI winTitle
    Call MsgBox(winBodyPrompt, vbOKOnly + vbExclamation, winTitle)
    'Retry when cable is connected
    Resume GPIBread
End Function

Private Function fmt(Str, args) As String
'''Formats A command string for GPIB transmission
'''
    On Error GoTo errHandler
    'Works like the printf-function in C.
    'Takes A string with formatting characters and an array to expand.
    'The formatting characters are always "%x", independent of the type.
    'Usage example:
    'Dim str As String
    'str = fmt( "hello, Mr. %x, today's date is %x.", Array("Miller",Date) )

    Dim cmdStr As String    'the command string.
    Dim pos As Long    'the Current position in the args array.
    Dim i As Long

    cmdStr = ""
    pos = 0

    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "%" Then
            'Process format character "%"
            If i < Len(Str) Then
                If Mid(Str, (i + 1), 1) = "%" Then
                    cmdStr = cmdStr & "%"
                    i = i + 1
                ElseIf Mid(Str, (i + 1), 1) = "x" Then
                    'Process format character "x"
                    cmdStr = cmdStr & CStr(args(pos))
                    pos = pos + 1
                    i = i + 1
                End If
            End If
        Else
            'Process normal character
            cmdStr = cmdStr & Mid(Str, i, 1)
        End If
    Next

    fmt = cmdStr

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "!!!!!! Command string formatting error"
    Resume Next

End Function

Private Sub tWait(waitSecs As Double)
'''Skip waiting when offline, otherwise wait for specified period.
'''
    If TheExec.TesterMode = testModeOnline Then
        thehdw.Wait waitSecs
    End If

End Sub

Private Function waitOut(ByVal startTime As Double, timeLimit As Double) As Double
'''If Thermostream    'At Temperature'bit is set, or time limit has elapsed, return true, otherwise false.
'''
    On Error GoTo errHandler
    Dim atTemp As Boolean
    Dim currTime As Double
    Dim elapsedTime As Double
    Dim cmdString As String
    Dim readBuf As String
    Dim tecrReg As Long

    waitOut = 0

    If TheExec.TesterMode = testModeOffline Then
        'Offline, return immediately.
        waitOut = 0
        Exit Function
    End If

    Do Until waitOut > 0 Or waitOut < -1

        thehdw.Wait 1
        cmdString = "TECR?"    'Returns event condition register.
        Call SendCommand(cmdString)
        'read back the response
        Call ReadData(readBuf)
        tecrReg = CLng(readBuf)
        atTemp = (tecrReg And &H1)    'Check LSB for 'At Temperature'status.
        elapsedTime = TheExec.Timer(startTime)

        If (elapsedTime > timeLimit) Then
            'Flag timeout occurred.
            waitOut = -999
            Exit Do
        End If

        If atTemp Then
            'Return time taken for transition + soak.
            waitOut = elapsedTime
            Exit Do
        End If

    Loop

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "!!!!!! waitOut error"
    Resume Next

End Function

Public Sub sendEmail()
'''Send an email using an external email client.
'''Using A batch file for this as passing parameters directly to the Windows shell involves complex syntax.
'''
    On Error GoTo errHandler

    Dim pgmpath As String
    Dim strProgramName As String
    Dim strArgs As String
    Dim shellcmd As String
    Dim PID As Variant

    pgmpath = ThisWorkbook.Path
    strProgramName = pgmpath & "\sendsmtp\automail_ended.bat"
    strArgs = ""
    shellcmd = ("""" & strProgramName & """ " & strArgs)
    PID = Shell(shellcmd, vbNormalFocus)

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    Resume Next
End Sub


