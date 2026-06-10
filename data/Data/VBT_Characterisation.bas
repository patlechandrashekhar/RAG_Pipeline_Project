Attribute VB_Name = "VBT_Characterisation"
Option Explicit
' This code facilitates testing over a range of temperatures, supplies, or both, using GPIB Thermostream control.
' Version history:
' Anon Oct 2007
' Donal Whelan Oct 2012
' Robert Drohan May 2014
' Robert Drohan Mar 2017
' Robert Drohan Aug 2017
' Robert Drohan May 2018
'
'~~~~~~~~~~~ For instructions on incorporating ThermoChar in a test program ~~~~~~~
'                          See Powerpoint file included.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Within the target program functions, cPar array values can be read and used to setup supplies etc.

Public cPar() As Double         'Holds all test condition parameters.
Public Const colTemp As Long = 1       'Column holding setpoint temperature.
Public Const colRamp As Long = 2       'Column holding ramp rate.
Public Const colSoak As Long = 3       'column holding soak time.
Public Const colWin As Long = 4        'Column holding window size.

Private Const retestEnable As Boolean = False        'Determines whether to prompt for retest on fail.
Private cParName() As String        'Names of char parameters
Private failBinCount As Long        'Count of fail bins through all test loops
Private currLoop As Long        'Current test loop.
Private totLoops As Long        'Total number of test loops.
Private currRow As Long        'Current char table row.
Private charDUTID As Long        'DUT ID number
Private wafer As Long           'Skew wafer number
Private selRetest As Boolean        'Retest part if true
Private forceEndPart As Boolean        'Force end part if true

Public Function charSetup()
''' Set up characterisation conditions for the current test run.
''' This function is used for table driven characterisation testing.
'''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Apply Voltages from "charConds" sheet.
'   The "CharConds" sheet needs to be updated before using this routine
'   This routine is Enabled by "CharEn" Enable word in Run Options
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandler
    Const titleRow As Long = 1
    Dim firstCondRow As Long
    Dim lastCondRow As Long
    Dim strow As Long
    Dim lastCondColumn As Long
    Dim wkst As Worksheet
    Dim site As Variant
    Dim headPos As String
    Dim thermoStatus As Long
    Dim i As Long
    Dim OverlayName As String
    Dim ContextName As String

    If Not SheetExists("charConds") Then
        TheExec.Datalog.WriteComment "!!!!! The charConds worksheet is missing and must be added."
        End
    End If

    Set wkst = ThisWorkbook.Worksheets("charConds")
    'Identify the number of used rows and columns.
    lastCondColumn = LastCol(wkst)
    lastCondRow = LastRow(wkst)
    strow = stopRow(wkst)
    If strow < lastCondRow Then
        lastCondRow = strow
    End If

    ReDim cParName(1 To lastCondColumn)
    ReDim cPar(1 To lastCondColumn)
    totLoops = lastCondRow - 1        'Total number of test loops to perform.
    firstCondRow = titleRow + 1        'Position of first condition row.

    currLoop = currLoop + 1
    currRow = (titleRow) + currLoop        'Select source row for test conditions.
    If currRow < firstCondRow Then currRow = firstCondRow        'Start at first condition.
    If currRow > lastCondRow Then currRow = firstCondRow        'Reset to start for the next part.

    If (currRow = firstCondRow And Not selRetest) Then
        failBinCount = 0
        wafer = 0

        ''=== Get device number from IGXL, e.g for sample testing.
        'charDUTID = TheExec.Datalog.setup.LotSetup.DeviceNumber

        ''=== Ask the user to enter device ID, e.g for characterisation parts.
        wafer = getNumFromUser("WAFER_NUMBER", 120)
        charDUTID = getNumFromUser("DEVICE_NUMBER", 200)


        TheExec.Datalog.WriteComment "****** " & Date & " " & Time
        TheExec.Datalog.WriteComment "****** First run on wafer " & wafer & " device " & charDUTID
        TheExec.Datalog.WriteComment "****** Connecting to Thermostream at address " & ThermoGPIBAddr
        Call ConnectDeviceAndPlaceOnline
        Call SetHead(HeadDown)
        headPos = ReadHeadPosition()
        If thermoEn And headPos <> HeadDown And headPos <> OfflineString Then        'Head problem
            TheExec.Datalog.WriteComment "!!!!!! Head down check fail, attempting to lower it again"
            Call SetHead(HeadDown)
        End If
    End If

    'Get conditions from the current row of the table.
    For i = 1 To lastCondColumn
        cParName(i) = wkst.Cells(titleRow, i).Value
        If IsNumeric(wkst.Cells(currRow, i).Value) Then
            cPar(i) = wkst.Cells(currRow, i).Value
        End If
    Next i
    
    
    '########################################################################################
    '######## SET DUT SUPPLY VALUES = cPar VALUES HERE, OR IN THE TESTS, AS NEEDED ########

    OverlayName = "Characterization"
    ContextName = "Char"
    
    With TheExec.Overlays
        If (.Contains(OverlayName) <> False) Then .Remove OverlayName
    Call .Add(OverlayName)
    End With
    With TheExec.Overlays(OverlayName)
        .Specs.Add("avdd3p3").Value = cPar(5)
        .Specs.Add("vddio_r").Value = cPar(6)
        .Specs.Add("dvdd0p9").Value = cPar(7)
        .Specs.Add("vddio_m").Value = cPar(8)
        .Specs.Add("vddio").Value = cPar(9)
    End With
    
    With TheExec.Contexts.Aliases
        If .Contains(ContextName) = True Then .Remove (ContextName)
        .Add (ContextName)
    End With
    
    With TheExec.Contexts.Aliases(ContextName)
        .DCCategory = "Char"
    End With
    
    
    With TheExec.Contexts(ContextName).Overlays
        If .Contains(OverlayName) = True Then .Remove (OverlayName)
        .Add OverlayName
    End With
    


    '########################################################################################
    
    



    If currRow = firstCondRow Or (currRow > firstCondRow And cPar(colTemp) <> wkst.Cells(currRow - 1, colTemp).Value) Then

        '        'Dry out the DIB before going to next temperature, if required.
        '        If wkst.Cells(currRow - 1, colTemp).Value < 15 And cPar(colTemp) < 0 Then
        '            TheExec.Datalog.WriteComment "****** Please wait: Heating to remove moisture."
        '            Call SetThermoStream(85, 9999, 20, 4)
        '        End If

        TheExec.Datalog.WriteComment "****** Going to temperature: " & cPar(colTemp) & "C on wafer " & wafer & " device " & charDUTID
        'Set Thermostream to new temperature.
        thermoStatus = SetThermoStream(cPar(colTemp), cPar(colRamp), cPar(colSoak), cPar(colWin))
    Else
        'Check Thermostream is still at the setpoint.
        thermoStatus = CheckTemperatureDrift(cPar(colTemp))
    End If

    'Return results to flow table.
    'First result is a dummy test in case a GPIB connect error message causes a fail. Ignore result but do not delete.
    Call TheExec.Flow.TestLimit(Resultval:=0, ScaleType:=scaleNoScaling, formatstr:="%1.0f", unit:=unitNone, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=thermoStatus, ScaleType:=scaleNoScaling, formatstr:="%1.0f", unit:=unitNone, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=wafer, ScaleType:=scaleNoScaling, formatstr:="%2.0f", unit:=unitNone, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=charDUTID, ScaleType:=scaleNoScaling, formatstr:="%5.0f", unit:=unitNone, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=cPar(colTemp), ScaleType:=scaleNoScaling, formatstr:="%3.1f", unit:=unitCustom, customUnit:="C", forceVal:=cPar(colTemp), ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=cPar(colRamp), ScaleType:=scaleNoScaling, formatstr:="%3.1f", unit:=unitCustom, customUnit:="#", forceVal:=cPar(colRamp), ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=cPar(colSoak), ScaleType:=scaleNoScaling, formatstr:="%3.1f", unit:=unitCustom, customUnit:="S", forceVal:=cPar(colSoak), ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=cPar(colWin), ScaleType:=scaleNoScaling, formatstr:="%3.1f", unit:=unitCustom, customUnit:="C", forceVal:=cPar(colWin), ForceResults:=tlForceFlow)

    For i = (colWin + 1) To lastCondColumn
        Call TheExec.Flow.TestLimit(Resultval:=cPar(i), ScaleType:=scaleNoScaling, formatstr:="%3.2f", Tname:="_" & cParName(i), ForceResults:=tlForceFlow)
    Next i


    Exit Function
errHandler:
    TheExec.Datalog.WriteComment ("!!!!!! Char setup error")

End Function

Public Function charCleanup()
''' Clean up after a characterisation test run.
'''
    On Error GoTo errHandler
    Dim site As Variant
    Dim userInput As Long        'User response to prompt.

    userInput = 0
    selRetest = False

    'Status notice
    TheExec.Datalog.WriteComment "****** " & Date & " " & Time
    TheExec.Datalog.WriteComment "****** Loop " & currLoop & " of " & totLoops & " on wafer " & wafer & " device " & charDUTID & " done."

    'Prompt for retest on fail, if enabled.
    If retestEnable And TheExec.Sites(site).BinNumber <> 1 And Not forceEnd Then
        userInput = MsgBox("Part Failed." & vbLf & "Yes to proceed." & vbLf & "No to retest using previous conditions.", vbYesNoCancel + vbQuestion, "Proceed?")
        'Definitions: vbYes= 6, vbNo= 7, vbCancel= 2
        If userInput = vbCancel Then End
        If userInput = vbNo Then
            currLoop = currLoop - 1        'retest
            selRetest = True
        End If
    End If

    If (currLoop < totLoops) Then
        'Thermostream operation continuing.
        TheExec.Datalog.WriteComment "****** Thermostream is under GPIB control."
    End If


    'This code is not executed when
    'ThermoChar: Handle DUT fail.
    ' Update the fail bin count.
    For Each site In TheExec.Sites.Starting
        If TheExec.Sites(site).BinNumber <> 1 Then
            incFailCount
        End If
    Next site


    If (forceEnd Or currLoop >= totLoops) Then
        'Thermostream operation ended.
        TheExec.Datalog.WriteComment "****** All testing complete on wafer " & wafer & " device " & charDUTID & ". Fail bin count = " & failBinCount
        Call EndThermoTest(cPar(colTemp))
        Call ReleaseThermoStreamGPIB

        If forceEnd Then
            'Use enable words to stop program looping.
            TheExec.EnableWord("CharEN") = False
            TheExec.EnableWord("Calok") = False
        End If

        clearForceEnd
        currLoop = 0
    End If

    If selRetest = True Then
        'Remind user of pending retest.
        TheExec.Datalog.WriteComment "****** Retest pending."
    End If

    TheExec.Datalog.WriteComment " "
    TheExec.Datalog.WriteComment " "

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment ("!!!!!! Cleanup error")

End Function

Public Function getNumFromUser(prompt As String, posit As Long) As Long
''' Ask user to input a device ID, check for valid number.
'''
    On Error GoTo errHandler

    Const MaxNum As Long = 16777215        'Only 24 bits available in STDF record.
    Dim askDevId As Boolean        'Whether to ask user for a device ID
    Dim winIDTitle As String
    Dim winIDBodyPrompt As String
    Dim response As Variant

    askDevId = True
    While askDevId = True
        winIDTitle = prompt
        winIDBodyPrompt = "Enter a number for " & prompt

        BoxTop_ADIUI winIDTitle
        response = Application.InputBox(winIDBodyPrompt, winIDTitle, "", posit, posit, "", "", Type:=2)

        If response = False Then
            'Cancel pressed or 0 entered, break out of the loop.
            getNumFromUser = -9999        'Will cause test to fail.
            askDevId = False
            TheExec.Datalog.WriteComment "!!!!!! Cancelling"
            End
        Else
            'Check response is a valid number
            If IsNumeric(response) = True And response > 0 And response <= MaxNum Then
                'Accept valid number as part ID.
                getNumFromUser = response
                askDevId = False
            Else
                TheExec.Datalog.WriteComment "****** Enter only valid numbers 1 to " & MaxNum
            End If
        End If

    Wend

    Exit Function
errHandler:
    Stop        'debug

End Function


Public Function SheetExists(wkshName As String) As Boolean
''' Test for existence of a worksheet
'''
    On Error Resume Next
    Dim wksh As Worksheet

    Set wksh = ThisWorkbook.Sheets(wkshName)
    On Error GoTo 0
    SheetExists = Not wksh Is Nothing

End Function

Public Function LastRow(ws As Object) As Long
''' Find last used row in worksheet.
''' More reliable than built in Excel function.
'''
    On Error GoTo errHandler
    Dim rLastCell As Object

    Set rLastCell = ws.Cells.Find("*", ws.Cells(1, 1), , , xlByRows, xlPrevious)
    LastRow = rLastCell.row

    Exit Function

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "LastRow()"
    Exit Function

End Function

Public Function LastCol(ws As Object) As Long
''' Find last used column in worksheet.
''' More reliable than built in Excel function.
'''
    On Error GoTo errHandler
    Dim rLastCell As Object

    Set rLastCell = ws.Cells.Find("*", ws.Cells(1, 1), , , xlByColumns, xlPrevious)
    LastCol = rLastCell.Column

    Exit Function

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "LastRow()"
    Exit Function

End Function

Public Function stopRow(ws As Object) As Long
''' Find which row to stop on.
''' Returns the number of the last row before a cell containing "STOP", if a stop exists.
'''
    On Error GoTo errHandler

    Dim stopRng As range
    stopRow = 999999
    Set stopRng = ws.range("A1:A1000").Find("STOP", lookat:=xlPart)
    If Not stopRng Is Nothing Then
        stopRow = stopRng.row - 1
    End If

    Exit Function

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "stopRow()"
    Exit Function

End Function

Public Sub incFailCount()
    TheExec.Datalog.WriteComment "****** Incrementing Fail bin counter."
    failBinCount = failBinCount + 1
End Sub

Public Sub setForceEnd()
    forceEndPart = True
End Sub

Public Sub clearForceEnd()
    forceEndPart = False
End Sub

Public Function forceEnd() As Boolean
    forceEnd = forceEndPart
End Function
