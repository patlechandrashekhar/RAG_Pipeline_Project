Attribute VB_Name = "Exec_IP_Module"
Option Explicit
 
' This module contains empty Exec Interpose functions (see online help
' for details).  These are here for convenience and are completely optional.
' It is not necessary to delete them if they are not being used, nor is it
' necessary that they exist in the program.



' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    ' OnTesterInitialized executes before TheExec is even established so nothing
    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
    ' need to be debugged carefully.
    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
        "VBT Error # " + Trim(Str(Err.Number)) + ": " + Err.Description
End Function
 
' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()

    On Error GoTo errHandler

    ' Put code here
    'Enable pattern simulation in offline, comment while going to tester
    If TheExec.TesterMode = testModeOffline Then
        TheExec.Simulator.ForceAllSimulation (tlSimDefault)
    End If
    
    'Turn on user supplies
    thehdw.DIB.powerOn = True


    'Enable datalog
    TheExec.Datalog.Setup.LotSetup.DatalogOn = True
    TheExec.Datalog.Setup.DatalogSetup.WindowOutput = True
    TheExec.Datalog.Setup.DatalogSetup.STDFOutput = False
    TheExec.Datalog.Setup.DatalogSetup.TextOutput = False

    'Datalog length
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Enable = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width = 50
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.TestName.Enable = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.TestName.Width = 50
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 100

    TheExec.Datalog.ApplySetup

    ' Allows time set and edge set sheets programming model where the same pin can be programmed
    ' multiple times taking only the value of the last entry on the sheet.
    thehdw.Digital.EnablePinRespecification = True
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramLoaded"
End Function

' OnDataToolStart executes before load test program

Function OnDataToolStart()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
     MsgBox "Error encountered in Exec Interpose Function OnDataToolStart" + vbCrLf + _
        "VBT Error # " + Trim(Str(Err.Number)) + ": " + Err.Description
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramValidated"
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramFailedValidation"
End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnTDRCalibrated"
End Function
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
    On Error GoTo errHandler

    ' Put code here
    
Dim ChanMap As String
Dim site As Variant
Dim i As Integer

    ChanMap = TheExec.CurrentChanMap

    'Enable/Disable Yoda for debug 'Adds ~500 ms test time
    #If 1 Then
   
        If MM Is Nothing Then Set MM = chipInstance
    
        If RM_GEPHY Is Nothing Then Set RM_GEPHY = MM.GEPHY.registerMaps.GEPHY
        If BF_GEPHY Is Nothing Then Set BF_GEPHY = MM.GEPHY.bitFields
        
        If RM_GESUB Is Nothing Then Set RM_GESUB = MM.GESubsys.registerMaps.GESubsys
        If BF_GESUB Is Nothing Then Set BF_GESUB = MM.GESubsys.bitFields
        
        MM.Initialize
        'Set A flag to indicate the device and software are in sync
        MM.Changed = False
    
        If MM Is Nothing Then    'detect loss of definitions.
            TheExec.Datalog.WriteComment "Public variables, Yoda memory map have been lost."
        End If
    
    #End If
    
    If ChanMap = "ChanMap_pkg64_eng" Then
        For Each site In TheExec.Sites.Selected
            If ((site <> 2) And TheExec.Sites.Selected(site)) Then
                TheExec.Sites.Starting = False
                TheExec.Datalog.WriteComment ("Site " & site & " is not Allowed. Site 2 can only be used")
                Call TheExec.AddOutput("Site " & site & " is not Allowed. Site 2 can only be used", vbRed, True)
            End If
        Next site
    End If
    
    If ChanMap = "ChanMap_pkg32_eng" Then
        For Each site In TheExec.Sites.Selected
            If ((site <> 3) And TheExec.Sites.Selected(site)) Then
                TheExec.Sites.Starting = False
                TheExec.Datalog.WriteComment ("Site " & site & " is not Allowed. Site 3 can only be used")
                Call TheExec.AddOutput("Site " & site & " is not Allowed. Site 3 can only be used", vbRed, True)
            End If
        Next site
    End If
    
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramStarted"
End Function
 
' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    On Error GoTo errHandler

    ' Put code here
    'ThermoChar: cleanup after test run.
    If TheExec.EnableWord("CharEN") = True Then
        charCleanup
    End If
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramEnded"
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnPreShutDownSite"
End Function
 
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
    On Error GoTo errHandler

    ' Put code here
    Dim site As Variant
    If TheExec.EnableWord("CharEN") = True Then
        ' Update the fail bin count.
        For Each site In TheExec.Sites.Starting
            If TheExec.Sites(site).BinNumber <> 1 Then
                incFailCount

                ' Stop_on_Fail: Enable the forceEnd line to end part testing on fail.
                ' If using this option, also set Run_Until_Fail on Run window.
                'setForceEnd  'Removed since testing stop if anything fails.  ' VS 052819

            End If
            
            'Stop testing for Contact related test failures
            ' Bin 5 failures is related to continuity, powersupply shorts & pin2pinshorts
            If TheExec.Sites(site).BinNumber = 5 Then

                ' Stop_on_Fail: Enable the forceEnd line to end part testing on fail.
                ' If using this option, also set Run_Until_Fail on Run window.
                setForceEnd  'Removed since testing stop if anything fails.  ' VS 052819

            End If
        Next site
    End If
    
    Exit Function
errHandler:
    HandleExecIPError "OnPostShutDownSite"
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalStarted"
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalCompleted"
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)

    On Error GoTo errHandler
    
'    UNCOMMENT TO THE FOLLOWING LINES TO PARSE ALARMS

'    Dim i As Long
'    Dim alarmArray() As String
'
'    ' The string is a tab delimited list of alarm error messages
'    alarmArray = Split(alarmList, vbTab)
'
'    ' This will loop through all the alarms
'    For i = LBound(alarmArray) To UBound(alarmArray)
'        ' Then you can print it
'        Debug.Print "Alarm " & i & ": " & alarmArray(i)
'
'        ' Or check for a specific error
'        If InStr(1, alarmArray(i), "DCVS:0001") Then
'            Debug.Print "Found DCVS Alarm 1!!"
'        End If
'    Next i

    Exit Function
errHandler:
    HandleExecIPError "OnAlarmOccurred"
End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnGlobalVariableReset"
End Function

' Immediately once Vaildation get started
Function OnValidationStart()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnValidationStart"
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (“X”) button of the workbook.
Function OnProgramClose()
    On Error GoTo errHandler

    ' Put code here


    Exit Function
errHandler:

    HandleExecIPError "OnProgramClose"

End Function
' This function will be automatically called anytime DSP Globally Accessible Variable
' values have been reset, such as when the VBA Stop button is pressed.  This function
' is also called one time, after first validation but before run to initialize the
' variables for the first time.  This function can thus be used to do one-time
' initialization of DSP Globally Accessible Variables. Any values assigned to variables
' here will be automatically synchronized to DSP Procedures whenever needed. This function
' should be used to initialize all DSP Global Variables, but not for other VBA global variables.
' This function is only called when DSP Globally Accessible Variables are used in the test program.
Function OnDSPGlobalVariableReset()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnDSPGlobalVariableReset"
End Function
