Attribute VB_Name = "VBT_Fail_Dump"
Option Explicit

Public Enum Fail_Log_Fmt
    Synopsys = 1
    Default = 0
End Enum

Public Function HRAM_Fail_Dump(pinlist As String, fail_cnt As Long, Optional fmt As Fail_Log_Fmt = 0, Optional Start_Vector As Long = 0) As Long

'**************************************************************************************************
'   Description:
'   To Dump HRAM failures for executed patterns
'
'   Method:
'   Call Function to display failures in datalog
'
'   E.g Call HRAM_Fail_Dump("MyPin1, MyPin2, ...", 2000)
'   Arg pinlist -> Comma delimited Pinname or Pin Group name whose failures are logged. Need to be part of pattern or Error is thrown
'   Arg fail_cnt -> Number of lines of failures to be logged in datalog window
'   Arg Start_Vector -> Optional Vector number to start capturing failures. Default set to first vector
'
'Rev 1.0 (vsomasun, 31st Jul 2018): Initial
'**************************************************************************************************

Dim Cap_type As Long, HRAM_Depth As Long, cap_cnt As Long, numPins As Long
Dim hram_index As Integer, pin_cnt As Integer
Dim TrigType As TrigType
Dim TrigWait As Boolean
Dim PreTrigCycleCount As Long
Dim Accumulated_fail  As Long

Dim Vector As Long, Cycle As Long
Dim Expect As String, Actual As String

Dim Capture_done As Boolean

Dim PinName() As String

Dim PatName As String, PatSetName As String, Label As String

Dim PinPF As New PinListData
Dim PinData As New PinListData, PatData As New PinListData

Dim pEnable As Boolean
Dim pPatSetName As String
Dim pPatName As String
Dim pLabel As String
Dim pOffset As Long
Dim pOccurrence As Long

Dim site As Variant

    

    Cap_type = thehdw.Digital.HRAM.CaptureType
    
    'Set HRAM Capture Type to CaptureFail
    thehdw.Digital.HRAM.CaptureType = captFail
       
    'Collect Information about HRAM, pattern and PinName
    HRAM_Depth = thehdw.Digital.HRAM.MaxDepth
''    PatName = thehdw.Digital.HRAM.LastCyclePatgenInfo(pgPattern)
''    PatSetName = thehdw.Digital.HRAM.LastCyclePatgenInfo(pgPatSet)
    Call TheExec.DataManager.DecomposePinList(pinlist, PinName, numPins)
    
    'Collect trigger Information
    Call thehdw.Digital.HRAM.GetTrigger(TrigType, TrigWait, PreTrigCycleCount, True)
    
    'Get Vector Event setup of Pattern generator
    thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    thehdw.Digital.Patgen.Events.GetVector pEnable, pPatSetName, pPatName, pLabel, pOffset, pOccurrence
    
    'Sometimes GetVector doesnt get Pattern Information
    If pPatSetName = "" Then
        pPatSetName = thehdw.Digital.HRAM.LastCyclePatgenInfo(pgPatSet)
    End If
    
    If pPatName = "" Then
        pPatName = thehdw.Digital.HRAM.LastCyclePatgenInfo(pgPattern)
    End If

    
    'Set HRAM Capture to user Vector to start capturing failures
    thehdw.Digital.Patgen.Events.SetVector True, pPatSetName, pPatName, pLabel, Start_Vector, pOccurrence
    thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    
    
    
    For Each site In TheExec.Sites.Active
    
        Capture_done = False
        Accumulated_fail = 0
        cap_cnt = thehdw.Digital.HRAM.CapturedCycles
        If cap_cnt = 0 Then
            Capture_done = True
        Else
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "***********************************************************************************"
            TheExec.Datalog.WriteComment "*************************  HRAM Failure Dump **************************************"
            TheExec.Datalog.WriteComment "**   Vector Name : " & pPatSetName
            TheExec.Datalog.WriteComment "*"
            TheExec.Datalog.WriteComment "***********************************************************************************"
            TheExec.Datalog.WriteComment ""
            
        End If
        
        
        'Loop Until Required failures or Full Pattern Failures are captured
        Do While (Not (Capture_done))
        
            'Extract Failure Data from HRAM
            hram_index = 0
            Do While ((hram_index < cap_cnt) And (Accumulated_fail < fail_cnt))
                
                Vector = thehdw.Digital.HRAM.PatGenInfo(hram_index, pgVector)
                Cycle = thehdw.Digital.HRAM.PatGenInfo(hram_index, pgCycle)
                
                PinPF = thehdw.Digital.Pins(pinlist).HRAM.PinPF(startIndex:=hram_index)
                PinData = thehdw.Digital.Pins(pinlist).HRAM.PinData(startIndex:=hram_index)
                PatData = thehdw.Digital.Pins(pinlist).HRAM.PatData(startIndex:=hram_index)
                
                For pin_cnt = 0 To numPins - 1
                    If (PinPF.Pins(pin_cnt).Value(site) = tlResultFail) And (Accumulated_fail < fail_cnt) Then
                        Expect = PatData.Pins(pin_cnt).Value(site)
                        Actual = PinData.Pins(pin_cnt).Value(site)
                        If fmt = Synopsys Then
                            If Actual = "H" Then
                                TheExec.Datalog.WriteComment (Cycle & " " & PinData.Pins(pin_cnt).Name & " 1")
                            Else
                                TheExec.Datalog.WriteComment (Cycle & " " & PinData.Pins(pin_cnt).Name & " 0")
                            End If
                        Else
                            TheExec.Datalog.WriteComment ("PinName: " & Format(PinData.Pins(pin_cnt).Name, "!@@@@@@@@@@@@@@@@@@@@") & " , Vector: " & Format(Vector, "!@@@@@@@@@") & " , Cycle:  " & Format(Cycle, "!@@@@@@@@@") & " , Expect: " & Expect & " , Actual: " & Actual)
                        End If
                        'TheExec.Datalog.WriteComment ("PinName: " & Format(pinName(pin_cnt), "!@@@@@@@@@@@@@@@@@@@@") & " , Vector: " & Format(Vector, "!@@@@@@@@@") & " , Cycle:  " & Format(Cycle, "!@@@@@@@@@") & " , Expect: " & Expect & " , Actual: " & Actual)
                        Accumulated_fail = Accumulated_fail + 1
                    End If
                
                Next pin_cnt
                
                hram_index = hram_index + 1
                
            'Next hram_index
            Loop
            
            
            
            'If HRAM is full, change event to accumulate next set of failures
            If (hram_index = HRAM_Depth) Then
                Vector = thehdw.Digital.HRAM.LastCyclePatgenInfo(pgVector)
                thehdw.Digital.Patgen.Events.SetVector Enable:=True, PatSetName:=pPatSetName, PatName:=pPatName, Label:=pLabel, Offset:=Vector + 1, occurrence:=pOccurrence
                thehdw.Digital.Patgen.Restart
                thehdw.Digital.Patgen.HaltWait
                thehdw.Wait 0.001
                cap_cnt = thehdw.Digital.HRAM.CapturedCycles
            
            'If HRAM Capture is not full, then capture is done
            ElseIf ((cap_cnt < HRAM_Depth) And (hram_index = cap_cnt)) Then
                Capture_done = True
                TheExec.Datalog.WriteComment "Next Capture Count " & cap_cnt
            End If
            
            'If No HRAM Data or Required Fail Count reached, then Capture is done
            If (cap_cnt = 0) Or (Accumulated_fail >= fail_cnt) Then
                Capture_done = True
            End If
        
        Loop
    
    Next site

    
    'Restore HRAM Capture to original settings after dumping failures
    thehdw.Digital.Patgen.Events.SetVector pEnable, pPatSetName, pPatName, pLabel, pOffset, pOccurrence
    thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    
    Call thehdw.Digital.HRAM.SetTrigger(TrigType, TrigWait, PreTrigCycleCount, True)
    thehdw.Digital.HRAM.CaptureType = Cap_type


End Function
