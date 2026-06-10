Attribute VB_Name = "VBT_Utilities"

Public Function Find_Passing_Region(Srch_pttrn As String, PinName As String, Start_Search As Double, End_Search As Double, resolution As Double, timeset As String, Optional number_of_pass As Double = 1) As SiteDouble

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Search for Passing Region in pattern for certain pins using Linear search
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim tmp_strobe As Double
Dim original_strobe As New SiteDouble, search_results As New SiteDouble
Dim pattern_results As New SiteBoolean, pattern_first_pass As New SiteBoolean
Dim TimeSetSheet As String
Dim passing_run As Integer

    Set Find_Passing_Region = New SiteDouble

    'Check the input parameters for Validity
    If (End_Search - Start_Search) <= resolution Then
            TheExec.Datalog.WriteComment ("The search parameters are invalid!")
    End If
    If number_of_pass < 1 Then
            TheExec.Datalog.WriteComment ("The number of pass parameter is invalid! Needs to have atleast 1 pass to search.")
    End If
    
    'Get Name of current Timeset
    TimeSetSheet = thehdw.Digital.Timing.TimeSetNameList

    'Load Search Pattern
    'Call TheHdw.Patterns(Srch_pttrn).Load
    Call thehdw.Patterns(Srch_pttrn).Start
    thehdw.Digital.Patgen.HaltWait
    
 
    'Store Original timings so that it can be restored at the end
    For Each local_site In TheExec.Sites.Active
        original_strobe = thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, chEdgeR0)
    Next local_site
    
    search_results = 999
    tmp_strobe = Start_Search
    pattern_first_pass = False
    thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, chEdgeR0) = tmp_strobe
    
    Do While tmp_strobe <= End_Search
    
        thehdw.Wait 1 * ms
        
        passing_run = 0
        pattern_results = True 'Assume Pattern pass
        
        'Make sure the pattern passes consistently
        Do While passing_run < number_of_pass
            Call thehdw.Digital.Patgen.Restart
            thehdw.Digital.Patgen.HaltWait
            
            pattern_results = pattern_results.LogicalAnd(thehdw.Digital.Patgen.PatternBurstPassedPerSite)
            passing_run = passing_run + 1
        Loop
        
        
        'Store the Compare strobe if the pattern passes consistently for the first time
        For Each local_site In TheExec.Sites.Active
            If pattern_results And Not (pattern_first_pass) Then search_results = tmp_strobe: pattern_first_pass = True
        Next local_site
        
        'If all sites have passing region exit search loop
        If Not (pattern_results.Any(False)) Then
            Exit Do
        Else
            'Change the strobe edge for the next search pass
            tmp_strobe = tmp_strobe + resolution
            thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, chEdgeR0) = tmp_strobe
        End If
        
    Loop
    
    'Restore Original Timing of patgen
    For Each local_site In TheExec.Sites.Active
        thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, chEdgeR0) = original_strobe
    Next local_site
    
    For Each local_site In TheExec.Sites.Active
        If search_results = 999 Then TheExec.Datalog.WriteComment ("No Passing region found for Site: " & local_site & " !")
    Next local_site
    
    Find_Passing_Region = search_results
    
    Exit Function

End Function


Public Function Binary_Edge_Search(Srch_pttrn As String, PinName As String, Start_Search As Double, End_Search As Double, resolution As Double, timeset As String, Edge As chEdge, Optional number_of_pass As Double = 1) As Double

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Search for Pass/Fail Region in pattern using Binary Search
'   Used for Edge search (D0, D1, D2, D3, R0 & R1)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim tmp_strobe As Double
Dim original_strobe As Double, search_results As Double
Dim start_state As Boolean, end_state As Boolean, temp_state As Boolean
Dim start_strobe As Double, end_strobe As Double
Dim TimeSetSheet As String
Dim passing_run As Integer


    'Check the input parameters for Validity
    If (End_Search - Start_Search) <= resolution Then
            TheExec.Datalog.WriteComment ("The search parameters are invalid!")
    End If
    If number_of_pass < 1 Then
            TheExec.Datalog.WriteComment ("The number of pass parameter is invalid! Needs to have atleast 1 pass to search.")
    End If
    
    'Get Name of current Timeset
    TimeSetSheet = thehdw.Digital.Timing.TimeSetNameList
    
    start_strobe = Start_Search
    end_strobe = End_Search
    

    'Load Search Pattern
    Call thehdw.Patterns(Srch_pttrn).Load
    Call thehdw.Patterns(Srch_pttrn).Start
    thehdw.Digital.Patgen.HaltWait
    
    original_strobe = thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, Edge)

    
    'Change the edge strobe to start and record results
    thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, Edge) = start_strobe
    
    thehdw.Wait 1 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    start_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    'Change the edge strobe to end and record results
    thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, Edge) = end_strobe
    
    thehdw.Wait 1 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    end_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    If start_state = end_state Then
        TheExec.Datalog.WriteComment ("The pattern results for the start and end strobes are the same! The pattern has to have a passing and failing region.")
        Binary_Edge_Search = 999
    Else
        Do While (end_strobe - start_strobe) >= resolution
            tmp_strobe = (start_strobe + end_strobe) / 2
            
            'Change the edge strobe to start and record results
            thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, Edge) = tmp_strobe
            
            thehdw.Wait 1 * ms
            
            passing_run = 0
            temp_state = True
            Do While (passing_run < number_of_pass) And temp_state
                Call thehdw.Digital.Patgen.Restart
                thehdw.Digital.Patgen.HaltWait
                temp_state = thehdw.Digital.Patgen.PatternBurstPassed
                passing_run = passing_run + 1
            Loop
            
            
            If (temp_state = start_state) Then start_strobe = tmp_strobe Else end_strobe = tmp_strobe
            
        Loop
        
        Binary_Edge_Search = (start_strobe + end_strobe) / 2
        
    End If
    

    
    
 
    'Restore Original Timing of patgen
    thehdw.Digital.Pins(PinName).Timing.EdgeTime(timeset, Edge) = original_strobe

End Function

Public Function Binary_Level_Search(Srch_pttrn As String, PinName As String, Start_Search As Double, End_Search As Double, resolution As Double, level As ChPinLevel, Optional number_of_pass As Double = 1) As Double

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Routine to Search for Pass/Fail Region in pattern using Binary Search
'   Used for Level search (VIL, VIH, VOL, VOH & VT)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim tmp_level As Double
Dim original_level As Double, search_results As Double
Dim start_state As Boolean, end_state As Boolean, temp_state As Boolean
Dim start_level As Double, end_level As Double
Dim TimeSetSheet As String
Dim passing_run As Integer


    'Check the input parameters for Validity
    If (End_Search - Start_Search) <= resolution Then
            TheExec.Datalog.WriteComment ("The search parameters are invalid!")
    End If
    If number_of_pass < 1 Then
            TheExec.Datalog.WriteComment ("The number of pass parameter is invalid! Needs to have atleast 1 pass to search.")
    End If
    
    
    start_level = Start_Search
    end_level = End_Search
    

    'Load Search Pattern
    Call thehdw.Patterns(Srch_pttrn).Load
    Call thehdw.Patterns(Srch_pttrn).Start
    thehdw.Digital.Patgen.HaltWait
    
    original_level = thehdw.Digital.Pins(PinName).Levels.Value(level)
    

    
    'Change the edge strobe to start and record results
    thehdw.Digital.Pins(PinName).Levels.Value(level) = start_level
    
    thehdw.Wait 1 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    start_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    'Change the edge strobe to end and record results
    thehdw.Digital.Pins(PinName).Levels.Value(level) = end_level
    
    thehdw.Wait 1 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    end_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    If start_state = end_state Then
        TheExec.Datalog.WriteComment ("The pattern results for the start and end Levelss are the same! The pattern has to have a passing and failing region.")
        Binary_Level_Search = 999
    Else
        Do While (end_level - start_level) >= resolution
            tmp_level = (start_level + end_level) / 2
            
            'Change the edge strobe to start and record results
            thehdw.Digital.Pins(PinName).Levels.Value(level) = tmp_level
            
            thehdw.Wait 1 * ms
            
            passing_run = 0
            temp_state = True
            Do While (passing_run < number_of_pass) And temp_state
                Call thehdw.Digital.Patgen.Restart
                thehdw.Digital.Patgen.HaltWait
                temp_state = thehdw.Digital.Patgen.PatternBurstPassed
                passing_run = passing_run + 1
            Loop
            
            
            If (temp_state = start_state) Then start_level = tmp_level Else end_level = tmp_level
            
        Loop
        
        Binary_Level_Search = (start_level + end_level) / 2
        
    End If
    

    
    
 
    'Restore Original Timing of patgen
    thehdw.Digital.Pins(PinName).Levels.Value(level) = original_level

End Function


Public Function VLV_Level_Search(Srch_pttrn As String, PinName As String, Start_Search As Double, End_Search As Double, resolution As Double, Optional number_of_pass As Double = 1) As Double

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Search for Pass/Fail Region in pattern using Binary Search
'   Used for Level search in DCVI supplies. For DCVS, command needs to be changed
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim tmp_level As Double
Dim original_level As Double, search_results As Double
Dim start_state As Boolean, end_state As Boolean, temp_state As Boolean
Dim start_level As Double, end_level As Double

Dim passing_run As Integer


    'Check the input parameters for Validity
    If (End_Search - Start_Search) >= resolution Then
            TheExec.Datalog.WriteComment ("The search parameters are invalid!")
    End If
    If number_of_pass < 1 Then
            TheExec.Datalog.WriteComment ("The number of pass parameter is invalid! Needs to have atleast 1 pass to search.")
    End If
    
    search_results = 0
    start_level = Start_Search
    end_level = End_Search
    

    'Load Search Pattern
    Call thehdw.Patterns(Srch_pttrn).Load
    Call thehdw.Patterns(Srch_pttrn).Start
    thehdw.Digital.Patgen.HaltWait
    
    original_level = thehdw.DCVI.Pins(PinName).Voltage
    

    
    'Change the edge strobe to start and record results
    thehdw.DCVI.Pins(PinName).Voltage = start_level
    
    thehdw.Wait 10 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    start_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    'Change the edge strobe to end and record results
    thehdw.DCVI.Pins(PinName).Voltage = end_level
    
    thehdw.Wait 10 * ms
    
    Call thehdw.Digital.Patgen.Restart
    thehdw.Digital.Patgen.HaltWait
    end_state = thehdw.Digital.Patgen.PatternBurstPassed
    
    tmp_level = start_level + resolution
    
    If start_state = end_state Then
        TheExec.Datalog.WriteComment ("The pattern results for the start and end Levels are the same! The pattern has to have a passing and failing region.")
        VLV_Level_Search = 999
    Else
        Do While (end_level - tmp_level) <= resolution
            
            
            
            'Change the edge strobe to start and record results
            thehdw.DCVI.Pins(PinName).Voltage = tmp_level
            
            thehdw.Wait 10 * ms
            
            passing_run = 0
            temp_state = True
            Do While (passing_run < number_of_pass) And temp_state
                Call thehdw.Digital.Patgen.Restart
                thehdw.Digital.Patgen.HaltWait
                temp_state = thehdw.Digital.Patgen.PatternBurstPassed
                passing_run = passing_run + 1
            Loop
            
            
            If (temp_state = start_state) Then search_results = tmp_level: tmp_level = tmp_level + resolution Else tmp_level = end_level
            
        Loop
        
        VLV_Level_Search = search_results
        
    End If
    

    
    
 
    'Restore Original Timing of patgen
    thehdw.DCVI.Pins(PinName).Voltage = original_level

End Function

Public Function Measure_Frequency(PinName As String, Interval As Double, Optional EventSrc As FreqCtrEventSrcSel = BOTH, Optional EventSlope As FreqCtrEventSlopeSel = Positive) As PinListData

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Frequency of Pin using HSD Counter
'   Counter must be enabled outside of this function
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Interval_Rd As Double
Dim Rd_Count As New PinListData

    Set Measure_Frequency = New PinListData

    
'    TheHdw.Digital.Pins(PinName).FreqCtr.Enable = IntervalEnable


    ' Reset the frequency counter.
    Call thehdw.Digital.Pins(PinName).FreqCtr.Clear

    ' Set the interval.
    thehdw.Digital.Pins(PinName).FreqCtr.Interval = Interval
    
    ' Select the counting event.
    thehdw.Digital.Pins(PinName).FreqCtr.EventSource = EventSrc
    thehdw.Digital.Pins(PinName).FreqCtr.EventSlope = EventSlope

    ' Read back the time interval from hardware
    ' to account for resolution rounding error.
    Interval_Rd = thehdw.Digital.Pins(PinName).FreqCtr.Interval
    
    'Wait for the setup
    thehdw.Wait 1 * ms

    ' Start the period counter.
    Call thehdw.Digital.Pins(PinName).FreqCtr.Start

    'Read the count value.
    Measure_Frequency = thehdw.Digital.Pins(PinName).FreqCtr.Read().Math.Divide(Interval_Rd)
    
'    TheHdw.Digital.Pins(PinName).FreqCtr.Enable = Disable


End Function
