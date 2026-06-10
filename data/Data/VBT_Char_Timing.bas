Attribute VB_Name = "VBT_Char_Timing"
Option Explicit


Public Function MDIO_timing() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and hold time for MDIO input timing
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim start_timing As Double, end_timing As Double, resolution As Double
Dim measurement As New SiteDouble, stp_time As New SiteDouble, hld_time As New SiteDouble
Dim mdc_edge As Double, Period As Double

Dim site As Variant

Dim Pat_name As String

On Error GoTo errHandler

    Pat_name = "./Patterns/PAT_MDIO_PHYID_Read.PAT"

    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    thehdw.Digital.Pins("MDIO").Timing.EdgeTime("timing_1_0", chEdgeD1) = Period
    thehdw.Digital.Pins("MDIO").Timing.EdgeTime("timing_1_0", chEdgeR0) = 1.5 * Period
    
    
    thehdw.Wait 0.01
    
    thehdw.Patterns(Pat_name).Load
    thehdw.Patterns(Pat_name).Start
    thehdw.Digital.Patgen.HaltWait

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   MDIO Timing Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    mdc_edge = thehdw.Digital.Pins("MDC").Timing.EdgeTime("timing_1_0", chEdgeD1)
      
    start_timing = Period
    end_timing = Period * 1.15
    resolution = 0.1 * ns
    
    For Each site In TheExec.Sites.Active
        measurement = Binary_Edge_Search(Pat_name, "MDIO", start_timing, end_timing, resolution, "timing_1_0", chEdgeD1, 5)
        stp_time = Period + mdc_edge - measurement
    Next site
    
    start_timing = Period * 0.04
    end_timing = Period * 0.1
    resolution = 0.1 * ns
    
    For Each site In TheExec.Sites.Active
        measurement = Binary_Edge_Search(Pat_name, "MDIO", start_timing, end_timing, resolution, "timing_1_0", chEdgeD1, 5)
        hld_time = measurement - mdc_edge
    Next site
    
    
    TheExec.Flow.TestLimit Resultval:=stp_time, unit:=unitTime, ScaleType:=scaleNano, ForceResults:=tlForceFlow, Tname:="Setup Time"
    TheExec.Flow.TestLimit Resultval:=hld_time, unit:=unitTime, ScaleType:=scaleNano, ForceResults:=tlForceFlow, Tname:="Hold Time"
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function




Public Function reset_dly_timing() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure time for the chip to come out of reset
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim CapturedWave As New DSPWave
    Dim rdata As New SiteLong
    Dim pattern_file As String
    Dim index As Long
    Dim reset_time As New SiteDouble
    Dim site As Variant
    
    On Error GoTo errHandler

    pattern_file = "./Patterns/Rst_release_dly_meas.PAT"

    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Call thehdw.Patterns(pattern_file).Load

    thehdw.DSSC.Pins("RX_CLK").Pattern(pattern_file).Capture.Signals.Add "ResetData"

    With thehdw.DSSC.Pins("RX_CLK").Pattern(pattern_file).Capture.Signals("ResetData")
        .SampleSize = 1500
        .SampleRate = 5000000#
        .LoadSettings
    End With

    'Run pattern
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    thehdw.Wait 2 * ms

    '*******************************************************************************************************
    ' extract data from captured dsp wave
    '*******************************************************************************************************


    CapturedWave = thehdw.DSSC.Pins("RX_CLK").Pattern(pattern_file).Capture.Signals("ResetData").DSPWave
'''''    'CapturedWave(0).Plot 'for debug


    'Amount of repeats after reset. This pattern has 12502 vectors after reset at 1uS per period
    reset_time = 1 * us
    
    For Each site In TheExec.Sites
        index = CapturedWave.FindIndex(OfFirstElement, EqualTo, 1)
        'If Low to High transition not found, fail the test with a very large number
        If index = -1 Then
            reset_time = 999
            TheExec.Datalog.WriteComment "Cannot find High Transition in RX_CLK Capture!"
        ElseIf index = 0 Then
            reset_time = 999
            TheExec.Datalog.WriteComment "First Data is High in RX_CLK Capture! No Low to High transition exists"
        Else
            reset_time = reset_time + (index * (200 * ns))
        End If
    Next site


    TheExec.Flow.TestLimit Resultval:=reset_time, lowVal:=13 * ms, hiVal:=14 * ms, ScaleType:=scaleMicro, unit:=unitTime, Tname:="Out_Of_Reset_Time", ForceResults:=tlForceFlow
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function
Public Function GMII_1G_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RGMII Gig Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double
    Dim measurement As New SiteDouble
    
'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
      
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    GTX_edge = thehdw.Digital.Pins("GTX_CLK").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    

    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), 0.000000008, 0.000000014, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge + Period - measurement
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), 0.000000002, 0.000000008, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Hold_Time.Pins(Test_Pin(i)).Value = measurement + Period - GTX_edge
            Next site
            
        End If
        
    Next i
    

    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)



    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RGMII_1G_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RGMII Gig Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double
    Dim measurement As New SiteDouble
    
'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
      
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    GTX_edge = thehdw.Digital.Pins("GTX_CLK").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    

    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), 0.000000008, 0.000000012, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge - measurement
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), 0.000000003, 0.000000008, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Hold_Time.Pins(Test_Pin(i)).Value = measurement + Period - GTX_edge
            Next site
            
        End If
        
    Next i
    

    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)



    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function RGMII_100M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist, CTL_Falling_edge As Boolean) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RGMII 100M Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double
    Dim measurement As New SiteDouble
    

'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    GTX_edge = thehdw.Digital.Pins("GTX_CLK").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    
    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
    
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - (3 * ns), GTX_edge + (4 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge - measurement
                    
                If Test_Pin(i) = "TX_CTL" And CTL_Falling_edge Then
                    'Falling Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - Period - (3 * ns), GTX_edge - (15 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = (measurement + Period) - GTX_edge
                Else
                    'Rising Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - (2 * Period) - (3 * ns), GTX_edge - Period - (16 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = (measurement + (2 * Period)) - GTX_edge

                End If
            Next site
            
        End If
        
    Next i
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RGMII_10M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist, CTL_Falling_edge As Boolean) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RGMII 10M Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



   On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double
    Dim measurement As New SiteDouble
    

'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    GTX_edge = thehdw.Digital.Pins("GTX_CLK").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    

    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - (3 * ns), GTX_edge + (5 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge - measurement
                If Test_Pin(i) = "TX_CTL" And CTL_Falling_edge Then
                    'Falling Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - Period - (3 * ns), GTX_edge - Period + (4 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = (measurement + Period) - GTX_edge
                Else
                    'Rising Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - (2 * Period) - (3 * ns), GTX_edge - (2 * Period) + (4 * ns), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = (measurement + (2 * Period)) - GTX_edge
                End If
            Next site
        
        End If
        
    Next i
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MII_TXCLK_100M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist, CLK_Pin As String) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for MII 100M TX_CLK Mode
'   In this mode, TX_CLK is an output
'   Setup and Hold time is measured against an output clock
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandler


    Dim temp_voh As Double, temp_vol As Double
    
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double, XTAL_edge As Double
    Dim measurement As New SiteDouble
    Dim Start_Search As Double, End_Search As Double


'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    GTX_edge = thehdw.Digital.Pins(CLK_Pin).Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    XTAL_edge = thehdw.Digital.Pins("XTAL_I").Timing.EdgeTime("timing_1_0", chEdgeD1)

    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If

  
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
   
    For Each site In TheExec.Sites.Active
    
        GTX_edge = Binary_Edge_Search(pattern_file, CLK_Pin, Period * 0.5, Period * 1.5, 0.000000000001, "timing_1_0", chEdgeR0, 3)
        
            For i = 0 To NumOfPin - 1
            
                'Dont test N/C Pins
                If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
                    Setup_Time.AddPin (Test_Pin(i))
                    Hold_Time.AddPin (Test_Pin(i))
                    
            
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), XTAL_edge + (Period * 2.5), XTAL_edge + (Period * 3.5), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge + (2 * Period) - measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), XTAL_edge + (Period * 0.5), XTAL_edge + (Period * 1.5), 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = measurement - GTX_edge
                End If
            
            Next i
        
    Next site
    
    



    
    Call TheExec.Flow.TestLimit(GTX_edge, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceNone, Tname:="TX_CLK_Edge")
    Call TheExec.Flow.TestLimit(GTX_edge - Period - XTAL_edge, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceNone, Tname:="TXCLK_XTAL_dly")
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    
    





    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MII_TXCLK_10M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for MII 10M TX_CLK Mode
'   In this mode, TX_CLK is an output
'   Setup and Hold time is measured against an output clock
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, GTX_edge As Double, XTAL_edge As Double
    Dim measurement As New SiteDouble
    

    


'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")

    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    
    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
        
    
    
    For Each site In TheExec.Sites.Active
    
        GTX_edge = Binary_Edge_Search(pattern_file, "GTX_CLK", Period * 1, Period * 2, 0.000000000001, "timing_1_0", chEdgeR0, 3)
        
        If (GTX_edge / ns) Mod 40 > 20 Then
            XTAL_edge = ((((CInt((GTX_edge / ns)) - (GTX_edge / ns) Mod 40) / 40) + 1) * (40 * ns)) + thehdw.Digital.Pins("XTAL_I").Timing.EdgeTime("timing_1_0", chEdgeD1)
        Else
            XTAL_edge = (((CInt((GTX_edge / ns)) - (GTX_edge / ns) Mod 40) / 40) * (40 * ns)) + thehdw.Digital.Pins("XTAL_I").Timing.EdgeTime("timing_1_0", chEdgeD1)
        End If
    
        For i = 0 To NumOfPin - 1
        
            'Dont test N/C Pins
            If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
                Setup_Time.AddPin (Test_Pin(i))
                Hold_Time.AddPin (Test_Pin(i))
            
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), Period * 1.25, Period * 2.5, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Setup_Time.Pins(Test_Pin(i)).Value = GTX_edge + (2 * Period) - measurement
                measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), GTX_edge - (5 * ns), GTX_edge + Period, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                Hold_Time.Pins(Test_Pin(i)).Value = measurement - GTX_edge
                
            End If
        
        Next i
    Next site
    

    
    Call TheExec.Flow.TestLimit(GTX_edge, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceNone, Tname:="TX_CLK_Edge")
    Call TheExec.Flow.TestLimit(GTX_edge - XTAL_edge, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceNone, Tname:="TXCLK_XTAL_dly")
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RMII_100M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RMII 100M

'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, XTAL_edge As Double
    Dim measurement As New SiteDouble
    

'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    XTAL_edge = thehdw.Digital.Pins("XTAL_I").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    
    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
    
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), Period, Period * 1.7, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Setup_Time.Pins(Test_Pin(i)).Value = XTAL_edge + Period - measurement

                    'Rising Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), Period * 0.3, Period, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = measurement - XTAL_edge


            Next site
            
        End If
        
    Next i
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function RMII_10M_Timing(Meas_Pins As pinlist, pat_file As PatternSet, Hi_Pins As pinlist, Lo_Pins As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure Setup and Hold Time for RMII 10M
'   There are 10 clock cycles per data. Any one of the clock cycle can latch the data
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   On Error GoTo errHandler

    Dim temp_voh As Double, temp_vol As Double
    
    Dim ppmu_voltage As New DSPWave
    Dim ppmu_state As New DSPWave
    
    Dim numPins As Long
    Dim PinName() As String
    Dim i As Integer
    
    
    Dim pattern_file As String
    Dim site As Variant
    Dim Setup_Time As New PinListData, Hold_Time As New PinListData
    Dim Test_Pin() As String, NumOfPin As Long
    
    Dim Period As Double, XTAL_edge As Double
    Dim measurement As New SiteDouble
    

'    Meas_Pins = "TXD_0,TXD_1,TXD_2,TXD_3,TXD_4,TXD_5,TXD_6,TXD_7, TX_CTL"
    Call TheExec.DataManager.DecomposePinList(Meas_Pins, Test_Pin, NumOfPin)

    pattern_file = pat_file
    
    
    'Apply Levels and TIming
    thehdw.Digital.ApplyLevelsTiming False, True, True
    
    Period = thehdw.Digital.Timing.Period("timing_1_0")
    XTAL_edge = thehdw.Digital.Pins("XTAL_I").Timing.EdgeTime("timing_1_0", chEdgeD1)
    
    If Hi_Pins <> "" Then
        thehdw.Digital.Pins(Hi_Pins).InitState = chInitHi
        thehdw.Digital.Pins(Hi_Pins).StartState = chStartHi
    End If
    
    If Lo_Pins <> "" Then
        thehdw.Digital.Pins(Lo_Pins).InitState = chInitLo
        thehdw.Digital.Pins(Lo_Pins).StartState = chStartLo
    End If
    
    
    'Test the pattern
    Call thehdw.Patterns(pattern_file).Load
    Call thehdw.Patterns(pattern_file).Start
    thehdw.Digital.Patgen.HaltWait
    
    
    For i = 0 To NumOfPin - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(Test_Pin(i)) = "I/O" Then
            Setup_Time.AddPin (Test_Pin(i))
            Hold_Time.AddPin (Test_Pin(i))
    
            For Each site In TheExec.Sites.Active
    
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), Period * 0.2, Period * 0.7, 0.000000000001, "timing_1_0", chEdgeD1, 3)
                    Setup_Time.Pins(Test_Pin(i)).Value = XTAL_edge - measurement

                    'Rising Edge Measurement
                    measurement = Binary_Edge_Search(pattern_file, Test_Pin(i), Period * 0.3, Period * 0.8, 0.000000000001, "timing_1_0", chEdgeD2, 3)
                    Hold_Time.Pins(Test_Pin(i)).Value = measurement - XTAL_edge


            Next site
            
        End If
        
    Next i
    
    Call TheExec.Flow.TestLimit(Setup_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Hold_Time, ScaleType:=scaleNano, unit:=unitTime, ForceResults:=tlForceFlow)
    


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


