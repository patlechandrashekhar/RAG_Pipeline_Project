Attribute VB_Name = "VBT_Char_Levels"
Option Explicit



Public Function MLS_Speed_Config_threshold()

Dim i As Integer

Dim avdd3p3 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim Pattern_Name As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim PinName As String

Dim GP_OUT_trip_lvl_1 As New SiteDouble, GP_OUT_trip_lvl_2 As New SiteDouble, GP_OUT_trip_lvl_3 As New SiteDouble
Dim LED_0_trip_lvl_1 As New SiteDouble, LED_0_trip_lvl_2 As New SiteDouble, LED_0_trip_lvl_3 As New SiteDouble

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    If package_type = "pkg32" Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        '   Vih = Mode_4 (Pattern will drive 1)
        '   Vt = Mode_3 (Pattern will drive X)
        '   Vil = Mode_1 (Pattern will drive 0)
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r * 0.85
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVih) = vddio_r
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVil) = 0
        thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = avdd3p3 * 0.85
        thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3
        thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = 0
        thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
        
        Pattern_Name = ".\Patterns\adin1200_adv_speed_config.PAT"
    
    Else
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        '   Vih = Mode_3 (Pattern will drive 1)
        '   Vt = Mode_2 (Pattern will drive X)
        '   Vil = Mode_1 (Pattern will drive 0)
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r * 0.15
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVih) = vddio_r * 0.85
        thehdw.Digital.Pins("GP_OUT").Levels.Value(chVil) = 0
        thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = avdd3p3 * 0.15
        thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3 * 0.85
        thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = 0
        thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
        
        
        
        Pattern_Name = ".\Patterns\adin1300_adv_frc_speed_config_3.PAT"
    
    
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   GP_OUT MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PinName = "GP_OUT"
    
    start_voltage = 0
    end_voltage = vddio_r * 0.15
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_OUT_trip_lvl_1 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = vddio_r * 0.15
    end_voltage = vddio_r * 0.85
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_OUT_trip_lvl_2 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVt, 5)
        
    Next site
    
    start_voltage = vddio_r * 0.85
    end_voltage = vddio_r
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_OUT_trip_lvl_3 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVih, 5)
        
    Next site
    
    
  
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   LED_0 MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PinName = "LED_0"
    
    start_voltage = 0
    end_voltage = avdd3p3 * 0.15
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_0_trip_lvl_1 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = avdd3p3 * 0.15
    end_voltage = avdd3p3 * 0.85
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_0_trip_lvl_2 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVt, 5)
        
    Next site
    
    start_voltage = avdd3p3 * 0.85
    end_voltage = avdd3p3
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_0_trip_lvl_3 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVih, 5)
        
    Next site
    

    TheExec.Flow.TestLimit Resultval:=GP_OUT_trip_lvl_1, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_OUT_Trip_Level_1"
    TheExec.Flow.TestLimit Resultval:=GP_OUT_trip_lvl_2, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_OUT_Trip_Level_2"
    TheExec.Flow.TestLimit Resultval:=GP_OUT_trip_lvl_3, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_OUT_Trip_Level_3"
    TheExec.Flow.TestLimit Resultval:=LED_0_trip_lvl_1, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_LED_0_Trip_Level_1"
    TheExec.Flow.TestLimit Resultval:=LED_0_trip_lvl_2, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_LED_0_Trip_Level_2"
    TheExec.Flow.TestLimit Resultval:=LED_0_trip_lvl_3, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_LED_0_Trip_Level_3"


    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVt) = 0
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVil) = vddio_r * 0.05
    thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3 * 0.95
    thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = avdd3p3 * 0.05
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeHiZ


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function MLS_auto_mdix_threshold()

Dim i As Integer


Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim Pattern_Name As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim PinName As String

Dim GP_CLK_trip_lvl_1 As New SiteDouble, GP_CLK_trip_lvl_2 As New SiteDouble, GP_CLK_trip_lvl_3 As New SiteDouble


On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage

    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   Vih = Mode_2 (Pattern will drive 1)
    '   Vil = Mode_1 (Pattern will drive 0)
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVih) = vddio_r * 0.85
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVil) = vddio_r * 0.15
    
    
    
    Pattern_Name = ".\Patterns\Multi_level_sense_vector_auto_mdix_config_1_2.PAT"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   GP_OUT MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PinName = "GP_CLK"
    
    start_voltage = 0
    end_voltage = vddio_r * 0.15
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_CLK_trip_lvl_1 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = vddio_r * 0.15
    end_voltage = vddio_r * 0.85
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_CLK_trip_lvl_2 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = vddio_r * 0.85
    end_voltage = vddio_r
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        GP_CLK_trip_lvl_3 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVih, 5)
        
    Next site
    

    

    TheExec.Flow.TestLimit Resultval:=GP_CLK_trip_lvl_1, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_1"
    TheExec.Flow.TestLimit Resultval:=GP_CLK_trip_lvl_2, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_2"
    TheExec.Flow.TestLimit Resultval:=GP_CLK_trip_lvl_3, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_3"



    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVil) = vddio_r * 0.05


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function MLS_Energy_Det_threshold()

Dim i As Integer


Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim Pattern_Name As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim PinName As String

Dim LED_1_trip_lvl_1 As New SiteDouble, LED_1_trip_lvl_2 As New SiteDouble, LED_1_trip_lvl_3 As New SiteDouble


On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio = thehdw.DCVI.Pins("VDDIO").Voltage

    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   Vih = Mode_2 (Pattern will drive 1)
    '   Vil = Mode_1 (Pattern will drive 0)
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    thehdw.Digital.Pins("LED_1").Levels.Value(chVih) = vddio * 0.85
    thehdw.Digital.Pins("LED_1").Levels.Value(chVil) = vddio * 0.15
    
    
    
    Pattern_Name = ".\Patterns\Multi_level_sense_vector_energy_det_config_1_2.PAT"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   LED_1 MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PinName = "LED_1"
    
    start_voltage = 0
    end_voltage = vddio * 0.15
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_1_trip_lvl_1 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = vddio * 0.15
    end_voltage = vddio * 0.85
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_1_trip_lvl_2 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVil, 5)
        
    Next site
    
    start_voltage = vddio * 0.85
    end_voltage = vddio
    resolution = 0.001
    For Each site In TheExec.Sites.Active
    
        LED_1_trip_lvl_3 = Binary_Level_Search(Pattern_Name, PinName, start_voltage, end_voltage, resolution, chVih, 5)
        
    Next site
    

    

    TheExec.Flow.TestLimit Resultval:=LED_1_trip_lvl_1, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_1"
    TheExec.Flow.TestLimit Resultval:=LED_1_trip_lvl_2, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_2"
    TheExec.Flow.TestLimit Resultval:=LED_1_trip_lvl_3, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="MLS_GP_CLK_Trip_Level_3"



    thehdw.Digital.Pins("LED_1").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("LED_1").Levels.Value(chVil) = vddio_r * 0.05


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function VIH_level(Pattern_Name As Pattern, test_pins As pinlist)

Dim i As Integer


Dim avdd3p3 As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim numPins As Long, pin_cnt As Long
Dim PinName() As String
Dim Pat_name As String

Dim measurement As New PinListData

Dim GP_CLK_trip_lvl_1 As New SiteDouble, GP_CLK_trip_lvl_2 As New SiteDouble, GP_CLK_trip_lvl_3 As New SiteDouble


On Error GoTo errHandler

    'Extract Pin Name from group
    Call TheExec.DataManager.DecomposePinList(test_pins, PinName, numPins)

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage

    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
        
    
    Pat_name = Pattern_Name     '".\Patterns\DFT\A0\pkgnon64\adin1300_bscan_jtag_non64_final_40ns_rz.PAT"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   GP_OUT MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    
    
    start_voltage = avdd3p3 * 0.05
    end_voltage = avdd3p3 * 0.6
    resolution = 0.001
    
    
    For pin_cnt = 0 To numPins - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(PinName(pin_cnt)) = "I/O" Then
        
            measurement.AddPin (PinName(pin_cnt))
            
            For Each site In TheExec.Sites.Active
                measurement.Pins(PinName(pin_cnt)) = Binary_Level_Search(Pat_name, PinName(pin_cnt), start_voltage, end_voltage, resolution, chVih, 5)
            Next site
            
            
        End If
    
    Next pin_cnt
        
    
 

    

    TheExec.Flow.TestLimit Resultval:=measurement, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="VIH_Level"


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function VIL_level(Pattern_Name As Pattern, test_pins As pinlist)

Dim i As Integer


Dim avdd3p3 As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim numPins As Long, pin_cnt As Long
Dim PinName() As String
Dim Pat_name As String

Dim measurement As New PinListData

Dim GP_CLK_trip_lvl_1 As New SiteDouble, GP_CLK_trip_lvl_2 As New SiteDouble, GP_CLK_trip_lvl_3 As New SiteDouble


On Error GoTo errHandler

    'Extract Pin Name from group
    Call TheExec.DataManager.DecomposePinList(test_pins, PinName, numPins)

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage

    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
        
    
    Pat_name = Pattern_Name     '".\Patterns\DFT\A0\pkgnon64\adin1300_bscan_jtag_non64_final_40ns_rz.PAT"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   GP_OUT MLS Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    
    
    start_voltage = 0
    end_voltage = avdd3p3 * 0.6
    resolution = 0.001
    
    
    For pin_cnt = 0 To numPins - 1
    
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(PinName(pin_cnt)) = "I/O" Then
        
            measurement.AddPin (PinName(pin_cnt))
            
            For Each site In TheExec.Sites.Active
                measurement.Pins(PinName(pin_cnt)) = Binary_Level_Search(Pat_name, PinName(pin_cnt), start_voltage, end_voltage, resolution, chVil, 5)
            Next site
            
            
        End If
    
    Next pin_cnt
        
    
 

    

    TheExec.Flow.TestLimit Resultval:=measurement, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="VIL_Level"


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function VOH_level(Pattern_Name As Pattern, test_pins As pinlist, current_load As Double, stop_label As String, Force_Hi As pinlist, Force_Lo As pinlist)

Dim i As Integer


Dim avdd3p3 As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim site As Variant
Dim start_voltage As Double, end_voltage As Double, resolution As Double
Dim numPins As Long, pin_cnt As Long
Dim PinName() As String
Dim Pat_name As String

Dim measurement As New PinListData

Dim GP_CLK_trip_lvl_1 As New SiteDouble, GP_CLK_trip_lvl_2 As New SiteDouble, GP_CLK_trip_lvl_3 As New SiteDouble


On Error GoTo errHandler

    'Extract Pin Name from group
    Call TheExec.DataManager.DecomposePinList(test_pins, PinName, numPins)

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage

    

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    If Force_Hi <> "" Then
        thehdw.Digital.Pins(Force_Hi).InitState = chInitHi
        thehdw.Digital.Pins(Force_Hi).StartState = chStartHi
    End If
    
    If Force_Lo <> "" Then
        thehdw.Digital.Pins(Force_Lo).InitState = chInitLo
        thehdw.Digital.Pins(Force_Lo).StartState = chStartLo
    End If
        
    
    Pat_name = Pattern_Name     '".\Patterns\DFT\A0\pkgnon64\adin1300_bscan_jtag_non64_final_40ns_rz.PAT"
    
    'Load Search Pattern
    Call thehdw.Patterns(Pattern_Name).Load
    Call thehdw.Patterns(Pattern_Name).Start("", stop_label, tlPatConcurrentModeNone)
    thehdw.Digital.Patgen.HaltWait
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   VOH Levels Measurement
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    thehdw.Digital.Pins(test_pins).Disconnect
    
    With thehdw.PPMU.Pins(test_pins)
        .Connect
        .Gate = tlOn
    End With
    
    thehdw.PPMU.Pins(test_pins).ForceI current_load, 0.05
    
    measurement = thehdw.PPMU.Pins(test_pins).Read
    

    TheExec.Flow.TestLimit Resultval:=measurement, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow

    With thehdw.PPMU.Pins(test_pins)
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins(test_pins).Connect


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function





