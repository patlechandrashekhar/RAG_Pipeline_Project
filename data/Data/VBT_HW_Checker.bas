Attribute VB_Name = "VBT_HW_Checker"
Option Explicit

Public Function Socket_IO_Leakage(Leakage_Pins As pinlist, Force_V As Double, Optional Current_Range As Double = 0) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Leakage_current As New PinListData

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").Connect
    
    'Force the pullup Voltage to 0V to measure resistance
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").InitState = chInitLo
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").StartState = chStartLo

    thehdw.Digital.Pins(Leakage_Pins).Disconnect
     
    With thehdw.PPMU.Pins(Leakage_Pins)
        .Connect
        .Gate = tlOn
    End With
    
    If Current_Range <> 0 Then
        thehdw.PPMU.Pins(Leakage_Pins).ForceV Force_V, Current_Range
    Else
        thehdw.PPMU.Pins(Leakage_Pins).ForceV Force_V
    End If
    
    Leakage_current = thehdw.PPMU.Pins(Leakage_Pins).Read
    
    
    

    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Leakage_current, lowVal:=-2 * uA, hiVal:=2 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceNone
    
    With thehdw.PPMU.Pins(Leakage_Pins)
        .Gate = tlOff
        .Disconnect
    End With
    thehdw.Digital.Pins(Leakage_Pins).Connect



End Function



Public Function Socket_Power_Shorts() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vdd0p9 As Double, avdd1p2 As Double, avdd3p3 As Double
Dim vddio_m As Double, vddio As Double, vddio_r As Double

Dim ChanMap As String
Dim package_type As String
Dim cond As String


Dim UVI80_current As New PinListData

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.01
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    
    'Read Programmed Voltages on the supply
    vdd0p9 = 0.9 * 1.1
    avdd1p2 = 1.2 * 1.1
    avdd3p3 = 3.3 * 1.1
    vddio_r = 3.3 * 1.05
    vddio_m = 3.3 * 1.03
    vddio = 3.3 * 1

    
    thehdw.Digital.Pins("REXT, MDI_Pins").Connect
    thehdw.Digital.Pins("alldigpins, REXT, MDI_Pins").InitState = chInitLo
    thehdw.Digital.Pins("alldigpins, REXT, MDI_Pins").StartState = chStartLo
    
    
    With thehdw.DCVI.Pins("VDD_Pins")
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Program Voltages on the supply
    thehdw.DCVI.Pins("VDD0P9").Voltage = vdd0p9
    thehdw.DCVI.Pins("AVDD3P3").Voltage = avdd3p3
    thehdw.DCVI.Pins("VDDIO_R").Voltage = vddio_r
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("AVDD1P2").Voltage = avdd1p2
        thehdw.DCVI.Pins("VDDIO_M").Voltage = vddio_m
        thehdw.DCVI.Pins("VDDIO").Voltage = vddio
    End If
    

    
    thehdw.Wait 0.02
    
    'Program the Current and Current Range
    thehdw.DCVI.Pins("VDD_Pins").SetCurrentAndRange 20 * uA, 20 * uA
    
    cond = "COND: " & "VDD0P9 = " & Format(vdd0p9, "0.000") _
                    & ", AVDD3P3 = " & Format(avdd3p3, "0.000") _
                    & ", VDDIO_R = " & Format(vddio_r, "0.000")

                                            
    If package_type = "pkg64" Then
        cond = cond & ", AVDD1P2 = " & Format(avdd1p2, "0.000") _
                    & ", VDDIO_M = " & Format(vddio_m, "0.000") _
                    & ", VDDIO = " & Format(vddio, "0.000")
    End If

    

    'Display Voltage Conditions in Datalog
    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment (cond)
    TheExec.Datalog.WriteComment ("")
    
    'wait for supply to stabilize
    thehdw.Wait 1
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins("VDD_Pins").Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=UVI80_current, lowVal:=-5 * uA, hiVal:=5 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceNone


    thehdw.Digital.Pins("REXT, MDI_Pins").Disconnect
    thehdw.Digital.Pins("alldigpins, REXT, MDI_Pins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins, REXT, MDI_Pins").StartState = chStartOff


End Function


Public Function Resistor_Measurement() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Res_Voltage As New PinListData
Dim test_current As Double

test_current = 100 * uA

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    'Turn on the 5V RLAG3VM relay
    thehdw.Utility.Pins("UTIL_K2, UTIL_K3, UTIL_K4").State = tlUtilBitOn
    
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").Connect
    
    'Force the pullup Voltage to 0V to measure resistance
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").InitState = chInitLo
    thehdw.Digital.Pins("INT_N_PULLUP_V, MDIO_PULLUP_V").StartState = chStartLo
    

    thehdw.Digital.Pins("INT_N, MDIO, REXT").Disconnect
     
    With thehdw.PPMU.Pins("INT_N, MDIO, REXT")
        .Connect
        .Gate = tlOn
    End With
    
    thehdw.PPMU.Pins("INT_N, MDIO, REXT").ForceI test_current, 200 * uA
    
    Res_Voltage = thehdw.PPMU.Pins("INT_N, MDIO, REXT").Read

    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Res_Voltage.Pins("REXT").Divide(test_current), lowVal:=3000, hiVal:=3200, unit:=unitCustom, customUnit:="Ohm", ScaleType:=scaleKilo, ForceResults:=tlForceNone, Tname:="Res_R2"
    TheExec.Flow.TestLimit Resultval:=Res_Voltage.Pins("INT_N").Divide(test_current), lowVal:=1400, hiVal:=1600, unit:=unitCustom, customUnit:="Ohm", ScaleType:=scaleKilo, ForceResults:=tlForceNone, Tname:="Res_R4"
    TheExec.Flow.TestLimit Resultval:=Res_Voltage.Pins("MDIO").Divide(test_current), lowVal:=1400, hiVal:=1600, unit:=unitCustom, customUnit:="Ohm", ScaleType:=scaleKilo, ForceResults:=tlForceNone, Tname:="Res_R6"
    
    
    With thehdw.PPMU.Pins("INT_N, MDIO, REXT")
        .Gate = tlOff
        .Disconnect
    End With
    thehdw.Digital.Pins("INT_N, MDIO, REXT").Connect
    
    'Turn off the 5V RLAG3VM relay
    thehdw.Utility.Pins("UTIL_K2, UTIL_K3, UTIL_K4").State = tlUtilBitOff



End Function

Public Function Relay_5V_Check() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim USR_5V_On_Voltage As New PinListData
Dim USR_5V_Off_Current As New SiteDouble

Dim Rly_G6K_Current As New PinListData, Rly_G3VM_Current As New PinListData
Dim test_current As Double

Dim Relay_G6K As String, Relay_G3VM As String
Dim Test_Pin() As String
Dim numPins As Long
Dim i As Integer
Dim site As Variant


test_current = 100 * uA

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001

    'Define Relays for test
    Relay_G6K = "UTIL_K1"
    Relay_G3VM = "UTIL_K2, UTIL_K3, UTIL_K4"
    

    'The DCVI Force channel is made HiZ for measureing voltage
    thehdw.DCVI.Pins("P5V_RLYCHK").Gate(tlDCVIGateHiZ) = False
    thehdw.DCVI.Pins("P5V_RLYCHK").Connect (tlDCVIConnectHighSense)


    'Measure Voltage of 5V_1 User supplies using DCVI
    thehdw.DCVI.Pins("P5V_RLYCHK").Meter.mode = tlDCVIMeterVoltage
    thehdw.Wait 0.02
    USR_5V_On_Voltage = thehdw.DCVI.Pins("P5V_RLYCHK").Meter.Read




    'Turn off 5V DIB Power supply
    thehdw.DIB.power.Item("5V_1").State = tlOff

    thehdw.Wait 0.05

    'Connect the Force and Sense of 5V DCVI
    thehdw.DCVI.Pins("P5V_RLYCHK").Connect (tlDCVIConnectHighForce)
    thehdw.DCVI.Pins("P5V_RLYCHK").Gate(tlDCVIGate) = True
    
    'Force 5V on powerplane
    thehdw.DCVI.Pins("P5V_RLYCHK").SetVoltageAndRange 4.9, 7
    thehdw.DCVI.Pins("P5V_RLYCHK").SetCurrentAndRange 0.4, 0.4



    'Measure no load Current of 5V_1 User supplies using DCVI
    thehdw.DCVI.Pins("P5V_RLYCHK").Meter.mode = tlDCVIMeterCurrent
    thehdw.Wait 1
    USR_5V_Off_Current = thehdw.DCVI.Pins("P5V_RLYCHK").Meter.Read
    

    Rly_G6K_Current.AddPin (Relay_G6K)

    'Turn on the 5V XTAL RLAG6K relay
    thehdw.Utility.Pins(Relay_G6K).State = tlUtilBitOn

    'Measure Current of 5V_1 User supplies using DCVI
    thehdw.Wait 0.2
    Rly_G6K_Current.Pins(Relay_G6K).Value = thehdw.DCVI.Pins("P5V_RLYCHK").Meter.Read

    'Turn off the 5V XTAL RLAG6K relay
    thehdw.Utility.Pins(Relay_G6K & ", " & Relay_G3VM).State = tlUtilBitOff
    
    
    
    
    'Decompose the Pin List
    Call TheExec.DataManager.DecomposePinList(Relay_G3VM, Test_Pin, numPins)
    
    For i = 0 To numPins - 1
        Rly_G3VM_Current.AddPin (Test_Pin(i))
    Next i
    
    For Each site In TheExec.Sites.Active
        For i = 0 To numPins - 1
        
            'Turn on the 5V RLAG3VM relay
            thehdw.Utility.Pins(Test_Pin(i)).State = tlUtilBitOn
                        
            thehdw.Wait 0.02
            
            'Measure Current of 5V_1 User supplies using DCVI
            Rly_G3VM_Current.Pins(Test_Pin(i)).Value = thehdw.DCVI.Pins("P5V_RLYCHK").Meter.Read
            
            'Turn off the 5V XTAL RLAG6K relay
            thehdw.Utility.Pins(Relay_G6K & ", " & Relay_G3VM).State = tlUtilBitOff
            
        
        Next i
    Next site



    TheExec.Flow.TestLimit Resultval:=USR_5V_On_Voltage, lowVal:=4.5, hiVal:=5.2, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceNone, Tname:="USR_5V_Voltage"
    TheExec.Flow.TestLimit Resultval:=USR_5V_Off_Current, lowVal:=0.005, hiVal:=0.018, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="USR_5V_No_load_I"
    TheExec.Flow.TestLimit Resultval:=Rly_G6K_Current.Math.Subtract(USR_5V_Off_Current), lowVal:=0.02, hiVal:=0.06, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="G6K_5V_Rly_Current"
    TheExec.Flow.TestLimit Resultval:=Rly_G3VM_Current.Math.Subtract(USR_5V_Off_Current), lowVal:=0.006, hiVal:=0.013, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="G3VM_5V_Rly_Current"



    'Turn off all the relays
    thehdw.Utility.Pins("UTIL_K1, UTIL_K2, UTIL_K3, UTIL_K4").State = tlUtilBitOff

    thehdw.DCVI.Pins("P5V_RLYCHK").Gate(tlDCVIGate) = False
    thehdw.DCVI.Pins("P5V_RLYCHK").Disconnect

    'Turn off 5V DIB Power supply
    thehdw.DIB.power.Item("5V_1").State = tlOn





End Function






Public Function Relay_12V_Check() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Relay_Pins As String
Dim USR_12V As Double
Dim USR_12V_On_Voltage As New PinListData, USR_12V_Off_Voltage As New PinListData
Dim Rly_Current As New PinListData
Dim R101_Value As Double
Dim Test_Pin() As String
Dim numPins As Long
Dim i As Integer
Dim site As Variant
Dim current As New SiteDouble, USR_12V_Off_Current As New SiteDouble
Dim ChanMap As String
Dim package_type As String


     R101_Value = 10   'Resistor Value 10 ohms

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    If package_type = "pkg40" Then
        Relay_Pins = "UTIL_KA1, UTIL_KA2, UTIL_KA3, UTIL_KA4," _
                     & "UTIL_KB1, UTIL_KB2, UTIL_KB3, UTIL_KB4," _
                     & "UTIL_KC1, UTIL_KC2, UTIL_KD1, UTIL_KE1," _
                     & "UTIL_KR1, UTIL_KR2, UTIL_KR3, UTIL_KR4, UTIL_KR9, UTIL_KR11"
    ElseIf package_type = "pkg32" Then
        Relay_Pins = "UTIL_KA1, UTIL_KA2," _
                     & "UTIL_KB1, UTIL_KB2," _
                     & "UTIL_KC1, UTIL_KE1," _
                     & "UTIL_KR1, UTIL_KR2, UTIL_KR3, UTIL_KR4, UTIL_KR9, UTIL_KR11"
    Else
        Relay_Pins = "UTIL_KA1, UTIL_KA2, UTIL_KA3, UTIL_KA4," _
                     & "UTIL_KB1, UTIL_KB2, UTIL_KB3, UTIL_KB4," _
                     & "UTIL_KC1, UTIL_KC2, UTIL_KD1, UTIL_KE1," _
                     & "UTIL_KR1, UTIL_KR2, UTIL_KR3, UTIL_KR4, UTIL_KR5, UTIL_KR6," _
                     & "UTIL_KR7, UTIL_KR8, UTIL_KR9, UTIL_KR10, UTIL_KR11, UTIL_KR12"
    End If
    
    USR_12V = thehdw.DIB.power.Item("12V").Reading
    
    'The DCVI Force channel is made HiZ for measureing voltage
    thehdw.DCVI.Pins("P12V_RLYCHK").Gate(tlDCVIGateHiZ) = False
    
    thehdw.DCVI.Pins("P12V_RLYCHK").Connect (tlDCVIConnectHighSense)
    
    'Measure Voltage of 5V_1 User supplies using DCVI
    thehdw.DCVI.Pins("P12V_RLYCHK").Meter.mode = tlDCVIMeterVoltage
    thehdw.Wait 0.02
    USR_12V_On_Voltage = thehdw.DCVI.Pins("P12V_RLYCHK").Meter.Read
    
    'Measure current going through 10 ohm resistor
    USR_12V_Off_Current = USR_12V_On_Voltage.Math.Multiply(2).Subtract(USR_12V).Divide(R101_Value * -1)
    
    'Decompose the Pin List
    Call TheExec.DataManager.DecomposePinList(Relay_Pins, Test_Pin, numPins)
    
    For i = 0 To numPins - 1
        Rly_Current.AddPin (Test_Pin(i))
    Next i
    
    For Each site In TheExec.Sites.Active
        For i = 0 To numPins - 1
        
            'Turn on the 12V RLAG6K relay
            thehdw.Utility.Pins(Test_Pin(i)).State = tlUtilBitOn
            
            'Measure Current of 5V_1 User supplies using DCVI
            thehdw.Wait 0.02
            
            current = thehdw.DCVI.Pins("P12V_RLYCHK").Meter.Read
            
            thehdw.Utility.Pins(Relay_Pins).State = tlUtilBitOff
            
            Rly_Current.Pins(Test_Pin(i)).Value = (USR_12V - (2 * current)) / R101_Value
        
        Next i
    Next site

    
    TheExec.Flow.TestLimit Resultval:=USR_12V_On_Voltage.Math.Multiply(2), lowVal:=11, hiVal:=12.5, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceNone, Tname:="USR_12V_Voltage"
    TheExec.Flow.TestLimit Resultval:=USR_12V_Off_Current, lowVal:=0, hiVal:=100, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="USR_12V_No_load_I"
    TheExec.Flow.TestLimit Resultval:=Rly_Current.Math.Subtract(USR_12V_Off_Current), lowVal:=0.006, hiVal:=0.012, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="G6K_12V_Rly_Current"

    
   
    'Turn off all the relays
    thehdw.Utility.Pins(Relay_Pins).State = tlUtilBitOff

    thehdw.DCVI.Pins("P12V_RLYCHK").Disconnect (tlDCVIConnectHighSense)
    
    

End Function

Public Function HSD_Trace(PinName As pinlist) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Test_Pin() As String
Dim numPins As Long
Dim i As Integer
Dim site As Variant
Dim trace_length As New PinListData

Dim ChanMap As String
Dim package_type As String


    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    If package_type = "pkg40" Then
        PinName = "GP_CLK,GP_OUT,LED_0,MDC,MDIO," _
                    & "INT_N,RESET_N,REF_CLK,REXT,TEST_EN," _
                    & "MDI_0_N,MDI_0_P,MDI_1_N,MDI_1_P,MDI_2_N,MDI_2_P,MDI_3_N,MDI_3_P," _
                    & "RXD_0,RXD_1,RXD_2,RXD_3," _
                    & "RX_CLK,RX_CTL," _
                    & "TXD_0,TXD_1,TXD_2,TXD_3," _
                    & "TX_CTL,GTX_CLK," _
                    & "XTAL_O,XTAL_I,"
    ElseIf package_type = "pkg32" Then
        PinName = "GP_CLK,GP_OUT,LED_0,MDC,MDIO," _
                    & "INT_N,RESET_N,REXT,TEST_EN," _
                    & "MDI_0_N,MDI_0_P,MDI_1_N,MDI_1_P," _
                    & "RXD_0,RXD_1,RXD_2,RXD_3," _
                    & "RX_CLK,RX_CTL," _
                    & "TXD_0,TXD_1,TXD_2,TXD_3," _
                    & "TX_CTL,GTX_CLK," _
                    & "XTAL_O,XTAL_I,"
    Else
        PinName = "GP_CLK,GP_OUT,LED_0,LED_1,LED_2,MDC,MDIO," _
                    & "INT_N,RESET_N,REF_CLK,REXT,TEST_EN," _
                    & "MDI_0_N,MDI_0_P,MDI_1_N,MDI_1_P,MDI_2_N,MDI_2_P,MDI_3_N,MDI_3_P," _
                    & "COL,CRS,RX_ER," _
                    & "RXD_0,RXD_1,RXD_2,RXD_3," _
                    & "RXD_4,RXD_5,RXD_6,RXD_7," _
                    & "RX_CLK,RX_CTL," _
                    & "TXD_0,TXD_1,TXD_2,TXD_3," _
                    & "TXD_4,TXD_5,TXD_6,TXD_7," _
                    & "TX_CTL,GTX_CLK," _
                    & "TX_CLK,TX_ER," _
                    & "XTAL_O,XTAL_I," _
                    & "TDI,TMS,TDO,TCK"
    End If


    'Decompose the Pin List
    Call TheExec.DataManager.DecomposePinList(PinName, Test_Pin, numPins)


    For i = 0 To numPins - 1
        trace_length.AddPin (Test_Pin(i))
    Next i
    
    For Each site In TheExec.Sites.Active
        For i = 0 To numPins - 1
            trace_length.Pins(Test_Pin(i)).Value = thehdw.Digital.Pins(Test_Pin(i)).Calibration.DIB.Trace
        Next i
    Next site
    
    TheExec.Flow.TestLimit Resultval:=trace_length, lowVal:=0, hiVal:=10 * ns, unit:=unitTime, ScaleType:=scaleNano, ForceResults:=tlForceNone, Tname:="HSD_Trace_length"


End Function

