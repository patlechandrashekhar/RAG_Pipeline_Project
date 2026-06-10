Attribute VB_Name = "VBT_Powersupply"

Option Explicit


Public Function Set_PowerSupply() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply Power to VDD Pins and Levels to I/O pins
'
'   Rev 1.0 (vsomasun, Feb 8th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vdd0p9 As Double, avdd1p2 As Double, avdd3p3 As Double
Dim vddio_m As Double, vddio_r As Double, vddio As Double
Dim ChanMap As String
Dim package_type As String
Dim cond As String

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    cond = "COND: " & "AVDD3P3 = " & Format(avdd3p3, "0.000") _
                        & ", VDDIO_R = " & Format(vddio_r, "0.000")
                        
    If package_type = "pkg40" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000")
    ElseIf package_type = "pkg64" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000") _
                    & ", VDDIO_M = " & Format(vddio_m, "0.000") _
                    & ", VDDIO = " & Format(vddio, "0.000")
    End If

    'Display Voltage Conditions in Datalog
    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment (cond)
    TheExec.Datalog.WriteComment ("")
    
    Exit Function

End Function


Public Function Powersupply_Shorts() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim cond As String

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    thehdw.Wait 0.005

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg32" Then
        power_pins = "AVDD3P3, VDDIO_R"
    ElseIf package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If


    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    thehdw.Digital.Pins("alldigpins").InitState = chInitLo
    thehdw.Digital.Pins("alldigpins").StartState = chStartLo
    
    'Read Programmed Voltages on the supply
    vdd0p9 = 0.9 * 0.08
    avdd3p3 = 3.3 * 0.15
    vddio_m = 3.3 * 0.2
    vddio_r = 3.3 * 0.18
    vddio = 3.3 * 0.16
    

    thehdw.Digital.Pins("alldigpins").InitState = chInitLo
    thehdw.Digital.Pins("alldigpins").StartState = chStartLo
'    thehdw.Digital.Pins("alldigpins").Disconnect
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Program Voltages on the supply
    thehdw.DCVI.Pins("AVDD3P3").Voltage = avdd3p3
    thehdw.DCVI.Pins("VDDIO_R").Voltage = vddio_r
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").Voltage = vdd0p9
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").Voltage = vdd0p9
        thehdw.DCVI.Pins("VDDIO_M").Voltage = vddio_m
        thehdw.DCVI.Pins("VDDIO").Voltage = vddio
    End If
    
    thehdw.Wait 0.005
    
    
    'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * uA, 200 * uA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * uA, 200 * uA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 2 * mA, 2 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 2 * mA, 2 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 20 * uA, 20 * uA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * uA, 200 * uA
    End If
    

    cond = "COND: " & "AVDD3P3 = " & Format(avdd3p3, "0.000") _
                        & ", VDDIO_R = " & Format(vddio_r, "0.000")
                        
    If package_type = "pkg40" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000")
    ElseIf package_type = "pkg64" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000") _
                    & ", VDDIO_M = " & Format(vddio_m, "0.000") _
                    & ", VDDIO = " & Format(vddio, "0.000")
    End If
    

    'Display Voltage Conditions in Datalog
    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment (cond)
    TheExec.Datalog.WriteComment ("")
    
    'wait for supply to stabilize
    thehdw.Wait 0.05
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    End If
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function pin_to_pin_shorts(alldigcont As pinlist, test_pins As pinlist, test_current As Double) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply small Voltages to VDD Pins and test current leakage between supplies
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim measurement As New PinListData
Dim numPins As Long, pin_cnt As Long
Dim PinName() As String

    On Error GoTo errHandler

    'Extract Pin Name from group
    Call TheExec.DataManager.DecomposePinList(test_pins, PinName, numPins)

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001

    thehdw.Digital.Pins(alldigcont).Disconnect
     
    With thehdw.PPMU.Pins(alldigcont)
        .Connect
        .Gate = tlOn
    End With
    
    thehdw.PPMU.Pins(alldigcont).ForceV 0, 200 * uA
    

 
    thehdw.Wait 0.005
    
    For pin_cnt = 0 To numPins - 1
        'Dont test N/C Pins
        If TheExec.DataManager.channelType(PinName(pin_cnt)) = "I/O" Then
    
            measurement.AddPin (PinName(pin_cnt))
        
            thehdw.PPMU.Pins(PinName(pin_cnt)).ForceI test_current, 0.0002
            
            thehdw.Wait 0.001
        
            measurement.Pins(PinName(pin_cnt)) = thehdw.PPMU.Pins(PinName(pin_cnt)).Read
            
            thehdw.PPMU.Pins(PinName(pin_cnt)).ForceV 0, 200 * uA
            
            thehdw.Wait 0.001
            
        End If
        
    Next pin_cnt
    
    thehdw.PPMU.Pins(alldigcont).ForceV 0, 200 * uA
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=measurement, lowVal:=-0.6, hiVal:=-0.3, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    
    With thehdw.PPMU.Pins(alldigcont)
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins(alldigcont).Disconnect
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Leakage(Leakage_Pins As pinlist, Force_V As Double, Force_Hi As pinlist, Force_Lo As pinlist, Optional Current_Range As Double = 0) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Measure Leakage on I/O pins (IIL & IIH) for pull-up, pull-down and NC.
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Leakage_current As New PinListData

    On Error GoTo errHandler

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
    
    thehdw.Wait 0.001

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
    
    thehdw.Wait 0.005
    
    Leakage_current = thehdw.PPMU.Pins(Leakage_Pins).Read
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Leakage_current, lowVal:=30 * uA, hiVal:=120 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    
    With thehdw.PPMU.Pins(Leakage_Pins)
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins(Leakage_Pins).Connect
    
    thehdw.Digital.Pins(Leakage_Pins).InitState = chInitoff
    thehdw.Digital.Pins(Leakage_Pins).StartState = chStartOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Leakage_Supply(Leakage_Pins As pinlist, Force_Hi As pinlist, Force_Lo As pinlist, Optional Current_Range As Double = 0) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Measure Leakage on I/O pins (IIL & IIH) for pull-up, pull-down and NC.
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Leakage_current As New PinListData
Dim Force_V As Double

    On Error GoTo errHandler

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
    
    Force_V = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    thehdw.Wait 0.001

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
    
    thehdw.Wait 0.005
    
    Leakage_current = thehdw.PPMU.Pins(Leakage_Pins).Read
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Leakage_current, lowVal:=30 * uA, hiVal:=120 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    
    With thehdw.PPMU.Pins(Leakage_Pins)
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins(Leakage_Pins).Connect
    
    thehdw.Digital.Pins(Leakage_Pins).InitState = chInitoff
    thehdw.Digital.Pins(Leakage_Pins).StartState = chStartOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function XTAL_Parametrics() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Measure Leakage on Crystal on XTAL_I and measure output amplitude on XTAL_O
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim avdd3p3 As Double
Dim XTALI_Lkg_HI As New SiteDouble, XTALI_Lkg_LO As New SiteDouble
Dim XTALO_Voltage_HI As New SiteDouble, XTALO_Voltage_LO As New SiteDouble

    On Error GoTo errHandler


    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    thehdw.Wait 0.001
    
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    

    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo
    
    thehdw.Digital.Pins("RESET_N").InitState = chInitHi
    thehdw.Digital.Pins("RESET_N").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_I, XTAL_O").Disconnect
     
    With thehdw.PPMU.Pins("XTAL_I, XTAL_O")
        .Connect
        .Gate = tlOn
    End With
    

    thehdw.PPMU.Pins("XTAL_I").ForceV avdd3p3, 200 * uA
    thehdw.PPMU.Pins("XTAL_O").ForceI 0

    thehdw.Wait 0.005
    
    XTALI_Lkg_HI = thehdw.PPMU.Pins("XTAL_I").Read
    XTALO_Voltage_LO = thehdw.PPMU.Pins("XTAL_O").Read
    
    thehdw.PPMU.Pins("XTAL_I").ForceV 0, 200 * uA
    thehdw.PPMU.Pins("XTAL_O").ForceI 0
    
    thehdw.Wait 0.005

    
    XTALI_Lkg_LO = thehdw.PPMU.Pins("XTAL_I").Read
    XTALO_Voltage_HI = thehdw.PPMU.Pins("XTAL_O").Read
    
    
    

    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=XTALI_Lkg_HI, lowVal:=0 * uA, hiVal:=120 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=XTALO_Voltage_LO, lowVal:=1 * uA, hiVal:=0.15, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=XTALI_Lkg_LO, lowVal:=0 * uA, hiVal:=120 * uA, unit:=unitAmp, ScaleType:=scaleMicro, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=XTALO_Voltage_HI, lowVal:=3.2 * uA, hiVal:=3.5, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    
    With thehdw.PPMU.Pins("XTAL_I, XTAL_O")
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins("XTAL_I, XTAL_O").Connect
    
    thehdw.Digital.Pins("XTAL_I, XTAL_O").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_I, XTAL_O").StartState = chStartOff

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function REXT_Voltage() As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Measure Analog Bias Voltage on REXT pin
'
'   Rev 1.0 (vsomasun, Feb 9th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim REXT_V As New SiteDouble

    On Error GoTo errHandler


    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    
    thehdw.Wait 0.01
    
    thehdw.Patterns("./Patterns/PAT_MDIO_PHYID_Read.PAT").Load
    thehdw.Patterns("./Patterns/PAT_MDIO_PHYID_Read.PAT").Start
    thehdw.Digital.Patgen.HaltWait

    thehdw.Digital.Pins("REXT").Disconnect
     
    With thehdw.PPMU.Pins("REXT")
        .Connect
        .Gate = tlOn
    End With
    
    thehdw.PPMU.Pins("REXT").ForceI 0

    thehdw.Wait 0.005
    
    REXT_V = thehdw.PPMU.Pins("REXT").Read


    'Datalog voltage measurements
    TheExec.Flow.TestLimit Resultval:=REXT_V, lowVal:=3.2 * uA, hiVal:=3.5, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    
    With thehdw.PPMU.Pins("REXT")
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins("REXT").Disconnect
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
    

End Function

Public Function Power_Extlpbk_GMII_1000T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for GMII in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData


    On Error GoTo errHandler
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)

    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE04) 'GeRgmiiCfg
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn



'    Result = Extlpbk_RGMII_1000T_Power_yoda_test

    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    thehdw.Wait 0.05
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    Temp = meas_temp_mdio(100 * uA)
    
    'Disable Frame generator
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, 0) ' FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, 0) ' FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2) 'DiagClkEn Disable
    
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function


errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function




Public Function Power_Extlpbk_RGMII_1000T()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for RGMII in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData


    On Error GoTo errHandler
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)

    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn



'    Result = Extlpbk_RGMII_1000T_Power_yoda_test

    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.05
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    Temp = meas_temp_mdio(100 * uA)
    
    'Disable Frame generator
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, 0) ' FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, 0) ' FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2) 'DiagClkEn Disable
    
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function


errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Power_Extlpbk_RGMII_100T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for RGMII 100M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_RGMII_100T_Power_yoda_test

    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    'DiagClkEn Disable
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)
    
    Temp = meas_temp_mdio(100 * uA)
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Power_Extlpbk_RGMII_10T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for RGMII 10M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff


    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_RGMII_10T_Power_yoda_test
    
    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    'DiagClkEn Disable
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)
    
    Temp = meas_temp_mdio(100 * uA)
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function


Public Function Power_Extlpbk_MII_100T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for MII 100M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_MII_100T_Power_yoda_test
    
    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE04) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site

    Temp = meas_temp_mdio(100 * uA)
    
    'Disable Frame generator
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, 0) ' FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, 0) ' FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2) 'DiagClkEn Disable
    
    
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function


Public Function Power_Extlpbk_MII_10T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for MII 10M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff


    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_MII_10T_Power_yoda_test
    
    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE04) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    'Disable Frame generator
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, 0) ' FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, 0) ' FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2) 'DiagClkEn Disable
    
    Temp = meas_temp_mdio(100 * uA)
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function Power_Extlpbk_RMII_100T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for RMII 100M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_RGMII_100T_Power_yoda_test

    'Enable 50Mhz clock to be divided by 2
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65288, &HC) 'GeXtalDiv2En  'GeXtalBypLpEn

    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE04) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65316, &H117) 'GeRmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    'DiagClkEn Disable
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)
    
    Temp = meas_temp_mdio(100 * uA)
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Power_Extlpbk_RMII_10T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement test for RMII 10M in Loopback mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff


    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)


'    Result = Extlpbk_RGMII_10T_Power_yoda_test

    'Enable 50Mhz clock to be divided by 2
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65288, &HC) 'GeXtalDiv2En  'GeXtalBypLpEn
    
    'Disable RefClk
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, &H4) 'GeRxTxExtLbEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, &H1) 'BrkLnkFrc

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE04) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65316, &H117) 'GeRmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H0) 'GeMiiUseGtxClk
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H100)

    'Wait for link
    GEPHY_MDIO_Wait 100 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)

    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37911, 1) 'FgContModeEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1500) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37910, &H1) 'FgCntrl
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    'DiagClkEn Disable
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)
    
    Temp = meas_temp_mdio(100 * uA)
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in LPBK Mode
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function



Public Function Power_IDDQ(IDDQ_Pattern As Pattern)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Power Measurement in IDDQ mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String

Dim UVI80_current As New PinListData

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff
    
    thehdw.Wait 0.01
    
    thehdw.Patterns(IDDQ_Pattern).Load
    thehdw.Patterns(IDDQ_Pattern).Start
    thehdw.Digital.Patgen.HaltWait

     
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 2 * mA, 2 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 20 * mA, 20 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 2 * mA, 2 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 2 * mA, 2 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=50)
    
    
    
    


    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function HVST()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   High Voltage Stress Test (HVST) for QA Reliability. Doc ADI1310 Rev.B
'   Test done to exercise Analog blocks
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ReferenceTime As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim Pattern_Name As String

Dim HVST_Time As New DSPWave

Dim UVI80_current As New PinListData

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

'''    If package_type = "pkg40" Then
'''        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
'''    Else
'''        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
'''    End If

    '
    Call HVST_Time.CreateConstant(0, 2, DspDouble)

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    
    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff
    thehdw.Wait 0.01
    
    'Wait to avoid mode alarm on VDDIO_M supply
    If package_type = "pkg64" Then
        thehdw.Wait 0.2
    End If
    

'''    'PHY and MAC Relay in LPBK Mode
'''    thehdw.Utility.Pins("UTIL_KA1, UTIL_KA2, UTIL_KA3, UTIL_KA4").State = tlUtilBitOff
'''    thehdw.Utility.Pins("UTIL_KB1, UTIL_KB2, UTIL_KB3, UTIL_KB4").State = tlUtilBitOff
'''    thehdw.Utility.Pins("UTIL_KR1, UTIL_KR2, UTIL_KR3, UTIL_KR4, UTIL_KR9, UTIL_KR11").State = tlUtilBitOff
'''
'''
'''
'''    'Apply Levels and Timing
'''    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, _
'''    TimeSetSheet:="TSB_Static_Stuck_At", ACCategory:="40ns", ACSelector:="Typ", InitPinsHi:="TEST_EN"
'''
'''    'Create Starting Reference Time
'''    ReferenceTime = TheExec.Timer
'''
'''    Pattern_Name = ".\Patterns\DFT\GePhy_Static_Comp_Logic.PAT"
'''
'''    thehdw.Patterns(Pattern_Name).Load
'''    thehdw.Patterns(Pattern_Name).Start
'''    thehdw.Digital.Patgen.HaltWait
'''
'''    'Measure Test Time and Reset Reference Time
'''    HVST_Time.Element(0) = TheExec.Timer(ReferenceTime)
'''    ReferenceTime = TheExec.Timer
'''
'''    Pattern_Name = ".\Patterns\DFT\ROC_Static_FS_Logic.PAT"
'''
'''    thehdw.Patterns(Pattern_Name).Load
'''    thehdw.Patterns(Pattern_Name).Start
'''    thehdw.Digital.Patgen.HaltWait
'''
'''    'Measure Test Time and Reset Reference Time
'''    HVST_Time.Element(1) = TheExec.Timer(ReferenceTime)
    
    thehdw.Patterns(".\Patterns\Grandlpbk_100T_200ns_rz.PAT").Load
    thehdw.Patterns(".\Patterns\Grandlpbk_1000T_200ns_rz.PAT").Load
    thehdw.Patterns(".\Patterns\Grandlpbk_10T_200ns_rz.PAT").Load
    
    
    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, _
    TimeSetSheet:="TSB_MDIO_DSSC", ACCategory:="200ns", ACSelector:="Typ", InitPinsLo:="RXD_6, CRS, TEST_EN", InitPinsHi:="LED_2"
    
    

    
    ReferenceTime = TheExec.Timer
    
    If package_type = "pkg32" Then
    
        Pattern_Name = ".\Patterns\Grandlpbk_100T_200ns_rz.PAT"
        
        thehdw.Patterns(Pattern_Name).Start
        thehdw.Digital.Patgen.HaltWait
        
    Else
    
        Pattern_Name = ".\Patterns\Grandlpbk_1000T_200ns_rz.PAT"
'        thehdw.Wait 0.09
        thehdw.Patterns(Pattern_Name).Start
        thehdw.Digital.Patgen.HaltWait
    
    End If
    
    'Measure Test Time and Reset Reference Time
    HVST_Time.Element(0) = TheExec.Timer(ReferenceTime)
    ReferenceTime = TheExec.Timer
    
    Pattern_Name = ".\Patterns\Grandlpbk_10T_200ns_rz.PAT"
    
    thehdw.Patterns(Pattern_Name).Start
    thehdw.Digital.Patgen.HaltWait
    
    'Measure Test Time and Reset Reference Time
    HVST_Time.Element(1) = TheExec.Timer(ReferenceTime)
    ReferenceTime = TheExec.Timer
    
    

'    TheExec.Flow.TestLimit resultVal:=HVST_Time.Element(0), unit:=unitTime, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="GePhy_Static_Comp_Logic"
'    TheExec.Flow.TestLimit resultVal:=HVST_Time.Element(1), unit:=unitTime, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="ROC_Static_FS_Logic"
    If package_type = "pkg32" Then
        TheExec.Flow.TestLimit Resultval:=HVST_Time.Element(0), unit:=unitTime, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="Grandlpbk_100T"
    Else
        TheExec.Flow.TestLimit Resultval:=HVST_Time.Element(0), unit:=unitTime, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="Grandlpbk_1000T"
    End If
    TheExec.Flow.TestLimit Resultval:=HVST_Time.Element(1), unit:=unitTime, ScaleType:=scaleMilli, ForceResults:=tlForceNone, Tname:="Grandlpbk_10T"


    'Reset PHY and MAC Relay
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function AMVR()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Absolute Maximum Voltage ratings (AMVR) for QA Reliability. Doc ADI1264 Rev.C
'   Test done to screen latent defect under AMVR conditions
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ReferenceTime As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim Pattern_Name As String
Dim cond As String


Dim UVI80_pre_current As New PinListData, UVI80_post_current As New PinListData
Dim delta_current As New PinListData

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg32" Then
        power_pins = "AVDD3P3, VDDIO_R"
    ElseIf package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    
    'Apply Levels and Timing

    
    If package_type = "pkg32" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg32", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR_pkg32", InitPinsHiZ:="alldigpins, MDI_Pins"
    ElseIf package_type = "pkg40" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg40", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR", InitPinsHiZ:="alldigpins, MDI_Pins"
    Else
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg64", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR", InitPinsHiZ:="alldigpins, MDI_Pins"
    End If
    
    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    cond = "COND: " & "AVDD3P3 = " & Format(avdd3p3, "0.000") _
                        & ", VDDIO_R = " & Format(vddio_r, "0.000")
                        
    If package_type = "pkg40" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000")
    ElseIf package_type = "pkg64" Then
        cond = cond & ", VDD0P9 = " & Format(vdd0p9, "0.000") _
                    & ", AVDD1P2 = " & Format(avdd1p2, "0.000") _
                    & ", VDDIO = " & Format(vddio, "0.000")
    End If
    

    'Display Voltage Conditions in Datalog
    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment (cond)
    TheExec.Datalog.WriteComment ("")
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 20 * mA, 20 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 20 * mA, 20 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 20 * mA, 20 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    thehdw.Wait 0.015
    
    'Measure current of supplies
    UVI80_pre_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
   
    'Elevate Supplies and I/O pins to AMVR levels
    If package_type = "pkg32" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Max", PinLevelsSheet:="Levels_AMVR_pkg32", InitPinsHi:="alldigpins, MDI_Pins"
    Else
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Max", PinLevelsSheet:="Levels_AMVR", InitPinsHi:="alldigpins, MDI_Pins"
    End If

    
    thehdw.Wait 0.01
    
    'Reduce I/O pins to minimum AMVR levels
    If package_type = "pkg32" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Max", PinLevelsSheet:="Levels_AMVR_pkg32", InitPinsLo:="alldigpins, MDI_Pins"
    Else
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Max", PinLevelsSheet:="Levels_AMVR", InitPinsLo:="alldigpins, MDI_Pins"
    End If

    
    thehdw.Wait 0.01
    
    

    'Set the supplies to Operating voltages
    If package_type = "pkg32" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg32", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR_pkg32", InitPinsHiZ:="alldigpins, MDI_Pins"
    ElseIf package_type = "pkg40" Then
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg40", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR", InitPinsHiZ:="alldigpins, MDI_Pins"
    Else
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="Zero", DCSelector:="Typ", PinLevelsSheet:="Levels_pkg64", InitPinsLo:="alldigpins, MDI_Pins"
        thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=False, RelayMode:=tlPowered, _
        DCCategory:="AMVR", DCSelector:="Typ", PinLevelsSheet:="Levels_AMVR", InitPinsHiZ:="alldigpins, MDI_Pins"
    End If
    
    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 20 * mA, 20 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 20 * mA, 20 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 20 * mA, 20 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 2 * mA, 2 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 20 * mA, 20 * mA
    End If
    
    thehdw.Wait 0.015
    
    'Measure current of supplies
    UVI80_post_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    delta_current = UVI80_post_current.Math.Subtract(UVI80_pre_current)
    
    
    
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_pre_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    
    TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_post_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=delta_current, unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow



    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function VLV_Threshold()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Test to measure Very Low Voltage Operating point for Scan patterns
'   Test meant for HTOL to detect shifts in VLV operating points
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ReferenceTime As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim Pattern_Name As String
Dim site As Variant

Dim GePhy_static_VLV_Voltage As New SiteDouble, GePhy_dynamic_VLV_Voltage As New SiteDouble
Dim ROC_VLV_Voltage As New SiteDouble

Dim HVST_Time As New DSPWave

Dim UVI80_current As New PinListData

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

'''    If package_type = "pkg40" Then
'''        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
'''    Else
'''        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
'''    End If

    '
    Call HVST_Time.CreateConstant(0, 4, DspDouble)

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    

    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, _
    TimeSetSheet:="TSB_Static_Stuck_At", ACCategory:="60ns", ACSelector:="Typ", InitPinsHi:="TEST_EN"
    
    
    Pattern_Name = "GePhy_Static_Comp_Scan"   '".\Patterns\DFT\GePhy_Static_Comp_Logic.PAT"
    
    For Each site In TheExec.Sites.Active
    
        GePhy_static_VLV_Voltage = VLV_Level_Search(Pattern_Name, "VDD0P9", 0.8, 0.65, -0.005, 3)
        
    Next site
    
    

    
    Pattern_Name = "ROC_Static_FS_Logic"    '".\Patterns\DFT\ROC_Static_FS_Logic.PAT"
    
    For Each site In TheExec.Sites.Active

        ROC_VLV_Voltage = VLV_Level_Search(Pattern_Name, "VDD0P9", 0.8, 0.65, -0.005, 3)
    
    Next site
    
    
    'Apply Levels and Timing
    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered, _
    TimeSetSheet:="TSB_GePHY_Dynamic_TDF", ACCategory:="60ns", ACSelector:="Typ", InitPinsHi:="TEST_EN"
    
    
    Pattern_Name = "GePhy_Dynamic_Comp_Logic"    '".\Patterns\DFT\GePhy_Dynamic_Comp_Logic.PAT"
    
    For Each site In TheExec.Sites.Active
    
        GePhy_dynamic_VLV_Voltage = VLV_Level_Search(Pattern_Name, "VDD0P9", 0.8, 0.65, -0.005, 3)
        
    Next site
    
    

    

    TheExec.Flow.TestLimit Resultval:=GePhy_static_VLV_Voltage, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="GePhy_Static_Comp_Logic"
    TheExec.Flow.TestLimit Resultval:=ROC_VLV_Voltage, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="ROC_Static_FS_Logic"
    TheExec.Flow.TestLimit Resultval:=GePhy_dynamic_VLV_Voltage, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceFlow, Tname:="GePhy_Dynamic_Comp_Logic"


    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function Power_SftPd()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Software Power down Power measurements
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    If package_type = "pkg40" Then
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R"
    Else
        power_pins = "VDD0P9, AVDD3P3, VDDIO_R, VDDIO_M, VDDIO"
    End If

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff
    
    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    'Set Chip to Powerdown mode . SftPd bit 11 = 1
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H3900)
    
    GEPHY_MDIO_Wait (0.01)
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    If package_type = "pkg40" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
    ElseIf package_type = "pkg64" Then
        vdd0p9 = thehdw.DCVI.Pins("VDD0P9").Voltage
        vddio_m = thehdw.DCVI.Pins("VDDIO_M").Voltage
        vddio = thehdw.DCVI.Pins("VDDIO").Voltage
    End If
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    If package_type = "pkg40" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    If package_type = "pkg64" Then
        thehdw.DCVI.Pins("VDD0P9").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO_M").SetCurrentAndRange 200 * mA, 200 * mA
        thehdw.DCVI.Pins("VDDIO").SetCurrentAndRange 200 * mA, 200 * mA
    End If
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    UVI80_current = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
'    For Each site In TheExec.Sites.Active
        power = power.Add(UVI80_current.Pins("AVDD3P3").Multiply(avdd3p3)).Add(UVI80_current.Pins("VDDIO_R").Multiply(vddio_r))
        If package_type = "pkg40" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
        ElseIf package_type = "pkg64" Then
            power = power.Add(UVI80_current.Pins("VDD0P9").Multiply(vdd0p9))
            power = power.Add(UVI80_current.Pins("VDDIO_M").Multiply(vddio_m))
            power = power.Add(UVI80_current.Pins("VDDIO").Multiply(vddio))
        End If
'    Next site
    
    
    Call meas_temp_mdio(20 * uA)
    Call meas_temp_mdio(50 * uA)
    Temp = meas_temp_mdio(100 * uA)
    
    
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=Temp, unit:=unitCustom, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    If package_type = "pkg40" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    ElseIf package_type = "pkg64" Then
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDD0P9"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO_M"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
        TheExec.Flow.TestLimit Resultval:=UVI80_current.Pins("VDDIO"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    End If
    
    TheExec.Flow.TestLimit Resultval:=power, unit:=unitCustom, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function meas_temp_mdio(test_current As Double) As SiteDouble
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply current to INT_N Pin and measure Voltage on LED_0 to compute jn temperature
'
'   Rev 1.0 (vsomasun, Nov 27th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'GeTmpSnsSel selects between which signal is connected
'to the LED_0 pad, to measure the voltage, when GeTmpSnsEn is set, as follows:
'1'b0 Internal ground, 1'b1 Temperature sensing diode

'GeTmpSnsEn sets the INT_N and LED_0 chip I/Os in high impedance and connects
'the INT_N pad to the temperature sensing diode in order to drive current into
'it for the temperature measurements. The LED_0 pad is also connected to either the diode or the
'internal ground, denpending on GeTmpSnsSel, in order to measure the diode or ground voltages.

Dim GeTmpSnsCntrl As Long, GeTmpSnsEn As Long, GeTmpSnsSel As Long

Dim diode_meas As New SiteDouble, gnd_meas As New SiteDouble
Dim Read As New SiteLong
Dim Tname As String
Dim Temp As New SiteDouble

    Set meas_temp_mdio = New SiteDouble

    'Enable Temp Sensing Diode Path
    GeTmpSnsEn = 1
    GeTmpSnsSel = 1
    GeTmpSnsCntrl = (GeTmpSnsSel * 2) + GeTmpSnsEn
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65352, GeTmpSnsCntrl)
    


    thehdw.Digital.Pins("INT_N, LED_0").Disconnect
     
    With thehdw.PPMU.Pins("INT_N, LED_0")
        .Connect
        .Gate = tlOn
    End With
    
    thehdw.PPMU.Pins("INT_N").ForceI test_current, 200 * uA
    
    thehdw.PPMU.Pins("LED_0").ForceI 0, 0.0002
    
    'Wait for diode to stabilize
    thehdw.Wait 0.005
    
    diode_meas = thehdw.PPMU.Pins("LED_0").Read(tlPPMUReadMeasurements, 20)
    
    'Enable Internal Ground Path
    GeTmpSnsEn = 1
    GeTmpSnsSel = 0
    GeTmpSnsCntrl = (GeTmpSnsSel * 2) + GeTmpSnsEn
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65352, GeTmpSnsCntrl)
       
    thehdw.PPMU.Pins("LED_0").ForceI 0, 0.0002
    
    'Wait for diode to stabilize
    thehdw.Wait 0.005
    
    gnd_meas = thehdw.PPMU.Pins("LED_0").Read(tlPPMUReadMeasurements, 20)
    
    Tname = "Diode_Volt_" & Format(test_current / (1 * uA)) & "uA"
    
    diode_meas = diode_meas.Subtract(gnd_meas)
    
    'Second order polynomial measurement
    Temp = Temp.Add(diode_meas.power(2).Multiply(-531.55)).Subtract(diode_meas.Multiply(133.63)).Add(452.5)
    
'''    'Linear Interpolation measurement
'''    Temp = Temp.Add(diode_meas.Multiply(-909.53)).Add(734.8)

    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=diode_meas, lowVal:=0.2, hiVal:=0.9, unit:=unitVolt, ScaleType:=scaleNone, ForceResults:=tlForceNone, Tname:=Tname
    
    With thehdw.PPMU.Pins("INT_N, LED_0")
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins("INT_N, LED_0").Disconnect
    
    
    GeTmpSnsEn = 0
    GeTmpSnsSel = 0
    GeTmpSnsCntrl = (GeTmpSnsSel * 2) + GeTmpSnsEn
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65352, GeTmpSnsCntrl)
    
    meas_temp_mdio = Temp



End Function


Public Function ADIN1200_Regulator()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Apply Power to VDD Pins and Levels to I/O pins
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim vdd0p9 As Double, avdd3p3 As Double, avdd1p2 As Double
Dim vddio As Double, vddio_m As Double, vddio_r As Double
Dim ChanMap As String
Dim package_type As String
Dim power_pins As String
Dim power As New SiteDouble
Dim Temp As New SiteDouble
Dim site As Variant

Dim UVI80_current As New PinListData
Dim I_UVI80_0mA As New PinListData, I_UVI80_load As New PinListData
Dim V_VDD0P9_0mA As New PinListData, V_VDD0P9_load As New PinListData

    On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    power_pins = "AVDD3P3, VDDIO_R"

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Wait 0.01
    
    thehdw.DCVI.Pins("VDD0P9").Alarm = tlAlarmOff

    Call GEPHY_MDIO_Init
    
''    'Disable RefClk
''    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 65311, &H0)
    
    'Set Chip to Powerdown mode . SftPd bit 11 = 1
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H3900)
    
    GEPHY_MDIO_Wait (0.01)
    
    
    With thehdw.DCVI.Pins(power_pins)
        .Meter.mode = tlDCVIMeterCurrent
    End With
    
    'Measure Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    
     'Program the Current and Current Range
    thehdw.DCVI.Pins("AVDD3P3").SetCurrentAndRange 200 * mA, 200 * mA
    thehdw.DCVI.Pins("VDDIO_R").SetCurrentAndRange 200 * mA, 200 * mA
    
    thehdw.Wait 0.005
    
    'Measure current of supplies
    I_UVI80_0mA = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    With thehdw.DCVI.Pins("VDD0P9")
        .mode = tlDCVIModeCurrent
        .Meter.mode = tlDCVIMeterVoltage
        .current = 0
        .Gate(tlDCVIGateHiZ) = False
        .Connect
    End With
    
    V_VDD0P9_0mA = thehdw.DCVI.Pins("VDD0P9").Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
        
    
    With thehdw.DCVI.Pins("VDD0P9")
        .mode = tlDCVIModeCurrent
        .Meter.mode = tlDCVIMeterVoltage
        .current = -0.07
        .Gate = True
    End With
    

    thehdw.Wait 0.01
    
    'Measure current of supplies
    I_UVI80_load = thehdw.DCVI.Pins(power_pins).Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    V_VDD0P9_load = thehdw.DCVI.Pins("VDD0P9").Meter.Read(StrobeOption:=tlStrobe, SampleSize:=20)
    
    

    With thehdw.DCVI.Pins("VDD0P9")
        .Disconnect
        .Gate = False
        .current = 0
        .Alarm = tlAlarmDefault
    End With
    
  
    
    

    Call GEPHY_MDIO_Halt
    
    'Datalog current measurements
    TheExec.Flow.TestLimit Resultval:=I_UVI80_0mA.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=I_UVI80_0mA.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    TheExec.Flow.TestLimit Resultval:=V_VDD0P9_0mA, unit:=unitVolt, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    TheExec.Flow.TestLimit Resultval:=I_UVI80_load.Pins("AVDD3P3"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow

    TheExec.Flow.TestLimit Resultval:=I_UVI80_load.Pins("VDDIO_R"), unit:=unitAmp, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    
    TheExec.Flow.TestLimit Resultval:=V_VDD0P9_load, unit:=unitVolt, ScaleType:=scaleMilli, ForceResults:=tlForceFlow
    

    
    
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function


