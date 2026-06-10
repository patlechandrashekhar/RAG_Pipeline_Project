Attribute VB_Name = "VBT_Tests"
Public Function MDIO_Read()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Test to check functionality of MDIO
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong




    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff


    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    
    Call MDIO_Access_test

    Call GEPHY_MDIO_Halt


End Function

Public Function ABIST()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   ABIST Test
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

    On Error GoTo errHandler


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)
    
    Result = ABIST_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 1
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
    
    


End Function

Public Function Grandlpbk_1000T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of Gigabit Mode
'   RGMII mode in dly/no dly tested
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

On Error GoTo errHandler


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    

    


    Result = Grandlpbk_1000T_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 9
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Grandlpbk_100T()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of 100M Mode
'   RGMII mode in dly/no dly tested. MII mode tested
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

On Error GoTo errHandler

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)
    


    Result = Grandlpbk_100T_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 15
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Grandlpbk_10T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of 10M Mode
'   RGMII mode in dly/no dly tested. MII mode tested
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

On Error GoTo errHandler


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)
    


    Result = Grandlpbk_10T_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 15
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC Connected to HSD
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function


Public Function Extlpbk_GMII_1000T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of GMII Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer




    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    


    Result = Extlpbk_GMII_1000T_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 5
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    


End Function


Public Function Extlpbk_RGMII_1000T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of RGMII Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer




    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff

    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    


    Result = Extlpbk_RGMII_1000T_yoda_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 5
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    


End Function

Public Function Extlpbk_RGMII_100T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of RGMII 100M Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer




    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



'    thehdw.Digital.Pins("MDIO_PULLUP_V").Connect
'    thehdw.Digital.Pins("MDIO_PULLUP_V").Levels.Value(chVih) = 3.3
'    thehdw.Digital.Pins("MDIO_PULLUP_V").InitState = chInitHi
'    thehdw.Digital.Pins("MDIO_PULLUP_V").StartState = chStartHi
    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    
    Result = Extlpbk_RGMII_100T_yoda_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 5
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff


End Function

Public Function Extlpbk_RGMII_10T()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Grandloopback test to check functionality of RGMII 10M Mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer




    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff



    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    
    Result = Extlpbk_RGMII_10T_yoda_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 5
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff


End Function

Public Function Process_Monitor()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to identify Skew information of Silicon
'   Skew shifts is identified from Frequency measured in the code
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

On Error GoTo errHandler

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.005


    Call GEPHY_MDIO_Init
    
    'Process monitor count over 81.96 us Period
    Result = Process_Monitor_test

    Call GEPHY_MDIO_Halt
    
    For i = 0 To 8
        TheExec.Flow.TestLimit Resultval:=Result.Element(i), unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    Next i
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function GeClkCfg_Frequency()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure different clock frequencies from PLL using test mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong

Dim i As Integer

Dim GeClkCfg As Long, GeClk25En As Long, GeClkHrtFreeEn As Long
Dim GeClkHrtRcvrEn As Long, GeRefClkEn As Long, GeClkFree125En As Long
Dim GeClkRcvr125En As Long, GeTclkEn As Long
Dim TstClkEn As Long

Dim ChanMap As String
Dim package_type As String



Dim sdclk_refclk_25MHz As New PinListData
Dim sdclk_ge_tclk_125MHz As New PinListData

Dim sdclk_25MHz As New PinListData
Dim sdclk_hrt_free_125MHz As New PinListData
Dim sdclk_hrt_rcvr_125MHz As New PinListData
Dim sdclk_free_125MHz As New PinListData
Dim sdclk_rcvr_125MHz As New PinListData

Dim freq As New PinListData

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)

    'Default Value
    GeClk25En = 0
    GeClkHrtFreeEn = 0
    GeClkHrtRcvrEn = 0
    GeRefClkEn = 1
    GeClkFree125En = 0
    GeClkRcvr125En = 0
    GeTclkEn = 0
    TstClkEn = 0
    
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkRcvr125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
            
  
    'Connect pins
    thehdw.Digital.Pins("GP_CLK,REF_CLK").Connect
    'Enable the frequency counter with the period counter.
    thehdw.Digital.Pins("GP_CLK,REF_CLK").FreqCtr.Enable = IntervalEnable

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo


    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init

    GeClk25En = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                
    'clk_25MHz
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_25MHz = Measure_Frequency("GP_CLK", 0.005)

    'clk_hrt_free_125MHz
    GeClk25En = 0
    GeClkHrtFreeEn = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_hrt_free_125MHz = Measure_Frequency("GP_CLK", 0.005)
    
    'clk_hrt_rcvr_125MHz
    GeClkHrtFreeEn = 0
    GeClkHrtRcvrEn = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_hrt_rcvr_125MHz = Measure_Frequency("GP_CLK", 0.005)
    
    'clk_hrt_rcvr_125MHz
    GeClkHrtRcvrEn = 0
    GeClkFree125En = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                
    'clk_25MHz
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_free_125MHz = Measure_Frequency("GP_CLK", 0.005)
    
    'clk_rcvr_125MHz
    GeClkFree125En = 0
    GeClkRcvr125En = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                
    'clk_25MHz
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_rcvr_125MHz = Measure_Frequency("GP_CLK", 0.005)
    
    'clk_rcvr_125MHz
    GeClkRcvr125En = 0
    GeTclkEn = 1
    TstClkEn = 1
    GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                
    'clk_ge_tclk_125MHz
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
    Call GEPhy_MDIO_CL_22_Write(&H0, 22, 776 + (TstClkEn * (2 ^ 5))) 'Default value + programmed value
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
    sdclk_ge_tclk_125MHz = Measure_Frequency("GP_CLK", 0.005)
    
    If package_type <> "pkg32" Then
    
    
        'clk_rcvr_125MHz
        GeRefClkEn = 1
        GeClkCfg = (GeTclkEn * (2 ^ 6)) + (GeClkRcvr125En * (2 ^ 5)) + (GeClkFree125En * (2 ^ 4)) + _
                    (GeRefClkEn * (2 ^ 3)) + (GeClkHrtRcvrEn * (2 ^ 2)) + (GeClkHrtFreeEn * (2 ^ 1)) + GeClk25En
                    
        'clk_ge_tclk_125MHz
        Call GEPhy_MDIO_CL_45_Write(0, 30, 65311, GeClkCfg)
    '    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)
        sdclk_refclk_25MHz = Measure_Frequency("REF_CLK", 0.005)
    
    End If
    
    



    
    'Datalogging
    Call TheExec.Flow.TestLimit(sdclk_25MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(sdclk_hrt_free_125MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(sdclk_hrt_rcvr_125MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(sdclk_free_125MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(sdclk_rcvr_125MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(sdclk_ge_tclk_125MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    
    If package_type <> "pkg32" Then
        Call TheExec.Flow.TestLimit(sdclk_refclk_25MHz, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    End If


    Call GEPHY_MDIO_Halt
    
    'Enable the frequency counter with the period counter.
    thehdw.Digital.Pins("GP_CLK,REF_CLK").FreqCtr.Enable = Disable
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function MII_TX_CLK_Frequency()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to measure TX_CLK Frequency
'   Test is to ensure that the TX_CLK frequency is coming out of the device
'   since Grandloopback test will still pass if there is no output
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong

Dim i As Integer

Dim TX_CLK_10M_Freq As New PinListData
Dim TX_CLK_100M_Freq As New PinListData

Dim ChanMap As String
Dim package_type As String
Dim Pin_Name As String

On Error GoTo errHandler

    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    If package_type = "pkg64" Then
        Pin_Name = "TX_CLK"
    Else
        Pin_Name = "GTX_CLK"
    End If

    'Connect pins
    thehdw.Digital.Pins(Pin_Name).Connect
    'Enable the frequency counter with the period counter.
    thehdw.Digital.Pins(Pin_Name).FreqCtr.Enable = IntervalEnable

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins").StartState = chStartOff

    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi


    Call GEPHY_MDIO_Init

    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2100)

    
    'Configuring GeSubsys for MII mode  GeMiiUseGtxClk = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE02) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H2) 'GeMiiUseGtxClk
    
    
'    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65311)

    
    TX_CLK_100M_Freq = Measure_Frequency(Pin_Name, 0.0005)
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H100)
    
    TX_CLK_10M_Freq = Measure_Frequency(Pin_Name, 0.0005)

    
    'Datalogging
    Call TheExec.Flow.TestLimit(TX_CLK_10M_Freq, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(TX_CLK_100M_Freq, ScaleType:=scaleMega, unit:=unitHz, ForceResults:=tlForceFlow)
    


    Call GEPHY_MDIO_Halt
    
    'Enable the frequency counter with the period counter.
    thehdw.Digital.Pins(Pin_Name).FreqCtr.Enable = Disable
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function




Public Function ADIN1200_MDI_Parametric()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Measure Analog Parametrics of the MDI Pin for 32 pin package
'   which supports only Dim 0 & 1 for 100M
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim avdd3p3 As Double
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim common_mode_v As New PinListData
Dim DIM_0_CM As New PinListData
Dim DIM_1_CM As New PinListData
Dim impedance As New PinListData
Dim DIM_0_imp As New PinListData
Dim DIM_1_imp As New PinListData
Dim V_0mA As New PinListData, I_CM As New PinListData
Dim V_CM As Double
Dim pin As Variant

On Error GoTo errHandler


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    
    'PHY MDI pins need to be connected to HSD
    thehdw.Utility.Pins("UTIL_KA1, UTIL_KA2").State = tlUtilBitOff

    thehdw.Digital.Pins("alldigpins, MDI_Pins").InitState = chInitoff
    thehdw.Digital.Pins("alldigpins, MDI_Pins").StartState = chStartOff

    thehdw.Digital.Pins("TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("TEST_EN").StartState = chStartLo

    thehdw.Wait 0.005
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (True)

    Call GEPHY_MDIO_Init
    
    
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    V_CM = avdd3p3 / 2
    thehdw.Digital.Pins("MDI_Pins").Levels.Value(chVt) = V_CM

    
'    'Sftpd = 1
'    Call GEPhy_MDIO_CL_22_Write(0, 0, &H3800)
'
'    Call GEPhy_MDIO_CL_22_Write(0, 9, &H2000)
    
    'Transmit long dwell time on dimension 0 (zero on dimension 1)
    'bfWrite BF_GEPHY.B100TxTstMode, MSV(&H3)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46099, &H3)
    
    'Enable active termination on RX dimension
    'bfWrite BF_GEPHY.B100ZptmEnDimrx, MSV(&H0)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46725, &H0)
    
   'Disable functional controls of TX DAC switches
    'bfWrite BF_GEPHY.FenTxSwAbEn, MSV(&H0)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46713, &H0)
    
'    'Sftpd = 0
'    Call GEPhy_MDIO_CL_22_Write(0, 0, &H3000)
    
    'May need to adjust this value?
    'bfWrite BF_GEPHY.B100TxDacLvl, MSV(&H9)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46086, &H9)
    
   'Enable TX DAC switches for both dimensions
    'bfWrite BF_GEPHY.TxSwAbEn, MSV(&H9)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46698, &H9)
    
   'Now the TX DAC should be enabled on both dimensions
   'And transmitting a "0" on dimension 1 (and some data on dimension 0)
   'So measurements can be performed on dimension 1

    With thehdw.PPMU.Pins("MDI_Pins")
        .ForceI 0, 2 * mA
        .Gate = tlOn
        .Connect
    End With
    thehdw.Digital.Pins("MDI_Pins").Levels.DriverMode = tlDriverModeVt
    thehdw.Digital.Pins("MDI_Pins").Connect
    thehdw.Wait 0.005
    
   DIM_1_CM = thehdw.PPMU.Pins("MDI_1_N, MDI_1_P").Read
   
   thehdw.Digital.Pins("MDI_1_N, MDI_1_P").Disconnect
   thehdw.Wait 0.005
   
    thehdw.PPMU.Pins("MDI_1_N, MDI_1_P").ForceV V_CM, 2 * mA
    thehdw.Wait 0.005

    
    I_CM = thehdw.PPMU.Pins("MDI_1_N, MDI_1_P").Read
    
    thehdw.PPMU.Pins("MDI_1_N, MDI_1_P").ForceI 0 * mA, 200 * uA
    thehdw.Wait 0.005
    
    V_0mA = thehdw.PPMU.Pins("MDI_1_N, MDI_1_P").Read
    
    DIM_1_imp = V_0mA.Math.Subtract(V_CM).Divide(I_CM).Negate
   
    
    
    'Transmit long dwell time on dimension 1 (zero on dimension 0)
    'bfWrite BF_GEPHY.B100TxTstMode, MSV(&H4)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46099, &H4)
    
    
   'Now the TX DAC should be enabled on both dimensions
   'And transmitting a "0" on dimension 0 (and some data on dimension 1)
   'So measurements can be performed on dimension 0
   
   DIM_0_CM = thehdw.PPMU.Pins("MDI_0_N, MDI_0_P").Read
   
   thehdw.Digital.Pins("MDI_0_N, MDI_0_P").Disconnect
   thehdw.Wait 0.005
   
    thehdw.PPMU.Pins("MDI_0_N, MDI_0_P").ForceV V_CM, 2 * mA
    thehdw.Wait 0.005
   
    I_CM = thehdw.PPMU.Pins("MDI_0_N, MDI_0_P").Read
    
    
    thehdw.PPMU.Pins("MDI_0_N, MDI_0_P").ForceI 0 * mA, 200 * uA
    thehdw.Wait 0.005
    
    V_0mA = thehdw.PPMU.Pins("MDI_0_N, MDI_0_P").Read
    
    DIM_0_imp = V_0mA.Math.Subtract(V_CM).Divide(I_CM).Negate
   


    Call GEPHY_MDIO_Halt
    
    'Consolidate Common mode voltage in single pinlistdata variable
    For Each pin In DIM_0_CM.Pins
        common_mode_v.AddPin (DIM_0_CM.Pins(pin))
        common_mode_v.Pins(pin) = DIM_0_CM.Pins(pin)
    Next pin
    
    For Each pin In DIM_1_CM.Pins
        common_mode_v.AddPin (DIM_1_CM.Pins(pin))
        common_mode_v.Pins(pin) = DIM_1_CM.Pins(pin)
    Next pin
    
    'Consolidate Impedance in single pinlistdata variable
    For Each pin In DIM_0_imp.Pins
        impedance.AddPin (DIM_0_imp.Pins(pin))
        impedance.Pins(pin) = DIM_0_imp.Pins(pin)
    Next pin
    
    For Each pin In DIM_1_imp.Pins
        impedance.AddPin (DIM_1_imp.Pins(pin))
        impedance.Pins(pin) = DIM_1_imp.Pins(pin)
    Next pin
    



    TheExec.Flow.TestLimit Resultval:=common_mode_v, unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=impedance, unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow



    With thehdw.PPMU.Pins("MDI_Pins")
        .ForceI 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
    
    thehdw.Digital.Pins("MDI_Pins").Levels.Value(chVt) = 0
    thehdw.PPMU.AllowPPMUFuncRelayConnection (False)
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function MDI_Parametric()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Routine to Measure Analog Parametrics of the MDI Pin in Gigabit mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim avdd3p3 As Double
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim i As Integer

Dim common_mode_v As New PinListData
Dim impedance As New PinListData
Dim V_0mA As New PinListData, I_CM As New PinListData
Dim V_CM As Double

On Error GoTo errHandler


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    
    'PHY MDI pins need to be connected to HSD
    thehdw.Utility.Pins("PHY_LPBK_RLY_A").State = tlUtilBitOff
    
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT, REF_CLK, MDI_Pins").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT, REF_CLK, MDI_Pins").StartState = chStartOff

    thehdw.Wait 0.005

    thehdw.PPMU.AllowPPMUFuncRelayConnection (True)

    Call GEPHY_MDIO_Init


    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    V_CM = avdd3p3 / 2
    thehdw.Digital.Pins("MDI_Pins").Levels.Value(chVt) = V_CM


    'Setup
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H1940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 9, &H2000) 'TstMode
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46704, &H0) 'FenTxDacDataB10B1000Dimx
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H1140)
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46807, &H9) 'B1000TxDacLvl
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46762, 57) 'TcPgaGainB1000
'
'    'Zero Scale
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H0) 'TxDacDataB10B1000Dim0
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H0) 'TxDacDataB10B1000Dim1
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46686, &H0) 'TxDacDataDim2
'    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46685, &H0) 'TxDacDataDim3
    
   'Now the TX DAC should be enabled on both dimensions
   'And transmitting a "0" on dimension 1 (and some data on dimension 0)
   'So measurements can be performed on dimension 1

    With thehdw.PPMU.Pins("MDI_Pins")
        .ForceI 0, 2 * mA
        .Gate = tlOn
        .Connect
    End With

    thehdw.Digital.Pins("MDI_Pins").Levels.DriverMode = tlDriverModeVt
    thehdw.Digital.Pins("MDI_Pins").Connect
    thehdw.Wait 0.01

   common_mode_v = thehdw.PPMU.Pins("MDI_Pins").Read

   thehdw.Digital.Pins("MDI_Pins").Disconnect
   thehdw.Wait 0.005

    thehdw.PPMU.Pins("MDI_Pins").ForceV V_CM, 2 * mA
    thehdw.Wait 0.005


    I_CM = thehdw.PPMU.Pins("MDI_Pins").Read

    thehdw.PPMU.Pins("MDI_Pins").ForceI 0 * mA, 200 * uA
    thehdw.Wait 0.005

    V_0mA = thehdw.PPMU.Pins("MDI_Pins").Read

    impedance = V_0mA.Math.Subtract(V_CM).Divide(I_CM).Negate

    With thehdw.PPMU.Pins("MDI_Pins")
        .ForceI 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
    
    Call GEPHY_MDIO_Halt

    TheExec.Flow.TestLimit Resultval:=common_mode_v, unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=impedance, unit:=unitNone, ScaleType:=scaleNone, ForceResults:=tlForceFlow

    thehdw.Digital.Pins("MDI_Pins").Levels.Value(chVt) = 0
    thehdw.Digital.Pins("MDI_Pins").Levels.DriverMode = tlDriverModeLargeHiZ


    thehdw.PPMU.AllowPPMUFuncRelayConnection (False)
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function




Public Function VOD_1000BaseT()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   DAC Measurements for pk-pk in Gigabit mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong

Dim VOD_ZS_1000BaseT As New PinListData
Dim VOD_NFS_1000BaseT As New PinListData
Dim VOD_PFS_1000BaseT As New PinListData

Dim VOD_ZS_Diff As New PinListData
Dim VOD_NFS_Diff As New PinListData
Dim VOD_PFS_Diff As New PinListData

Dim PinName_Pos As String, PinName_Neg As String

    On Error GoTo errHandler



    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A").State = tlUtilBitOff




    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff


    thehdw.Wait 0.01

    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
    
    'Setup
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H1940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 9, &H2000) 'TstMode
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46704, &H0) 'FenTxDacDataB10B1000Dimx
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H1140)
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46807, &H9) 'B1000TxDacLvl
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46762, 57) 'TcPgaGainB1000


    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (True)
    
    'Force 0uA Current to MDI_Pins and measure voltage
    With thehdw.PPMU.Pins("mdi_pins")
        .ForceI 0, 2 * mA
        .Gate = tlOn
        .Connect
    End With
    
    thehdw.Digital.Pins("mdi_pins").Connect
     
    thehdw.Digital.Pins("mdi_pins").InitState = chInitoff
    thehdw.Digital.Pins("mdi_pins").StartState = chStartOff
    thehdw.Wait 0.01
    
    thehdw.Digital.Pins("mdi_pins").Levels.DriverMode = tlDriverModeVt
    thehdw.Digital.Pins("mdi_pins").Levels.Value(chVt) = 1.65 * v
    
    'Zero Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H0) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H0) 'TxDacDataB10B1000Dim1
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46686, &H0) 'TxDacDataDim2
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46685, &H0) 'TxDacDataDim3
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_ZS_1000BaseT = thehdw.PPMU.Pins("mdi_pins").Read
    
    'Negative Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H18) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H18) 'TxDacDataB10B1000Dim1
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46686, &H18) 'TxDacDataDim2
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46685, &H18) 'TxDacDataDim3
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_NFS_1000BaseT = thehdw.PPMU.Pins("mdi_pins").Read
    
    'Positive Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H8) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H8) 'TxDacDataB10B1000Dim1
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46686, &H8) 'TxDacDataDim2
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46685, &H8) 'TxDacDataDim3
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_PFS_1000BaseT = thehdw.PPMU.Pins("mdi_pins").Read
    
    thehdw.Digital.Pins("mdi_pins").Levels.Value(chVt) = 0 * v
    thehdw.Digital.Pins("mdi_pins").Levels.DriverMode = tlDriverModeHiZ
    thehdw.Digital.Pins("mdi_pins").Disconnect
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (False)
    
    'Force 0uA Current and disconnect
    With thehdw.PPMU.Pins("mdi_pins")
        .ForceI 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
    

    
    Call GEPHY_MDIO_Halt
    

    'Compute Diffrential Levels
    For i = 0 To 3
    
        PinName_Pos = "MDI_" & i & "_P"
        PinName_Neg = "MDI_" & i & "_N"
        
        VOD_ZS_Diff.AddPin (PinName_Pos)
        VOD_NFS_Diff.AddPin (PinName_Pos)
        VOD_PFS_Diff.AddPin (PinName_Pos)
        
        For Each site In TheExec.Sites
            VOD_ZS_Diff.Pins(PinName_Pos).Value = VOD_ZS_1000BaseT.Pins(PinName_Pos).Value - VOD_ZS_1000BaseT.Pins(PinName_Neg).Value
            VOD_NFS_Diff.Pins(PinName_Pos).Value = VOD_NFS_1000BaseT.Pins(PinName_Pos).Value - VOD_NFS_1000BaseT.Pins(PinName_Neg).Value
            VOD_PFS_Diff.Pins(PinName_Pos).Value = VOD_PFS_1000BaseT.Pins(PinName_Pos).Value - VOD_PFS_1000BaseT.Pins(PinName_Neg).Value
        Next site
        
    Next i


    
    'Datalog measurements
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_1000BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_1000BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_1000BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow


    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

    
End Function

Public Function VOD_100BaseT()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   DAC Measurements for pk-pk in 100M mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong

Dim VOD_ZS_100BaseT As New PinListData
Dim VOD_NFS_100BaseT As New PinListData
Dim VOD_PFS_100BaseT As New PinListData

Dim VOD_ZS_Diff As New PinListData
Dim VOD_NFS_Diff As New PinListData
Dim VOD_PFS_Diff As New PinListData

Dim PinName_Pos As String, PinName_Neg As String
Dim Meas_Pins As String

    On Error GoTo errHandler


    Meas_Pins = "MDI_0_P, MDI_0_N, MDI_1_P, MDI_1_N"

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    
    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A").State = tlUtilBitOff


'    thehdw.Digital.Pins("MDIO_PULLUP_V").Connect
'    thehdw.Digital.Pins("MDIO_PULLUP_V").Levels.Value(chVih) = 3.3
'    thehdw.Digital.Pins("MDIO_PULLUP_V").InitState = chInitHi
'    thehdw.Digital.Pins("MDIO_PULLUP_V").StartState = chStartHi
    thehdw.Digital.Pins("RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RESET_N, TEST_EN").StartState = chStartLo

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01
    
    thehdw.Digital.Pins("RESET_N").InitState = chInitHi
    thehdw.Digital.Pins("RESET_N").StartState = chStartHi
    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init
    
'    Call mapReset(RM_GEPHY)
'    Call mapReset(RM_GESUB)
    
    
    'Setup
    '// Set active termination on RX dimension (so both dims can measured)
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46725, &H0) 'B100ZptmEnDimrx
    
    '// Set a 100BASE-TX test mode
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46099, &H1) 'B100TxTstMode
    
    '// Disable functional control of required signals
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46708, &H0) 'FenTxB100En
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46786, &H2) 'FenTxVregEnDimxy
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46710, &H2) 'FenTxDacEnDimxy
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46709, &H0) 'FenTxDacDllReset
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46713, &H6) 'FenTxSwAbEn
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46704, &HC) 'FenTxDacDataB10B1000Dimx
    
    '// Since TxB100En is going to be disabled, the B1000TxDacLvl setting
    '// will be used, so the 100BASE-TX default TX DAC level value is written
    '// Other TX DAC settings are the same in 1000BASE-T and 100BASE-TX
    '// so they do not need to be changed
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46087, &H8) 'B1000TxDacLvl
    
    '//loc.GEPhy.MDIOMap.SftPd.Write(1'b0);
    
    '// Disable TxB100En so that TxDacDataB10B100Dim0/1 can be forced
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46692, &H0) 'TxB100En
    
    '// Since TxB100En is disabled the expected functional values need to be set
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46785, &H1) 'TxVregEnDimxy
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46695, &H1) 'TxDacEnDimxy
    
    '// Since TxB100En is disabled we need to de-assert the DLL reset
    '// Note that this must be done after the previous two writes
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46693, &H0) 'TxDacDllReset
    
    '// Enable the DAC to line driver connection on both dimensions
    '// so that measurements can be taken simulataneously on dim0 and dim1
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46698, &H9) 'TxSwAbEn

    


    
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (True)
    
    'Force 0uA Current to MDI_Pins and measure voltage
    With thehdw.PPMU.Pins(Meas_Pins)
        .ForceI 0, 2 * mA
        .Gate = tlOn
        .Connect
    End With
    
    thehdw.Digital.Pins(Meas_Pins).Connect
     
    thehdw.Digital.Pins(Meas_Pins).InitState = chInitoff
    thehdw.Digital.Pins(Meas_Pins).StartState = chStartOff
    thehdw.Wait 0.01
    
    thehdw.Digital.Pins(Meas_Pins).Levels.DriverMode = tlDriverModeVt
    thehdw.Digital.Pins(Meas_Pins).Levels.Value(chVt) = 1.65 * v
    
    'Zero Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H0) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H0) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_ZS_100BaseT = thehdw.PPMU.Pins(Meas_Pins).Read
    
    'Negative Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H18) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H18) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_NFS_100BaseT = thehdw.PPMU.Pins(Meas_Pins).Read
    
    'Positive Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H8) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H8) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_PFS_100BaseT = thehdw.PPMU.Pins(Meas_Pins).Read
    
    thehdw.Digital.Pins("mdi_pins").Levels.Value(chVt) = 0 * v
    thehdw.Digital.Pins("mdi_pins").Levels.DriverMode = tlDriverModeHiZ
    thehdw.Digital.Pins("mdi_pins").Disconnect
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (False)
    
    'Force 0uA Current and disconnect
    With thehdw.PPMU.Pins("mdi_pins")
        .ForceI 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
    
    Call GEPHY_MDIO_Halt
    
    Dim i As Integer
    
    'Datalog Raw measurements
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_100BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_100BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_100BaseT, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    
    For i = 0 To 1
    
        PinName_Pos = "MDI_" & i & "_P"
        PinName_Neg = "MDI_" & i & "_N"
        
        VOD_ZS_Diff.AddPin (PinName_Pos)
        VOD_NFS_Diff.AddPin (PinName_Pos)
        VOD_PFS_Diff.AddPin (PinName_Pos)
        
        For Each site In TheExec.Sites
            VOD_ZS_Diff.Pins(PinName_Pos).Value = VOD_ZS_100BaseT.Pins(PinName_Pos).Value - VOD_ZS_100BaseT.Pins(PinName_Neg).Value
            VOD_NFS_Diff.Pins(PinName_Pos).Value = VOD_NFS_100BaseT.Pins(PinName_Pos).Value - VOD_NFS_100BaseT.Pins(PinName_Neg).Value
            VOD_PFS_Diff.Pins(PinName_Pos).Value = VOD_PFS_100BaseT.Pins(PinName_Pos).Value - VOD_PFS_100BaseT.Pins(PinName_Neg).Value
        Next site
        
    Next i

    'Find difference values and display
    For Each site In TheExec.Sites
        For i = 0 To 1
            VOD_ZS_100BaseT.Pins(i).Value = (VOD_ZS_100BaseT.Pins((i * 2) + 1).Value - VOD_ZS_100BaseT.Pins(i * 2).Value)
            VOD_NFS_100BaseT.Pins(i).Value = (VOD_NFS_100BaseT.Pins((i * 2) + 1).Value - VOD_NFS_100BaseT.Pins(i * 2).Value)
            VOD_PFS_100BaseT.Pins(i).Value = (VOD_PFS_100BaseT.Pins((i * 2) + 1).Value - VOD_PFS_100BaseT.Pins(i * 2).Value)
        Next i
    Next site
    
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow

'    'Datalog measurements
'    For Each site In TheExec.Sites
'        For i = 0 To 1
'            TheExec.Flow.TestLimit resultVal:=VOD_ZS_100BaseT.Pins(i).Value, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
'        Next i
'        For i = 0 To 1
'            TheExec.Flow.TestLimit resultVal:=VOD_NFS_100BaseT.Pins(i).Value, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
'        Next i
'        For i = 0 To 1
'            TheExec.Flow.TestLimit resultVal:=VOD_PFS_100BaseT.Pins(i).Value, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
'        Next i
'    Next site

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function VOD_10BaseT()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   DAC Measurements for pk-pk in 10M mode
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Read As New SiteLong

Dim VOD_ZS_10BaseT_Dim0 As New PinListData
Dim VOD_NFS_10BaseT_Dim0 As New PinListData
Dim VOD_PFS_10BaseT_Dim0 As New PinListData

Dim VOD_ZS_10BaseT_Dim1 As New PinListData
Dim VOD_NFS_10BaseT_Dim1 As New PinListData
Dim VOD_PFS_10BaseT_Dim1 As New PinListData

Dim VOD_ZS_10BaseT As New PinListData
Dim VOD_NFS_10BaseT As New PinListData
Dim VOD_PFS_10BaseT As New PinListData

Dim VOD_ZS_Diff As New PinListData
Dim VOD_NFS_Diff As New PinListData
Dim VOD_PFS_Diff As New PinListData

Dim PinName_Pos As String, PinName_Neg As String

Dim mdi_10b_pins As String
Dim mdi_dim_0 As String, mdi_dim_1 As String

Dim i As Integer

    On Error GoTo errHandler

    mdi_10b_pins = "MDI_0_N, MDI_0_P, MDI_1_N, MDI_1_P"
    mdi_dim_0 = "MDI_0_N, MDI_0_P"
    mdi_dim_1 = "MDI_1_N, MDI_1_P"


    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn
    
    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A").State = tlUtilBitOff


    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01
    

    Call GEPHY_MDIO_Init
    

    '// Configure PHY for 10BASE-Te VOD measurements
    
    '// Set a 10BASE-T dim0 test mode for DIM0 VOD measurements
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46098, &H1) 'B10TxTstMode
    
    '// Disable functional control of DAC data signals
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46704, &HC) 'FenTxDacDataB10B1000Dimx

    
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (True)
    
    'Force 0uA Current to MDI_Pins and measure voltage
    With thehdw.PPMU.Pins(mdi_10b_pins)
        .ForceI 0, 2 * mA
        .Gate = tlOn
        .Connect
    End With
    
    thehdw.Digital.Pins(mdi_10b_pins).Connect
     
    thehdw.Digital.Pins(mdi_10b_pins).InitState = chInitoff
    thehdw.Digital.Pins(mdi_10b_pins).StartState = chStartOff
    thehdw.Wait 0.01
    
    thehdw.Digital.Pins(mdi_10b_pins).Levels.DriverMode = tlDriverModeVt
    thehdw.Digital.Pins(mdi_10b_pins).Levels.Value(chVt) = 1.65 * v
    
    'Zero Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H0) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H0) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_ZS_10BaseT_Dim0 = thehdw.PPMU.Pins(mdi_dim_0).Read
    
    'Negative Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H18) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H18) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_NFS_10BaseT_Dim0 = thehdw.PPMU.Pins(mdi_dim_0).Read
    
    'Positive Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H8) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H8) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_PFS_10BaseT_Dim0 = thehdw.PPMU.Pins(mdi_dim_0).Read
    
    '// Set a 10BASE-T dim1 test mode  for DIM1 VOD measurements
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46098, &H2) 'B10TxTstMode
    
    thehdw.Wait 10 * ms
    
    'Zero Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H0) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H0) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_ZS_10BaseT_Dim1 = thehdw.PPMU.Pins(mdi_dim_1).Read
        
    'Positive Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H8) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H8) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_PFS_10BaseT_Dim1 = thehdw.PPMU.Pins(mdi_dim_1).Read
    
    'Negative Full Scale
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46688, &H18) 'TxDacDataB10B1000Dim0
    Call GEPhy_MDIO_CL_45_Write(&H0, 30, 46687, &H18) 'TxDacDataB10B1000Dim1
    
    'Measure
    thehdw.Wait 5 * ms
    VOD_NFS_10BaseT_Dim1 = thehdw.PPMU.Pins(mdi_dim_1).Read
    
    thehdw.Digital.Pins("mdi_pins").Levels.Value(chVt) = 0 * v
    thehdw.Digital.Pins("mdi_pins").Levels.DriverMode = tlDriverModeHiZ
    thehdw.Digital.Pins("mdi_pins").Disconnect
    
    thehdw.PPMU.AllowPPMUFuncRelayConnection (False)
    
    'Force 0uA Current and disconnect
    With thehdw.PPMU.Pins(mdi_10b_pins)
        .ForceI 0, 2 * mA
        .Gate = tlOff
        .Disconnect
    End With
    
    Call GEPHY_MDIO_Halt
    
    'Datalog Raw measurements
    
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_10BaseT_Dim0, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_10BaseT_Dim1, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_10BaseT_Dim0, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_10BaseT_Dim1, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_10BaseT_Dim0, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_10BaseT_Dim1, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    
    For i = 0 To 1
    
        PinName_Pos = "MDI_" & i & "_P"
        PinName_Neg = "MDI_" & i & "_N"
        
        VOD_ZS_Diff.AddPin (PinName_Pos)
        VOD_NFS_Diff.AddPin (PinName_Pos)
        VOD_PFS_Diff.AddPin (PinName_Pos)
        
        If i = 0 Then
        
            For Each site In TheExec.Sites
                VOD_ZS_Diff.Pins(PinName_Pos).Value = VOD_ZS_10BaseT_Dim0.Pins(PinName_Pos).Value - VOD_ZS_10BaseT_Dim0.Pins(PinName_Neg).Value
                VOD_NFS_Diff.Pins(PinName_Pos).Value = VOD_NFS_10BaseT_Dim0.Pins(PinName_Pos).Value - VOD_NFS_10BaseT_Dim0.Pins(PinName_Neg).Value
                VOD_PFS_Diff.Pins(PinName_Pos).Value = VOD_PFS_10BaseT_Dim0.Pins(PinName_Pos).Value - VOD_PFS_10BaseT_Dim0.Pins(PinName_Neg).Value
            Next site
        
        Else
        
            For Each site In TheExec.Sites
                VOD_ZS_Diff.Pins(PinName_Pos).Value = VOD_ZS_10BaseT_Dim1.Pins(PinName_Pos).Value - VOD_ZS_10BaseT_Dim1.Pins(PinName_Neg).Value
                VOD_NFS_Diff.Pins(PinName_Pos).Value = VOD_NFS_10BaseT_Dim1.Pins(PinName_Pos).Value - VOD_NFS_10BaseT_Dim1.Pins(PinName_Neg).Value
                VOD_PFS_Diff.Pins(PinName_Pos).Value = VOD_PFS_10BaseT_Dim1.Pins(PinName_Pos).Value - VOD_PFS_10BaseT_Dim1.Pins(PinName_Neg).Value
            Next site
        
        End If
        
    Next i


    'Datalog measurements
    TheExec.Flow.TestLimit Resultval:=VOD_ZS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_NFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit Resultval:=VOD_PFS_Diff, unit:=unitVolt, ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow


    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

    
End Function


Public Function ADC_DC_Offset()
Dim Read As New SiteLong

    Dim Iadc_signals As Long
    Dim Iadc_signals_end As Long
    Dim code_incr As Long
    Dim array_count As Long
    Dim dim_seperate As Long
    Dim dim_sep_trim As Long
    Dim slDgOutLsb1(0 To 3) As New SiteLong
    Dim slDgOutLsb2(0 To 3) As New SiteLong
    Dim slDgOutMid(0 To 3) As New SiteLong
    Dim slDgOutMsb(0 To 3) As New SiteLong

    Dim slall_adc_offs(0 To 1000) As New SiteLong

    Dim sdold_adc_offs_final(0 To 256) As New SiteDouble
    Dim sdnew_adc_offs_final(0 To 256) As New SiteDouble
    
    Dim sladc_old_dc_offs(0 To 3) As New SiteLong
    Dim sladc_new_dc_offs(0 To 3) As New SiteLong
    
    Dim sladc_offs_default(0 To 8) As New SiteLong
    
    Dim sdWindow_Length As New SiteDouble
    
On Error GoTo errHandler
    
    Dim sdadc_old_dim0_dc_offs_dsp As New DSPWave
    Dim sdadc_new_dim0_dc_offs_dsp As New DSPWave
    Dim sdadc_old_dim1_dc_offs_dsp As New DSPWave
    Dim sdadc_new_dim1_dc_offs_dsp As New DSPWave
    Dim sdadc_old_dim2_dc_offs_dsp As New DSPWave
    Dim sdadc_new_dim2_dc_offs_dsp As New DSPWave
    Dim sdadc_old_dim3_dc_offs_dsp As New DSPWave
    Dim sdadc_new_dim3_dc_offs_dsp As New DSPWave

    Call sdadc_old_dim0_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_new_dim0_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_old_dim1_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_new_dim1_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_old_dim2_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_new_dim2_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_old_dim3_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    Call sdadc_new_dim3_dc_offs_dsp.CreateConstant(0, 64, DspDouble)
    
    ChanMap = TheExec.CurrentChanMap
    package_type = Mid(ChanMap, 9, 5)
    
    If package_type = "pkg32" Then
        Iadc_signals_end = 1
        dim_sep_trim = 2
    Else
        Iadc_signals_end = 3
        dim_sep_trim = 4
    End If

     'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff
    
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi

    thehdw.Digital.Pins("XTAL_O, REXT").InitState = chInitoff
    thehdw.Digital.Pins("XTAL_O, REXT").StartState = chStartOff
    thehdw.Wait 0.01
    
'    thehdw.Digital.Pins("RESET_N").InitState = chInitHi
'    thehdw.Digital.Pins("RESET_N").StartState = chStartHi
'    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init

    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)

    'Offset readback
    'Release reset and speed-up PLL lock
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    bfWrite BF_GEPHY.LinkEn, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim0Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim1Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim2Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim3Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxEnDim, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacEnDim01Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacEnDim23Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDrvEnDim01Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDrvEnDim23Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxSw00EnEmi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxSw01EnEmi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxSw10EnEmi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxSw11EnEmi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxVregEnDim01Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxVregEnDim23Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacDataB10B1000Dim0Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacDataB10B1000Dim1Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacDataDim2Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenTxDacDataDim3Emi, MSV(&H0)
    '*************************************************
    bfWrite BF_GEPHY.SftPd, MSV(&H0)

    '*************************************************
    'Wait for Phy in Stdby
    thehdw.Wait 100 * ms
    Read = bfRead(BF_GEPHY.PhyInStndby)

    'Logger.Info ("Configure AFE for 1000BASE-T mode, transmitting 0")
    bfWrite BF_GEPHY.TxVregEnDim01, MSV(&H1)
    bfWrite BF_GEPHY.TxVregEnDim23, MSV(&H1)
    bfWrite BF_GEPHY.TxSw00En, MSV(&H1)
    bfWrite BF_GEPHY.TxSw01En, MSV(&H0)
    bfWrite BF_GEPHY.TxSw10En, MSV(&H0)
    bfWrite BF_GEPHY.TxSw11En, MSV(&H1)
    bfWrite BF_GEPHY.TxDacEnDim01, MSV(&H1)
    bfWrite BF_GEPHY.TxDacEnDim23, MSV(&H1)
    bfWrite BF_GEPHY.TxDacDataB10B1000Dim0, MSV(&H0)
    bfWrite BF_GEPHY.TxDacDataB10B1000Dim1, MSV(&H0)
    bfWrite BF_GEPHY.TxDacDataDim2, MSV(&H0)
    bfWrite BF_GEPHY.TxDacDataDim3, MSV(&H0)
    bfWrite BF_GEPHY.TxDrvEnDim01, MSV(&H1)
    bfWrite BF_GEPHY.TxDrvEnDim23, MSV(&H1)
    bfWrite BF_GEPHY.RxEnDim, MSV(&HF)
    thehdw.Wait 20 * ms
    
    sladc_offs_default(0) = bfRead(BF_GEPHY.DcOffOldDim0)
    sladc_offs_default(1) = bfRead(BF_GEPHY.DcOffNewDim0)
    sladc_offs_default(2) = bfRead(BF_GEPHY.DcOffOldDim1)
    sladc_offs_default(3) = bfRead(BF_GEPHY.DcOffNewDim1)
    sladc_offs_default(4) = bfRead(BF_GEPHY.DcOffOldDim2)
    sladc_offs_default(5) = bfRead(BF_GEPHY.DcOffNewDim2)
    sladc_offs_default(6) = bfRead(BF_GEPHY.DcOffOldDim3)
    sladc_offs_default(7) = bfRead(BF_GEPHY.DcOffNewDim3)
    
    For Each site In TheExec.Sites
        For i = 0 To 7
            If (sladc_offs_default(i) < 15) Then
                sladc_offs_default(i) = sladc_offs_default(i)
            Else
                sladc_offs_default(i) = sladc_offs_default(i) - 32
            End If
        Next i
    Next site
    
    'Reset Configs
    bfWrite BF_GEPHY.SigAnlCfgRst, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlMathEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlMathWinLenm1, MSV(65535)
    bfWrite BF_GEPHY.SigAnlMathSig1En, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlMathMltEn, MSV(&H0)
    '*************************************************

    '******************* Read values *********************************************
    'Performing sweep across PGA gain
    
    sdWindow_Length = bfRead(BF_GEPHY.SigAnlSmpHldWinLenm1)

    array_count = 0

    For Each site In TheExec.Sites
        For code_incr = 0 To 63
            Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &HB63B, code_incr) 'RxPgaGainDim0
            Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &HB63A, code_incr) 'RxPgaGainDim1
            Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &HB639, code_incr) 'RxPgaGainDim2
            Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &HB638, code_incr) 'RxPgaGainDim3

                For Iadc_signals = 0 To Iadc_signals_end
                    Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &H9001, Iadc_signals + 4) 'SigAnlSel0
                    Call GEPhy_MDIO_CL_45_Write(&H0, &H1E, &H9002, Iadc_signals + 8) 'SigAnlSel1
                    thehdw.Wait 1 * ms

                    slDgOutLsb1(Iadc_signals) = bfRead(BF_GEPHY.SigAnlMathMeanAccLo)
                    slDgOutMsb(Iadc_signals) = bfRead(BF_GEPHY.SigAnlMathMeanAccHi)
                    
                    slDgOutLsb2(Iadc_signals) = bfRead(BF_GEPHY.SigAnlMathProdAccLo)
                    slDgOutMid(Iadc_signals) = bfRead(BF_GEPHY.SigAnlMathProdAccMid)
                    
                    sladc_old_dc_offs(Iadc_signals) = slDgOutMsb(Iadc_signals).BitwiseAnd(&H7FFF).ShiftLeft(16).Add(slDgOutLsb1(Iadc_signals))
                    sladc_old_dc_offs(Iadc_signals) = sladc_old_dc_offs(Iadc_signals).Subtract(sladc_old_dc_offs(Iadc_signals).BitwiseAnd(&H40000000).ShiftLeft(1))
                                        
                    sladc_new_dc_offs(Iadc_signals) = slDgOutMid(Iadc_signals).ShiftLeft(15).Add(slDgOutLsb2(Iadc_signals).BitwiseAnd(&H7FFF))
                    sladc_new_dc_offs(Iadc_signals) = sladc_new_dc_offs(Iadc_signals).Subtract(sladc_new_dc_offs(Iadc_signals).BitwiseAnd(&H40000000).ShiftLeft(1))
                    
                    sdold_adc_offs_final(array_count) = sladc_old_dc_offs(Iadc_signals).Divide(256).Divide(sdWindow_Length).Multiply(128)
                    sdnew_adc_offs_final(array_count) = sladc_new_dc_offs(Iadc_signals).Divide(256).Divide(sdWindow_Length).Multiply(128)
                    
                    array_count = array_count + 1
                Next Iadc_signals
        Next code_incr
    Next site
    
    For Each local_site In TheExec.Sites.Active
        For dim_seperate = 0 To 63

            sdadc_old_dim0_dc_offs_dsp(local_site).Element(dim_seperate) = sdold_adc_offs_final((dim_sep_trim * dim_seperate) + 0)
            sdadc_new_dim0_dc_offs_dsp(local_site).Element(dim_seperate) = sdnew_adc_offs_final((dim_sep_trim * dim_seperate) + 0)
            sdadc_old_dim1_dc_offs_dsp(local_site).Element(dim_seperate) = sdold_adc_offs_final((dim_sep_trim * dim_seperate) + 1)
            sdadc_new_dim1_dc_offs_dsp(local_site).Element(dim_seperate) = sdnew_adc_offs_final((dim_sep_trim * dim_seperate) + 1)
            
            If package_type <> "pkg32" Then
                sdadc_old_dim2_dc_offs_dsp(local_site).Element(dim_seperate) = sdold_adc_offs_final((dim_sep_trim * dim_seperate) + 2)
                sdadc_new_dim2_dc_offs_dsp(local_site).Element(dim_seperate) = sdnew_adc_offs_final((dim_sep_trim * dim_seperate) + 2)
                sdadc_old_dim3_dc_offs_dsp(local_site).Element(dim_seperate) = sdold_adc_offs_final((dim_sep_trim * dim_seperate) + 3)
                sdadc_new_dim3_dc_offs_dsp(local_site).Element(dim_seperate) = sdnew_adc_offs_final((dim_sep_trim * dim_seperate) + 3)
            End If

        Next dim_seperate
    Next local_site
    
Call GEPHY_MDIO_Halt

    'Datalogging
    For i = 0 To 3
        TheExec.Flow.TestLimit Resultval:=sladc_offs_default(i), ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
    Next i
    
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim0_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim0_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim0_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim0_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim0_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim0_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)

    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim1_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim1_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim1_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim1_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim1_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim1_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)

    If package_type <> "pkg32" Then
        
        For i = 4 To 7
            TheExec.Flow.TestLimit Resultval:=sladc_offs_default(i), ScaleType:=scaleNoScaling, ForceResults:=tlForceFlow
        Next i
        
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim2_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim2_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim2_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim2_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim2_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim2_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
    
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim3_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim3_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_old_dim3_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim3_dc_offs_dsp.CalcMinimumValue, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim3_dc_offs_dsp.CalcMean, ForceResults:=tlForceFlow)
        Call TheExec.Flow.TestLimit(Resultval:=sdadc_new_dim3_dc_offs_dsp.CalcMaximumValue, ForceResults:=tlForceFlow)
    End If
    
    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Phase_Sel()
Dim Read As New SiteLong
    
    Dim Iadc_signals As Long
    Dim code_incr As Long
    Dim array_count As Long
    Dim dim_seperate As Long
    
On Error GoTo errHandler

    'Turn on REXT resistor
    thehdw.Utility.Pins("REXT_RLY").State = tlUtilBitOn

    'PHY and MAC in LPBK Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOn
    thehdw.Utility.Pins("PHY_LPBK_RLY_B").State = tlUtilBitOff
    
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").InitState = chInitLo
    thehdw.Digital.Pins("RX_Pins, RESET_N, TEST_EN").StartState = chStartLo
    thehdw.Digital.Pins("LED_Pins, RXD_5").InitState = chInitHi
    thehdw.Digital.Pins("LED_Pins, RXD_5").StartState = chStartHi
    thehdw.Wait 0.01
    
'''    thehdw.Digital.Pins("RESET_N").InitState = chInitHi
'''    thehdw.Digital.Pins("RESET_N").StartState = chStartHi
'''    thehdw.Wait 0.01


    Call GEPHY_MDIO_Init

    'Call mapReset(RM_GEPHY)
    'Call mapReset(RM_GESUB)

    'Run cable diagnostics
    bfWrite BF_GEPHY.LinkEn, MSV(&H0)
    
    'Wait for Phy in Stdby
    thehdw.Wait 100 * ms
    Read = bfRead(BF_GEPHY.PhyInStndby)
    
    bfWrite BF_GEPHY.CdiagRun, MSV(&H1)
    
    thehdw.Wait 100 * ms
    Read = bfRead(BF_GEPHY.CdiagRun)
    
    'Read cable diagnostics
    thehdw.Wait 100 * ms
    Read = bfRead(BF_GEPHY.CdiagRslt0Xshrt1)
    Read = bfRead(BF_GEPHY.CdiagRslt1Xshrt0)
    
    Read = bfRead(BF_GEPHY.CdiagFltDist0)
    Read = bfRead(BF_GEPHY.CdiagFltDist1)
    
    'Configuring PHY for 100BASE-TX T/X test mode (tx_dim_rsp=1)
    'rx_dim_rsp = 1+((tx_dim_rsp>>1)<<1)-(tx_dim_rsp&1)
    'rx_dim_rsp=0
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.B100TxTstMode, MSV(&H4) '4
    bfWrite BF_GEPHY.FenRxEnDim, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    thehdw.Wait 1 * ms
    bfWrite BF_GEPHY.RxEnDim, MSV(&H1)
    
    Dim hyb_rx_gain, pga_gain_min, pga_gain_max, hyb_rx_gain_db, pga_gain_min_db, pga_gain_max_db, rx_gains(0 To 63), coomp_val As Double
    
    Dim y, trgt_adc_max_abs, rx_gains_comp(0 To 63) As Integer
    Dim pga_lvl As Long
    Dim len_b100_tst_pat As Long
    
    len_b100_tst_pat = 56
    
    trgt_adc_max_abs = 124
    
    comp_val = ((0.5 / 1.75) * (trgt_adc_max_abs / 128))
    
    hyb_rx_gain_db = -0.67
    pga_gain_min_db = -2.4
    pga_gain_max_db = -17.6
    
    'Set initial PGA level assuming 1.75V peak at MDI
    hyb_rx_gain = 10 ^ (hyb_rx_gain_db / 20)
    pga_gain_min = 10 ^ (pga_gain_min_db / 20)
    pga_gain_max = 10 ^ (pga_gain_max_db / 20)
    
    pga_lvl = 0
    
    '*****************
    For y = 0 To 63
    
        rx_gains(y) = hyb_rx_gain * (pga_gain_min * pga_gain_max) / (pga_gain_min + (y / 63) * (pga_gain_max - pga_gain_min))
        
        
    Next y
    
     For y = 0 To 63
        
        If (rx_gains(y) <= comp_val) Then
            rx_gains_comp(y) = 1
        Else
            rx_gains_comp(y) = 0
            pga_lvl = y
            y = 64
        End If
        
    Next y
    'pga_lvl = numpy.max(numpy.nonzero(rx_gains<=(0.5/1.75)*float(trgt_adc_max_abs)/128)[0].astype(int))
    '******************
    
    bfWrite BF_GEPHY.FenRxPgaGainDim0Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim1Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim2Emi, MSV(&H0)
    bfWrite BF_GEPHY.FenRxPgaGainDim3Emi, MSV(&H0)
    
    bfWrite BF_GEPHY.RxPgaGainDim0, MSV(pga_lvl)
    thehdw.Wait 1 * ms
    
    'Configuring signal analysis to measure maximum absolute ADC value
    'Reset Configs
    bfWrite BF_GEPHY.SigAnlCfgRst, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlSel0, MSV(&H8)
    bfWrite BF_GEPHY.SigAnlSel1, MSV(&H4)
    bfWrite BF_GEPHY.SigAnlStatsContEn, MSV(&H0)
    bfWrite BF_GEPHY.SigAnlStatsWinLenm1, MSV(len_b100_tst_pat * 256)
    
    Read = bfRead(BF_GEPHY.PllPhase)
    
    'Debug.Print "PllPhase_Read:"; Read(0)
    
    Dim slStatsMax0(0 To 127) As New SiteLong
    Dim slStatsMax1(0 To 127) As New SiteLong
    
    Dim slStatsMax(0 To 127) As New SiteLong
    
    Dim abs_max_adc_dsp As New DSPWave
    
    Dim index As Long
    Dim index1 As Long
    Dim num_dly_dec As Long

    Call abs_max_adc_dsp.CreateConstant(0, 128, DspDouble)
    
    Dim abs_max_value As Long
    Dim max_abs_rx_vltg As Double
    
    
    For Each site In TheExec.Sites
        For code_incr = 0 To 127

            'Read = bfRead(BF_GEPHY.SigAnlStatsRdy)
            
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36887) 'SigAnlStatsRdyMax1
            slStatsMax1(code_incr) = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlStatsMax1)
            slStatsMax1(code_incr) = (slStatsMax1(code_incr) - (slStatsMax1(code_incr).BitwiseAnd(&H4000).ShiftLeft(1)))
            'Debug.Print "SigAnlStatsRdyNoClr:"; Read.BitwiseAnd(&H8000).Subtract(32767)
            
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36885) 'SigAnlStatsRdyMax1
            thehdw.Wait 1 * ms
            slStatsMax0(code_incr) = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlStatsMax0)
            slStatsMax0(code_incr) = (slStatsMax0(code_incr) - (slStatsMax0(code_incr).BitwiseAnd(&H4000).ShiftLeft(1)))
            'Debug.Print "SigAnlStatsRdy:"; Read.BitwiseAnd(&H8000).Subtract(32767)
            
            If (slStatsMax0(code_incr) > slStatsMax1(code_incr)) Then
                slStatsMax(code_incr) = slStatsMax0(code_incr)
            Else
                slStatsMax(code_incr) = slStatsMax1(code_incr)
            End If
            
            slStatsMax(code_incr) = slStatsMax(code_incr).Multiply(128).Divide(256)
            
            'Debug.Print "Maximum absolute ADC value:"; slStatsMax(code_incr)
            
            bfWrite BF_GEPHY.PllDlyInc, MSV(&H1)
            
            abs_max_adc_dsp(site).Element(code_incr) = slStatsMax(code_incr)
            
        Next code_incr
    Next site
    
    abs_max_value = abs_max_adc_dsp.CalcMaximumValue
    
    Call abs_max_adc_dsp.CalcMaximumValue(index)
    
    For Each site In TheExec.Sites
        num_dly_dec = 127 - index
            For y = 0 To num_dly_dec
                bfWrite BF_GEPHY.PllDlyDec, MSV(&H1)
            Next y
    Next site
    
    Read = bfRead(BF_GEPHY.PllPhase)
    
    'Debug.Print "PllDly_Meas:"; index
    'Debug.Print "PllPhase_Read:"; Read(0)
    
    If rx_gains(pga_lvl) = 0 Then
        rx_gains(pga_lvl) = 999
    End If
    
    max_abs_rx_vltg = (0.5 * abs_max_value) / (128 * rx_gains(pga_lvl))
    
    Dim comp_val_new As Double
    
    If max_abs_rx_vltg = 0 Then
        max_abs_rx_vltg = 999
    End If
    
    comp_val_new = (0.5 * trgt_adc_max_abs) / (128 * max_abs_rx_vltg)
    
    pga_lvl = 0
    
    For y = 0 To 63
        
        If (rx_gains(y) <= comp_val_new) Then
            rx_gains_comp(y) = 1
        Else
            rx_gains_comp(y) = 0
            pga_lvl = y
            y = 64
        End If
        
    Next y
    
    'Write updated PGA value
    bfWrite BF_GEPHY.RxPgaGainDim0, MSV(pga_lvl)
    thehdw.Wait 1 * ms
    
    Dim pga_lvl_incr As Long
    
    'Read back ADC max value now agter pga_lvl change
    Dim slStatsMax_Final As New SiteLong
    Dim slStatsMax0_Final As New SiteLong
    Dim slStatsMax1_Final As New SiteLong
    
    For Each site In TheExec.Sites
        For pga_lvl_incr = pga_lvl To 63
        
            'Write updated PGA value
            bfWrite BF_GEPHY.RxPgaGainDim0, MSV(pga_lvl_incr)
            thehdw.Wait 1 * ms
            
            Read = bfRead(BF_GEPHY.SigAnlStatsRdy)
            
            thehdw.Wait 1 * ms
            
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36887) 'SigAnlStatsRdyMax1
            slStatsMax1_Final = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlStatsMax1)
            slStatsMax1_Final = (slStatsMax1_Final - (slStatsMax1_Final.BitwiseAnd(&H4000).ShiftLeft(1)))
            'Debug.Print "SigAnlStatsRdyNoClr:"; Read.BitwiseAnd(&H8000).Subtract(32767)
            
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36885) 'SigAnlStatsRdyMax1
            thehdw.Wait 1 * ms
            slStatsMax0_Final = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlStatsMax0)
            slStatsMax0_Final = (slStatsMax0_Final - (slStatsMax0_Final.BitwiseAnd(&H4000).ShiftLeft(1)))
            'Debug.Print "SigAnlStatsRdy:"; Read.BitwiseAnd(&H8000).Subtract(32767)
            
            'Debug.Print "pga_lvl_incr:"; pga_lvl_incr
        
            'slStatsMax0_Final = bfRead(BF_GEPHY.SigAnlStatsMax0)
            'slStatsMax0_Final = (slStatsMax0_Final - (slStatsMax0_Final.BitwiseAnd(&H4000).ShiftLeft(1)))
            
            'slStatsMax1_Final = bfRead(BF_GEPHY.SigAnlStatsMax0)
            'slStatsMax1_Final = (slStatsMax1_Final - (slStatsMax1_Final.BitwiseAnd(&H4000).ShiftLeft(1)))
            
            If (slStatsMax0_Final > slStatsMax1_Final) Then
                slStatsMax_Final = slStatsMax0_Final
            Else
                slStatsMax_Final = slStatsMax1_Final
            End If
            
            slStatsMax_Final = slStatsMax_Final.Multiply(128).Divide(256)
            
            'Debug.Print "Maximum absolute ADC value:"; slStatsMax_Final
            
            If (slStatsMax_Final < trgt_adc_max_abs) Then
                pga_lvl = pga_lvl_incr
            Else
                pga_lvl_incr = 64
            End If
            
        Next pga_lvl_incr
    Next site
    
    'Write final value
    bfWrite BF_GEPHY.RxPgaGainDim0, MSV(pga_lvl - 1) '-1
    thehdw.Wait 1 * ms
    
    'Return PllDly to Zero and check PllPhase
    For Each site In TheExec.Sites
        num_dly_dec = index - 1
            For y = 0 To num_dly_dec
                bfWrite BF_GEPHY.PllDlyDec, MSV(&H1)
            Next y
    Next site
    
    thehdw.Wait 1 * ms
    Read = bfRead(BF_GEPHY.PllPhase)
    
    'Debug.Print "PllPhase_Read:"; Read(0)
    
    'Acquire test pattern response
    'Reset Configs
    bfWrite BF_GEPHY.SigAnlCfgRst, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlSel0, MSV(&H8)
    bfWrite BF_GEPHY.SigAnlStatsAbsEn, MSV(&H0)
    bfWrite BF_GEPHY.SigAnlStatsContEn, MSV(&H0)
    bfWrite BF_GEPHY.SigAnlStatsWinLenm1, MSV(len_b100_tst_pat * 64)
    
    Dim abs_max_adc_new_dsp As New DSPWave
    
    Dim index_new_adc As Long
    Dim abs_max_value_new As Long

    Call abs_max_adc_new_dsp.CreateConstant(0, 128, DspDouble)
    
    Dim slStatsMax0_new(0 To 127) As New SiteLong
    Dim slStatsMax1_new(0 To 127) As New SiteLong
    
    Dim slStatsMax_new(0 To 127) As New SiteLong
    
    For Each site In TheExec.Sites
        For code_incr = 0 To 127

            thehdw.Wait 1 * ms
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36885) 'SigAnlStatsRdyMax1
            slStatsMax_new(code_incr) = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlStatsMax0)
            slStatsMax_new(code_incr) = (slStatsMax_new(code_incr) - (slStatsMax_new(code_incr).BitwiseAnd(&H4000).ShiftLeft(1)))
            slStatsMax_new(code_incr) = slStatsMax_new(code_incr).Multiply(128).Divide(256)
            
            'Debug.Print "Maximum absolute ADC value:"; slStatsMax_new(code_incr)
            
            bfWrite BF_GEPHY.PllDlyInc, MSV(&H1)
            
            abs_max_adc_new_dsp(site).Element(code_incr) = slStatsMax_new(code_incr)
            
        Next code_incr
    Next site
    
    abs_max_value_new = abs_max_adc_new_dsp.CalcMaximumValue
    
    Call abs_max_adc_new_dsp.CalcMaximumValue(index_new_adc)
    
    'Debug.Print "Maximum absolute ADC value Final:"; abs_max_value_new
    'Debug.Print "PllDly_Meas:"; index_new_adc
    
    For Each site In TheExec.Sites
        num_dly_dec = 127 - index_new_adc
            For y = 0 To num_dly_dec
                bfWrite BF_GEPHY.PllDlyDec, MSV(&H1)
            Next y
    Next site
    
    Read = bfRead(BF_GEPHY.PllPhase)
    
    'Debug.Print "PllDly_Meas:"; index_new_adc
    'Debug.Print "PllPhase_Read:"; Read(0)
    
    'Configuring signal analysis to capture ADC signals (enable SampleHold SigAnl); Capturing raw new ADC signal
    bfWrite BF_GEPHY.SigAnlCfgRst, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlSel0, MSV(&H8)
    bfWrite BF_GEPHY.SigAnlSel1, MSV(&H4)
    bfWrite BF_GEPHY.SigAnlSmpHldEn, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlSmpHldWinLenm1, MSV(len_b100_tst_pat - 1)
    
    'Return PllDly to Zero and check PllPhase
'    For Each site In TheExec.Sites
'        num_dly_dec = index_new_adc - 1
'            For y = 0 To num_dly_dec
'                bfWrite BF_GEPHY.PllDlyDec, MSV(&H1)
'            Next y
'    Next site
'
'    TheHdw.Wait 1 * mS
'    Read = bfRead(BF_GEPHY.PllPhase)
'
'    Debug.Print "PllPhase_Read:"; Read(0)
    
    Dim raw_adc_new_dsp As New DSPWave

    Call raw_adc_new_dsp.CreateConstant(0, 56, DspDouble)
    
    Dim slRaw_ADC(0 To 127) As New SiteLong
    Dim sdRaw_ADC(0 To 127) As New SiteDouble
    
    Dim index_max1 As Long
    Dim index_max2 As Long
    
    Dim index_min1 As Long
    Dim index_min2 As Long
    
    For Each site In TheExec.Sites
        For code_incr = 0 To len_b100_tst_pat - 1

            'Read = bfRead(BF_GEPHY.SigAnlSmpHldIndx)
            'Debug.Print "SigAnlSmpHldIndx  :"; Read

            'Read = bfRead(BF_GEPHY.SigAnlSmpHldRdy)
            'TheHdw.Wait 1 * mS
            Read = GEPhy_MDIO_CL_45_Read(0, 30, 36875) 'SigAnlSmpHldRdyVal0

            'slRaw_ADC(code_incr) = bfRead(BF_GEPHY.SigAnlSmpHldVal0)

            slRaw_ADC(code_incr) = Read.BitwiseAnd(&H7FFF)
            slRaw_ADC(code_incr) = (slRaw_ADC(code_incr) - (slRaw_ADC(code_incr).BitwiseAnd(&H4000).ShiftLeft(1)))
            'sdRaw_ADC(code_incr) = slRaw_ADC(code_incr).Divide(256)

            'Debug.Print "SigAnlSmpHldRdy:"; Read.BitwiseAnd(&H8000).Subtract(32767)
            'Debug.Print "Raw ADC value:"; sdRaw_ADC(code_incr)

            raw_adc_new_dsp(site).Element(code_incr) = slRaw_ADC(code_incr)

        Next code_incr
    Next site

    raw_adc_new_dsp = raw_adc_new_dsp.Divide(256)
    
    ''''''''''raw_adc_new_dsp.Plot "Raw New ADC Signal"
    
    thehdw.Wait 1 * ms
    
    Read = bfRead(BF_GEPHY.SigAnlSmpHldIndx)
    'Debug.Print "SigAnlSmpHldIndx  :"; Read(0)

    Call raw_adc_new_dsp.CalcMaximumValue(index_max1)
    index_max2 = index_max1 - 1 '1

    Call raw_adc_new_dsp.CalcMinimumValue(index_min1)
    index_min2 = index_min1 - 1 '1

    'Debug.Print "Maximum index 1:"; index_max1
    'Debug.Print "Maximum index 2:"; index_max2

    'Debug.Print "Minimum index 1:"; index_min1
    'Debug.Print "Minimum index 2:"; index_min2

    Dim new_adc_peak_array(0 To 3) As Long

    new_adc_peak_array(0) = index_max2
    new_adc_peak_array(1) = index_max1
    new_adc_peak_array(2) = index_min2
    new_adc_peak_array(3) = index_min1

    'Debug.Print "Array Point 0:"; new_adc_peak_array(0)
    'Debug.Print "Array Point 1:"; new_adc_peak_array(1)
    'Debug.Print "Array Point 2:"; new_adc_peak_array(2)
    'Debug.Print "Array Point 3:"; new_adc_peak_array(3)

    'Configuring signal analysis to measure averaged ADC signals
    'Computing averaged ADC signals
    bfWrite BF_GEPHY.SigAnlSmpHldWaitRdEn, MSV(&H0)
    bfWrite BF_GEPHY.SigAnlMathSig1En, MSV(&H1)
    bfWrite BF_GEPHY.SigAnlMathMltEn, MSV(&H0)
    bfWrite BF_GEPHY.SigAnlMathWinLenm1, MSV(255)

    'bfWrite BF_GEPHY.SigAnlMathEn, MSV(&H1)

    Dim slDgOutLsb1(0 To 3) As New SiteLong
    Dim slDgOutLsb2(0 To 3) As New SiteLong
    Dim slDgOutMid(0 To 3) As New SiteLong
    Dim slDgOutMsb(0 To 3) As New SiteLong


    Dim sdold_adc_offs_final1(0 To 512) As New SiteDouble
    Dim sdnew_adc_offs_final1(0 To 512) As New SiteDouble
    Dim sdold_adc_offs_final2(0 To 512) As New SiteDouble
    Dim sdnew_adc_offs_final2(0 To 512) As New SiteDouble
    Dim sdold_adc_offs_final3(0 To 512) As New SiteDouble
    Dim sdnew_adc_offs_final3(0 To 512) As New SiteDouble
    Dim sdold_adc_offs_final4(0 To 512) As New SiteDouble
    Dim sdnew_adc_offs_final4(0 To 512) As New SiteDouble

    Dim sladc_old_dc_offs(0 To 3) As New SiteLong
    Dim sladc_new_dc_offs(0 To 3) As New SiteLong

    Dim sladc_offs_default(0 To 8) As New SiteLong

    Dim sdWindow_Length As New SiteDouble

    Dim new_adc_avg_vals1_dsp As New DSPWave
    Dim old_adc_avg_vals1_dsp As New DSPWave
    Dim new_adc_avg_vals2_dsp As New DSPWave
    Dim old_adc_avg_vals2_dsp As New DSPWave
    Dim new_adc_avg_vals3_dsp As New DSPWave
    Dim old_adc_avg_vals3_dsp As New DSPWave
    Dim new_adc_avg_vals4_dsp As New DSPWave
    Dim old_adc_avg_vals4_dsp As New DSPWave

    Dim new_adc_avg_vals_comb1_dsp As New DSPWave
    Dim old_adc_avg_vals_comb1_dsp As New DSPWave
    Dim new_adc_avg_vals_comb2_dsp As New DSPWave
    Dim old_adc_avg_vals_comb2_dsp As New DSPWave
    Dim new_adc_avg_vals_comb3_dsp As New DSPWave
    Dim old_adc_avg_vals_comb3_dsp As New DSPWave

    Dim new_adc_avg_vals_comball_dsp As New DSPWave
    Dim old_adc_avg_vals_comball_dsp As New DSPWave

    Call new_adc_avg_vals1_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals1_dsp.CreateConstant(0, 128, DspDouble)
    Call new_adc_avg_vals2_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals2_dsp.CreateConstant(0, 128, DspDouble)
    Call new_adc_avg_vals3_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals3_dsp.CreateConstant(0, 128, DspDouble)
    Call new_adc_avg_vals4_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals4_dsp.CreateConstant(0, 128, DspDouble)

    Call new_adc_avg_vals_comb1_dsp.CreateConstant(0, 256, DspDouble)
    Call old_adc_avg_vals_comb1_dsp.CreateConstant(0, 256, DspDouble)
    Call new_adc_avg_vals_comb2_dsp.CreateConstant(0, 256, DspDouble)
    Call old_adc_avg_vals_comb2_dsp.CreateConstant(0, 256, DspDouble)
    Call new_adc_avg_vals_comb3_dsp.CreateConstant(0, 256, DspDouble)
    Call old_adc_avg_vals_comb3_dsp.CreateConstant(0, 256, DspDouble)

    Call new_adc_avg_vals_comball_dsp.CreateConstant(0, 512, DspDouble)
    Call old_adc_avg_vals_comball_dsp.CreateConstant(0, 512, DspDouble)

    Dim adc_offs_old As New SiteDouble
    Dim adc_offs_new As New SiteDouble

    Dim symm_err_nrg_new As New SiteDouble
    Dim symm_err_nrg_old As New SiteDouble

    Dim symm_err_nrg_new_db As New SiteDouble
    Dim symm_err_nrg_old_db As New SiteDouble

    sdWindow_Length = bfRead(BF_GEPHY.SigAnlMathWinLenm1)
    
    For Each site In TheExec.Sites
        If (sdWindow_Length = 0) Then
            sdWindow_Length = 999999999
        End If
    Next site

    array_count = 0

    For Each site In TheExec.Sites
        For code_incr = 0 To 127

                For Iadc_signals = 0 To 3
                    bfWrite BF_GEPHY.SigAnlSmpHldUsrIndx, MSV(new_adc_peak_array(Iadc_signals))
                    'TheHdw.Wait 1 * mS

                    Read = bfRead(BF_GEPHY.SigAnlMathRdy)
                    'Debug.Print "SigAnlMathRdy  :"; Read
                    thehdw.Wait 1 * ms
                    
                    Read = GEPhy_MDIO_CL_45_Read(0, 30, 36902) 'SigAnlMathRdyProdAccLo
                    slDgOutLsb2(Iadc_signals) = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlMathProdAccLo)
                    
                    'Debug.Print "SigAnlMathRdyNoClr:"; Read.BitwiseAnd(&H8000).Subtract(32767)
                    
                    Read = GEPhy_MDIO_CL_45_Read(0, 30, 36904) 'SigAnlMathMeanAccLo
                    slDgOutLsb1(Iadc_signals) = Read 'bfRead(BF_GEPHY.SigAnlMathMeanAccLo)
                    
                    Read = GEPhy_MDIO_CL_45_Read(0, 30, 36903) 'SigAnlMathMeanAccHi
                    slDgOutMsb(Iadc_signals) = Read.BitwiseAnd(&H7FFF) 'bfRead(BF_GEPHY.SigAnlMathMeanAccHi)
                    
                    Read = GEPhy_MDIO_CL_45_Read(0, 30, 36901) 'SigAnlMathProdAccMid
                    slDgOutMid(Iadc_signals) = Read 'bfRead(BF_GEPHY.SigAnlMathProdAccMid)

                    sladc_old_dc_offs(Iadc_signals) = slDgOutMsb(Iadc_signals).BitwiseAnd(&H7FFF).ShiftLeft(16).Add(slDgOutLsb1(Iadc_signals))
                    sladc_old_dc_offs(Iadc_signals) = sladc_old_dc_offs(Iadc_signals).Subtract(sladc_old_dc_offs(Iadc_signals).BitwiseAnd(&H40000000).ShiftLeft(1))

                    sladc_new_dc_offs(Iadc_signals) = slDgOutMid(Iadc_signals).ShiftLeft(15).Add(slDgOutLsb2(Iadc_signals).BitwiseAnd(&H7FFF))
                    sladc_new_dc_offs(Iadc_signals) = sladc_new_dc_offs(Iadc_signals).Subtract(sladc_new_dc_offs(Iadc_signals).BitwiseAnd(&H40000000).ShiftLeft(1))

                    sdold_adc_offs_final1(code_incr) = sladc_old_dc_offs(0).Divide(256).Divide(sdWindow_Length)
                    sdnew_adc_offs_final1(code_incr) = sladc_new_dc_offs(0).Divide(256).Divide(sdWindow_Length)

                    sdold_adc_offs_final2(code_incr) = sladc_old_dc_offs(1).Divide(256).Divide(sdWindow_Length)
                    sdnew_adc_offs_final2(code_incr) = sladc_new_dc_offs(1).Divide(256).Divide(sdWindow_Length)

                    sdold_adc_offs_final3(code_incr) = sladc_old_dc_offs(2).Divide(256).Divide(sdWindow_Length)
                    sdnew_adc_offs_final3(code_incr) = sladc_new_dc_offs(2).Divide(256).Divide(sdWindow_Length)

                    sdold_adc_offs_final4(code_incr) = sladc_old_dc_offs(3).Divide(256).Divide(sdWindow_Length)
                    sdnew_adc_offs_final4(code_incr) = sladc_new_dc_offs(3).Divide(256).Divide(sdWindow_Length)

                    new_adc_avg_vals1_dsp(site).Element(code_incr) = sdnew_adc_offs_final1(code_incr)
                    old_adc_avg_vals1_dsp(site).Element(code_incr) = sdold_adc_offs_final1(code_incr)

                    new_adc_avg_vals2_dsp(site).Element(code_incr) = sdnew_adc_offs_final2(code_incr)
                    old_adc_avg_vals2_dsp(site).Element(code_incr) = sdold_adc_offs_final2(code_incr)

                    new_adc_avg_vals3_dsp(site).Element(code_incr) = sdnew_adc_offs_final3(code_incr)
                    old_adc_avg_vals3_dsp(site).Element(code_incr) = sdold_adc_offs_final3(code_incr)

                    new_adc_avg_vals4_dsp(site).Element(code_incr) = sdnew_adc_offs_final4(code_incr)
                    old_adc_avg_vals4_dsp(site).Element(code_incr) = sdold_adc_offs_final4(code_incr)
                    
                    'Read = bfRead(BF_GEPHY.SigAnlSmpHldUsrIndx)
                    'Debug.Print "SigAnlSmpHldUsrIndx  :"; Read

                    array_count = array_count + 1
                Next Iadc_signals
                
            'Read = bfRead(BF_GEPHY.PllPhase)
            'Debug.Print "PllPhase_Read:"; Read

            bfWrite BF_GEPHY.PllDlyInc, MSV(&H1)
        Next code_incr
    Next site
    
    Dim sdsum_comb1_new, sdsum_comb1_old As New SiteDouble
    

    new_adc_avg_vals_comb1_dsp = new_adc_avg_vals1_dsp.Concatenate(new_adc_avg_vals2_dsp)
    new_adc_avg_vals_comb2_dsp = new_adc_avg_vals3_dsp.Concatenate(new_adc_avg_vals4_dsp)
    new_adc_avg_vals_comball_dsp = new_adc_avg_vals_comb1_dsp.Concatenate(new_adc_avg_vals_comb2_dsp)
    adc_offs_new = new_adc_avg_vals_comball_dsp.CalcMean
    new_adc_avg_vals_comb1_dsp = new_adc_avg_vals_comb1_dsp.Subtract(adc_offs_new)
    new_adc_avg_vals_comb2_dsp = new_adc_avg_vals_comb2_dsp.Subtract(adc_offs_new)
    new_adc_avg_vals_comb3_dsp = new_adc_avg_vals_comb1_dsp.Add(new_adc_avg_vals_comb2_dsp)
    new_adc_avg_vals_comb3_dsp = new_adc_avg_vals_comb3_dsp.Multiply(new_adc_avg_vals_comb3_dsp)
    symm_err_nrg_new = new_adc_avg_vals_comb3_dsp.CalcSum
    sdsum_comb1_new = new_adc_avg_vals_comb1_dsp.CalcSum
    For Each site In TheExec.Sites
        If (sdsum_comb1_new = 0) Then
            sdsum_comb1_new = 0.000000000001
        End If
    Next site
    symm_err_nrg_new_db = symm_err_nrg_new.Divide(sdsum_comb1_new)
    symm_err_nrg_new_db = symm_err_nrg_new_db.Log10.Multiply(10)

    'Debug.Print "symm_err_nrg_new:"; symm_err_nrg_new(0)
    'Debug.Print "symm_err_nrg_new_db:"; symm_err_nrg_new_db(0)

    'Added for plotting
    ''''''''''new_adc_avg_vals_comb2_dsp = new_adc_avg_vals_comb2_dsp.Negate
    ''''''''''new_adc_avg_vals_comb1_dsp.Plot "Average Measured Signal at New ADC"
    ''''''''''new_adc_avg_vals_comb2_dsp.Plot "+Average Measured Signal at New ADC"

    old_adc_avg_vals_comb1_dsp = old_adc_avg_vals1_dsp.Concatenate(old_adc_avg_vals2_dsp)
    old_adc_avg_vals_comb2_dsp = old_adc_avg_vals3_dsp.Concatenate(old_adc_avg_vals4_dsp)
    old_adc_avg_vals_comball_dsp = old_adc_avg_vals_comb1_dsp.Concatenate(old_adc_avg_vals_comb2_dsp)
    adc_offs_old = old_adc_avg_vals_comball_dsp.CalcMean
    old_adc_avg_vals_comb1_dsp = old_adc_avg_vals_comb1_dsp.Subtract(adc_offs_old)
    old_adc_avg_vals_comb2_dsp = old_adc_avg_vals_comb2_dsp.Subtract(adc_offs_old)
    old_adc_avg_vals_comb3_dsp = old_adc_avg_vals_comb1_dsp.Add(old_adc_avg_vals_comb2_dsp)
    old_adc_avg_vals_comb3_dsp = old_adc_avg_vals_comb3_dsp.Multiply(old_adc_avg_vals_comb3_dsp)
    symm_err_nrg_old = old_adc_avg_vals_comb3_dsp.CalcSum
    sdsum_comb1_old = old_adc_avg_vals_comb1_dsp.CalcSum
    For Each site In TheExec.Sites
        If (sdsum_comb1_old = 0) Then
            sdsum_comb1_old = 0.000000000001
        End If
    Next site
    symm_err_nrg_old_db = symm_err_nrg_old.Divide(sdsum_comb1_old)
    symm_err_nrg_old_db = symm_err_nrg_old_db.Log10.Multiply(10)

    'Debug.Print "symm_err_nrg_old:"; symm_err_nrg_old(0)
    'Debug.Print "symm_err_nrg_old_db:"; symm_err_nrg_old_db(0)

    'Added for plotting
    ''''''''''old_adc_avg_vals_comb2_dsp = old_adc_avg_vals_comb2_dsp.Negate
    ''''''''''old_adc_avg_vals_comb1_dsp.Plot "Average Measured Signal at Old ADC"
    ''''''''''old_adc_avg_vals_comb2_dsp.Plot "+Average Measured Signal at Old ADC"

    Dim adc_new_max_val As New SiteDouble
    Dim adc_old_max_val As New SiteDouble

    Dim vmid_new, ileft_new, iright_new, vleft_new, vright_new, adc_new_edge_x As New SiteDouble
    Dim vmid_old, ileft_old, iright_old, vleft_old, vright_old, adc_old_edge_x As New SiteDouble

    Dim edge_dly_ns As New SiteDouble

    Dim ph_sel_res As Double

    ph_sel_res = 0.0000000000625 '62.5pSec

    adc_new_max_val = new_adc_avg_vals_comb1_dsp.CalcMaximumValue
    vmid_new = adc_new_max_val.Divide(2)
    ileft_new = new_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, vmid_new)
    ileft_new = ileft_new - 1
    vleft_new = new_adc_avg_vals_comb1_dsp.Element(ileft_new)
    vright_new = new_adc_avg_vals_comb1_dsp.Element(ileft_new + 1)
    
    If (vright_new = vleft_new) Then
        vright_new = vleft_new = 999
    End If
    adc_new_edge_x = ileft_new + (vmid_new - vleft_new) / (vright_new - vleft_new)

    adc_old_max_val = old_adc_avg_vals_comb1_dsp.CalcMaximumValue
    vmid_old = adc_old_max_val.Divide(2)
    ileft_old = old_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, vmid_old)
    ileft_old = ileft_old - 1
    vleft_old = old_adc_avg_vals_comb1_dsp.Element(ileft_old)
    vright_old = old_adc_avg_vals_comb1_dsp.Element(ileft_old + 1)
    
    If (vright_old = vleft_old) Then
        vright_old = vleft_old = 999
    End If
    adc_old_edge_x = ileft_old + (vmid_old - vleft_old) / (vright_old - vleft_old)

    edge_dly_ns = ph_sel_res * (Abs(adc_old_edge_x - adc_new_edge_x)) / 0.000000001

    'Debug.Print "edge_dly_ns:"; edge_dly_ns(0)
    
    ''''''''''new_adc_avg_vals_comb1_dsp.Plot "Overlay of New and Old"
    ''''''''''old_adc_avg_vals_comb1_dsp.Plot "+Overlay of New and Old"

    Dim new_adc_avg_vals_rising1_dsp As New DSPWave
    Dim old_adc_avg_vals_rising1_dsp As New DSPWave

    Dim new_adc_avg_vals_rising2_dsp As New DSPWave
    Dim old_adc_avg_vals_rising2_dsp As New DSPWave

    Dim new_adc_avg_vals_rising_diff_dsp As New DSPWave
    Dim old_adc_avg_vals_rising_diff_dsp As New DSPWave

    Call new_adc_avg_vals_rising1_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals_rising1_dsp.CreateConstant(0, 128, DspDouble)

    Call new_adc_avg_vals_rising2_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals_rising2_dsp.CreateConstant(0, 128, DspDouble)

    Call new_adc_avg_vals_rising_diff_dsp.CreateConstant(0, 128, DspDouble)
    Call old_adc_avg_vals_rising_diff_dsp.CreateConstant(0, 128, DspDouble)

    Dim vstp_avg_new, vstp_avg_old As New SiteDouble
    Dim vstp_max_new, vstp_max_old As New SiteDouble
    Dim vstp_min_new, vstp_min_old As New SiteDouble

    ileft_new = new_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, 0.1 * adc_new_max_val)
    iright_new = new_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, 0.9 * adc_new_max_val)
    new_adc_avg_vals_rising1_dsp = new_adc_avg_vals_comb1_dsp.Select(ileft_new, 1, iright_new - ileft_new)
    new_adc_avg_vals_rising2_dsp = new_adc_avg_vals_rising1_dsp.Select(1, 1)
    new_adc_avg_vals_rising_diff_dsp = new_adc_avg_vals_rising2_dsp.Subtract(new_adc_avg_vals_rising1_dsp)
    vstp_avg_new = new_adc_avg_vals_rising_diff_dsp.CalcMean
    vstp_max_new = new_adc_avg_vals_rising_diff_dsp.CalcMaximumValue
    vstp_min_new = new_adc_avg_vals_rising_diff_dsp.CalcMinimumValue

    ''''''''''new_adc_avg_vals_rising1_dsp.Plot "Rising Edge at New ADC"

    ileft_old = old_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, 0.1 * adc_old_max_val)
    iright_old = old_adc_avg_vals_comb1_dsp.FindIndex(OfFirstElement, GreaterThan, 0.9 * adc_old_max_val)
    old_adc_avg_vals_rising1_dsp = old_adc_avg_vals_comb1_dsp.Select(ileft_old, 1, iright_old - ileft_old)
    old_adc_avg_vals_rising2_dsp = old_adc_avg_vals_rising1_dsp.Select(1, 1)
    old_adc_avg_vals_rising_diff_dsp = old_adc_avg_vals_rising2_dsp.Subtract(old_adc_avg_vals_rising1_dsp)
    vstp_avg_old = old_adc_avg_vals_rising_diff_dsp.CalcMean
    vstp_max_old = old_adc_avg_vals_rising_diff_dsp.CalcMaximumValue
    vstp_min_old = old_adc_avg_vals_rising_diff_dsp.CalcMinimumValue

    ''''''''''old_adc_avg_vals_rising1_dsp.Plot "Rising Edge at Old ADC"

Call GEPHY_MDIO_Halt

    'Datalogging
    Call TheExec.Flow.TestLimit(symm_err_nrg_new_db, ScaleType:=scaleNoScaling, unit:=unitDb, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(symm_err_nrg_old_db, ScaleType:=scaleNoScaling, unit:=unitDb, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(edge_dly_ns, ScaleType:=scaleNoScaling, unit:=unitTime, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(iright_new - ileft_new, ScaleType:=scaleNoScaling, unit:=unitNone, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(iright_old - ileft_old, ScaleType:=scaleNoScaling, unit:=unitNone, ForceResults:=tlForceFlow)
    
    Call TheExec.Flow.TestLimit(vstp_min_new - (0.25 * vstp_avg_new), ScaleType:=scaleMilli, unit:=unitVolt, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit((2.25 * vstp_avg_new) - vstp_max_new, ScaleType:=scaleMilli, unit:=unitVolt, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit(vstp_min_old - (0.25 * vstp_avg_old), ScaleType:=scaleMilli, unit:=unitVolt, ForceResults:=tlForceFlow)
    Call TheExec.Flow.TestLimit((2.25 * vstp_avg_old) - vstp_max_old, ScaleType:=scaleMilli, unit:=unitVolt, ForceResults:=tlForceFlow)

    'PHY and MAC in Digital Mode
    thehdw.Utility.Pins("PHY_LPBK_RLY_A, MAC_LPBK_RLY").State = tlUtilBitOff

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

