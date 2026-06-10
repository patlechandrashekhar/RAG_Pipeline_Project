Attribute VB_Name = "RunVBT"
' This ALWAYS GENERATED file contains wrappers for VBT tests.
' Do not edit.

Private Sub HandleUntrappedError()
    ' Sanity clause
    If TheExec Is Nothing Then
        MsgBox "IG-XL is not running!  VBT tests cannot execute unless IG-XL is running."
        Exit Sub
    End If
    ' If the last site has failed out, let's ignore the error
    If TheExec.Sites.Active.Count = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 6.10 aborts gracefully.)
    Do While TheExec.Sites.InSiteLoop
        Call TheExec.Sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Select all active sites in case a subset of sites was selected when error occurred.
    TheExec.Sites.Selected = TheExec.Sites.Active
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    AbortTest
End Sub

Public Function DCVIPowerSupply_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As New Pattern
    p8.Value = v(7)
    Dim p9 As New pinlist
    p9.Value = v(8)
    Dim p10 As New pinlist
    p10.Value = v(9)
    Dim p11 As New pinlist
    p11.Value = v(10)
    Dim p12 As New pinlist
    p12.Value = v(11)
    Dim p13 As New pinlist
    p13.Value = v(17)
    Dim p14 As New pinlist
    p14.Value = v(18)
    Dim p15 As tlPSSource
    p15 = v(19)
    Dim p16 As tlRelayMode
    p16 = v(34)
    Dim p17 As New pinlist
    p17.Value = v(35)
    Dim p18 As New pinlist
    p18.Value = v(36)
    Dim p19 As tlPSTestControl
    p19 = v(37)
    Dim p20 As New InterposeName
    p20.Value = v(39)
    Dim p21 As tlWaitVal
    p21 = v(41)
    Dim p22 As tlWaitVal
    p22 = v(42)
    Dim p23 As tlWaitVal
    p23 = v(43)
    Dim p24 As tlWaitVal
    p24 = v(44)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVIPowerSupply_T__ = Template.VBT_DCVIPowerSupply_T.DCVIPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, CDbl(v(12)), CLng(v(13)), CStr(v(14)), CDbl(v(15)), CDbl(v(16)), p13, p14, p15, CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CStr(v(32)), CBool(v(33)), p16, p17, p18, p19, CBool(v(38)), p20, CStr(v(40)), p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(46)), , CStr(v(47)), CBool(v(48)), CBool(v(49)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVSPowerSupply_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As New Pattern
    p8.Value = v(7)
    Dim p9 As New pinlist
    p9.Value = v(8)
    Dim p10 As New pinlist
    p10.Value = v(9)
    Dim p11 As New pinlist
    p11.Value = v(10)
    Dim p12 As New pinlist
    p12.Value = v(11)
    Dim p13 As New pinlist
    p13.Value = v(12)
    Dim p14 As New pinlist
    p14.Value = v(16)
    Dim p15 As tlPSSource
    p15 = v(17)
    Dim p16 As tlRelayMode
    p16 = v(31)
    Dim p17 As New pinlist
    p17.Value = v(32)
    Dim p18 As New pinlist
    p18.Value = v(33)
    Dim p19 As tlPSTestControl
    p19 = v(34)
    Dim p20 As tlWaitVal
    p20 = v(35)
    Dim p21 As tlWaitVal
    p21 = v(36)
    Dim p22 As tlWaitVal
    p22 = v(37)
    Dim p23 As tlWaitVal
    p23 = v(38)
    Dim p24 As New FormulaArg
    p24.Value = v(40)
    Dim p25 As New FormulaArg
    p25.Value = v(41)
    Dim p26 As New FormulaArg
    p26.Value = v(42)
    Dim p27 As New FormulaArg
    p27.Value = v(43)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVSPowerSupply_T__ = Template.VBT_DCVSPowerSupply_T.DCVSPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CDbl(v(13)), CLng(v(14)), CStr(v(15)), p14, p15, CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(30)), p16, p17, p18, p19, p20, p21, p22, p23, CBool(v(UBound(v))), p24, p25, p26, p27, , CStr(v(44)), CBool(v(45)), CBool(v(46)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Empty_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New InterposeName
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New pinlist
    p7.Value = v(12)
    Dim p8 As New pinlist
    p8.Value = v(13)
    Dim p9 As New pinlist
    p9.Value = v(14)
    Dim p10 As New pinlist
    p10.Value = v(15)
    Dim p11 As New pinlist
    p11.Value = v(16)
    Dim p12 As New pinlist
    p12.Value = v(17)
    Dim p13 As New pinlist
    p13.Value = v(18)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Empty_T__ = Template.VBT_Empty_T.Empty_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, p13, pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New pinlist
    p10.Value = v(9)
    Dim p11 As New pinlist
    p11.Value = v(10)
    Dim p12 As New pinlist
    p12.Value = v(11)
    Dim p13 As New pinlist
    p13.Value = v(12)
    Dim p14 As New pinlist
    p14.Value = v(13)
    Dim p15 As New pinlist
    p15.Value = v(20)
    Dim p16 As New pinlist
    p16.Value = v(21)
    Dim p17 As New InterposeName
    p17.Value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Functional_T__ = Template.VBT_Functional_T.Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PinPMU_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New InterposeName
    p1.Value = v(1)
    Dim p2 As New InterposeName
    p2.Value = v(2)
    Dim p3 As New InterposeName
    p3.Value = v(3)
    Dim p4 As New InterposeName
    p4.Value = v(4)
    Dim p5 As New InterposeName
    p5.Value = v(5)
    Dim p6 As New InterposeName
    p6.Value = v(6)
    Dim p7 As New Pattern
    p7.Value = v(7)
    Dim p8 As New Pattern
    p8.Value = v(8)
    Dim p9 As New pinlist
    p9.Value = v(10)
    Dim p10 As New pinlist
    p10.Value = v(11)
    Dim p11 As New pinlist
    p11.Value = v(12)
    Dim p12 As New pinlist
    p12.Value = v(13)
    Dim p13 As New pinlist
    p13.Value = v(14)
    Dim p14 As New pinlist
    p14.Value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.Value = v(18)
    Dim p17 As New FormulaArg
    p17.Value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New pinlist
    p19.Value = v(36)
    Dim p20 As New pinlist
    p20.Value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.Value = v(52)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p27 As New pinlist
    p27.Value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.Value = v(55)
    PinPMU_T__ = Template.VBT_PinPmu_T.PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MtoMemory_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As New pinlist
    p9.Value = v(8)
    Dim p10 As New pinlist
    p10.Value = v(9)
    Dim p11 As New pinlist
    p11.Value = v(10)
    Dim p12 As New pinlist
    p12.Value = v(11)
    Dim p13 As New pinlist
    p13.Value = v(12)
    Dim p14 As New pinlist
    p14.Value = v(19)
    Dim p15 As New pinlist
    p15.Value = v(20)
    Dim p16 As New InterposeName
    p16.Value = v(21)
    Dim p17 As tlRelayMode
    p17 = v(24)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim ExtraArgs(0 To 49) As Variant
    Dim i As Integer
    For i = 0 To 49
        ExtraArgs(i) = v(51 + i)
    Next i
    MtoMemory_T__ = Template.VBT_MTOMemory_T.MtoMemory_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(23)), p17, CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CLng(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CLng(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CStr(v(38)), CLng(v(39)), CLng(v(40)), CBool(v(UBound(v))), pStep, ExtraArgs, CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MLS_Speed_Config_threshold__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MLS_Speed_Config_threshold__ = VBAProject.VBT_Char_Levels.MLS_Speed_Config_threshold()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MLS_auto_mdix_threshold__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MLS_auto_mdix_threshold__ = VBAProject.VBT_Char_Levels.MLS_auto_mdix_threshold()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MLS_Energy_Det_threshold__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MLS_Energy_Det_threshold__ = VBAProject.VBT_Char_Levels.MLS_Energy_Det_threshold()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VIH_level__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(1)
    VIH_level__ = VBAProject.VBT_Char_Levels.VIH_level(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VIL_level__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(1)
    VIL_level__ = VBAProject.VBT_Char_Levels.VIL_level(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOH_level__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(4)
    Dim p4 As New pinlist
    p4.Value = v(5)
    VOH_level__ = VBAProject.VBT_Char_Levels.VOH_level(p1, p2, CDbl(v(2)), CStr(v(3)), p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MDIO_timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MDIO_timing__ = VBAProject.VBT_Char_Timing.MDIO_timing()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function reset_dly_timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    reset_dly_timing__ = VBAProject.VBT_Char_Timing.reset_dly_timing()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GMII_1G_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    GMII_1G_Timing__ = VBAProject.VBT_Char_Timing.GMII_1G_Timing(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RGMII_1G_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    RGMII_1G_Timing__ = VBAProject.VBT_Char_Timing.RGMII_1G_Timing(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RGMII_100M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    RGMII_100M_Timing__ = VBAProject.VBT_Char_Timing.RGMII_100M_Timing(p1, p2, p3, p4, CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RGMII_10M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    RGMII_10M_Timing__ = VBAProject.VBT_Char_Timing.RGMII_10M_Timing(p1, p2, p3, p4, CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MII_TXCLK_100M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    MII_TXCLK_100M_Timing__ = VBAProject.VBT_Char_Timing.MII_TXCLK_100M_Timing(p1, p2, p3, p4, CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MII_TXCLK_10M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    MII_TXCLK_10M_Timing__ = VBAProject.VBT_Char_Timing.MII_TXCLK_10M_Timing(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RMII_100M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    RMII_100M_Timing__ = VBAProject.VBT_Char_Timing.RMII_100M_Timing(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RMII_10M_Timing__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New PatternSet
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Dim p4 As New pinlist
    p4.Value = v(3)
    RMII_10M_Timing__ = VBAProject.VBT_Char_Timing.RMII_10M_Timing(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function charSetup__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    charSetup__ = VBAProject.VBT_Characterisation.charSetup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function charCleanup__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    charCleanup__ = VBAProject.VBT_Characterisation.charCleanup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function getNumFromUser__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    getNumFromUser__ = VBAProject.VBT_Characterisation.getNumFromUser(CStr(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SheetExists__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' SheetExists__ = VBAProject.VBT_Characterisation.SheetExists(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LastRow__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' LastRow__ = VBAProject.VBT_Characterisation.LastRow(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LastCol__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' LastCol__ = VBAProject.VBT_Characterisation.LastCol(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function stopRow__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' stopRow__ = VBAProject.VBT_Characterisation.stopRow(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function incFailCount__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Call VBAProject.VBT_Characterisation.incFailCount
    incFailCount__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function setForceEnd__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Call VBAProject.VBT_Characterisation.setForceEnd
    setForceEnd__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function clearForceEnd__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Call VBAProject.VBT_Characterisation.clearForceEnd
    clearForceEnd__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function forceEnd__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' forceEnd__ = VBAProject.VBT_Characterisation.forceEnd(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function HRAM_Fail_Dump__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As Fail_Log_Fmt
    p1 = v(2)
    HRAM_Fail_Dump__ = VBAProject.VBT_Fail_Dump.HRAM_Fail_Dump(CStr(v(0)), CLng(v(1)), p1, CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Socket_IO_Leakage__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Socket_IO_Leakage__ = VBAProject.VBT_HW_Checker.Socket_IO_Leakage(p1, CDbl(v(1)), CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Socket_Power_Shorts__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Socket_Power_Shorts__ = VBAProject.VBT_HW_Checker.Socket_Power_Shorts()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Resistor_Measurement__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Resistor_Measurement__ = VBAProject.VBT_HW_Checker.Resistor_Measurement()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_5V_Check__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Relay_5V_Check__ = VBAProject.VBT_HW_Checker.Relay_5V_Check()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_12V_Check__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Relay_12V_Check__ = VBAProject.VBT_HW_Checker.Relay_12V_Check()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HSD_Trace__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    HSD_Trace__ = VBAProject.VBT_HW_Checker.HSD_Trace(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Retry__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Retry__ = VBAProject.VBT_Interpose.Retry(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_4_3_1__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_GP_OUT_LED_0_MLS_MODE_4_3_1__ = VBAProject.VBT_Interpose.Set_GP_OUT_LED_0_MLS_MODE_4_3_1(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_4_3_2__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_GP_OUT_LED_0_MLS_MODE_4_3_2__ = VBAProject.VBT_Interpose.Set_GP_OUT_LED_0_MLS_MODE_4_3_2(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_3_2_1__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_GP_OUT_LED_0_MLS_MODE_3_2_1__ = VBAProject.VBT_Interpose.Set_GP_OUT_LED_0_MLS_MODE_3_2_1(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Reset_GP_OUT_LED_0_MLS_MODE__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Reset_GP_OUT_LED_0_MLS_MODE__ = VBAProject.VBT_Interpose.Reset_GP_OUT_LED_0_MLS_MODE(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_GP_CLK_MLS_MODE_0_3__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_GP_CLK_MLS_MODE_0_3__ = VBAProject.VBT_Interpose.Set_GP_CLK_MLS_MODE_0_3(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_GP_CLK_MLS_MODE_1_2__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_GP_CLK_MLS_MODE_1_2__ = VBAProject.VBT_Interpose.Set_GP_CLK_MLS_MODE_1_2(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Reset_GP_CLK_MLS_MODE__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Reset_GP_CLK_MLS_MODE__ = VBAProject.VBT_Interpose.Reset_GP_CLK_MLS_MODE(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_LED_1_MLS_MODE_0_3__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_LED_1_MLS_MODE_0_3__ = VBAProject.VBT_Interpose.Set_LED_1_MLS_MODE_0_3(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_LED_1_MLS_MODE_1_2__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_LED_1_MLS_MODE_1_2__ = VBAProject.VBT_Interpose.Set_LED_1_MLS_MODE_1_2(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Reset_LED_1_MLS_MODE__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Reset_LED_1_MLS_MODE__ = VBAProject.VBT_Interpose.Reset_LED_1_MLS_MODE(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_MLS_Spd_Config_Idx_0__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Set_MLS_Spd_Config_Idx_0__ = VBAProject.VBT_Interpose.Set_MLS_Spd_Config_Idx_0(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function GEPHY_MDIO_Init__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GEPHY_MDIO_Init__ = VBAProject.VBT_MDIO.GEPHY_MDIO_Init()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPhy_MDIO_CL_45_Write__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GEPhy_MDIO_CL_45_Write__ = VBAProject.VBT_MDIO.GEPhy_MDIO_CL_45_Write(CLng(v(0)), CLng(v(1)), CLng(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPhy_MDIO_CL_45_Read__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' GEPhy_MDIO_CL_45_Read__ = VBAProject.VBT_MDIO.GEPhy_MDIO_CL_45_Read(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPhy_MDIO_CL_22_Write__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GEPhy_MDIO_CL_22_Write__ = VBAProject.VBT_MDIO.GEPhy_MDIO_CL_22_Write(CLng(v(0)), CLng(v(1)), CLng(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPhy_MDIO_CL_22_Read__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' GEPhy_MDIO_CL_22_Read__ = VBAProject.VBT_MDIO.GEPhy_MDIO_CL_22_Read(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPHY_MDIO_Wait__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GEPHY_MDIO_Wait__ = VBAProject.VBT_MDIO.GEPHY_MDIO_Wait(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GEPHY_MDIO_Halt__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GEPHY_MDIO_Halt__ = VBAProject.VBT_MDIO.GEPHY_MDIO_Halt()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Set_PowerSupply__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Set_PowerSupply__ = VBAProject.VBT_Powersupply.Set_PowerSupply()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Powersupply_Shorts__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Powersupply_Shorts__ = VBAProject.VBT_Powersupply.Powersupply_Shorts()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function pin_to_pin_shorts__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(1)
    pin_to_pin_shorts__ = VBAProject.VBT_Powersupply.pin_to_pin_shorts(p1, p2, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Leakage__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(2)
    Dim p3 As New pinlist
    p3.Value = v(3)
    Leakage__ = VBAProject.VBT_Powersupply.Leakage(p1, CDbl(v(1)), p2, p3, CDbl(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Leakage_Supply__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pinlist
    p1.Value = v(0)
    Dim p2 As New pinlist
    p2.Value = v(1)
    Dim p3 As New pinlist
    p3.Value = v(2)
    Leakage_Supply__ = VBAProject.VBT_Powersupply.Leakage_Supply(p1, p2, p3, CDbl(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function XTAL_Parametrics__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    XTAL_Parametrics__ = VBAProject.VBT_Powersupply.XTAL_Parametrics()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function REXT_Voltage__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    REXT_Voltage__ = VBAProject.VBT_Powersupply.REXT_Voltage()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_GMII_1000T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_GMII_1000T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_GMII_1000T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_RGMII_1000T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_RGMII_1000T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_RGMII_1000T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_RGMII_100T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_RGMII_100T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_RGMII_100T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_RGMII_10T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_RGMII_10T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_RGMII_10T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_MII_100T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_MII_100T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_MII_100T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_MII_10T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_MII_10T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_MII_10T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_RMII_100T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_RMII_100T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_RMII_100T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Extlpbk_RMII_10T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_Extlpbk_RMII_10T__ = VBAProject.VBT_Powersupply.Power_Extlpbk_RMII_10T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_IDDQ__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New Pattern
    p1.Value = v(0)
    Power_IDDQ__ = VBAProject.VBT_Powersupply.Power_IDDQ(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HVST__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    HVST__ = VBAProject.VBT_Powersupply.HVST()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AMVR__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    AMVR__ = VBAProject.VBT_Powersupply.AMVR()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VLV_Threshold__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    VLV_Threshold__ = VBAProject.VBT_Powersupply.VLV_Threshold()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_SftPd__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Power_SftPd__ = VBAProject.VBT_Powersupply.Power_SftPd()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function meas_temp_mdio__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' meas_temp_mdio__ = VBAProject.VBT_Powersupply.meas_temp_mdio(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADIN1200_Regulator__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ADIN1200_Regulator__ = VBAProject.VBT_Powersupply.ADIN1200_Regulator()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MDIO_Read__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MDIO_Read__ = VBAProject.VBT_Tests.MDIO_Read()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ABIST__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ABIST__ = VBAProject.VBT_Tests.ABIST()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Grandlpbk_1000T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Grandlpbk_1000T__ = VBAProject.VBT_Tests.Grandlpbk_1000T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Grandlpbk_100T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Grandlpbk_100T__ = VBAProject.VBT_Tests.Grandlpbk_100T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Grandlpbk_10T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Grandlpbk_10T__ = VBAProject.VBT_Tests.Grandlpbk_10T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Extlpbk_GMII_1000T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Extlpbk_GMII_1000T__ = VBAProject.VBT_Tests.Extlpbk_GMII_1000T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Extlpbk_RGMII_1000T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Extlpbk_RGMII_1000T__ = VBAProject.VBT_Tests.Extlpbk_RGMII_1000T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Extlpbk_RGMII_100T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Extlpbk_RGMII_100T__ = VBAProject.VBT_Tests.Extlpbk_RGMII_100T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Extlpbk_RGMII_10T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Extlpbk_RGMII_10T__ = VBAProject.VBT_Tests.Extlpbk_RGMII_10T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Process_Monitor__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Process_Monitor__ = VBAProject.VBT_Tests.Process_Monitor()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GeClkCfg_Frequency__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    GeClkCfg_Frequency__ = VBAProject.VBT_Tests.GeClkCfg_Frequency()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MII_TX_CLK_Frequency__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MII_TX_CLK_Frequency__ = VBAProject.VBT_Tests.MII_TX_CLK_Frequency()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADIN1200_MDI_Parametric__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ADIN1200_MDI_Parametric__ = VBAProject.VBT_Tests.ADIN1200_MDI_Parametric()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MDI_Parametric__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    MDI_Parametric__ = VBAProject.VBT_Tests.MDI_Parametric()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOD_1000BaseT__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    VOD_1000BaseT__ = VBAProject.VBT_Tests.VOD_1000BaseT()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOD_100BaseT__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    VOD_100BaseT__ = VBAProject.VBT_Tests.VOD_100BaseT()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOD_10BaseT__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    VOD_10BaseT__ = VBAProject.VBT_Tests.VOD_10BaseT()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADC_DC_Offset__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ADC_DC_Offset__ = VBAProject.VBT_Tests.ADC_DC_Offset()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Phase_Sel__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Phase_Sel__ = VBAProject.VBT_Tests.Phase_Sel()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Find_Passing_Region__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Find_Passing_Region__ = VBAProject.VBT_Utilities.Find_Passing_Region(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Binary_Edge_Search__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Binary_Edge_Search__ = VBAProject.VBT_Utilities.Binary_Edge_Search(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Binary_Level_Search__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Binary_Level_Search__ = VBAProject.VBT_Utilities.Binary_Level_Search(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VLV_Level_Search__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' VLV_Level_Search__ = VBAProject.VBT_Utilities.VLV_Level_Search(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Measure_Frequency__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    ' Measure_Frequency__ = VBAProject.VBT_Utilities.Measure_Frequency(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































