Attribute VB_Name = "MDIO_test"
Public Function Process_Monitor_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant
Dim GePmCntN320nsm1 As Long, GePmRoSel As Long, GePmReq As Long
Dim GePmCntrl As Long

    Set Process_Monitor_test = New DSPWave
    
    Call Result.CreateConstant(0, 9, DspDouble)
    
    'Setups
    'This register controls the time the process monitor counter is enabled and counting the
    'number of cycles of the process monitor ring oscillator. It is programmed in units of 320 ns.
    GePmCntN320nsm1 = &HFF
    'The process monitor request causes the GE sub-system to perform a process monitor count to be initiated.
    GePmReq = 1
    
    'RVT ring oscillator
    GePmRoSel = &H0
    GePmCntrl = (GePmCntN320nsm1 * (2 ^ 3)) + (GePmRoSel * (2 ^ 1)) + GePmReq

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65350, GePmCntrl)
    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65351)
    
    'Get PM Results
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H1)
     Next site
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H2).ShiftRight(1)
     Next site

     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read.ShiftRight(2).Divide(((320 * (GePmCntN320nsm1 + 1)) + 40) * ns)
     Next site
     

    'LVT ring oscillator
    GePmRoSel = &H1
    GePmCntrl = (GePmCntN320nsm1 * (2 ^ 3)) + (GePmRoSel * (2 ^ 1)) + GePmReq

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65350, GePmCntrl)
    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65351)
    
    'Get PM Results
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read.BitwiseAnd(&H1)
     Next site

     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read.BitwiseAnd(&H2).ShiftRight(1)
     Next site

     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read.ShiftRight(2).Divide(((320 * (GePmCntN320nsm1 + 1)) + 40) * ns)
     Next site
     
    'HVT ring oscillator
    GePmRoSel = &H2
    GePmCntrl = (GePmCntN320nsm1 * (2 ^ 3)) + (GePmRoSel * (2 ^ 1)) + GePmReq

    Call GEPhy_MDIO_CL_45_Write(0, 30, 65350, GePmCntrl)
    Read = GEPhy_MDIO_CL_45_Read(0, 30, 65351)
    
    'Get PM Results
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read.BitwiseAnd(&H1)
     Next site

     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read.BitwiseAnd(&H2).ShiftRight(1)
     Next site

     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read.ShiftRight(2).Divide(((320 * (GePmCntN320nsm1 + 1)) + 40) * ns)
     Next site
        

    For Each site In TheExec.Sites.Active
        Process_Monitor_test = Result.Copy
    Next site

End Function




Public Function Extlpbk_RGMII_1000T_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_1000T_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 9, &H1800) 'ManMstrSlvEnAdv 'ManMstrAdv
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, 4) 'SftPdPllPdEn
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 41477, 1) 'MsqlSkip
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34355, 0) 'Dfe1gCfg0Coef0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34356, 0) 'Dfe1gCfg1Coef0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 41483, 0) 'ClenEstEn
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34312, 3) 'CanNcycm1Ndd2
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34322, 0) 'CanNcycm1Dd1
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34326, 0) 'CanNcycm1Dd2
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34332, 3) 'CanNcycm1Fb1
    Call GEPhy_MDIO_CL_45_Write(0, 30, 34336, 0) 'CanNcycm1Fb2
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 40449, 1) 'PgaLvlFixEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 40450, &H1F) 'PgaLvlFval
    Call GEPhy_MDIO_CL_45_Write(0, 30, 43010, 3355) 'FreqOffsFrcVal
    Call GEPhy_MDIO_CL_45_Write(0, 30, 43008, &H1) 'FreqOffsFrcEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H1048) 'LinkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms

    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site
     
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE07) 'GeRgmiiCfg
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1250) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37917, 10) 'FgNfrmL
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H3)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
     
     Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H1)   'DiagClkEn
     
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site

    

    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_1000T_test = Result.Copy
    Next site

End Function


Public Function Extlpbk_GMII_1000T_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_GMII_1000T_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
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
    
    'Read Link status
    GEPHY_MDIO_Wait 150 * ms

    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site
     
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1250) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37917, 10) 'FgNfrmL
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
     
     Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H1)   'DiagClkEn
 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site

    

    For Each site In TheExec.Sites.Active
        Extlpbk_GMII_1000T_test = Result.Copy
    Next site

End Function




Public Function Grandlpbk_1000T_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_1000T_test = New DSPWave
    
    Call Result.CreateConstant(0, 10, DspLong)
    
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
'    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
'    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HF5)   'LbAllDigSel 'LbExtEn
'    Call GEPhy_MDIO_CL_22_Write(&H0, 9, &H1800) 'ManMstrSlvEnAdv 'ManMstrAdv
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 65330, 4) 'GeRxTxExtLbEn
'
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 41477, 1) 'MsqlSkip
'
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34355, 0) 'Dfe1gCfg0Coef0
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34356, 0) 'Dfe1gCfg1Coef0
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 41483, 0) 'ClenEstEn
'
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34312, 3) 'CanNcycm1Ndd2
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34322, 0) 'CanNcycm1Dd1
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34326, 0) 'CanNcycm1Dd2
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34332, 3) 'CanNcycm1Fb1
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 34336, 0) 'CanNcycm1Fb2
'
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 40449, 1) 'PgaLvlFixEn
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 40450, &H1F) 'PgaLvlFval
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 43010, 3355) 'FreqOffsFrcVal
'    Call GEPhy_MDIO_CL_45_Write(0, 30, 43008, &H1) 'FreqOffsFrcEn
'    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H3048) 'LinkEn
'    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)

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
    
    'Read Link status
    GEPHY_MDIO_Wait 180 * ms
    
    'Read Link status
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site

    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg


    
    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1250) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37917, 10) 'FgNfrmL
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
    

 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20) 'RxErrCnt
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898) 'FcFrmCntH
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899) 'FcFrmCntL
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     

     
    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE03) 'GeRgmiiCfg
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    
    
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
     
     Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
     
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(9) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '1
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
         


    

    For Each site In TheExec.Sites.Active
        Grandlpbk_1000T_test = Result.Copy
    Next site

End Function



Public Function Grandlpbk_100T_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_100T_test = New DSPWave
    
    Call Result.CreateConstant(0, 16, DspLong)
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn 'AutoMdiEn 'ManMdix
    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H3048) 'LinkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H2100)

    
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    
    'Read Link status
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site

    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg


    
    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1250) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37917, 10) 'FgNfrmL
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
    

 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20) 'RxErrCnt
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898) 'FcFrmCntH
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899) 'FcFrmCntL
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     

     
    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE03) 'GeRgmiiCfg
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    
    
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
     
     Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
     
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(9) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '1
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site

    'force MDI
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H202) 'DiagClkEn 'AutoMdiEn 'ManMdix


    
    'Read Link status
    GEPHY_MDIO_Wait 40 * ms
    
    'Read Link status
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(10) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(11) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site
         
    'Configuring GeSubsys for MII mode  GeMiiUseGtxClk = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE02) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H2) 'GeMiiUseGtxClk

    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H206) 'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn

     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(12) = Read
     Next site
    
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H202)   'DiagClkEn
    
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(13) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(14) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(15) = Read
     Next site

    

    For Each site In TheExec.Sites.Active
        Grandlpbk_100T_test = Result.Copy
    Next site

End Function




Public Function Grandlpbk_10T_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_10T_test = New DSPWave
    
    Call Result.CreateConstant(0, 16, DspLong)
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49152, 0) 'SftPdPllPdEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H900)
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HC1)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn 'AutoMdiEn 'ManMdix
    Call GEPhy_MDIO_CL_45_Write(0, 30, 33281, 1) 'BrkLnkFrc
    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H3048) 'LinkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H100)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 45057, 262) 'LnkPulDlyInact
    Call GEPhy_MDIO_CL_45_Write(0, 30, 49667, &H8E30) 'MsTickN100usm1Emi 'P1msTickNusm1Emi 'UsTickN40nsm1Emi
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    
    'Read Link status
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site
     
     Call GEPhy_MDIO_CL_45_Write(0, 30, 49667, &HC3F9) 'MsTickN100usm1Emi 'P1msTickNusm1Emi 'UsTickN40nsm1Emi

    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE05) 'GeRgmiiCfg


    
    'Configuring frame generator to send packets
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37914, 1250) 'FgFrmLen
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37917, 10) 'FgNfrmL
    
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37895, &H1) 'FcTxSel
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    'Read FG status
    GEPHY_MDIO_Wait 10 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
    

 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20) 'RxErrCnt
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898) 'FcFrmCntH
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899) 'FcFrmCntL
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     

     
    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE03) 'GeRgmiiCfg
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn
    
     
    
    
    'Read FG status
    GEPHY_MDIO_Wait 10 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
     
     Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H2)   'DiagClkEn
     
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(9) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '1
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site

    'force MDI
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H202) 'DiagClkEn 'AutoMdiEn 'ManMdix


    
    'Read Link status
'    GEPHY_MDIO_Wait 40 * mS
    
    'Read Link status
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
    Read = GEPhy_MDIO_CL_22_Read(&H0, 1)
     For Each site In TheExec.Sites.Active
        Result.Element(10) = Read.BitwiseAnd(&H4).ShiftRight(2)
     Next site
     
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(11) = Read.BitwiseAnd(&H380).ShiftRight(7)
     Next site
         
    'Configuring GeSubsys for MII mode  GeMiiUseGtxClk = 0
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65315, &HE02) 'GeRgmiiCfg
    Call GEPhy_MDIO_CL_45_Write(0, 30, 65319, &H2) 'GeMiiUseGtxClk

    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H206) 'DiagClkEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H0) 'FgEn
    Call GEPhy_MDIO_CL_45_Write(0, 30, 37909, &H1) 'FgEn

     
    'Read FG status
    GEPHY_MDIO_Wait 10 * ms
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37918) 'FgDone
     For Each site In TheExec.Sites.Active
        Result.Element(12) = Read
     Next site
    
    
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H202)   'DiagClkEn
    
    Read = GEPhy_MDIO_CL_22_Read(&H0, 20)
     For Each site In TheExec.Sites.Active
        Result.Element(13) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37898)
     For Each site In TheExec.Sites.Active
        Result.Element(14) = Read
     Next site
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 37899)
     For Each site In TheExec.Sites.Active
        Result.Element(15) = Read
     Next site

    

    For Each site In TheExec.Sites.Active
        Grandlpbk_10T_test = Result.Copy
    Next site

End Function




Public Function ABIST_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set ABIST_test = New DSPWave
    
    Call Result.CreateConstant(0, 2, DspLong)

    'Setups to start AFE bIST
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)

    Call GEPhy_MDIO_CL_45_Write(0, 30, 46083, 0) 'B10eEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 19, &HF5)   'LbAllDigSel 'LbExtEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H3048) 'LinkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    
    'Wait for Phy in Stdby
    GEPHY_MDIO_Wait 35 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read.BitwiseAnd(&H8000).ShiftRight(15)
     Next site
     
    'Test 2 B1000 PGA Gain (TcPgaGainB1000 [30:63], TcAdcTolB1000 [2:12])
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46762, 56) 'TcPgaGainB1000 'Default is 57
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46753, 10) 'TcAdcTolB1000 'Default is 7
    
    'Test 3 B10 PGA Gain (TcPgaGainB10 [20:30], TcAdcTolB10 [1:11])
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46763, 28) 'TcPgaGainB10 'Default is 27
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46754, 10) 'TcAdcTolB10 'Default is 7
    
    'Test 4 PGA Tolerance Self test (TcLpbkDlyB1100[3:5], TcAdcTolPga[1:31])
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46756, 22) 'TcAdcTolPga 'Default is 8, 14 in test case
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46759, 4) 'TcLpbkDlyB1100 ''Default is 4
    
    'Test 6 Phmix tolerance (TcPgaGainPhmix[50:63], TcAdcTolPhmix[1:15])
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46757, 13) 'TcAdcTolPhmix 'Default is 12, 31 in test case
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46765, 52) 'TcPgaGainPhmix 'Default is 52, 58 in test case
    
    'Test 5 has no setting - DC Offset
    'Test 7 default settings are fine - Hybrid mode (TcPgaGainHybm[45:63], TcHybAdcThr [1:31])
    
    'For test_7 increase the tol hybm & increase the gain
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46758, 31) 'TcHybAdcThr 'Default is 13
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46766, 48) 'TcPgaGainHybm 'Default is 57

    Call GEPhy_MDIO_CL_45_Write(0, 30, 46747, 0) 'TcInst

    
    '//now set the tc_inst_emi
    '//bit 0 = 1 which is the enable bit
    '//bit1 = 0
    '//bit2-7 = 1(all the testmodes)
    '//bits8-9 for debug = 0
    '// Expect T6 to fail - PhaseMixer as using perfect channel model, no slewing just | line rise time
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46747, &HFD) 'TcInst

    
    GEPHY_MDIO_Wait 100 * ms
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 46748) 'TcRslt
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    
    'Setups to start 10BTe AFE bIST
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H940)
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46083, &H1) 'B10eEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 23, &H3048) 'LinkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 18, &H6)   'DiagClkEn
    Call GEPhy_MDIO_CL_22_Write(&H0, 0, &H140)

    
    'Wait for Phy in Stdby
    GEPHY_MDIO_Wait 10 * ms
    Read = GEPhy_MDIO_CL_22_Read(&H0, 26)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read.BitwiseAnd(&H8000).ShiftRight(15)
     Next site
    
    'Test 3 B10 PGA Gain (TcPgaGainB10 [20:30], TcAdcTolB10 [1:11])
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46763, 28) 'TcPgaGainB10 'Default is 27
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46754, 10) 'TcAdcTolB10 'Default is 7

    Call GEPhy_MDIO_CL_45_Write(0, 30, 46764, 41) 'TcPgaGainB10Te 'Default is 45
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46755, 8) 'TcAdcTolB10Te 'Default is 7

       
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46747, 0) 'TcInst

    
    '//now set the tc_inst_emi
    '//bit 0 = 1 which is the enable bit
    '//bit1 = 0
    '//bit2-7 = 1(all the testmodes)
    '//bits8-9 for debug = 0
    '// Expect T6 to fail - PhaseMixer as using perfect channel model, no slewing just | line rise time
    Call GEPhy_MDIO_CL_45_Write(0, 30, 46747, &H9) 'TcInst
    
    thehdw.Wait 20 * ms
    
    Read = GEPhy_MDIO_CL_45_Read(&H0, 30, 46748) 'TcRslt
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site



    For Each site In TheExec.Sites.Active
        ABIST_test = Result.Copy
    Next site



End Function
