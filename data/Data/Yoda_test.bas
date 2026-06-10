Attribute VB_Name = "Yoda_test"
Option Explicit

'Public variables to use with Yoda memory map
Public MM As CInstance

Public RM_GEPHY As CRM_GEPhy
Public BF_GEPHY As CBFS_MdioMap_GEPhy

Public RM_GESUB As CRM_GESubsys
Public BF_GESUB As CBFS_MdioMap_GESubsys


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module Description:
'    Module to store MDIO transactions done through Yoda
'
'   Requirement:
'    Ycomms and VBT_MDIO module with either MDIO_DSSC or MDIO_PA class.
'   ADIN1300 program uses the ADIN1300_yda.xla reference in IGXL
'
'   Usage:
'   -  To access GEPHY regsiter functions, use RM_GEPHY and BF_GEPHY object
'   -  To access GE Subsystem regsiter functions, use RM_GESUB and BF_GESUB object
'   -  This Module uses works with YComms and VBT_MDIO module
'
'   Rev 1.0 (vsomasun, Feb 8th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Function MDIO_Access_test()

Dim Read As New SiteLong


    Read = regRead(RM_GEPHY.PhyId1)
    
    Read = regRead(RM_GEPHY.Miicontrol)
    TheExec.Datalog.WriteComment ("Miireg:" & Hex(Read(0)))
    
    Call bfWrite(BF_GESUB.GeClkFree125En, MSV(&H1))
    
    Read = regRead(RM_GEPHY.Miicontrol)
    

End Function


Public Function Grandlpbk_1000T_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_1000T_yoda_test = New DSPWave
    
    Call mapReset(RM_GEPHY)
    Call mapReset(RM_GESUB)
    
    Call Result.CreateConstant(0, 10, DspLong)
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H1)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.ManMstrAdv, MSV(&H1)
    bfWrite BF_GEPHY.ManMstrSlvEnAdv, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)
    
    bfWrite BF_GEPHY.MsqlSkip, MSV(&H1)
    
    bfWrite BF_GEPHY.Dfe1gCfg0Coef0, MSV(&H0)
    bfWrite BF_GEPHY.Dfe1gCfg1Coef0, MSV(&H0)
    bfWrite BF_GEPHY.ClenEstEn, MSV(&H0)
    
    bfWrite BF_GEPHY.CanNcycm1Ndd2, MSV(&H3)
    bfWrite BF_GEPHY.CanNcycm1Dd1, MSV(&H0)
    bfWrite BF_GEPHY.CanNcycm1Dd2, MSV(&H0)
    bfWrite BF_GEPHY.CanNcycm1Fb1, MSV(&H3)
    bfWrite BF_GEPHY.CanNcycm1Fb2, MSV(&H0)
    
    bfWrite BF_GEPHY.PgaLvlFixEn, MSV(&H1)
    bfWrite BF_GEPHY.PgaLvlFval, MSV(&H1F)
    bfWrite BF_GEPHY.FreqOffsFrcVal, MSV(3355)
    bfWrite BF_GEPHY.FreqOffsFrcEn, MSV(&H1)
    bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H0)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H1)


    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(1250)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    

 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In TheExec.Sites.Active
'''''        TheExec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
     
    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H0)
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    thehdw.Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
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
        Grandlpbk_1000T_yoda_test = Result.Copy
    Next site

End Function

Public Function Grandlpbk_100T_yoda_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_100T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 16, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H1)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    
    'force MDI
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    
    'bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site
     
    
    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H0)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H1)
    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(1250)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site

    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H0)
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    thehdw.Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(9) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site

    

   'force MDI
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H1)
    
    Read = bfRead(BF_GEPHY.LinkStatLat)
    
    'Read Link status
    GEPHY_MDIO_Wait 40 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(10) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(11) = Read
     Next site
     
    
    'Configuring GeSubsys for MII mode  GeMiiUseGtxClk = 0
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H0)
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
     
    'Read FG status
    thehdw.Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(12) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(13) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(14) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(15) = Read
     Next site

'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeMiiUseGtxClk)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeMiiUseGtxClk: " & Read(site)
'''''     Next site
        
    
    For Each site In TheExec.Sites.Active
        Grandlpbk_100T_yoda_test = Result.Copy
    Next site

End Function
Public Function Grandlpbk_10T_yoda_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Grandlpbk_10T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 16, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    
    'force MDI
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    
    'bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    bfWrite BF_GEPHY.LnkPulDlyInact, MSV(262)
    bfWrite BF_GEPHY.MsTickN100usm1Emi, MSV(&H0)
    bfWrite BF_GEPHY.P1msTickNusm1Emi, MSV(99)
    bfWrite BF_GEPHY.UsTickN40nsm1Emi, MSV(&H11)
    
 
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site
    
    bfWrite BF_GEPHY.MsTickN100usm1Emi, MSV(&H9)
    bfWrite BF_GEPHY.P1msTickNusm1Emi, MSV(63)
    bfWrite BF_GEPHY.UsTickN40nsm1Emi, MSV(&H18)
    
    'Configuring GeSubsys for RGMII mode  GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 1
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H0)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H1)
    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(125)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site

    'Configuring GeSubsys for RGMII mode with GeRgmiiTxIdEn = 1 & GeRgmiiRxIdEn = 0
    bfWrite BF_GESUB.GeRgmiiTxIdEn, MSV(&H1)
    bfWrite BF_GESUB.GeRgmiiRxIdEn, MSV(&H0)
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    thehdw.Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(9) = Read
     Next site
     
     
'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site

    

   'force MDI
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H1)
    
    Read = bfRead(BF_GEPHY.LinkStatLat)
    
    'Read Link status
    GEPHY_MDIO_Wait 40 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(10) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(11) = Read
     Next site
     
    
    'Configuring GeSubsys for MII mode  GeMiiUseGtxClk = 0
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H0)
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
     
    'Read FG status
    thehdw.Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(12) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(13) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(14) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(15) = Read
     Next site

'''''     'Debug code
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.AutoMdiEn)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "AutoMdiEn: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GEPHY.ManMdix)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "ManMdix: " & Read(site)
'''''     Next site
'''''
'''''     Read = bfRead(BF_GESUB.GeMiiUseGtxClk)   '0
'''''     For Each site In theexec.Sites.Active
'''''        theexec.Datalog.WriteComment "GeMiiUseGtxClk: " & Read(site)
'''''     Next site
     
        
    
    For Each site In TheExec.Sites.Active
        Grandlpbk_10T_yoda_test = Result.Copy
    Next site

End Function

Public Function ABIST_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set ABIST_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 2, DspLong)

    'Setups to start AFE bIST
    'bfWrite BF_GEPHY.SftPd, MSV(&H1)
    'bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    'bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    'bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H1)
    'bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    bfWrite BF_GEPHY.B10eEn, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LinkEn, MSV(&H0)
    
     
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    
    'Wait for Phy in Stdby
    thehdw.Wait 10 * ms
    Read = bfRead(BF_GEPHY.PhyInStndby)
    
    'Test 2 B1000 PGA Gain (TcPgaGainB1000 [30:63], TcAdcTolB1000 [2:12])
    bfWrite BF_GEPHY.TcPgaGainB1000, MSV(56) 'Default is 57
    bfWrite BF_GEPHY.TcAdcTolB1000, MSV(10) 'Default is 7
    
    'Test 3 B10 PGA Gain (TcPgaGainB10 [20:30], TcAdcTolB10 [1:11])
    bfWrite BF_GEPHY.TcPgaGainB10, MSV(28) 'Default is 27
    bfWrite BF_GEPHY.TcAdcTolB10, MSV(10) 'Default is 7
    
    'Test 4 PGA Tolerance Self test (TcLpbkDlyB1100[3:5], TcAdcTolPga[1:31])
    bfWrite BF_GEPHY.TcAdcTolPga, MSV(22) 'Default is 8, 14 in test case
    bfWrite BF_GEPHY.TcLpbkDlyB1100, MSV(4) 'Default is 4
    
    'Test 6 Phmix tolerance (TcPgaGainPhmix[50:63], TcAdcTolPhmix[1:15])
    bfWrite BF_GEPHY.TcAdcTolPhmix, MSV(13) 'Default is 12, 31 in test case
    bfWrite BF_GEPHY.TcPgaGainPhmix, MSV(52) 'Default is 52, 58 in test case
    
    'Test 5 has no setting - DC Offset
    'Test 7 default settings are fine - Hybrid mode (TcPgaGainHybm[45:63], TcHybAdcThr [1:31])
    
    'For test_7 increase the tol hybm & increase the gain
    bfWrite BF_GEPHY.TcHybAdcThr, MSV(31) 'Default is 13
    bfWrite BF_GEPHY.TcPgaGainHybm, MSV(48) 'Default is 57
   
    bfWrite BF_GEPHY.TcInst, MSV(&H0)
    
    '//now set the tc_inst_emi
    '//bit 0 = 1 which is the enable bit
    '//bit1 = 0
    '//bit2-7 = 1(all the testmodes)
    '//bits8-9 for debug = 0
    '// Expect T6 to fail - PhaseMixer as using perfect channel model, no slewing just | line rise time
    bfWrite BF_GEPHY.TcInst, MSV(&HFD)
    
    thehdw.Wait 100 * ms
    
    Read = bfRead(BF_GEPHY.TcRslt)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
     
     

    'Setups to start 10BTe AFE bIST
    
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.B10eEn, MSV(&H1)
    'bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    'bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LinkEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
     
    'bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    
    'Wait for Phy in Stdby
    thehdw.Wait 10 * ms
    Read = bfRead(BF_GEPHY.PhyInStndby)
    
    'Test 3 B10 PGA Gain (TcPgaGainB10 [20:30], TcAdcTolB10 [1:11])
    bfWrite BF_GEPHY.TcPgaGainB10, MSV(28) 'Default is 27
    bfWrite BF_GEPHY.TcAdcTolB10, MSV(10) 'Default is 7
    
    bfWrite BF_GEPHY.TcPgaGainB10Te, MSV(41) 'Default is 45
    bfWrite BF_GEPHY.TcAdcTolB10Te, MSV(8) 'Default is 7
       
    bfWrite BF_GEPHY.TcInst, MSV(&H0)
    
    '//now set the tc_inst_emi
    '//bit 0 = 1 which is the enable bit
    '//bit1 = 0
    '//bit2-7 = 1(all the testmodes)
    '//bits8-9 for debug = 0
    '// Expect T6 to fail - PhaseMixer as using perfect channel model, no slewing just | line rise time
    bfWrite BF_GEPHY.TcInst, MSV(&H9)
    
    thehdw.Wait 20 * ms
    
    Read = bfRead(BF_GEPHY.TcRslt)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    
    
    For Each site In TheExec.Sites.Active
        ABIST_yoda_test = Result.Copy
    Next site



End Function
Public Function Extlpbk_RGMII_1000T_Power_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_1000T_Power_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 5, DspLong)
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)
    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    

    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgContModeEn, MSV(1)
    bfWrite BF_GEPHY.FgFrmLen, MSV(1500)
    bfWrite BF_GEPHY.FgCntrl, MSV(1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
     
    


    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
     


    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_1000T_Power_yoda_test = Result.Copy
    Next site

End Function
Public Function Extlpbk_RGMII_100T_Power_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_100T_Power_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 5, DspLong)
    
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H1)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)

    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
 
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site


    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgContModeEn, MSV(1)
    bfWrite BF_GEPHY.FgFrmLen, MSV(1500)
    bfWrite BF_GEPHY.FgCntrl, MSV(1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)

    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site


    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_100T_Power_yoda_test = Result.Copy
    Next site

End Function


Public Function Extlpbk_RGMII_10T_Power_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_10T_Power_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 5, DspLong)
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)

    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
 
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site


    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgContModeEn, MSV(1)
    bfWrite BF_GEPHY.FgFrmLen, MSV(1500)
    bfWrite BF_GEPHY.FgCntrl, MSV(1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)

    
    
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site

    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_10T_Power_yoda_test = Result.Copy
    Next site

End Function

Public Function Extlpbk_MII_100T_Power_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_MII_100T_Power_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 5, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H1)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)

    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H0)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
 
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site


    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgContModeEn, MSV(1)
    bfWrite BF_GEPHY.FgFrmLen, MSV(1500)
    bfWrite BF_GEPHY.FgNfrmL, MSV(1024)
    bfWrite BF_GEPHY.FgCntrl, MSV(1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)


    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
     
     
    For Each site In TheExec.Sites.Active
        Extlpbk_MII_100T_Power_yoda_test = Result.Copy
    Next site

End Function

Public Function Extlpbk_MII_10T_Power_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_MII_10T_Power_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 5, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)

    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H0)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
 
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site


    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgContModeEn, MSV(1)
    bfWrite BF_GEPHY.FgFrmLen, MSV(1500)
    bfWrite BF_GEPHY.FgNfrmL, MSV(1024)
    bfWrite BF_GEPHY.FgCntrl, MSV(1)
    bfWrite BF_GEPHY.FgEn, MSV(&H0)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)


    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site

    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site

    For Each site In TheExec.Sites.Active
        Extlpbk_MII_10T_Power_yoda_test = Result.Copy
    Next site

End Function
Public Function Extlpbk_GMII_1000T_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_GMII_1000T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.AutoMdiEn, MSV(&H0)
    bfWrite BF_GEPHY.ManMdix, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.LbRemoteEn, MSV(&H0)
    bfWrite BF_GEPHY.Loopback, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.IsolateRx, MSV(&H0)
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)
    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H0)
    
    bfWrite BF_GESUB.GeMiiUseGtxClk, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    Read = bfRead(BF_GEPHY.SftPd)
    
    'Read Link status
    GEPHY_MDIO_Wait 150 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    

    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(1250)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
   

    For Each site In TheExec.Sites.Active
        Extlpbk_GMII_1000T_yoda_test = Result.Copy
    Next site
    
End Function


Public Function Extlpbk_RGMII_1000T_yoda_test() As DSPWave

Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_1000T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H1)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    bfWrite BF_GEPHY.ManMstrAdv, MSV(&H1)
    bfWrite BF_GEPHY.ManMstrSlvEnAdv, MSV(&H1)
    bfWrite BF_GESUB.GeRxTxExtLbEn, MSV(&H1)
    
    bfWrite BF_GEPHY.MsqlSkip, MSV(&H1)
    
    bfWrite BF_GEPHY.Dfe1gCfg0Coef0, MSV(&H0)
    bfWrite BF_GEPHY.Dfe1gCfg1Coef0, MSV(&H0)
    bfWrite BF_GEPHY.ClenEstEn, MSV(&H0)
    
    bfWrite BF_GEPHY.CanNcycm1Ndd2, MSV(&H3)
    bfWrite BF_GEPHY.CanNcycm1Dd1, MSV(&H0)
    bfWrite BF_GEPHY.CanNcycm1Dd2, MSV(&H0)
    bfWrite BF_GEPHY.CanNcycm1Fb1, MSV(&H3)
    bfWrite BF_GEPHY.CanNcycm1Fb2, MSV(&H0)
    
    bfWrite BF_GEPHY.PgaLvlFixEn, MSV(&H1)
    bfWrite BF_GEPHY.PgaLvlFval, MSV(&H1F)
    bfWrite BF_GEPHY.FreqOffsFrcVal, MSV(3355)
    bfWrite BF_GEPHY.FreqOffsFrcEn, MSV(&H1)
    bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    'Read Link status
    GEPHY_MDIO_Wait 100 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)

    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(1250)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 1 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
''     'Debug code
''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
''     Next site


    

    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_1000T_yoda_test = Result.Copy
    Next site

End Function


Public Function Extlpbk_RGMII_100T_yoda_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_100T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H1)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    
    'bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site
    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(1250)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
''     'Debug code
''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
''     Next site
     
        
    
    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_100T_yoda_test = Result.Copy
    Next site

End Function


Public Function Extlpbk_RGMII_10T_yoda_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Extlpbk_RGMII_10T_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 6, DspLong)
    
    'Setups
    bfWrite BF_GEPHY.SftPdPllPdEn, MSV(&H0)
    bfWrite BF_GEPHY.SftPd, MSV(&H1)
    bfWrite BF_GEPHY.AutonegEn, MSV(&H0)
    bfWrite BF_GEPHY.SpeedSelLsb, MSV(&H0)
    bfWrite BF_GEPHY.DplxMode, MSV(&H1)
    bfWrite BF_GEPHY.SpeedSelMsb, MSV(&H0)
    bfWrite BF_GEPHY.LbAllDigSel, MSV(&H0)
    bfWrite BF_GEPHY.LbExtEn, MSV(&H1)
    
    bfWrite BF_GEPHY.BrkLnkFrc, MSV(&H1)
    
    'bfWrite BF_GEPHY.LinkEn, MSV(&H1)
    bfWrite BF_GEPHY.SftPd, MSV(&H0)
    
    bfWrite BF_GEPHY.LnkPulDlyInact, MSV(262)
    bfWrite BF_GEPHY.MsTickN100usm1Emi, MSV(&H0)
    bfWrite BF_GEPHY.P1msTickNusm1Emi, MSV(99)
    bfWrite BF_GEPHY.UsTickN40nsm1Emi, MSV(&H11)
    
 
    'Read Link status
    GEPHY_MDIO_Wait 50 * ms
    Read = bfRead(BF_GEPHY.LinkStatLat)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.HcdTech)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site
    
    bfWrite BF_GEPHY.MsTickN100usm1Emi, MSV(&H9)
    bfWrite BF_GEPHY.P1msTickNusm1Emi, MSV(63)
    bfWrite BF_GEPHY.UsTickN40nsm1Emi, MSV(&H18)
    
    bfWrite BF_GESUB.GeRgmiiEn, MSV(&H1)
    
    'Configuring frame generator to send packets
    bfWrite BF_GEPHY.FgFrmLen, MSV(125)
    bfWrite BF_GEPHY.FgNfrmL, MSV(10)
    
    bfWrite BF_GEPHY.FcTxSel, MSV(&H1)
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H1)
    bfWrite BF_GEPHY.FgEn, MSV(&H1)
    
     
    'Read FG status
    GEPHY_MDIO_Wait 20 * ms
    Read = bfRead(BF_GEPHY.FgDone)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
    
    bfWrite BF_GEPHY.DiagClkEn, MSV(&H0)
    
    'Configuring GeSubsys for RGMII mode with GeRxTxExtLbEn = 1, GeRgmiiTxIdEn = 0 & GeRgmiiRxIdEn = 0
 
    Read = bfRead(BF_GEPHY.RxErrCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntH)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site
    
    Read = bfRead(BF_GEPHY.FcFrmCntL)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
''     'Debug code
''     Read = bfRead(BF_GESUB.GeRgmiiRxIdEn)   '1
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiRxSel)    '7
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiRxSel: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiTxIdEn)    '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiTxIdEn: " & Read(site)
''     Next site
''
''     Read = bfRead(BF_GESUB.GeRgmiiGtxSel)   '0
''     For Each site In theexec.Sites.Active
''        theexec.Datalog.WriteComment "GeRgmiiGtxSel: " & Read(site)
''     Next site
     
        
    
    For Each site In TheExec.Sites.Active
        Extlpbk_RGMII_10T_yoda_test = Result.Copy
    Next site

End Function


Public Function Process_Monitor_yoda_test() As DSPWave
Dim Read As New SiteLong
Dim Result As New DSPWave
Dim site As Variant

    Set Process_Monitor_yoda_test = New DSPWave
    
    Call Result.CreateConstant(0, 9, DspLong)
    
    'Setups
    
    bfWrite BF_GESUB.GePmRoSel, MSV(&H0) 'Select RVT ring Oscillator
    bfWrite BF_GESUB.GePmReq, MSV(&H1)   'Request Process Monitor Count
    
    'Get PM Results
    Read = bfRead(BF_GESUB.GePmCntDone)
     For Each site In TheExec.Sites.Active
        Result.Element(0) = Read
     Next site
    Read = bfRead(BF_GESUB.GePmCntErr)
     For Each site In TheExec.Sites.Active
        Result.Element(1) = Read
     Next site

    Read = bfRead(BF_GESUB.GePmCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(2) = Read
     Next site
     

    
    bfWrite BF_GESUB.GePmRoSel, MSV(&H1) 'Select LVT ring Oscillator
    bfWrite BF_GESUB.GePmReq, MSV(&H1)   'Request Process Monitor Count
    
    'Get PM Results
    Read = bfRead(BF_GESUB.GePmCntDone)
     For Each site In TheExec.Sites.Active
        Result.Element(3) = Read
     Next site
    Read = bfRead(BF_GESUB.GePmCntErr)
     For Each site In TheExec.Sites.Active
        Result.Element(4) = Read
     Next site

    Read = bfRead(BF_GESUB.GePmCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(5) = Read
     Next site
     
    bfWrite BF_GESUB.GePmRoSel, MSV(&H2) 'Select HVT ring Oscillator
    bfWrite BF_GESUB.GePmReq, MSV(&H1)   'Request Process Monitor Count
    
    'Get PM Results
    Read = bfRead(BF_GESUB.GePmCntDone)
     For Each site In TheExec.Sites.Active
        Result.Element(6) = Read
     Next site
    Read = bfRead(BF_GESUB.GePmCntErr)
     For Each site In TheExec.Sites.Active
        Result.Element(7) = Read
     Next site

    Read = bfRead(BF_GESUB.GePmCnt)
     For Each site In TheExec.Sites.Active
        Result.Element(8) = Read
     Next site
        

    For Each site In TheExec.Sites.Active
        Process_Monitor_yoda_test = Result.Copy
    Next site

End Function
