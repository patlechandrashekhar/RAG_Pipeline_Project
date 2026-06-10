VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDIO_DSSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Description:
'   MDIO Transaction Class using DSSC Instrument for UltraFlex
'
'   Requirements:
'   -  Require DSSC Pattern with MDC, MDIO and Reset defined. DSSC Transactions
'   happens through MDIO
'   -  Pin Name is Hardcoded to "MDIO" in the class
'
'   Rev 1.0 (vsomasun, Feb 8th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim MDIOWrDSSC_Src As New DSPWave
Dim MDIORdDSSC_Src As New DSPWave
Dim MDIO_Pin As String
Dim DSSC_Pattern As String
Private Init_Engine As Long
Private En_Keepalive As Boolean


Public Enum Wr_OP_DSSC
    Address = 0
    Wrt = 1
End Enum

Public Enum Rd_OP_DSSC
    Read_Addr = 2
    Rd_Reg = 3
End Enum


Public Function Init(Pattern_Name As String) As Long
Dim site As Variant

    MDIO_Pin = "MDIO"
    DSSC_Pattern = Pattern_Name

    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    
    thehdw.Patterns(DSSC_Pattern).Load
    
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 'DSSC Source Setup
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For Each site In TheExec.Sites.Active
    
        Call MDIOWrDSSC_Src.CreateConstant(0, 30, DspLong)
        
        'Create Wave Defn and Add Signal Name
        Call TheExec.WaveDefinitions.CreateWaveDefinition("MDIO_Write_SrcWavedef", MDIOWrDSSC_Src, True)
        Call thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Add("MDIOWrDSSC_Src_Sig")
        
        'Bind Waveform DSSC Source to dspwave
        With thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Item("MDIOWrDSSC_Src_Sig")
            .WaveDefinitionName = "MDIO_Write_SrcWavedef"
            .SampleSize = 30
            .SampleRate = 5 * MHz
        End With
    
    Next site
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 'DSSC Capture Setup
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Capture.Signals.Add "MDIORdDSSC_Cap_Sig"
    With thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Capture.Signals("MDIORdDSSC_Cap_Sig")
        .SampleSize = 16
        .SampleRate = 5 * MHz
        .LoadSettings
    End With
    
    'Start the Pattern and get in to idle loop
    thehdw.Patterns(DSSC_Pattern).Load
    thehdw.Patterns(DSSC_Pattern).Start
    Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
    
'    thehdw.Digital.Patgen.Continue 0, cpuA
    
'    Call thehdw.Digital.Patgen.HaltWait

    Init_Engine = 1

End Function


Private Function Reg_Write(st As Long, op As Long, phyadd As Long, reg_dev As Long, Data As Long) As Long
Dim i As Long
Dim temp_data As New SiteLong
Dim temp_wave As New DSPWave
Dim Wr_Data(30) As Long

    
    Call temp_wave.CreateConstant(0, 1, DspLong)
    
    temp_wave.Element(0) = st * (2 ^ 28) + op * (2 ^ 26) + phyadd * (2 ^ 21) + reg_dev * (2 ^ 16) + Data
    
    MDIOWrDSSC_Src = temp_wave.ConvertStreamTo(tldspSerial, 30, 0, Bit0IsLsb).Copy
    
    Call TheExec.WaveDefinitions.CreateWaveDefinition("MDIO_Write_SrcWavedef", MDIOWrDSSC_Src, True)
    Call thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Add("MDIOWrDSSC_Src_Sig")
    
    With thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Item("MDIOWrDSSC_Src_Sig")
        .WaveDefinitionName = "MDIO_Write_SrcWavedef"
        .SampleSize = 30
        .SampleRate = 5 * MHz
        .LoadSettings
        .LoadSamples
    End With
    
    'Set the Reg Write Module Name and clear cpu flag to move to next srm module
    Call thehdw.Digital.Patgen.ModuleJump("srm_dssc_mdio_reg_write")
    Call thehdw.Digital.Patgen.Continue(0, cpuA)
    Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
    
 

End Function


Private Function Reg_Read(st As Long, op As Long, phyadd As Long, reg_dev As Long) As SiteLong

Dim temp_wave As New DSPWave
Dim Rd_Data As New DSPWave
Dim Temp As New SiteLong
Dim site As Variant

    Set Reg_Read = New SiteLong
    
    Call temp_wave.CreateConstant(0, 1, DspLong)
    
    'Create DSSC Read Array
    temp_wave.Element(0) = st * (2 ^ 12) + op * (2 ^ 10) + phyadd * (2 ^ 5) + reg_dev
    
    MDIOWrDSSC_Src = temp_wave.ConvertStreamTo(tldspSerial, 14, 0, Bit0IsLsb).Copy
    

    Call TheExec.WaveDefinitions.CreateWaveDefinition("MDIO_Write_SrcWavedef", MDIOWrDSSC_Src, True)
    Call thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Add("MDIOWrDSSC_Src_Sig")
    
    With thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Source.Signals.Item("MDIOWrDSSC_Src_Sig")
        .WaveDefinitionName = "MDIO_Write_SrcWavedef"
        .SampleSize = 14
        .SampleRate = 5 * MHz
        .LoadSettings
        .LoadSamples
    End With
    
    
    'Set the Reg Write Module Name and clear cpu flag to move to next srm module
    Call thehdw.Digital.Patgen.ModuleJump("srm_dssc_mdio_reg_read")
    Call thehdw.Digital.Patgen.Continue(0, cpuA)
    Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
 
    Rd_Data = thehdw.DSSC.Pins(MDIO_Pin).Pattern(DSSC_Pattern).Capture.Signals("MDIORdDSSC_Cap_Sig").DSPWave.Copy
    
    For Each site In TheExec.Sites.Active
        Temp = Rd_Data.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsLsb).Element(0)
    Next site
    
    Reg_Read = Temp
    
    Exit Function

    
    
End Function
Public Function CL22_Reg_Write(phyadd As Long, regAdd As Long, Data As Long) As Long

    'In Hexadecimal Declaration, If bit 15 is high Data becomes negative
    'Convert Data to Type Long
    If Data < 0 Then
        Data = (Data And &H7FFF) + (2 ^ 15)
    End If
    
    If Init_Engine = 1 Then
        Call Reg_Write(&H1, &H1, phyadd, regAdd, Data)
    Else
        TheExec.Datalog.WriteComment "MDIO DSSC Engine must be initialized! Use the init function before transactions"
    End If

End Function

Public Function CL22_Reg_Read(phyadd As Long, regAdd As Long) As SiteLong

    Set CL22_Reg_Read = New SiteLong

    If Init_Engine = 1 Then
        CL22_Reg_Read = Reg_Read(&H1, &H2, phyadd, regAdd)
    Else
        TheExec.Datalog.WriteComment "MDIO DSSC Engine must be initialized! Use the init function before transactions"
    End If

    

End Function
Public Function CL45_Reg_Write(phyadd As Long, devtype As Long, Data As Long, op_wr As Wr_OP_DSSC) As Long

    'In Hexadecimal Declaration, If bit 15 is high Data becomes negative
    'Convert Data to Type Long
    If Data < 0 Then
        Data = (Data And &H7FFF) + (2 ^ 15)
    End If

    If Init_Engine = 1 Then
        Call Reg_Write(&H0, op_wr, phyadd, devtype, Data)
    Else
        TheExec.Datalog.WriteComment "MDIO DSSC Engine must be initialized! Use the init function before transactions"
    End If
    

End Function

Public Function CL45_Reg_Read(phyadd As Long, devtype As Long, op_rd As Rd_OP_DSSC) As SiteLong

    Set CL45_Reg_Read = New SiteLong

    If Init_Engine = 1 Then
        CL45_Reg_Read = Reg_Read(&H0, op_rd, phyadd, devtype)
    Else
        TheExec.Datalog.WriteComment "MDIO DSSC Engine must be initialized! Use the init function before transactions"
    End If
    

End Function


Public Function Wait(wait_time As Double) As Long

    thehdw.Wait wait_time

End Function

Public Function Halt() As Long

    'Exit Pattern Loop
    Call thehdw.Digital.Patgen.ModuleJump("srm_dssc_mdio_halt")
    Call thehdw.Digital.Patgen.Continue(0, cpuA)
    Call thehdw.Digital.Patgen.HaltWait
    
    Init_Engine = 0
    
    
End Function

Property Get Engine_Type() As String
    Engine_Type = "DSSC"
End Property

Property Get KeepAlive() As Boolean
    KeepAlive = En_Keepalive
End Property

Property Let KeepAlive(Enable As Boolean)
    En_Keepalive = Enable
End Property

