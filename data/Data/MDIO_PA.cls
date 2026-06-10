VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDIO_PA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Description:
'   MDIO Transaction Class using PA Instrument for UltraFlex
'
'   Requirements:
'   -  Require Port definition in IGXL program
'   -  Port Name is Hardcoded to "MDIO_Port" in the class
'   -  Setup pattern is Optional
'   -  If Setup pattern uses Keepalive, Enable Keepalive in Class Object
'
'   Rev 1.0 (vsomasun, Feb 8th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim MDIOWrDSSC_Src As New DSPWave
Dim MDIORdDSSC_Src As New DSPWave
Dim MDIO_Port As String
Dim DSSC_Pattern As String
Private Init_Engine As Long
Private En_Keepalive As Boolean



Public Function Init(Setup_Pattern As String) As Long

    MDIO_Port = "MDIO_Port"

    thehdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, LoadLevels:=True, LoadTiming:=True, RelayMode:=tlPowered
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 'Execute Setup Pattern
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If En_Keepalive = True Then
        thehdw.Digital.Patgen.KeepAlive.flag = cpuA
        thehdw.Digital.Patgen.KeepAlive.Enable = True
        thehdw.Digital.Patgen.Continue cpuA, 0
    End If
    
    If Setup_Pattern <> "" Then
        'Start the Pattern
        thehdw.Patterns(Setup_Pattern).Load
        thehdw.Patterns(Setup_Pattern).Start
        If En_Keepalive = False Then
            Call thehdw.Digital.Patgen.HaltWait
        End If
    End If
    

    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 'Enable Nwire Port
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    thehdw.Protocol.Ports(MDIO_Port).Enabled = True
    thehdw.Protocol.Ports(MDIO_Port).NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
    thehdw.Protocol.Ports(MDIO_Port).NWire.HRAM.Setup.WaitForEvent = False
    
    thehdw.Protocol.Ports(MDIO_Port).NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_MoveLink
    
    Init_Engine = 1

End Function


Private Function Reg_Write(st As Long, op As Long, phyadd As Long, reg_dev As Long, Data As Long) As Long

    With thehdw.Protocol.Ports(MDIO_Port).NWire.Frames("Write")
        .Fields("ST").Value = st
        .Fields("OP").Value = op
        .Fields("PHYAD").Value = phyadd
        .Fields("DEV_REG").Value = reg_dev
        .Fields("DATA").Value = Data
        Call .Execute
    End With
    
    thehdw.Protocol.Ports(MDIO_Port).IdleWait
 

End Function

Private Function Reg_Read(st As Long, op As Long, phyadd As Long, reg_dev As Long, Optional Capture_CMEM As Boolean = True) As SiteLong

Dim Data As New SiteLong
Dim site As Variant
Dim myPLD As New PinListData
Dim Capture_type As tlNWireExecutionType

    'Initialize data to 0
    Data = 0
    
    'Capture Read data in either HRAM or CMEM
    If Capture_CMEM = True Then
        Capture_type = tlNWireExecutionType_CaptureInCMEM
    Else
        Capture_type = tlNWireExecutionType_Default
    End If
    
    With thehdw.Protocol.Ports(MDIO_Port).NWire.Frames("Read")
        .Fields("ST").Value = st
        .Fields("OP").Value = op
        .Fields("PHYAD").Value = phyadd
        .Fields("DEV_REG").Value = reg_dev
        Call .Execute(Capture_type)
    End With
    
    thehdw.Protocol.Ports(MDIO_Port).IdleWait
    
    If Not (thehdw.Protocol.Ports(MDIO_Port).ModuleFiles.IsLoading) Then
    
         'Capture Read data in either HRAM or CMEM
        If Capture_CMEM = True Then
            myPLD = thehdw.Protocol.Ports(MDIO_Port).NWire.CMEM.Transactions.Read()
            
            For Each site In TheExec.Sites.Active
                Dim CMEM_indexes As INWireCMEMTransactionIndexes
                Set CMEM_indexes = myPLD.Pins(MDIO_Port).Value
                Data = CMEM_indexes(0).Fields("DATA").Value
            Next site
        Else
            myPLD = thehdw.Protocol.Ports(MDIO_Port).NWire.HRAM.Transactions.Read()
            
            For Each site In TheExec.Sites.Active
                Dim HRAM_indexes As INWireHRAMTransactionIndexes
                Set HRAM_indexes = myPLD.Pins(MDIO_Port).Value
                Data = HRAM_indexes(0).Fields("DATA").Value
            Next site
        End If
    
    End If
    
    Set Reg_Read = Data
    
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
        TheExec.Datalog.WriteComment "MDIO PA Engine must be initialized! Use the init function before transactions"
    End If


End Function

Public Function CL22_Reg_Read(phyadd As Long, regAdd As Long, Optional Capture_CMEM As Boolean = True) As SiteLong

    Set CL22_Reg_Read = New SiteLong

    If Init_Engine = 1 Then
        CL22_Reg_Read = Reg_Read(&H1, &H2, phyadd, regAdd, Capture_CMEM)
    Else
        TheExec.Datalog.WriteComment "MDIO PA Engine must be initialized! Use the init function before transactions"
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
        TheExec.Datalog.WriteComment "MDIO PA Engine must be initialized! Use the init function before transactions"
    End If

    

End Function

Public Function CL45_Reg_Read(phyadd As Long, devtype As Long, op_rd As Rd_OP_DSSC, Optional Capture_CMEM As Boolean = True) As SiteLong

    Set CL45_Reg_Read = New SiteLong
    
    If Init_Engine = 1 Then
       CL45_Reg_Read = Reg_Read(&H0, op_rd, phyadd, devtype, Capture_CMEM)
    Else
        TheExec.Datalog.WriteComment "MDIO PA Engine must be initialized! Use the init function before transactions"
    End If

    

End Function


Public Function Wait(wait_time As Double) As Long

Dim wait_loop As Integer, i As Integer
Dim remainder As Double

    'Nwire has a maximum timeout of 125 ms.
    'Wait time over 125ms needs to be repeated
    If wait_time > 0.125 Then
        wait_loop = Floor(wait_time / 0.125)
        remainder = wait_time - (wait_loop * 0.125)
        
        For i = 0 To wait_loop - 1
            thehdw.Protocol.Ports(MDIO_Port).NWire.Wait (0.125)
            thehdw.Protocol.Ports(MDIO_Port).IdleWait
        Next i
        
        If remainder <> 0 Then
            thehdw.Protocol.Ports(MDIO_Port).NWire.Wait (remainder)
        End If
    Else
        thehdw.Protocol.Ports(MDIO_Port).NWire.Wait (wait_time)
    End If
    
    thehdw.Protocol.Ports(MDIO_Port).IdleWait
        

End Function

Public Function Halt() As Long

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 'Disable Nwire Port
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    thehdw.Protocol.Ports(MDIO_Port).Enabled = False

    'Exit Pattern
    Call thehdw.Digital.Patgen.Halt
    
    If En_Keepalive = True Then
        
        thehdw.Digital.Patgen.KeepAlive.Enable = False
        thehdw.Digital.Patgen.Continue 0, cpuA

    End If
    
    Init_Engine = 0
    
    
End Function

Property Get Engine_Type() As String
    Engine_Type = "PA"
End Property

Property Get KeepAlive() As Boolean
    KeepAlive = En_Keepalive
End Property

Property Let KeepAlive(Enable As Boolean)
    En_Keepalive = Enable
End Property
