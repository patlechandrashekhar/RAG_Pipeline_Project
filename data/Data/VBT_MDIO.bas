Attribute VB_Name = "VBT_MDIO"
Dim MDIO As New MDIO_DSSC

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module Description:
'   Module to use MDIO class to facilitate MDIO transactions
'
'   Usage:
'   -  Pattern Path for MDIO class needs to be passed through init function
'   -  Change the Class to switch between PA and DSSC. To use PA, use class
'   "MDIO_PA". To use DSSC, use class "MDIO_DSSC"
'
'   Rev 1.0 (vsomasun, Feb 8th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GEPHY_MDIO_Init()

    If MDIO.Engine_Type = "DSSC" Then
        'Call DSSC Pattern
        Call MDIO.Init("./Patterns/DSSC_MDIO_Transaction.PAT")
    Else
        'Call PA Pattern
        MDIO.KeepAlive = True
        Call MDIO.Init("./Patterns/PA_MDIO_Transaction.PAT")
    End If
    
    

End Function
Public Function GEPhy_MDIO_CL_45_Write(phyadd As Long, Dev_Addr As Long, Addr As Long, Data As Long) As Long

    Call MDIO.CL45_Reg_Write(phyadd, Dev_Addr, Addr, Address)
    Call MDIO.CL45_Reg_Write(phyadd, Dev_Addr, Data, Wrt)

End Function

Public Function GEPhy_MDIO_CL_45_Read(phyadd As Long, Dev_Addr As Long, Addr As Long) As SiteLong

Set GEPhy_MDIO_CL_45_Read = New SiteLong

    Call MDIO.CL45_Reg_Write(phyadd, Dev_Addr, Addr, Address)
    GEPhy_MDIO_CL_45_Read = MDIO.CL45_Reg_Read(phyadd, Dev_Addr, Rd_Reg)
    
End Function

Public Function GEPhy_MDIO_CL_22_Write(phyadd As Long, Addr As Long, Data As Long) As Long

    Call MDIO.CL22_Reg_Write(phyadd, Addr, Data)

End Function

Public Function GEPhy_MDIO_CL_22_Read(phyadd As Long, Addr As Long) As SiteLong

Set GEPhy_MDIO_CL_22_Read = New SiteLong

    GEPhy_MDIO_CL_22_Read = MDIO.CL22_Reg_Read(phyadd, Addr)
    

End Function

Public Function GEPHY_MDIO_Wait(wait_time As Double)

    Call MDIO.Wait(wait_time)

End Function

Public Function GEPHY_MDIO_Halt()

    Call MDIO.Halt

End Function
