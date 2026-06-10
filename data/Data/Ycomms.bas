Attribute VB_Name = "Ycomms"
Option Explicit
'''Communications with the DUT via Yoda object are managed in this module.
'''
'''
'A communications stack is implemented as: Yoda object layer, intermediate layer, SPI (or other protocol) layer.
'Note .Value is the value held by Yoda object (software value).
'Note .readValue is the value read from the DUT (hardware value).

Public Const DUT As Boolean = True    'Use to select the DUT as the read/ write target.
Public Const YDA As Boolean = False    'Use to select the Yoda object as the read/write target.
Public Const CONFIRM As Boolean = True    'Use to force A readback and confirm after write.
Public Const COMPARE As Boolean = True    'Use to perform A comparison with A previous system snapshot.

'DUT specific constants and variables may be declared here.
Public Const logdir As String = "C:\temp\"    'Set this to preferred directory for datalogs.

'Set pin names and pattern signal names.
Private Const writePinName As String = "HSD_MOSI"
Private Const readPinName As String = "HSD_MISO"
Private Const sigNameCap As String = "digcapture"

'########################################################################################
'========================================================================================
'Yoda object layer
'
'Facilitates reading and writing to the part using high level Yoda object functions.
'========================================================================================
'########################################################################################

Public Function bfWrite(iBF As IBitField, Optional wrval As SiteLong, Optional writeDUT As Boolean = True, Optional verifyWrite As Boolean = False) As SiteLong
'''Write A value to the specified bitfield on the DUT. Optionally, read back and verify after write.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim iReg As Iregister
    Dim rColWrite As RegCollection
    Dim doWrite As Boolean
    Dim commStatus As New SiteLong

    Set commStatus = MSV(0)
    Set rColWrite = iBF.hostRegisters

    doWrite = False

    If wrval Is Nothing Then
        'When no value to write is provided, use existing Yoda object value.
        Set wrval = MSV(0)
        For Each site In TheExec.Sites
            wrval = iBF.Value
        Next site
        doWrite = True
    End If

    For Each site In TheExec.Sites
        'Check value to write is within permissible bounds.
        If wrval >= 0 And wrval < (2 ^ iBF.Width) Then
            'Update Yoda object bitfield to specified value.
            iBF.Value = wrval
        Else
            Err.Raise Number:=vbObjectError + 1300, Source:="bfWrite", Description:="invalid: " & iBF.Name & " Value " & wrval
        End If
    Next site

    If writeDUT Then
        For Each iReg In rColWrite
            'Write to each register hosting bits from the bitfield. Optionally, skip writing
            'where value is unchanged to save time. This also avoids any possible VBA stack overflow.
            'For Each site In TheExec.Sites
            'If (iReg.Value <> iReg.readValue) Then
            doWrite = True
            'End If
            'Next site


            If doWrite Then
                'Write to DUT register.
                interRegWrite iReg, verifyWrite

                If verifyWrite = True Then
                    'Check the register contents match the write.
                    For Each site In TheExec.Sites
                        If iReg.readValue <> iReg.Value Then
                            iReg.Changed = True
                            commStatus = commStatus + 1
                            'TheExec.dataLog.WriteComment "!!! Write to register " & ireg.Name & " failed!"
                        End If
                    Next site
                End If
            End If
        Next iReg
    End If

    'Return pass|fail status of this write.
    Set bfWrite = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "bfWrite " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function bfRead(iBF As IBitField, Optional DUTread As Boolean = True) As SiteLong
'''Read the value of the specified bitfield, from the DUT, or from Yoda object if A DUT read is not required.
'''
    On Error GoTo 0
    Dim iReg As Iregister
    Dim rColRead As RegCollection
    Dim site As Variant
    Dim bfVal As New SiteLong
    Set rColRead = iBF.hostRegisters

    If DUTread = True Then
        'Read from the DUT
        For Each iReg In rColRead
            'Read from each register hosting bits from specified bitfield.
            Call interRegRead(iReg, DUTread)
        Next iReg
    End If

    'Return bitfield contents
    For Each site In TheExec.Sites
        'Return the value read from the DUT, or the Yoda object value if no DUT read occurred.
        bfVal = iBF.readValue
    Next site
    Set bfRead = bfVal

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "bfRead Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function regWrite(iReg As Iregister, Optional wrval As SiteLong, Optional verifyWrite As Boolean = False) As SiteLong
'''Update the specified register value in Yoda object, then write it to the DUT.
'''Use bfWrite instead if possible. regWrite should only be needed in the fuseblow routine.
'''
    On Error GoTo 0
    Const noVer As Boolean = False
    Dim site As Variant
    Dim commStatus As New SiteLong

    Set commStatus = MSV(0)

    If wrval Is Nothing Then
        'When no value to write is provided, use existing Yoda object value.
        Set wrval = MSV(0)
        For Each site In TheExec.Sites
            wrval = iReg.Value
        Next site
    Else

        For Each site In TheExec.Sites
            'Check value to write is within permissible bounds.
            If wrval >= 0 And wrval < (2 ^ iReg.Width) Then
                'Update Yoda register object to specified value.
                iReg.Value = wrval
            Else
                Err.Raise Number:=vbObjectError + 1310, Source:="regWrite", Description:="invalid: " & iReg.Name & " Value " & wrval
            End If
        Next site

    End If

    'Write the Yoda object value to the DUT register.
    interRegWrite iReg, verifyWrite

    If verifyWrite = True Then
        'Read back the register from the DUT and check the contents match the write.
        For Each site In TheExec.Sites
            If iReg.readValue <> iReg.Value Then
                iReg.Changed = True
                commStatus = commStatus + 1
                'TheExec.dataLog.WriteComment "!!! Write to register " & ireg.Name & " failed!"
            End If
        Next site
    End If

    'Return pass|fail status of this write.
    Set regWrite = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "regWrite Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function regRead(iReg As Iregister, Optional DUTread As Boolean = True) As SiteLong
'''Read the value of the specified register, from the DUT, or from Yoda object if A DUT read is not required.
'''Use bfRead instead if possible. regRead should only be needed in the fuseblow routine.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim msread As New SiteLong
    Dim regResult As New SiteLong

    If DUTread = True Then
        'Read from the DUT register
        Call interRegRead(iReg, DUTread)
    End If

    'Return the value read from DUT, or the Yoda object value if no read from the DUT occurred.
    For Each site In TheExec.Sites
        regResult = iReg.readValue
    Next site

    Set regRead = regResult

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "regRead Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function mapWrite(iRegMap As IRegisterMap, Optional verifyWrite As Boolean = False) As SiteLong
'''Write all Yoda object register values in the specified map-collection to the DUT. Optionally, read back and verify after write.
'''This function does not facilitate control of the order the registers are written to.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim commStatus As New SiteLong
    Dim iReg As Iregister
    Dim iRegCol As RegCollection
    Dim numOfRegs As Long
    Set commStatus = MSV(0)


    'Write Yoda object values which have changed to DUT registers.
    Set iRegCol = iRegMap.registers.FilterByStatus(True)
    numOfRegs = iRegCol.Count

    If numOfRegs > 0 Then

        interRegCollWrite iRegCol, verifyWrite

        If verifyWrite = True Then
            For Each iReg In iRegCol
                'Check register read matches the value written.
                For Each site In TheExec.Sites
                    If iReg.readValue <> iReg.Value Then
                        TheExec.Datalog.WriteComment "!!! Write to register " & iReg.Name & " failed!"
                        commStatus = commStatus + 1
                    End If
                Next site
            Next iReg
        End If    'verifyWrite
    End If

    Set mapWrite = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapwrite Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function mapRead(iRegMap As IRegisterMap, Optional logToFile As Long = 0) As SiteLong
'''Read all DUT register values in the specified map and store them in the Yoda object readValue field.
'''Note that the mapRead return value will not be 0 if there are any write-only or volatile fields defined.
'''
    On Error GoTo 0

    Dim iReg As Iregister
    Dim site As Variant
    Dim commStatus As New SiteLong
    Dim regContents As New SiteLong
    Dim logname As String
    Dim fileID As Long

    Set commStatus = MSV(0)

    If logToFile > 0 Then
        logname = "mapRead"
        logname = logname & Str(logToFile)
        logname = logname & ".txt"
        fileID = txtFileOpen(logname)
    End If

    interMapRead iRegMap

    'Read register contents from DUT and store in Yoda object
    For Each iReg In iRegMap.registers
        For Each site In TheExec.Sites
            If iReg.readValue <> iReg.Value Then
                iReg.Changed = True
                commStatus = commStatus + 1
                If logToFile > 0 Then txtFileWrite fileID, iReg.Name & vbTab & " is " & iReg.readValue & vbTab & " should be " & iReg.Value
            End If
        Next site
    Next iReg

    If logToFile > 0 Then txtFileClose fileID

    Set mapRead = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapRead Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Sub mapReset(iRegMap As IRegisterMap)
'''Reset the Yoda object readValues in the specified map to their default values and flag as changed.
'''This is useful when the part has been reset and the DUT is out of sync with the Yoda object values.
'''Fuse maps should not be reset in this way after fuse programming.
'''This function does not write to the DUT.
'''
    On Error GoTo 0

    Dim site As Variant
    Dim i As Long
    Dim numRg As Long
    Dim iReg As Iregister

    For Each iReg In iRegMap.registers
        For Each site In TheExec.Sites
            iReg.readValue = iReg.DefaultValue
            iReg.Changed = True
        Next site
    Next iReg

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "mapReset Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Private Sub mapFill(iRegMap As IRegisterMap, fillVal As Long)
'''Fill all Yoda object registers with the same specified value, up to bit width of register.
'''This function dos not write to the DUT.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim i As Long

    Dim numOfRegs As Integer
    Dim regArray() As Long
    Dim mask As Long
    Dim valToWrite As Long

    numOfRegs = iRegMap.registers.Count
    ReDim regArray(0 To numOfRegs - 1)

    For i = 1 To numOfRegs
        mask = 2 ^ iRegMap.registers.Item(i).Width - 1
        valToWrite = fillVal And mask
        regArray(i - 1) = valToWrite
    Next i

    'Fill Yoda object map.
    For Each site In TheExec.Sites
        iRegMap.values = regArray
    Next site

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "mapfill Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Function mapBFMatch(iRegMap As IRegisterMap, Optional logToFile As Long = 0) As SiteLong
'''Check for each Yoda object bitfield: value read from DUT matches Yoda object value.
'''Optionally, log to file.
'''
    On Error GoTo 0

    Dim logname As String
    Dim fileID As Long
    Dim iBF As IBitField
    Dim site As Variant
    Dim baseMsg As String
    Dim errMsg As String
    Dim i As Long
    Dim commStatus As New SiteLong
    Dim numBF As Long

    Set commStatus = MSV(0)

    If logToFile > 0 Then
        logname = "mapBFMatch"
        logname = logname & Str(logToFile)
        logname = logname & ".txt"
        fileID = txtFileOpen(logname)
    End If

    'numBF = MM.bitFields.bitFields.Count

    For i = 1 To numBF
        Set iBF = MM.bitFields(i)

        If iBF.access = BitFieldAccess_ReadWrite And iBF.Volatile = False And iBF.memMapName = iRegMap.Name Then
            For Each site In TheExec.Sites
                baseMsg = "Site " & site & "  " & iBF.Name & " = " & vbTab & Str(iBF.readValue)
                errMsg = ""
                If iBF.readValue <> iBF.Value And TheExec.TesterMode = testModeOnline Then
                    commStatus = commStatus + 1
                    errMsg = "   !!! " & vbTab & " should be " & Str(iBF.Value)
                    TheExec.Datalog.WriteComment baseMsg & errMsg
                    If logToFile > 0 Then txtFileWrite fileID, baseMsg & errMsg
                End If
                TheExec.Datalog.WriteComment "Site " & site & " bitfield match errors= " & commStatus
                If logToFile > 0 Then txtFileWrite fileID, "Site " & site & " bitfield match errors= " & commStatus
            Next site
        End If
    Next i

    If logToFile > 0 Then txtFileClose fileID

    Set mapBFMatch = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapBFMatch Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function mapBFStuckBit(iRegMap As IRegisterMap) As SiteLong
'''Write bit patterns to each bitfield, read back from DUT and verify.
'''Note this only works on R/W bitfields.
'''Exclusions should be added for any bitfield that would reset the part or initiate fuse programming.
'''
    On Error GoTo 0

    Dim tstVal As Long
    Dim site As Variant
    Dim i As Long
    Dim iBF As IBitField
    Dim mask As Long
    Dim readresult As New SiteLong
    Dim baseMsg As String
    Dim logStr As String
    Dim commStatus As New SiteLong
    Dim numBF As Long

    Set commStatus = MSV(0)

    'numBF = MM.bitFields.bitFields.Count

    For i = 1 To numBF
        Set iBF = MM.bitFields(i)
        If iBF.access = BitFieldAccess_ReadWrite And iBF.Volatile = False And iBF.memMapName = iRegMap.Name Then
            mask = 2 ^ iBF.Width - 1
            baseMsg = "Bitfield " & i & " " & iBF.Name & " "

            'Define first bit pattern, write to part and read back.
            tstVal = &HAAAAAAA And mask

            bfWrite iBF, MSV(tstVal)
            Set readresult = bfRead(iBF, True)

            For Each site In TheExec.Sites
                If iBF.readValue <> tstVal And TheExec.TesterMode = testModeOnline Then
                    commStatus = commStatus + 1
                    logStr = baseMsg & " site " & site & vbTab & vbTab & " wrote " & Str(tstVal) & " | " & " read " & Str(readresult)
                    TheExec.Datalog.WriteComment logStr

                End If
            Next site

            'Define second bit pattern, write to part and read back.
            tstVal = &H5555555 And mask
            bfWrite iBF, MSV(tstVal)
            readresult = bfRead(iBF, True)

            For Each site In TheExec.Sites
                If iBF.readValue <> tstVal And TheExec.TesterMode = testModeOnline Then
                    commStatus = commStatus + 1
                    logStr = baseMsg & " site " & site & vbTab & vbTab & " wrote " & Str(tstVal) & " | " & " read " & Str(readresult)
                    TheExec.Datalog.WriteComment logStr

                End If
            Next site

        End If    'R|W bitfields only.
    Next i

    For Each site In TheExec.Sites
        TheExec.Datalog.WriteComment "Site " & site & " bitfield stuck errors= " & commStatus
    Next site

    Set mapBFStuckBit = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapBFStuckBit Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function mapRegMatch(iRegMap As IRegisterMap, Optional logToFile As Long = 0) As SiteLong
'''Check for each Yoda object register: value read from DUT matches Yoda object value.
'''Optionally, log  to file.
'''Note that only registers with all bitfields = R/W will pass.
'''
    On Error GoTo 0

    Dim logname As String
    Dim fileID As Long
    Dim iReg As Iregister
    Dim site As Variant
    Dim baseMsg As String
    Dim errMsg As String
    Dim DUTregErrCount As New SiteLong

    For Each site In TheExec.Sites
        DUTregErrCount = 0
    Next site

    If logToFile > 0 Then
        logname = "mapRegMatch"
        logname = logname & Str(logToFile)
        logname = logname & ".txt"
        fileID = txtFileOpen(logname)
    End If


    For Each site In TheExec.Sites
        For Each iReg In iRegMap.registers
            baseMsg = iReg.Name & "(Addr: " & iReg.Address & ") = " & " reads " & vbTab & Str(iReg.readValue)
            errMsg = ""
            If iReg.readValue <> iReg.Value And TheExec.TesterMode = testModeOnline Then
                DUTregErrCount = DUTregErrCount + 1
                errMsg = "   !!! should read " & Str(iReg.Value)
                TheExec.Datalog.WriteComment baseMsg & errMsg
                If logToFile > 0 Then txtFileWrite fileID, baseMsg & errMsg
            End If
        Next iReg
        TheExec.Datalog.WriteComment "Site " & site & " register match errors= " & DUTregErrCount
        If logToFile > 0 Then txtFileWrite fileID, "Site " & site & " register match errors= " & DUTregErrCount
    Next site

    If logToFile > 0 Then txtFileClose fileID

    Set mapRegMatch = DUTregErrCount

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapRegMatch Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function mapRegStuckBit(iRegMap As IRegisterMap) As SiteLong
'''Check all DUT registers can be read and written to.
'''Check for stuck bits, by writing A checkerboard pattern of alternating 1010 values, reading back and verifying.
'''Note this test will report A fail on any registers which contain read-only or write-only bitfields.
'''
    On Error GoTo 0
    Dim regErrs As New SiteLong
    Set regErrs = MSV(0)

    mapFill iRegMap, &H55555555
    mapWrite iRegMap, True
    mapRead iRegMap
    Set regErrs = regErrs.Add(mapRegMatch(iRegMap, &H55555555))

    mapFill iRegMap, &HAAAAAAAA
    mapWrite iRegMap, True
    mapRead iRegMap
    Set regErrs = regErrs.Add(mapRegMatch(iRegMap, &HAAAAAAAA))

    Set mapRegStuckBit = regErrs

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "mapRegStuckBit Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function chipWriteChanged(Optional verifyWrite As Boolean = False) As SiteLong
'''For each DUT register in all maps: If there has been A Yoda object change since last write, write the current Yoda object value to the DUT.
'''Optionally, verify register contents after write.
'''This function does not facilitate control of the order the registers are written to.
'''Change Yoda object bitfield / register values as required, then call this function to perform A bulk update of the DUT.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim commStatus As New SiteLong
    Dim iReg As Iregister
    Dim doWrite As Boolean
    Dim rcol As RegCollection
    Dim iRegMap As IRegisterMap

    Set commStatus = MSV(0)

    For Each iRegMap In MM.registerMaps
        'Debug.Print iRegMap.memMapName
        If iRegMap.Changed Then
            Set rcol = iRegMap.registers.FilterByStatus(True)
            interRegCollWrite rcol, verifyWrite
        End If
    Next iRegMap


    If verifyWrite = True Then
        For Each iReg In chipInstance.changedRegisters
            'Check register read matches the value written.
            For Each site In TheExec.Sites
                If iReg.readValue <> iReg.Value Then
                    TheExec.Datalog.WriteComment "!!! Write to register " & iReg.Name & " failed!"
                    commStatus = commStatus + 1
                End If
            Next site
        Next iReg
    End If    'verifyWrite

    Set chipWriteChanged = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "chipWriteChanged Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function snapshot(Optional diff As Boolean = False, Optional iRegMap As IRegisterMap) As SiteLong
'''This is A debug tool to compare snapshots of the system state.
'''When diff = False, read each bitfield from part, store value.
'''When diff = True, read each bitfield from part, compare with previously stored value.
'''If no map name is passed, all maps are included in the snapshot.
'''
    On Error GoTo 0

    Dim site As Variant
    Dim i As Long
    Dim iBF As IBitField
    Dim diffStatus As New SiteLong
    Dim numBF As Long
    Static snapStore() As New SiteLong

    Set diffStatus = MSV(0)
    'numBF = MM.bitFields.bitFields.Count

    ReDim snapStore(numBF)

    For i = 1 To numBF
        Set iBF = MM.bitFields(i)
        If iBF.memMapName = iRegMap.Name Or iRegMap Is Nothing Then
            Call bfRead(iBF, True)

            If Not diff Then
                Set snapStore(i) = bfRead(iBF, True)
            Else
                For Each site In TheExec.Sites
                    If iBF.readValue <> snapStore(i) Then
                        diffStatus = diffStatus + 1
                        TheExec.Datalog.WriteComment "Difference on site " & site & " " & iBF.Name & " was " & snapStore(i) & vbTab & " is " & iBF.readValue
                    End If
                Next site
            End If

        End If
    Next i

    Set snapshot = diffStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "bfStuckBit Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

'########################################################################################################
'========================================================================================================
'Intermediate layer
'
'Translates between Yoda object level and data frames, which can then be sent using the appropriate protocol.
'Some part-specific customisation may be needed in this section.
'
'========================================================================================================
'########################################################################################################

Public Function interRegWrite(iReg As Iregister, Optional verifyWrite As Boolean = False) As SiteLong
'''Call appropriate write function for ths register.
'''Verify write succeeded if required.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim commStatus As New SiteLong
    Dim regAddr As Long, devAddr As Long
    Dim regValToWrite() As Long    'Use array, not SiteLong
    Dim rMapName As String
    Dim modAddr As New SiteLong
    Dim actSite As Long

    ReDim regValToWrite(0 To TheExec.Sites.Existing.Count - 1)
    Set commStatus = MSV(0)

    'Translate Yoda object to address and data to write.
    regAddr = iReg.Address   'iReg.relativeAddress
    regAddr = iReg.Address And &HFFFF&     'Use 16 LSBs.
    devAddr = (iReg.Address And &H1F0000) / (2 ^ 16)      'Use 5 MSBs.
    For Each site In TheExec.Sites
        regValToWrite(site) = iReg.Value
    Next site

    'DEBUG: to record all bits exercised by the program and identify any coverage gaps:
    'record A 1 for each bit position which has been written with A value different to its default value.
    'writeArr(ireg.Address) = writeArr(ireg.Address) Or (ireg.value(0) Xor ireg.DefaultValue)

    'Perform any writes needed to set up access to the RegMap.
    rMapName = iReg.regMapName    'Get name of RegMap this register is in.
 

    Select Case rMapName
        'Part specific comms selection.
    Case "GESubsys"
        For Each site In TheExec.Sites
            Call GEPhy_MDIO_CL_45_Write(&H0, devAddr, regAddr, regValToWrite(site))
        Next site
    Case Else
        If ((devAddr = 0) And (regAddr < 32)) Then
            For Each site In TheExec.Sites
                Call GEPhy_MDIO_CL_22_Write(&H0, regAddr, regValToWrite(site))
            Next site
        Else
            For Each site In TheExec.Sites
                Call GEPhy_MDIO_CL_45_Write(&H0, devAddr, regAddr, regValToWrite(site))
            Next site
        End If
    End Select


    If TheExec.TesterMode = testModeOffline Then
        'Not checking if write to DUT succeeded.
        For Each site In TheExec.Sites
            iReg.readValue = iReg.Value
            iReg.Changed = False
        Next site
    ElseIf verifyWrite Then
        'Check register readback matches value written.
        Call interRegRead(iReg, True)
        For Each site In TheExec.Sites
            If iReg.readValue = iReg.Value Then
                iReg.Changed = False
            Else
                iReg.Changed = True
                commStatus = commStatus + 1
                TheExec.Datalog.WriteComment "!!! Write to register " & iReg.Name & " failed!"
            End If
        Next site
    End If    'verifyWrite

    Set interRegWrite = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "interRegWrite Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function interRegCollWrite(iRegColl As RegCollection, Optional verifyWrite As Boolean = False) As SiteLong
'''Write all changed registers to the DUT.
'''Verify write succeeded if required.
'''
    On Error GoTo 0
    Dim site As Variant
    Dim commStatus As New SiteLong
    Dim i As Long
    Dim collSize As Long
    Dim regAddr() As Long, devAddr() As Long
    Dim regValToWrite() As Variant    'Use array, not SiteLong
    Dim regWval() As Long
    Dim rMapName As String
    Dim modAddr As New SiteLong
    Dim actSite As Long

    collSize = iRegColl.Count

    ReDim regAddr(collSize)
    ReDim regValToWrite(collSize, 0 To TheExec.Sites.Selected.Count - 1)
    ReDim regWval(0 To TheExec.Sites.Selected.Count - 1)
    Set commStatus = MSV(0)

    If collSize > 0 Then
        'Perform any writes needed to set up access to this memory map.
        rMapName = iRegColl.Item(1).memMapName    'Get name of RegMap first register is in.
 

        'Translate Yoda object information into address and data to write.
                'Example: needs to be customised for part.
        For i = 1 To collSize
            regAddr(i) = iRegColl.Item(i).Address
            devAddr(i) = (regAddr(i) And &H1F0000) / (2 ^ 16)     'Use 5 MSBs.
            regAddr(i) = regAddr(i) And &HFFFF&    'Use 16 LSBs.
            For Each site In TheExec.Sites
                regValToWrite(i, site) = iRegColl.Item(i).Value(site)
                regWval(site) = CLng(regValToWrite(i, site))
            Next site

            'DEBUG: to record all bits exercised by the program and identify any coverage gaps:
            'record A 1 for each bit position which has been written with A value different to its default value.
            'writeArr(iRegColl(i).Address) = writeArr(iRegColl(i).Address) Or (iRegColl(i).value(0) Xor iRegColl(i).DefaultValue)

            'Sections below may be optimised for multiple register read/write.

''''            Select Case rMapName
''''                'Part specific comms selection.
''''            Case "XXXmodule"
''''                SPIwrite24 regAddr(i), regWval()
''''            Case Else
''''                SPIwrite16 regAddr(i), regWval()
''''            End Select

            Select Case rMapName
                'Part specific comms selection.
            Case "GESubsys"
                For Each site In TheExec.Sites
                    Call GEPhy_MDIO_CL_45_Write(&H0, devAddr(i), regAddr(i), regWval(site))
                Next site
            Case Else
                If ((devAddr(i) = 0) And (regAddr(i) < 32)) Then
                    For Each site In TheExec.Sites
                        Call GEPhy_MDIO_CL_22_Write(&H0, regAddr(i), regWval(site))
                    Next site
                Else
                    For Each site In TheExec.Sites
                        Call GEPhy_MDIO_CL_45_Write(&H0, devAddr(i), regAddr(i), regWval(site))
                    Next site
                End If
            End Select

        Next i

        If TheExec.TesterMode = testModeOffline Then
            'Not checking if write to DUT succeeded.
            For i = 1 To collSize
                For Each site In TheExec.Sites
                    iRegColl.Item(i).readValue = iRegColl.Item(i).Value
                    iRegColl.Item(i).Changed = False
                Next site
            Next i

        ElseIf verifyWrite Then
            'Check register readback matches value written.
            Call interRegRead(iRegColl.Item(i), True)

            For i = 1 To collSize
                For Each site In TheExec.Sites
                    If iRegColl.Item(i).readValue = iRegColl.Item(i).Value Then
                        iRegColl.Item(i).Changed = False
                    Else
                        iRegColl.Item(i).Changed = True
                        commStatus = commStatus + 1
                        TheExec.Datalog.WriteComment "!!! Write to register " & iRegColl.Item(i).Name & " failed!"
                    End If
                Next site
            Next i
        End If    'verifyWrite

    End If
    Set interRegCollWrite = commStatus

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "interRegWrite Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function interRegRead(iReg As Iregister, Optional DUTread As Boolean = True) As SiteLong
'''Calls appropriate comms function for this register. Read from the DUT, or from Yoda object if A DUT read is not required.
'''

    On Error GoTo 0
    Dim site As Variant
    Dim readresult As New SiteLong
    Dim regAddr As Long, devAddr As Long
    Dim regValRead As New SiteLong
    Dim actSite As Long
    Dim modAddr As New SiteLong

    'Translate Yoda object to address to read
    regAddr = iReg.Address    'iReg.relativeAddress
    devAddr = (iReg.Address And &H1F0000) / (2 ^ 16)     'Use 5 MSBs.
    regAddr = regAddr And &HFFFF&    'Use 16 LSBs.

    
    If DUTread = True Then    'Set Yoda object readValues to default.
        For Each site In TheExec.Sites
            iReg.readValue = -1
        Next site


''''        Select Case iReg.regMapName
''''            'Part specific comms selection.
''''        Case "XXXmodule"
''''            Set regValRead = SPIread24(regAddr)
''''        Case "YYYmodule"
''''            Set regValRead = SPIread24(regAddr)
''''        Case "ZZZmodule"
''''            Set regValRead = SPIread24(regAddr)
''''        Case Else
''''            Set regValRead = SPIread16(regAddr)
''''        End Select
        
        Select Case iReg.regMapName
            'Part specific comms selection.
        Case "GESubsys"
                regValRead = GEPhy_MDIO_CL_45_Read(&H0, devAddr, regAddr)
        Case Else
            If ((devAddr = 0) And (regAddr < 32)) Then
                regValRead = GEPhy_MDIO_CL_22_Read(&H0, regAddr)
            Else
                regValRead = GEPhy_MDIO_CL_45_Read(&H0, devAddr, regAddr)
            End If
        End Select

        For Each site In TheExec.Sites
            iReg.readValue = regValRead
        Next site

    End If

    'Return last value read from DUT, or the Yoda object value if no read occurred.
    For Each site In TheExec.Sites
        readresult = iReg.readValue
    Next site

    Set interRegRead = readresult

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "interRegread Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub interMapRead(iRegMap As IRegisterMap)
'''Read all registers in this map from the DUT, or from Yoda object if A DUT read is not required.
'''
    On Error GoTo 0
    Dim site As Variant

    Dim i As Long
    Dim mapSize As Long
    Dim regAddr() As Long, devAddr() As Long
    Dim regValRead() As New SiteLong
    Dim actSite As Long
    Dim modAddr As New SiteLong
    mapSize = iRegMap.registers.Count
    ReDim regAddr(mapSize)
    ReDim regValRead(mapSize)

    'Translate Yoda object to address to read.
    For i = 1 To mapSize
        regAddr(i) = iRegMap.registers(i).Address And &HFFFF&    'Use 16 LSBs.   '.relativeAddress
        devAddr(i) = (iRegMap.registers(i).Address And &H1F0000) / (2 ^ 16)     'Use 5 MSBs.
    Next i

    'Sections below may be optimised for multiple register reads.


    'Set Yoda object readValues to default.
    For i = 1 To mapSize
        For Each site In TheExec.Sites
            iRegMap.registers(i).readValue = -1
        Next site

''''        Select Case iRegMap.name
''''            'Part specific comms selection.
''''        Case "XXXmodule"
''''            Set regValRead(i) = SPIread24(regAddr(i))
''''        Case "YYYmodule"
''''            Set regValRead(i) = SPIread24(regAddr(i))
''''        Case "ZZZmodule"
''''            Set regValRead(i) = SPIread24(regAddr(i))
''''        Case Else
''''            Set regValRead(i) = SPIread16(regAddr(i))
''''        End Select
        
       Select Case iRegMap.Name
            'Part specific comms selection.
        Case "GESubsys"
            For Each site In TheExec.Sites
                regValRead = GEPhy_MDIO_CL_45_Read(&H0, devAddr(i), regAddr(i))
            Next site
        Case Else
            If ((devAddr(i) = 0) And (regAddr(i) < 32)) Then
                For Each site In TheExec.Sites
                    regValRead = GEPhy_MDIO_CL_22_Read(&H0, regAddr(i))
                Next site
            Else
                For Each site In TheExec.Sites
                    regValRead = GEPhy_MDIO_CL_45_Read(&H0, devAddr(i), regAddr(i))
                Next site
            End If
        End Select

        For Each site In TheExec.Sites
            iRegMap.registers(i).readValue = regValRead(i)
        Next site
    Next i


    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "interRegread Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Sub Else Resume Next

End Sub

'''''##########################################################################################################
'''''==========================================================================================================
'''''SPI layer
'''''Translates data frames to SPI read | write patterns.
'''''Functions in this section are entirely part-specific and provided only as examples.
'''''Optimisation of SPI writes should take place before program release, for example using flat patterns.
'''''==========================================================================================================
'''''##########################################################################################################
''''
''''Private Sub SPIwrite16(regAdd As Long, ByRef writeWord() As Long)
'''''''Write specified value to specified DUT register.
'''''''
''''    On Error GoTo 0
''''    Const patName As String = "SPIwrite16.pat"
''''    Const numPatBits As Long = 24
''''    Const wrVecOffset As Long = 5    'Define offset of first vector in the pattern to be changed.
''''    Dim site As Variant
''''    Dim frmstring() As String
''''    Dim writeFrame() As Long
''''
''''    ReDim writeFrame(0 To TheExec.Sites.Selected.Count - 1)
''''
''''    'Assemble the SPI frame.
''''    For Each site In TheExec.Sites
''''        writeFrame(site) = shl(regAdd, 16)    '8 bit register address
''''        writeFrame(site) = writeFrame(site) + writeWord(site)    '16 bit data
''''    Next site
''''
''''    'Convert the writeFrame to string array, per site
''''    For Each site In TheExec.Sites
''''        If writeFrame(site) < 0 Or writeFrame(site) > &HFFFFFF Then
''''            Err.Raise Number:=vbObjectError + 1400, Source:="SPIWrite16", Description:="Invalid write frame."
''''        End If
''''        frmstring = num2StrArr(writeFrame(site), numPatBits)
''''
''''        'Insert the data into the pattern
''''
''''        Call TheHdw.Digital.Patterns.Pat(patName).ModifyPinVectorBlockDataSite("", wrVecOffset, writePinName, frmstring, CLng(site))
''''    Next site
''''
''''    'Run the pattern
''''    TheHdw.Digital.Patterns.Pat(patName).Start
''''    TheHdw.Digital.Patgen.HaltWait
''''
''''    Exit Sub
''''errHandler:
''''    TheExec.Datalog.WriteComment Err.Source & " " & Err.Number & " " & Err.Description
''''    TheExec.Sites(-1).BinNumber = 11
''''    If AbortTest Then Exit Sub Else Resume Next
''''
''''End Sub
''''
''''Private Sub SPIwrite24(regAdd As Long, ByRef writeWord() As Long)
'''''''Write specified value to specified DUT register.
'''''''
''''    On Error GoTo 0
''''    Const patName As String = "SPIwrite24.pat"
''''
''''    Const numAddrBits As Long = 8
''''    Const numDataBits As Long = 24
''''
''''    Const wrVecOffsetAdd As Long = 5    'Define offsets of vectors in the pattern to be changed.
''''    Const wrVecOffsetDat As Long = 13    'Define offsets of vectors in the pattern to be changed.
''''    Dim site As Variant
''''    Dim frmstring() As String
''''    Dim writeFrame() As Long
''''
''''    ReDim writeFrame(0 To TheExec.Sites.Selected.Count - 1)
''''
''''    'Assemble the SPI frame: address.
''''    If regAdd < 0 Or regAdd > &HFFFF& Then
''''        Err.Raise Number:=vbObjectError + 1400, Source:="SPIWrite24", Description:="Invalid write frame."
''''    End If
''''    For Each site In TheExec.Sites
''''        writeFrame(site) = regAdd    '8 bit register address
''''    Next site
''''
''''    'Convert the address to string array, per site
''''    For Each site In TheExec.Sites
''''        frmstring = num2StrArr(writeFrame(site), numAddrBits)
''''        'Insert the data into the pattern
''''        Call TheHdw.Digital.Patterns.Pat(patName).ModifyPinVectorBlockDataSite("", wrVecOffsetAdd, writePinName, frmstring, CLng(site))
''''    Next site
''''
''''    'Assemble the SPI frame: data.
''''    For Each site In TheExec.Sites
''''        writeFrame(site) = writeWord(site)    '24 bit data
''''    Next site
''''
''''    'Convert the Data to string array, per site
''''    For Each site In TheExec.Sites
''''        If writeFrame(site) < 0 Then
''''            Err.Raise Number:=vbObjectError + 1400, Source:="SPIWrite24", Description:="Invalid write frame."
''''        End If
''''        frmstring = num2StrArr(writeFrame(site), numDataBits)
''''
''''        'Insert the data into the pattern
''''        Call TheHdw.Digital.Patterns.Pat(patName).ModifyPinVectorBlockDataSite("", wrVecOffsetDat, writePinName, frmstring, CLng(site))
''''    Next site
''''
''''    'Run the pattern
''''    TheHdw.Digital.Patterns.Pat(patName).Start
''''    TheHdw.Digital.Patgen.HaltWait
''''
''''    Exit Sub
''''errHandler:
''''    TheExec.Datalog.WriteComment Err.Source & " " & Err.Number & " " & Err.Description
''''    TheExec.Sites(-1).BinNumber = 11
''''    If AbortTest Then Exit Sub Else Resume Next
''''
''''End Sub
''''
''''
''''Private Function SPIread16(regAdd As Long) As SiteLong
'''''''Read from specified DUT register.
'''''''
''''    On Error GoTo 0
''''    Const numPatBitsReq As Long = 24    'For data request frame.
''''    Const patName As String = "SPIread16.pat"
''''    Const wrVecOffset As Long = 5    'Offset for first write vector.
''''    Const numCapBytes As Long = 3    'Match number of stores in pattern.
''''    Const RDADDR As Long = &HF2&    'Address of RDADDR in Yoda SPI module
''''
''''    Dim site As Variant
''''    Dim dataReqFrame As New SiteLong
''''    Dim readFrame As New SiteLong
''''    Dim regData As New SiteLong
''''    Dim fus As New SiteLong
''''    Dim frmOK As New SiteBoolean
''''    Dim echoAddr As New SiteLong
''''    Dim frmstring() As String
''''    Dim DSSCCaptureData As New DSPWave
''''    Dim failflag As Boolean
''''    Dim MSBextra As New SiteLong
''''    Dim status As New SiteLong
''''
''''    Set status = MSV(-1)
''''    Set regData = MSV(-1)
''''
''''    If regAdd < 0 Or regAdd > &HFF& Then
''''        Err.Raise Number:=vbObjectError + 1000, Source:="SPIread16", Description:="Invalid read address."
''''    End If
''''
''''    'Assemble the SPI frame.
''''    For Each site In TheExec.Sites
''''        dataReqFrame = shl(RDADDR, 16)    'Add RDADDR address
''''        dataReqFrame = dataReqFrame + regAdd    'Add address of reg to read.
''''    Next site
''''
''''    For Each site In TheExec.Sites
''''        'Convert the dataReqFrame to string array.
''''        If dataReqFrame < 0 Or dataReqFrame > &HFFFFFF Then
''''            Err.Raise Number:=vbObjectError + 1005, Source:="SPIread16", Description:="Invalid read request frame."
''''        End If
''''        frmstring = num2StrArr(dataReqFrame(site), numPatBitsReq)
''''
''''        'Insert the data into the pattern
''''        Call TheHdw.Digital.Patterns.Pat(patName).ModifyPinVectorBlockDataSite("", wrVecOffset, writePinName, frmstring, CLng(site))
''''    Next site
''''
''''    'Set up DSSC capture
''''    With TheHdw.DSSC.Pins(readPinName).Pattern(patName).Capture
''''        .Signals.Add (sigNameCap)
''''        .Signals(sigNameCap).SampleSize = numCapBytes
''''        .Signals(sigNameCap).LoadSettings
''''    End With
''''
''''    'Run the pattern
''''    TheHdw.Digital.Patterns.Pat(patName).Start
''''    TheHdw.Digital.Patgen.HaltWait
''''    failflag = TheHdw.Digital.Patgen.failflag    'For all sites
''''
''''    If failflag Then
''''        TheExec.Datalog.WriteComment "SPIread16 Invalid logic levels read from DUT."
''''    End If
''''
''''    'Move captured data from DSSC to dspwave
''''    If TheExec.TesterMode = testModeOnline Then
''''        DSSCCaptureData = TheHdw.DSSC.Pins(readPinName).Pattern(patName).Capture.Signals(sigNameCap).DSPWave
''''        For Each site In TheExec.Sites
''''            regData = 0
''''            status = DSSCCaptureData.Element(0)
''''
''''            regData = regData + shl(DSSCCaptureData.Element(1), 8)
''''            regData = regData + DSSCCaptureData.Element(2)
''''
''''
''''            If (status And &HF8&) <> &H80& Then    'mask reserved 3 LSBs.
''''                'Flag bad response frame here.
''''                'regData = &H10000
''''            End If
''''
''''        Next site
''''    Else
''''        Set status = MSV(&H80&)
''''        Set regData = MSV(&HDDDD&)
''''    End If
''''
''''    Set SPIread16 = regData
''''
''''    Exit Function
''''errHandler:
''''    TheExec.Datalog.WriteComment Err.Source & " " & Err.Number & " " & Err.Description
''''    TheExec.Sites(-1).BinNumber = 11
''''    If AbortTest Then Exit Function Else Resume Next
''''
''''End Function
''''
''''
''''Private Function SPIread24(regAdd As Long) As SiteLong
'''''''Read from specified DUT register.
''''
''''    On Error GoTo 0
''''    Const numPatBitsReq As Long = 24    'For data request frame, NOT data readback frame.
''''    Const patName As String = "SPIread24.pat"
''''    Const wrVecOffset As Long = 5    'Offset for first write vector.
''''    Const numCapBytes As Long = 4    'Match number of stores in pattern.
''''    Const RDADDR24 As Long = &HF3&    'Address of RDADDR in Yoda SPI module
''''
''''    Dim site As Variant
''''    Dim dataReqFrame As New SiteLong
''''    Dim readFrame As New SiteLong
''''    Dim regData As New SiteLong
''''    Dim status As New SiteLong
''''    Dim frmOK As New SiteBoolean
''''    Dim echoAddr As New SiteLong
''''    Dim frmstring() As String
''''    Dim DSSCCaptureData As New DSPWave
''''    Dim failflag As Boolean
''''
''''    Set status = MSV(-1)
''''    Set regData = MSV(-1)
''''
''''    If regAdd < 0 Or regAdd > &HFFFF& Then
''''        Err.Raise Number:=vbObjectError + 1000, Source:="SPIread24", Description:="Invalid read address."
''''    End If
''''
''''    'Assemble the SPI frame.
''''    For Each site In TheExec.Sites
''''        dataReqFrame = shl(RDADDR24, 16)    'Add RDADDR address
''''        dataReqFrame = dataReqFrame + regAdd    'Add address of reg to read.
''''    Next site
''''
''''    'Set up DSSC capture
''''    With TheHdw.DSSC.Pins(readPinName).Pattern(patName).Capture
''''        .Signals.Add (sigNameCap)
''''        .Signals(sigNameCap).SampleSize = numCapBytes
''''        .Signals(sigNameCap).LoadSettings
''''    End With
''''
''''    For Each site In TheExec.Sites
''''        'Convert the 24 bit dataReqFrame to string array.
''''        If dataReqFrame < 0 Or dataReqFrame > &HFFFFFF Then
''''            Err.Raise Number:=vbObjectError + 1005, Source:="SPIread24", Description:="Invalid read request frame."
''''        End If
''''        frmstring = num2StrArr(dataReqFrame(site), numPatBitsReq)
''''
''''        'Insert the data into the pattern
''''        Call TheHdw.Digital.Patterns.Pat(patName).ModifyPinVectorBlockDataSite("", wrVecOffset, writePinName, frmstring, CLng(site))
''''    Next site
''''
''''    'Run the pattern
''''    TheHdw.Digital.Patterns.Pat(patName).Start
''''    TheHdw.Digital.Patgen.HaltWait
''''    failflag = TheHdw.Digital.Patgen.failflag    'For all sites
''''
''''    If failflag Then
''''        TheExec.Datalog.WriteComment "SPIread24 Invalid logic levels read from DUT."
''''    End If
''''
''''    'Move captured data from DSSC to dspwave
''''    If TheExec.TesterMode = testModeOnline Then
''''        DSSCCaptureData = TheHdw.DSSC.Pins(readPinName).Pattern(patName).Capture.Signals(sigNameCap).DSPWave
''''        For Each site In TheExec.Sites
''''            status = DSSCCaptureData.Element(0)
''''
''''            regData = shl(DSSCCaptureData.Element(1), 16)
''''            regData = regData + shl(DSSCCaptureData.Element(2), 8)
''''            regData = regData + DSSCCaptureData.Element(3)
''''
''''            If status <> &H80& Then
''''                'Flag bad response frame here.
''''                'regData = &H1000000
''''            End If
''''
''''        Next site
''''    Else
''''        Set status = MSV(&H80)
''''        Set regData = MSV(&HDDDDDD)
''''    End If
''''
''''    Set SPIread24 = regData
''''
''''    Exit Function
''''errHandler:
''''    TheExec.Datalog.WriteComment Err.Source & " " & Err.Number & " " & Err.Description
''''    TheExec.Sites(-1).BinNumber = 11
''''    If AbortTest Then Exit Function Else Resume Next
''''
''''End Function

'''########################################################################################
'''========================================================================================
'''Comms Supporting functions
'''========================================================================================
'''########################################################################################

Public Function num2StrArr(writeWord As Long, numPatBits As Long) As String()
'''Convert an integer to A binary string array, for insertion in A pattern.
'''
    On Error GoTo 0
    Dim DecimalIn As Long
    Dim i As Long
    Dim frmstring() As String
    ReDim frmstring(0 To numPatBits - 1)

    If writeWord < 0 Or writeWord > (2 ^ numPatBits) - 1 Then
        TheExec.Datalog.WriteComment "!!! Invalid Data"
        GoTo errHandler
    End If

    DecimalIn = writeWord
    'Conversion to string array of bits.
    For i = 0 To (numPatBits - 1)
        frmstring((numPatBits - 1) - i) = Right$(Str(DecimalIn And &H1), 1)
        DecimalIn = shr(DecimalIn, 1)
    Next i
    num2StrArr = frmstring

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    TheExec.Sites(-1).BinNumber = 11
    If AbortTest Then Exit Function Else Resume Next

End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Added by Vikram from ADuCM410 program to make Ycomms Work
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Public Function MSV(wrval As Long) As SiteLong
'''Convert A Long to A SiteLong.
'''
'DanT: Replaced site loop with simpler assignment of Long to SiteLong
'The site loop is redundant for (and slower than) direct [multisite] assignment

    On Error GoTo 0
    Dim valTmp As New SiteLong
    Dim site As Variant

    'For Each site In TheExec.Sites
    'valTmp = wrval
    'Next site

    valTmp = wrval

    Set MSV = valTmp

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function shr(ByVal Value As Long, ByVal shift As Byte) As Long
'''Return A bitwise right shift on the input value.
'''
    shr = Value
    If shift > 0 Then
        shr = Int(shr / (2 ^ shift))
    End If

End Function

Public Function shl(ByVal Value As Long, ByVal shift As Byte) As Long
'''Return A bitwise left shift on the input value.
'''
    shl = Value
    If shift > 0 Then
        Dim i As Byte
        Dim M As Long
        For i = 1 To shift
            M = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If M <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If

End Function



Public Function dirExists(xdir As String) As Boolean
'''Check that the specified directory exists.
'''
    On Error GoTo 0

    Dim attrib As Long
    dirExists = False

    If Dir(xdir, vbDirectory) <> "" Then
        attrib = GetAttr(xdir)
        'Check xdir is A directory and not read-only.
        If ((attrib And vbDirectory) = vbDirectory) And ((attrib And vbReadOnly) = 0) Then
            dirExists = True
        End If
    End If

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function txtFileOpen(filename As String) As Long
'''Create A text file to write data, or close it.
'''
    On Error GoTo 0

    Dim txtNum As Long
    Dim filePath As String

    'Define file location
    filePath = logdir & filename

    'Get A new file number
    txtNum = FreeFile
    'Open text file for writing
    Open filePath For Output As txtNum
    txtFileOpen = txtNum

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub txtFileClose(txtNum As Long)
'''Save & Close Text File
'''
    On Error GoTo 0
    Close txtNum

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Sub txtFileWrite(txtNum As Long, newText As String)
'''Write specified text to file.
'''
    On Error GoTo 0

    Print #txtNum, newText

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description
    If AbortTest Then Exit Sub Else Resume Next

End Sub
