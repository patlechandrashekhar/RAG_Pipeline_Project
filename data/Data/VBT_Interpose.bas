Attribute VB_Name = "VBT_Interpose"
Public Function Retry(argc As Long, argv() As String) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Test Description:
'   Interpose Function to retry pattern if it fails in a functional template
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim site As Variant
Dim no_of_Retry As Long
Dim passing_site As New SiteBoolean
Dim active_site As New SiteBoolean
Dim i As Long
  

    If argc <> 0 Then
        no_of_Retry = CInt(argv(0))
    Else
        TheExec.Datalog.WriteComment "Atleast one argument is needed to use this interpose function. Need number of retries argument!"
        Exit Function
    End If
    
    'Store active and site pass/fail information
    For Each site In TheExec.Sites.Active
        passing_site = thehdw.Digital.Patgen.PatternBurstPassed(site)
        active_site = True
    Next site
    
    'If Any of the active site is failing, run the retry routine
    If passing_site.LogicalXor(active_site).Any(True) Then
    
        'Clear Interpose Function so that the function doesnt get called recursively
        Call TheExec.Flow.ClearInterpose(TL_C_POSTPATF)
        
        'Retry on failing sites
        For Each site In TheExec.Sites.Active
            If Not (passing_site) Then
                i = 0
                Do While (i < no_of_Retry) And Not (passing_site)
                    thehdw.Digital.Patgen.Restart
                    thehdw.Digital.Patgen.HaltWait
                    thehdw.Wait 0.001   ' 1 ms wait between pattern runs
                    
                    If thehdw.Digital.Patgen.PatternBurstPassed(site) Then
                        passing_site = True
                    End If
                    i = i + 1
                Loop
            End If
            TheExec.Datalog.WriteComment ("Site: " & site & ", Number of Retries: " & i)
        Next site
        
        'set the interpose function again
        Call tl_setinterpose(TL_C_POSTPATF, "Retry", argv(0))
        
    End If
    

End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_4_3_1(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_4 (Pattern will drive 1)
'   Vt = Mode_3 (Pattern will drive X)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double, avdd3p3 As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r * 0.85
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVih) = vddio_r
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVil) = 0
    thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = avdd3p3 * 0.85
    thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3
    thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = 0
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
    
End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_4_3_2(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_4 (Pattern will drive 1)
'   Vt = Mode_3 (Pattern will drive X)
'   Vil = Mode_2 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double, avdd3p3 As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r * 0.85
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVih) = vddio_r
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVil) = vddio_r * 0.15
    thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = avdd3p3 * 0.85
    thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3
    thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = avdd3p3 * 0.15
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
    
End Function

Public Function Set_GP_OUT_LED_0_MLS_MODE_3_2_1(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vt = Mode_2 (Pattern will drive X)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double, avdd3p3 As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r * 0.15
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVih) = vddio_r * 0.85
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVil) = 0
    thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = avdd3p3 * 0.15
    thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3 * 0.85
    thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = 0
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
    
End Function

Public Function Reset_GP_OUT_LED_0_MLS_MODE(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vt = Mode_2 (Pattern will drive X)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double, avdd3p3 As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVt) = 0
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.Value(chVil) = vddio_r * 0.05
    thehdw.Digital.Pins("LED_0").Levels.Value(chVih) = avdd3p3 * 0.95
    thehdw.Digital.Pins("LED_0").Levels.Value(chVil) = avdd3p3 * 0.05
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeHiZ
    
End Function

Public Function Set_GP_CLK_MLS_MODE_0_3(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vil = Mode_0 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage

    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVih) = vddio_r
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVil) = 0

    
End Function

Public Function Set_GP_CLK_MLS_MODE_1_2(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_2 (Pattern will drive 1)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage

    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVih) = vddio_r * 0.85
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVil) = vddio_r * 0.15

    
End Function

Public Function Reset_GP_CLK_MLS_MODE(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("GP_CLK").Levels.Value(chVil) = vddio_r * 0.05

    
End Function


Public Function Set_LED_1_MLS_MODE_0_3(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vil = Mode_0 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim avdd3p3 As Double

    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage

    thehdw.Digital.Pins("LED_1").Levels.Value(chVih) = avdd3p3
    thehdw.Digital.Pins("LED_1").Levels.Value(chVil) = 0

    
End Function

Public Function Set_LED_1_MLS_MODE_1_2(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_2 (Pattern will drive 1)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim avdd3p3 As Double

    avdd3p3 = thehdw.DCVI.Pins("AVDD3P3").Voltage

    thehdw.Digital.Pins("LED_1").Levels.Value(chVih) = avdd3p3 * 0.85
    thehdw.Digital.Pins("LED_1").Levels.Value(chVil) = avdd3p3 * 0.15

    
End Function

Public Function Reset_LED_1_MLS_MODE(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'   Vih = Mode_3 (Pattern will drive 1)
'   Vil = Mode_1 (Pattern will drive 0)
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
     'GP_OUT = MODE_1 , LED_0 = MODE_4
    thehdw.Digital.Pins("LED_1").Levels.Value(chVih) = vddio_r * 0.95
    thehdw.Digital.Pins("LED_1").Levels.Value(chVil) = vddio_r * 0.05

    
End Function



Public Function Set_MLS_Spd_Config_Idx_0(argc As Long, argv() As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Interpose Function to set MLS pins to certain state
'
'   Rev 1.0 (vsomasun, May 28th 2019): Initial
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim vddio_r As Double

    vddio_r = thehdw.DCVI.Pins("VDDIO_R").Voltage
    
    'GP_OUT = MODE_4 , LED_0 = MODE_4
    thehdw.Digital.Pins("GP_OUT").Levels.Value(chVt) = vddio_r
    thehdw.Digital.Pins("LED_0").Levels.Value(chVt) = vddio_r
    thehdw.Digital.Pins("GP_OUT, LED_0").Levels.DriverMode = tlDriverModeLargeVt
    
End Function

