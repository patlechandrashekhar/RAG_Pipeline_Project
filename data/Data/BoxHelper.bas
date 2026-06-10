Attribute VB_Name = "BoxHelper"
Option Explicit
'''This module defines GUI message box functions for user data entry.
'''

Private Type XRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias _
                                    "FindWindowA" (ByVal lpClassName As String, _
                                                   ByVal lpWindowName As String) As Long
Public Declare Function SetTimer& Lib "user32" _
                                  (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal _
                                                                                 lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" _
                                    (ByVal hwnd&, ByVal nIDEvent&)
Const DXB_BOX As Long = &H6666&
Const DXB_BOX1 As Long = &H6667&

'Constants for topmost.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSENDCHANGING = &H400
Private Const SWP_DEFERERASE = &H2000
Private Const SWP_ASYNCWINDOWPOS = &H4000
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_ASYNCWINDOWPOS
Private Const FLAGS1 = SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_ASYNCWINDOWPOS
Private Declare Function SetWindowPos Lib "user32" _
                                      (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                       ByVal x As Long, ByVal y As Long, _
                                       ByVal cx As Long, ByVal cy As Long, _
                                       ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As XRECT) As Long


Private Const MyVersion = "1.0.0"
Private MyIBName As String

' Start the timer looking for the IBName InputBox Window
Public Sub BoxTop(IBName As String)
    MyIBName = IBName
    SetTimer Application.hwnd, DXB_BOX, 100, AddressOf TimerProc
End Sub

' Start the timer looking for the IBName InputBox Window
Public Sub BoxTop_ADIUI(IBName As String)
    MyIBName = IBName
    SetTimer Application.hwnd, DXB_BOX1, 100, AddressOf TimerProc1
End Sub

Public Function GetBoxTopVersion() As String
    GetBoxTopVersion = MyVersion
End Function

Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, _
                     ByVal idEvent&, ByVal dwTime&)

    Call SetWindowPos(Application.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

    Dim h As Long
    h = FindWindow(vbNullString, MyIBName)

    Call SetWindowPos(Application.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

    If (h = 0) Then
        Exit Sub
    End If

    ' Move the Box to the Top most
    Call SetWindowPos(h, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

    KillTimer Application.hwnd, DXB_BOX
End Sub

Public Sub TimerProc1(ByVal hwnd&, ByVal uMsg&, _
                      ByVal idEvent&, ByVal dwTime&)

    Call SetWindowPos(Application.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

    Dim h As Long
    h = FindWindow(vbNullString, MyIBName)

    Call SetWindowPos(Application.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

    If (h = 0) Then
        Exit Sub
    End If

    Dim hADIUI As Long
    hADIUI = FindWindow(vbNullString, "Teradyne Flex Production User Interface")

    If (hADIUI <> 0) Then
        Dim r As XRECT
        Call GetClientRect(hADIUI, r)

        Dim x As Long
        Dim y As Long
        x = (r.Right - r.Left) / 2
        y = (r.Bottom - r.Top) / 2
        ' Move the Box to the Top most
        Call SetWindowPos(h, HWND_TOPMOST, x, y, 0, 0, FLAGS1)
    Else
        ' Move the Box to the Top most
        Call SetWindowPos(h, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If

    KillTimer Application.hwnd, DXB_BOX1
End Sub


