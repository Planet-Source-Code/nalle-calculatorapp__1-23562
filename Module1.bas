Attribute VB_Name = "Module1"
Option Explicit

'**************************************
'Windows API/Global Declarations for :Set Form on top easily put this into a module

Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_SHOWWINDOW = &H40
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2

Function FormOnTop(hwnd As Integer, OnTop As Boolean)
    Dim wFlags As Long, PosFlag As Long
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

    Select Case OnTop
        Case True
        PosFlag = HWND_TOPMOST
        Case False
        PosFlag = HWND_NOTOPMOST
    End Select
SetWindowPos hwnd, PosFlag, 0, 0, 0, 0, wFlags
End Function


