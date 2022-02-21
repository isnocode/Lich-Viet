Attribute VB_Name = "Module1"


Global Const WM_LBUTTONUP = &H202
Global Const WM_SYSCOMMAND = &H112
Global Const SC_MOVE = &HF010
Global Const MOUSE_MOVE = &HF012

#If Win32 Then
    Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
#Else
    Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#End If


