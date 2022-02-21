VERSION 5.00
Begin VB.UserControl CheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "CheckBox.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   1200
   End
End
Attribute VB_Name = "CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tac gia : Duong Quoc Hung
'Email : Myspecialbox@yahoo.com.vn
'Button With Unicode and StyleXP
'Create on 21/08/2008


Private Const DFC_CAPTION = 1 ' Caption
Private Const DFC_MENU = 2 ' Menubar
Private Const DFC_SCROLL = 3 ' ScrollBar
Private Const DFC_BUTTON = 4 ' Button
Private Const DFC_POPUPMENU = 5 ' Popup Menu
Private Const DFCS_BUTTON3STATE = &H8
Private Const DFCS_BUTTONCHECK = &H0
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_BUTTONRADIO = &H4
Private Const DFCS_BUTTONRADIOIMAGE = &H1
Private Const DFCS_BUTTONRADIOMASK = &H2
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_MENUARROW = &H0
Private Const DFCS_MENUARROWRIGHT = &H4
Private Const DFCS_MENUBULLET = &H2
Private Const DFCS_MENUCHECK = &H1
Private Const DFCS_SCROLLCOMBOBOX = &H5
Private Const DFCS_SCROLLDOWN = &H1
Private Const DFCS_SCROLLLEFT = &H2
Private Const DFCS_SCROLLRIGHT = &H3
Private Const DFCS_SCROLLSIZEGRIP = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Private Const DFCS_SCROLLUP = &H0
Private Const DFCS_CHECKED = &H400
Private Const DFCS_FLAT = &H4000
Private Const DFCS_INACTIVE = &H100
Private Const DFCS_MONO = &H8000
Private Const DFCS_PUSHED = &H200
Private Const DFCS_TRANSPARENT = &H800
Private Const DFCS_HOT = &H1000
Private Const DFCS_ADJUSTRECT = &H2000

Public Enum Checkstate
    vbUnchecked = 0
    vbChecked = 1
    vbGrayed = 2
End Enum

Private Type Rect
        left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long


'Default Property Values:
Const m_def_Value = 0
Const m_def_Caption = ""
'Property Variables:
Dim m_Value As Checkstate
Dim m_Caption As String
Dim Chuot As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           rtRect As Rect) As Boolean
  Dim hTheme  As Long '* hTheme Handle.
  Dim lResult As Long '* Temp Variable.
    On Error GoTo NoXP
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    If (hTheme) Then
        lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False
    End If
    Call CloseThemeData(hTheme)
    Exit Function
NoXP:
    DrawTheme = False
End Function

Private Sub DrawOption(Optional ByVal State As Integer = 0)
    Dim MRC As Rect, j As Long
    With MRC
        .left = 2
        .Top = (ScaleHeight / 2) - 7
        .Right = .left + 13
        .Bottom = .Top + 13
    End With
    Cls
    If Me.Value = vbUnchecked Then
        If State = 1 Then
           j = 3
        ElseIf State = 0 Then
           If MDMouseOver(hwnd) = False Then
              j = 1
           Else
              j = 2
           End If
        End If
      ElseIf Me.Value = vbChecked Then
        If State = 1 Then
           j = 7
        ElseIf State = 0 Then
           If MDMouseOver(hwnd) = False Then
              j = 5
           Else
              j = 6
           End If
        End If
      ElseIf Me.Value = vbGrayed Then
        If State = 1 Then
           j = 11
        ElseIf State = 0 Then
           If MDMouseOver(hwnd) = False Then
              j = 9
           Else
              j = 10
           End If
        End If
   End If
   
    If DrawTheme("Button", 3, j, MRC) = False Then
        If Me.Value = vbUnchecked Then
             If State = 1 Then
                 j = DFCS_BUTTONCHECK Or DFCS_PUSHED
             ElseIf State = 0 Then
                 j = DFCS_BUTTONCHECK
             End If
         ElseIf Me.Value = vbChecked Then
             If State = 1 Then
                 j = DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_PUSHED
             ElseIf State = 0 Then
                 j = DFCS_BUTTONCHECK Or DFCS_CHECKED
             End If
         ElseIf Me.Value = vbGrayed Then
             If State = 1 Then
                 j = DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_PUSHED Or DFCS_BUTTON3STATE
             ElseIf State = 0 Then
                 j = DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_BUTTON3STATE
             End If
        End If
        DrawFrameControl hDC, MRC, DFC_BUTTON, j
    End If
    
    With MRC
        .left = 19: .Top = (ScaleHeight / 2) - Me.Font.Size + 1
        .Bottom = (ScaleHeight / 2) + Me.Font.Size - 2
        .Right = ScaleWidth - 1
    End With
    DrawTextW hDC, StrPtr(cUni(m_Caption)), -1, MRC, &H10 Or &H2000
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  If Chuot = True Then
     SetCapture UserControl.hwnd     ' Send MouseUp event to the control
     UserControl_MouseDown 1, 0, 2, 2
  End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then
      DrawOption 1
   End If
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Timer1_Timer()
   If MDMouseOver(hwnd) = False Then
      DrawOption 0
      Timer1.Enabled = False
      Timer1.Interval = 0
   End If
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button <> 1 Then
      Timer1.Enabled = True
      Timer1.Interval = 100
      DrawOption 0
   Else
      Timer1.Enabled = False
      Timer1.Interval = 0
      If MDMouseOver(hwnd) = True Then
         DrawOption 1
      Else
         DrawOption 0
      End If
   End If
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MDMouseOver(hwnd) = True Then
       If Me.Value = vbChecked Then
          Me.Value = vbUnchecked
       ElseIf Me.Value = vbUnchecked Then
          Me.Value = vbChecked
       Else
          Me.Value = vbUnchecked
       End If
       RaiseEvent Click
    End If
    Timer1.Enabled = True
    Timer1.Interval = 1
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Function MDMouseOver(ByVal HandleWindow As Long) As Boolean

Dim pt As POINTAPI

    GetCursorPos pt
    MDMouseOver = (WindowFromPoint(pt.x, pt.y) = HandleWindow)

End Function

Private Sub UserControl_EnterFocus()
   bFocus = True
End Sub

Private Sub UserControl_ExitFocus()
   bFocus = False
   DrawOption 0
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    DrawOption
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal new_Font As Font)
    Set UserControl.Font = new_Font
    DrawOption
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    DrawOption
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    DrawOption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
    m_Value = 0
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
    DrawOption
End Sub

Private Sub UserControl_Show()
    DrawOption
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Checkstate
Attribute Value.VB_MemberFlags = "122c"
    Value = m_Value
End Property

Public Property Let Value(ByVal new_value As Checkstate)
    m_Value = new_value
    DrawOption
    PropertyChanged "Value"
End Property



