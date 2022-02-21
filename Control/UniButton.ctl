VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   DefaultCancel   =   -1  'True
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ToolboxBitmap   =   "UniButton.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1275
      Top             =   1650
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tac gia : Duong Quoc Hung
'Email : Myspecialbox@yahoo.com.vn
'Button With Unicode and StyleXP
'Create on 12/05/2007

'API

Dim bFocus As Boolean
Dim Chuot As Boolean
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum IconAlign
    [iLeft] = 0
    [iTop] = 1
    [iRight] = 2
    [iBottom] = 3
End Enum


Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

'Default Property Values:
Const m_def_IconAlignment = 0
Const m_def_PictureSize = 16
Const m_def_Caption = "UniButton"
'Property Variables:
Dim m_IconAlignment As IconAlign
Dim m_PictureSize As Integer
Dim m_Picture As Picture
Dim m_Caption As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           rtRect As RECT) As Boolean
  Dim hTheme  As Long '* hTheme Handle.
  Dim lResult As Long '* Temp Variable.
    On Error GoTo NoXP
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
    If (hTheme) Then
        lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, rtRect, rtRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False
    End If
    Call CloseThemeData(hTheme)
    Exit Function
NoXP:
    DrawTheme = False
End Function

Private Sub DrawButton(Optional State As Long = 1)
    Dim MRC As RECT
    
    With MRC
        .Left = 0
        .Top = 0
        .Bottom = ScaleHeight: .Right = ScaleWidth
    End With
    Cls

    If DrawTheme("Button", 1, IIf(Enabled, State, 4), MRC) = False Then
       If State = 3 Then
          DrawFrameControl hdc, MRC, 4, 16 Or 512
        ElseIf (sState = hPressed) Then
           DrawFrameControl hdc, MRC, 4, 16
        End If
    End If
    
    If Not (m_Picture Is Nothing) Then
        If m_IconAlignment = 0 Then
                If m_Caption <> "" Then
                    PaintPicture m_Picture, (ScaleWidth / 2) - (TextWidth(cUni(m_Caption)) / 2) - (m_PictureSize) + 2, (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                Else
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                End If
    
            With MRC
                .Left = 0
                .Top = 0
                .Bottom = ScaleHeight: .Right = ScaleWidth + (m_PictureSize)
            End With
        ElseIf m_IconAlignment = 1 Then
                If m_Caption <> "" Then
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) - (TextHeight(cUni(m_Caption)) / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                Else
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                End If
    
            With MRC
                .Left = 0
                .Top = m_PictureSize
                .Bottom = ScaleHeight: .Right = ScaleWidth
            End With
        ElseIf m_IconAlignment = 2 Then
                If m_Caption <> "" Then
                    PaintPicture m_Picture, (ScaleWidth / 2) + (TextWidth(cUni(m_Caption)) / 2) - 4, (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                Else
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                End If
    
            With MRC
                .Left = -m_PictureSize
                .Top = 0
                .Bottom = ScaleHeight: .Right = ScaleWidth
            End With
        ElseIf m_IconAlignment = 3 Then
                If m_Caption <> "" Then
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) + (TextHeight(cUni(m_Caption)) / 2) - 6, m_PictureSize, m_PictureSize
                Else
                    PaintPicture m_Picture, (ScaleWidth / 2) - (m_PictureSize / 2), (ScaleHeight / 2) - (m_PictureSize / 2), m_PictureSize, m_PictureSize
                End If
    
            With MRC
                .Left = 0
                .Top = -m_PictureSize
                .Bottom = ScaleHeight: .Right = ScaleWidth
            End With
        End If
    End If
    
    If Enabled = False Then
        SetTextColor hdc, &H808080
    End If
    DrawTextW hdc, StrPtr(cUni(m_Caption)), Len(cUni(m_Caption)), MRC, &H4 & &H15

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    DrawButton
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    DrawButton
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    DrawButton
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
   If MDMouseOver(hWnd) = False Then
       Timer1.Interval = 0
       If bFocus = True Then
          DrawButton 5
       Else
          DrawButton 1
       End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  If Chuot = True Then
     SetCapture UserControl.hWnd    ' Send MouseUp event to the control
     UserControl_MouseDown 1, 0, 0, 0
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
       Chuot = True
       DrawButton 3
    Else
       Chuot = False
    End If
End Sub

Private Function MDMouseOver(ByVal HandleWindow As Long) As Boolean

Dim pt As POINTAPI

    GetCursorPos pt
    MDMouseOver = (WindowFromPoint(pt.X, pt.Y) = HandleWindow)

End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button = 1 Then
      If MDMouseOver(hWnd) = True Then
         DrawButton 3
      Else
         DrawButton 2
      End If
    ElseIf Button <> 2 Then
       Timer1.Interval = 100
       DrawButton 2
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If bFocus = True Then
       DrawButton 5
    Else
       DrawButton 1
    End If
End Sub

Private Sub UserControl_EnterFocus()
   bFocus = True
   If Chuot = False Then
     DrawButton 5
   End If
End Sub

Private Sub UserControl_ExitFocus()
   bFocus = False
   DrawButton 1
   Chuot = False
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
'    m_Caption = m_def_Caption
    m_Caption = m_def_Caption
    Set m_Picture = Nothing
    m_PictureSize = m_def_PictureSize
    m_IconAlignment = m_def_IconAlignment
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureSize = PropBag.ReadProperty("PictureSize", m_def_PictureSize)
    m_IconAlignment = PropBag.ReadProperty("IconAlignment", m_def_IconAlignment)
End Sub

Private Sub UserControl_Resize()
    DrawButton
End Sub

Private Sub UserControl_Show()
    DrawButton
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureSize", m_PictureSize, m_def_PictureSize)
    Call PropBag.WriteProperty("IconAlignment", m_IconAlignment, m_def_IconAlignment)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    DrawButton
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    DrawButton
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,16
Public Property Get PictureSize() As Integer
    PictureSize = m_PictureSize
End Property

Public Property Let PictureSize(ByVal New_PictureSize As Integer)
    m_PictureSize = New_PictureSize
    PropertyChanged "PictureSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get IconAlignment() As IconAlign
    IconAlignment = m_IconAlignment
End Property

Public Property Let IconAlignment(ByVal New_IconAlignment As IconAlign)
    m_IconAlignment = New_IconAlignment
    DrawButton
    PropertyChanged "IconAlignment"
End Property

