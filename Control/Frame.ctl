VERSION 5.00
Begin VB.UserControl Frame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Frame.ctx":0000
End
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tac gia : Duong Quoc Hung
'Email : Myspecialbox@yahoo.com.vn
'Button With Unicode and StyleXP
'Create on 21/08/2008

'API

Private Declare Function ApiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawFrameControl Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawTextW Lib "User32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "User32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Default Property Values:
Const m_def_Caption = ""
'Property Variables:
Dim m_Caption As String

Public Function TranslateColor(ByVal clrColor As OLE_COLOR) As Long
    '--- handle invalid (none) color
    If clrColor = -1 Then
        TranslateColor = -1
    Else
        Call OleTranslateColor(clrColor, m_MemoryPal, TranslateColor)
        TranslateColor = TranslateColor And &HFFFFFF
    End If
End Function

Private Sub DrawASquare(DestDC As Long, rc As RECT, oColor As OLE_COLOR, Optional bFillRect As Boolean)
Dim iBrush As Long
Dim i(0 To 3) As Long
oColor = TranslateColor(oColor)
    i(0) = rc.Top
    i(1) = rc.Left
    i(2) = rc.Right
    i(3) = rc.Bottom
    iBrush = CreateSolidBrush(oColor)
    If bFillRect = True Then
        FillRect DestDC, rc, iBrush
    Else
        FrameRect DestDC, rc, iBrush
    End If
    ApiDeleteObject iBrush
End Sub

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

Private Sub DrawFrame()
    Dim MRC As RECT, txtRC As RECT
    With MRC
        .Left = 0
        .Top = 5
        .Right = ScaleWidth
        .Bottom = ScaleHeight
    End With
    Cls
    If DrawTheme("Button", 4, 1, MRC) = False Then
        MRC.Right = MRC.Right - 1
        MRC.Bottom = MRC.Bottom - 1
        DrawASquare UserControl.hdc, MRC, vbGrayText, False
        MRC.Top = MRC.Top + 1
        MRC.Left = MRC.Left + 1
        MRC.Right = MRC.Right + 1
        MRC.Bottom = MRC.Bottom + 1
        DrawASquare UserControl.hdc, MRC, vbWhite, False
    End If
    With txtRC
        .Left = 12
        .Top = 0
        .Right = TextWidth(cUni(m_Caption)) + 17
        .Bottom = 20
    End With
    FillRect hdc, txtRC, CreateSolidBrush(TranslateColor(vbButtonFace))
    With txtRC
        .Left = 14
    End With
    DrawTextW hdc, StrPtr(cUni(m_Caption)), -1, txtRC, 0
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    DrawFrame
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = Extender.Name
    Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub

Private Sub UserControl_Resize()
    DrawFrame
End Sub

Private Sub UserControl_Show()
    DrawFrame
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    DrawFrame
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    DrawFrame
    PropertyChanged "ForeColor"
End Property

