VERSION 5.00
Begin VB.UserControl UniTrayIcon 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   525
   ToolboxBitmap   =   "UniTrayIcon.ctx":0000
   Begin VB.Frame hWndHolder 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   -150
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "UniTrayIcon.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "UniTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_BALLOONLCLK = &H405
Private Const WM_BALLOONRCLK = &H404
Private Const WM_BALLOONXCLK = WM_BALLOONRCLK

Private Const sANSI = "a1,a2,a3,a4,a5,a8,a81,a82,a83,a84,a85,a6,a61,a62,a63,a64,a65,e1,e2,e3,e4,e5,e6,e61,e62,e63,e64,e65,i1,i2,i3,i4,i5,o1,o2,o3,o4,o5,o6,o61,o62,o63,o64,o65,o7,o71,o72,o73,o74,o75,u1,u2,u3,u4,u5,u7,u71,u72,u73,u74,u75,y1,y2,y3,y4,y5,d9,A1,A2,A3,A4,A5,A8,A81,A82,A83,A84,A85,A6,A61,A62,A63,A64,A65,E1,E2,E3,E4,E5,E6,E61,E62,E63,E64,E65,I1,I2,I3,I4,I5,O1,O2,O3,O4,O5,O6,O61,O62,O63,O64,O65,O7,O71,O72,O73,O74,O75,U1,U2,U3,U4,U5,U7,U71,U72,U73,U74,U75,Y1,Y2,Y3,Y4,Y5,D9"
Private Const sUNICODE = "00E1,00E0,1EA3,00E3,1EA1,0103,1EAF,1EB1,1EB3,1EB5,1EB7,00E2,1EA5,1EA7,1EA9,1EAB,1EAD,00E9,00E8,1EBB,1EBD,1EB9,00EA,1EBF,1EC1,1EC3,1EC5,1EC7,00ED,00EC,1EC9,0129,1ECB,00F3,00F2,1ECF,00F5,1ECD,00F4,1ED1,1ED3,1ED5,1ED7,1ED9,01A1,1EDB,1EDD,1EDF,1EE1,1EE3,00FA,00F9,1EE7,0169,1EE5,01B0,1EE9,1EEB,1EED,1EEF,1EF1,00FD,1EF3,1EF7,1EF9,1EF5,0111,00C1,00C0,1EA2,00C3,1EA0,0102,1EAE,1EB0,1EB2,1EB4,1EB6,00C2,1EA4,1EA6,1EA8,1EAA,1EAC,00C9,00C8,1EBA,1EBC,1EB8,00CA,1EBE,1EC0,1EC2,1EC4,1EC6,00CD,00CC,1EC8,0128,1ECA,00D3,00D2,1ECE,00D5,1ECC,00D4,1ED0,1ED2,1ED4,1ED6,1ED8,01A0,1EDA,1EDC,1EDE,1EE0,1EE2,00DA,00D9,1EE6,0168,1EE4,01AF,1EE8,1EEA,1EEC,1EEE,1EF0,00DD,1EF2,1EF6,1EF8,1EF4,0110"

Private ArrFromCode() As String
Private ArrToCode() As String

Public bEnable As Boolean   'luu thuoc tinh enable cua UniXPFrame
Public bListViewUnicode As Boolean 'cho biet co su dung font chu unicode trong dieu khien khong

Private Type NOTIFYICONDATA
    cbSize As Long                  ' 4
    hwnd As Long                    ' 8
    uID As Long                     ' 12
    uFlags As Long                  ' 16
    uCallbackMessage As Long        ' 20
    hIcon As Long                   ' 24
    szTip(0 To 255) As Byte         ' 280
    dwState As Long                 ' 284
    dwStateMask As Long             ' 288
    szInfo(0 To 511) As Byte        ' 800
    uTimeOutOrVersion As Long       ' 804
    szInfoTitle(0 To 127) As Byte   ' 932
    dwInfoFlags As Long             ' 936
End Type

Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4

Private m_TrayIcon As StdPicture
Private bUnicode As Boolean
Private m_sTooltipText As String

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private m_IconData As NOTIFYICONDATA

Public Enum BalloonTipStyle
    btsNoIcon = NIIF_NONE
    btsWarning = NIIF_WARNING
    btsError = NIIF_ERROR
    btsInfo = NIIF_INFO
End Enum

Public Enum stMouseEvent
    stLeftButtonDown = WM_LBUTTONDOWN
    stLeftButtonUp = WM_LBUTTONUP
    stLeftButtonDoubleClick = WM_LBUTTONDBLCLK
    stLeftButtonClick = WM_LBUTTONUP
    stRightButtonDown = WM_RBUTTONDOWN
    stRightButtonUp = WM_RBUTTONUP
    stRightButtonDoubleClick = WM_RBUTTONDBLCLK
    stRightButtonClick = WM_RBUTTONUP
End Enum

Public Enum stBalloonClickType
    stbLeftClick = WM_BALLOONLCLK
    stbRightClick = WM_BALLOONRCLK
    stbXClick = WM_BALLOONXCLK
End Enum

Public Event TrayClick(Button As stMouseEvent)
Public Event BalloonClick(ClickType As stBalloonClickType)

Private Sub StringToArray(ByVal sString As String, bArray() As Byte, ByVal lMaxSize As Long)
    Dim b() As Byte
    Dim i As Long
    Dim j As Long
    If Len(sString) > 0 Then
        b = sString
        For i = LBound(b) To UBound(b)
            bArray(i) = b(i)
            If (i = (lMaxSize - 2)) Then
                Exit For
            End If
        Next i
        For j = i To lMaxSize - 1
            bArray(j) = 0
        Next j
    End If
End Sub

Private Sub hWndHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Msg As Long
    
        Msg = x / Screen.TwipsPerPixelX
        If Msg >= WM_LBUTTONDOWN And Msg <= WM_RBUTTONDBLCLK Then
            RaiseEvent TrayClick(Msg)
        ElseIf Msg >= WM_BALLOONXCLK And Msg <= WM_BALLOONLCLK Then
            RaiseEvent BalloonClick(Msg)
        End If
End Sub

Private Sub picIcon_Resize()
    UserControl.Size picIcon.Width, picIcon.Height
End Sub

Private Sub UserControl_InitProperties()
    bUnicode = True
    m_sTooltipText = Ambient.DisplayName
    Set Icon = UserControl.Parent.Icon
    Set picIcon.Picture = m_TrayIcon
    UserControl.Size picIcon.Width, picIcon.Height
End Sub

Private Sub UserControl_Resize()
    UserControl.Size picIcon.Width, picIcon.Height
End Sub

Public Sub Create(ToolTipText As String, Optional Icon As StdPicture)
On Error Resume Next
    Dim m_sText As String
    
        m_sText = IIf(bUnicode, zToUnicode(ToolTipText), ToolTipText)
    
        If Not Icon Is Nothing Then Set m_TrayIcon = Icon
        With m_IconData
            .cbSize = Len(m_IconData)
            .hwnd = hWndHolder.hwnd
            .uID = m_TrayIcon ' vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
            .uCallbackMessage = WM_MOUSEMOVE
            If Not m_TrayIcon Is Nothing Then .hIcon = m_TrayIcon
    '        If Not IsMissing(ToolTipText) Then .szTip = ToolTipText & vbNullChar
            If Not IsMissing(ToolTipText) Then StringToArray m_sText, .szTip, 256
            
            .dwState = 0
            .dwStateMask = 0
'    '        .szInfo = "" & Chr(0)
'    '        .szInfoTitle = "" & Chr(0)
            .dwInfoFlags = NIIF_NONE
        End With
        Shell_NotifyIcon NIM_ADD, m_IconData
End Sub

Public Sub Remove()
    Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Public Sub BalloonTip(Prompt As String, Optional Style As BalloonTipStyle = btsNoIcon, Optional Title As String, Optional Timeout As Long = 2000)
    If Title = Empty Then Title = App.Title
    If Prompt = Empty Then Prompt = " "
    With m_IconData
'        .szInfo = Prompt & Chr(0)
'        .szInfoTitle = Title & Chr(0)
        StringToArray IIf(bUnicode, zToUnicode(Prompt), Prompt), .szInfo, 512
        StringToArray IIf(bUnicode, zToUnicode(Title), Title), .szInfoTitle, 128

        .dwInfoFlags = Style
        .uTimeOutOrVersion = Timeout
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Public Sub PopupMenu(Menu As Object, Optional flags, Optional DefaultMenu)
    SetForegroundWindow Menu.Parent.hwnd
    If IsMissing(flags) And IsMissing(DefaultMenu) Then
        Menu.Parent.PopupMenu Menu
    ElseIf IsMissing(flags) Then
        Menu.Parent.PopupMenu Menu, , , , DefaultMenu
    Else
        Menu.Parent.PopupMenu Menu, flags, , , DefaultMenu
    End If
End Sub

Property Set Icon(new_Icon As StdPicture)
    Set m_TrayIcon = new_Icon
    Set picIcon.Picture = new_Icon
    With m_IconData
        .hIcon = m_TrayIcon
'        .szInfo = "" & Chr(0)
        StringToArray "", .szInfo, 512
'        .szInfoTitle = "" & Chr(0)
        StringToArray "", .szInfoTitle, 128
        .dwInfoFlags = NIIF_NONE
        .uTimeOutOrVersion = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
    PropertyChanged "Icon"
End Property

Property Get Icon() As StdPicture
    Set Icon = m_TrayIcon
End Property

Public Property Get AutoUnicode() As Boolean
    AutoUnicode = bUnicode
End Property

Public Property Let AutoUnicode(ByVal Auto_Uni As Boolean)
    bUnicode = Auto_Uni
    PropertyChanged "AutoUnicode"
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_sTooltipText
End Property

Public Property Let ToolTipText(ByVal new_text As String)
    m_sTooltipText = new_text
    
    With m_IconData
        .hIcon = m_TrayIcon
        StringToArray IIf(bUnicode, zToUnicode(m_sTooltipText), m_sTooltipText), .szTip, 256 '128
'        .szInfo = "" & Chr(0)
        StringToArray "", .szInfo, 512
'        .szInfoTitle = "" & Chr(0)
        StringToArray "", .szInfoTitle, 128
        .dwInfoFlags = NIIF_NONE
        .uTimeOutOrVersion = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
    
    PropertyChanged "TooltipText"
End Property

Private Sub UserControl_Terminate()
    Remove
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Size picIcon.Width, picIcon.Height
    bUnicode = PropBag.ReadProperty("AutoUnicode", True)
    m_sTooltipText = PropBag.ReadProperty("TooltipText", Ambient.DisplayName)
    Set m_TrayIcon = PropBag.ReadProperty("Icon", UserControl.Parent.Icon)
    Set picIcon.Picture = m_TrayIcon
    
    If Ambient.UserMode Then Call Create(m_sTooltipText, m_TrayIcon)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TooltipText", m_sTooltipText, Ambient.DisplayName
    PropBag.WriteProperty "AutoUnicode", bUnicode, True
    PropBag.WriteProperty "Icon", m_TrayIcon, UserControl.Parent.Icon
End Sub



Public Function zToUnicode(ByRef sString As String) As String
    Dim sTam As String
    Dim i As Long, ArrChuoiXuLy() As String
    Dim k As Long, j As Long
    Dim sKyTu1 As String, sKyTu2 As String, sKyTu3 As String, sKyTu4 As String
    Dim sChuoiBenPhai As String, sKhoangTrang As String
    Dim iVitri As Integer
        
        If Trim$(sString) = "" Then zToUnicode = sString:    Exit Function
        
        ArrFromCode = Split(sANSI, ",")
        ArrToCode = Split(sUNICODE, ",")
        ArrChuoiXuLy = Split(sString, " ")
        For i = 0 To UBound(ArrChuoiXuLy)
            j = HaveNumber(ArrChuoiXuLy(i))
            If (j > 1) And (Not IsNumeric(ArrChuoiXuLy(i))) Then
                If j > 2 Then sTam = sTam & Left$(ArrChuoiXuLy(i), j - 2)
                sKyTu1 = Mid$(ArrChuoiXuLy(i), j - 1, 1)
                sKyTu2 = Mid$(ArrChuoiXuLy(i), j, 1)
                For k = j To Len(ArrChuoiXuLy(i))
                    If IsNumeric(Mid$(ArrChuoiXuLy(i), k + 1, 1)) And sChuoiBenPhai = "" Then
                        sKyTu3 = sKyTu3 & Mid$(ArrChuoiXuLy(i), k + 1, 1)
                    Else
                        sChuoiBenPhai = sChuoiBenPhai & Mid$(ArrChuoiXuLy(i), k + 1, 1)
                    End If
                Next
                If Trim$(sChuoiBenPhai) <> "" Then If HaveNumber(sChuoiBenPhai) > 0 Then sChuoiBenPhai = Trim$(zToUnicode(sChuoiBenPhai))
                sKyTu4 = sKyTu1 & sKyTu2 & sKyTu3
                sTam = sTam & ChangeString(sKyTu4)
            Else
                sTam = sTam & ArrChuoiXuLy(i) & " "
                GoTo TT
            End If
            sTam = sTam & sChuoiBenPhai & " "
TT:
            sKyTu1 = "":    sKyTu2 = "":    sKyTu3 = "":    sChuoiBenPhai = ""
        Next
        zToUnicode = IIf(Right$(sTam, 1) = " ", Left$(sTam, Len(sTam) - 1), sTam)
End Function

'HAM CHUYEN DOI CHUOI UNICODE SANG CHUOI ANSI
Private Function UNICODE_To_ANSI(ByRef sString As String) As String
    Dim sTam As String
    Dim i As Long, ArrChuoiXuLy() As String
    Dim k As Long, j As Long
    Dim bThay As Boolean
        
        If Trim$(sString) = "" Then UNICODE_To_ANSI = sString:   Exit Function
        
        ArrFromCode = Split(sUNICODE, ",")
        ArrToCode = Split(sANSI, ",")
        ArrChuoiXuLy = Split(sString, " ")
        
        For i = 0 To UBound(ArrChuoiXuLy)   'cho vong lap chay den het cac tu trong 1 chuoi (cac tu cach nhau 1 khoang trang)
            For j = 1 To Len(ArrChuoiXuLy(i))  'cho vong lap chay tu ky tu trong 1 tu
                'chi kiem tra cac ky tu nam sau z ma thoi neu la cac ky tu tu A -> Z, a -> z thi khong kiem tra
                If AscW(Mid$(ArrChuoiXuLy(i), j, 1)) = Asc("?") Or AscW(Mid$(ArrChuoiXuLy(i), j, 1)) > Asc("z") Then
                    For k = 0 To UBound(ArrFromCode)
                        If Mid$(ArrChuoiXuLy(i), j, 1) = ChrW$(CLng("&H" & ArrFromCode(k))) Then      'neu tim thay ky tu can thay the trong chuoi can chuyen doi
                            bThay = True 'tha^'y ky tu can chuyen doi
                            Exit For
                        End If
                    Next
                End If
                
                If bThay = True Then
                    sTam = sTam & ArrToCode(k)  'thay ky tu trong chuoi can chuyen doi thanh ky tu trong bang ma sau khi chuyen
                    bThay = False
                Else
                    sTam = sTam & Mid$(ArrChuoiXuLy(i), j, 1)
                End If
            Next
            sTam = sTam & " "   'sau khi kiem tra xong 1 chu thi them vao sau no 1 khoang trang
        Next
        
'cat bo 1 ky tu khoang trang du ra phia sau chuoi sau khi xu ly xong
        UNICODE_To_ANSI = IIf(Right$(sTam, 1) = " ", Left$(sTam, Len(sTam) - 1), sTam) ' Replace$(sTam, "  ", " ")
End Function

Private Function HaveNumber(sString As String) As Long
    Dim i As Long
    Dim sKytu As String
    
        For i = 1 To Len(sString)
            sKytu = Mid(sString, i, 1)
            If IsNumeric(sKytu) Then HaveNumber = i
            If HaveNumber > 0 Then Exit Function
        Next
End Function

Private Function ChangeString(sString As String) As String
    Dim k As Long, bThayDoi As Boolean

        For k = 0 To UBound(ArrToCode)
            If sString = ArrFromCode(k) Then
                ChangeString = ChangeString & ChrW$(CLng("&H" & ArrToCode(k)))
                bThayDoi = True
                Exit For
            End If
        Next
        If bThayDoi = False Then ChangeString = ChangeString & sString
End Function
