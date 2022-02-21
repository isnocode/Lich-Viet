VERSION 5.00
Begin VB.UserControl cmdButton 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   ScaleHeight     =   3900
   ScaleWidth      =   7695
   Begin VB.PictureBox picNormal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   3960
      ScaleHeight     =   1200
      ScaleWidth      =   3705
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   3705
   End
   Begin VB.PictureBox picOver 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   3960
      Picture         =   "cmdButton.ctx":0000
      ScaleHeight     =   1200
      ScaleWidth      =   3705
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1275
      Width           =   3705
   End
   Begin VB.PictureBox picClicked 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   3960
      Picture         =   "cmdButton.ctx":E8C2
      ScaleHeight     =   1200
      ScaleWidth      =   3705
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2550
      Width           =   3705
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2010
      Top             =   2385
   End
   Begin LunarCalendar.ucImage picIcono 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      bData           =   0   'False
      Filename        =   ""
      eScale          =   1
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   660
      Left            =   1140
      TabIndex        =   1
      Top             =   390
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   1140
      TabIndex        =   0
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "cmdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



'---------------------------------------------------------------------------------------
' Module        : cmdButton 2 Caption
' DateTime      : 08/03/2011
' Author        : ENTER
' Mail          : enterariel@msn.com
'
' Credits       : Covein        = ucImage
'               : Raul338       = MouseOut, LoadImageFromFile
'               : Lolabyte      = ImageFromFile
'               : Y a todos los foreros de http://www.leandroascierto.com.ar/foro
'---------------------------------------------------------------------------------------

Dim m_cImageFromFile As String

'Event Declarations:
Public Event Click()

'Api GetCursorPos para recuperar las coordenadas del puntero
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
'Api WindowFromPoint para determinar si el puntero se encuentra en el area del Picture
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Type Para el Api GetcursorPos ( Recupera la posición x y posición y)
Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim Posicion_Cursor As POINTAPI
Dim Ret_Hwnd As Long

'Flag para el área
Private Enum picEstado
    Normal
    Clicked
    Sobre
End Enum

Private picAcutal As picEstado

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption1() As String
    Caption1 = Label1.Caption
End Property

Public Property Let Caption1(ByVal New_Caption1 As String)
    Label1.Caption() = New_Caption1
    PropertyChanged "Caption1"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor1() As OLE_COLOR
    ForeColor1 = Label1.ForeColor
End Property

Public Property Let ForeColor1(ByVal New_ForeColor1 As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor1
    PropertyChanged "ForeColor1"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font1() As Font
    Set Font1 = Label1.Font
End Property

Public Property Set Font1(ByVal New_Font1 As Font)
    Set Label1.Font = New_Font1
    PropertyChanged "Font1"
End Property

Private Sub Image1_Click()
    RaiseEvent Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MouseAbajo
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Mouse_Arriba
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MouseAbajo
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Mouse_Arriba
End Sub

Private Sub Label2_Click()
    RaiseEvent Click
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MouseAbajo
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Mouse_Arriba
End Sub

Private Sub Timer1_Timer()
    Call MouseOut
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    UserControl.Height = 1200
    UserControl.Width = 3705
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MouseAbajo
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Mouse_Arriba
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Call Mouse_Arriba
    
    Label1.Caption = PropBag.ReadProperty("Caption1", "Label1")
    Label1.ForeColor = PropBag.ReadProperty("ForeColor1", &H80000012)
    Set Label1.Font = PropBag.ReadProperty("Font1", Ambient.Font)
    
    Label2.Caption = PropBag.ReadProperty("Caption2", "Label2")
    Label2.ForeColor = PropBag.ReadProperty("ForeColor2", &H80000012)
    Set Label2.Font = PropBag.ReadProperty("Font2", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    ImageFromFile = PropBag.ReadProperty("ImageFromFile", "")
    
    
    If Not Ambient.UserMode Then
        'Aca se puede cambiar al picture que quieras que aparesca a la hora de diseño asi para tener una idea de como va ir quedando
        'UserControl.Picture = picNormal.Picture
        'UserControl.Picture = picOver.Picture
        'UserControl.Picture = picClicked.Picture
        UserControl.Picture = picOver.Picture
        Timer1.Enabled = False
    End If
    
End Sub

Public Property Get ImageFromFile() As String
    ImageFromFile = m_cImageFromFile
End Property

Public Property Let ImageFromFile(c As String)
    m_cImageFromFile = c
    LoadImageFromFile c
    PropertyChanged "ImageFromFile"
End Property

Private Sub UserControl_Resize()
    UserControl.Height = 1200
    UserControl.Width = 3705
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption1", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("ForeColor1", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font1", Label1.Font, Ambient.Font)
    
    Call PropBag.WriteProperty("Caption2", Label2.Caption, "Label2")
    Call PropBag.WriteProperty("ForeColor2", Label2.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font2", Label2.Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    
    Call PropBag.WriteProperty("ImageFromFile", m_cImageFromFile, "")
    
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label2,Label2,-1,Caption
Public Property Get Caption2() As String
    Caption2 = Label2.Caption
End Property

Public Property Let Caption2(ByVal New_Caption2 As String)
    Label2.Caption() = New_Caption2
    PropertyChanged "Caption2"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label2,Label2,-1,ForeColor
Public Property Get ForeColor2() As OLE_COLOR
    ForeColor2 = Label2.ForeColor
End Property

Public Property Let ForeColor2(ByVal New_ForeColor2 As OLE_COLOR)
    Label2.ForeColor() = New_ForeColor2
    PropertyChanged "ForeColor2"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label2,Label2,-1,Font
Public Property Get Font2() As Font
    Set Font2 = Label2.Font
End Property

Public Property Set Font2(ByVal New_Font2 As Font)
    Set Label2.Font = New_Font2
    PropertyChanged "Font2"
End Property

'//----------------------------------------------------------------------

Private Sub MouseOut()
    'Obtenemos la Posicion del cursor con el Api getCursorPos
    Call GetCursorPos(Posicion_Cursor)

    'Obtenemos el Hwnd de la ventana o control donde esta el mouse
    Ret_Hwnd = WindowFromPoint(Posicion_Cursor.x, Posicion_Cursor.y)

    'Si el Hwnd no es del UserControl, Lo mostramos normal
    If Ret_Hwnd <> hwnd Then
        If picAcutal <> Normal Then
            UserControl.Picture = picNormal.Picture
            picAcutal = Normal
        End If
    Else
        ' Sino, mostramos cuando tiene el mouse encima
        If picAcutal <> Sobre Then
            UserControl.Picture = picOver.Picture
            picAcutal = Sobre
        End If
    End If
End Sub

Private Sub MouseAbajo()
    ' CLICK! Actualizamos, desactivamos el timer hasta que suelte el mouse :E
    UserControl.Picture = picClicked.Picture
    picAcutal = Clicked
    'Timer1.Enabled = False
    If Ambient.UserMode Then Timer1.Enabled = False
End Sub

Private Sub Mouse_Arriba()
    ' Solto el click, activamos el timer!!!
    UserControl.Picture = picOver.Picture
    picAcutal = Sobre
    'Timer1.Enabled = True
    If Ambient.UserMode Then Timer1.Enabled = True
End Sub

''Public Function LoadImageFromFile(ByVal sFile As String) As Boolean
''    LoadImageFromFile = picIcono.LoadImageFromFile(sFile)
''End Function

Public Function LoadImageFromFile(ByVal sFile As String) As Boolean

    LoadImageFromFile = picIcono.LoadImageFromFile(App.Path & "\iconos\" & sFile)
End Function


