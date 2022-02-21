VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "Lich Viet"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   6315
      TabIndex        =   18
      Top             =   6360
      Width           =   6375
      Begin VB.ComboBox TxtYear 
         Height          =   360
         Left            =   1680
         TabIndex        =   29
         Text            =   "0000"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox TxtMonth 
         Height          =   360
         Left            =   960
         TabIndex        =   28
         Text            =   "00"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Txtday 
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Text            =   "00"
         Top             =   480
         Width           =   735
      End
      Begin LunarCalendar.Label LblDate 
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nga2y:"
         AutoSize        =   -1  'True
      End
      Begin LunarCalendar.OptionBox OptAmDuong 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "A6m Li5ch"
      End
      Begin LunarCalendar.Button CmdTaoLich 
         Height          =   855
         Left            =   4800
         TabIndex        =   21
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
      End
      Begin LunarCalendar.OptionBox OptAmDuong 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   23
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Du7o7ng Li5ch"
         Value           =   -1  'True
      End
      Begin LunarCalendar.Label LblDate 
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   25
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tha1ng:"
         AutoSize        =   -1  'True
      End
      Begin LunarCalendar.Label LblDate 
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   26
         Top             =   120
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Na8m:"
         AutoSize        =   -1  'True
      End
      Begin VB.Line Line3 
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   960
      End
   End
   Begin VB.PictureBox PicCalenda 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   6255
      Left            =   10
      ScaleHeight     =   6195
      ScaleWidth      =   6420
      TabIndex        =   0
      Top             =   10
      Width           =   6480
      Begin LunarCalendar.Label LblTucNgu1 
         Height          =   255
         Left            =   -480
         TabIndex        =   4
         Top             =   1680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Kho6ng co1 vie65c gi2 kho1"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblTucNgu2 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "chi3 so75 lo2ng kho6ng be62n"
         Alignment       =   1
      End
      Begin LunarCalendar.Label lblNam 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Na8m 2017"
         Alignment       =   1
      End
      Begin LunarCalendar.Label Lblthang 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tha1ng Gie6ng - January"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblNgayDL 
         Height          =   1215
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2143
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "09"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblTacgia 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tên"
         Alignment       =   2
      End
      Begin LunarCalendar.Label LblThuAL 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2880
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Thu71 Tu7"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblThangsoAL 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tha1ng Cha5p"
         Alignment       =   1
      End
      Begin LunarCalendar.Label lblThangDu 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "(D9u3)"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblNgaysoAL 
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "09"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblTietKhi 
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   4680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tie61t: Tie63u ha2n"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblGioAL 
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Gio72: Gia1p Ty1"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblNgayAL 
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   3540
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Nga2y: Gia1p Thi2n"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblThangAL 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   3840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tha1ng: Ta6n Su73u"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblNamAL 
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   4140
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Na8m: Bi1nh Tha6n"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblHoangdao 
         Height          =   1095
         Left            =   3000
         TabIndex        =   16
         Top             =   4560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Gio72 Hoa2ng D9a5o:"
         Alignment       =   1
         WordWrap        =   -1  'True
      End
      Begin LunarCalendar.Label lblThuDL 
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2880
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Wednesday"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblNgayHD 
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   5160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Nga2y Hoa2ng D9a5o"
         Alignment       =   1
      End
      Begin LunarCalendar.Label LblTruc 
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   4920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Tru75c:"
         Alignment       =   1
      End
      Begin LunarCalendar.ucImageList ImgListMoon 
         Left            =   2760
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         bvData          =   "Form1.frx":27A2
         bData           =   -1  'True
      End
      Begin LunarCalendar.ucImage imgMoon 
         Height          =   1455
         Left            =   2160
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2566
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         X1              =   0
         X2              =   6000
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Shape Shape1 
         Height          =   300
         Left            =   0
         Top             =   2860
         Width           =   6015
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   3000
         Y1              =   2880
         Y2              =   3120
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dD As Double, dM As Double, dy As Double

Private Sub CmdTaoLich_Click()
    On Error Resume Next
    Dim iResFile As Integer
    If OptAmDuong(1).Value = True Then
    dD = CDbl(Txtday.Text)
    dM = CDbl(TxtMonth.Text)
    dy = CDbl(TxtYear.Text)
    Else
    dD = AmLich2DuongLich(Txtday.Text, TxtMonth.Text, TxtYear.Text, 1, 7).fdNgay
    dM = AmLich2DuongLich(Txtday.Text, TxtMonth.Text, TxtYear.Text, 1, 7).fdThang
    dy = AmLich2DuongLich(Txtday.Text, TxtMonth.Text, TxtYear.Text, 1, 7).fdNam
    End If
    '-----------------
    With GetThongTinAmLich(dD, dM, dy, 7)
        lblNam.Caption = "Na8m " & dy
        Lblthang.Caption = DoiThangAL(dM) & " - " & DoiquathangEnglish(dM)
        LblNgayDL.Caption = Format(dD, "00")
        LblTucNgu1.Caption = MyEvent(dD, dM, dy).Dong1
        LblTucNgu2.Caption = MyEvent(dD, dM, dy).Dong2
        LblTacgia.Caption = MyEvent(dD, dM, dy).Tacgia
        LblTucNgu1.ForeColor = MyEvent(dD, dM, dy).lColor
        LblTucNgu2.ForeColor = MyEvent(dD, dM, dy).lColor
        
        LblThuAL.Caption = .sThuTrongTuanVN
        lblThuDL.Caption = .sThuTrongTuanEng
        LblThangsoAL.Caption = UNI(.sDoiThangAL)
        LblNgaysoAL.Caption = Format(.fdNgayAL.fdNgay, "00")
        If GetDaysOfMonth(dM, dy) > 30 Then
            lblThangDu.Caption = "( D9u3 )"
        Else
            lblThangDu.Caption = "( Thie61u )"
        End If
        
        LblNamAL.Caption = "Na8m: " & .sNamAmLich
        LblThangAL.Caption = "Tha1ng: " & .sThangAmLich
        LblNgayAL.Caption = "Nga2y: " & .sNgayAmLich
        LblGioAL.Caption = "Gio72: " & .sGioAmLich
        LblTietKhi.Caption = "Tie61t " & .sTietKhi & "(" & .fdThoigianbatDauTietKhiAL.fdNgay & "/" & .fdThoigianbatDauTietKhiAL.fdThang & ")" ' & vbCrLf & .fdThoigianbatDauTietKhiDL.fdNgay & "/" & .fdThoigianbatDauTietKhiDL.fdThang & "(DL)"
        LblTruc.Caption = ""
        
        LblNgayHD.Caption = sNgayHoangDao(dD, dM, dy)
        LblHoangdao.Caption = "Gio72 hoa2ng d9a5o: " & vbCrLf & Gio_Hoang_Dao(dD, dM, dy)
    End With
    iResFile = Angle(Txtday & "/" & TxtMonth & "/" & TxtYear) / 12
    If iResFile > 29 Then iResFile = 0
    imgMoon.LoadImageFromStream ImgListMoon.GetStream(iResFile)
End Sub
Private Sub Form_Load()
Dim i
For i = 1 To 31 'ngay
Txtday.AddItem i
Next i
For i = 1 To 12 'thang
TxtMonth.AddItem i
Next i
For i = 1 To 3000 'nam
TxtYear.AddItem Format(i, "0000")
Next i
    Txtday = Day(Date): TxtMonth = Month(Date): TxtYear = Year(Date)
    Call CmdTaoLich_Click
End Sub

