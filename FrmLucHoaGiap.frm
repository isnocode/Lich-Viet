VERSION 5.00
Begin VB.Form FrmLucHoaGiap 
   Caption         =   "Bang Luc Hoa Giap"
   ClientHeight    =   8220
   ClientLeft      =   2130
   ClientTop       =   1725
   ClientWidth     =   10005
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   10005
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Gia1p"
   End
   Begin LunarCalendar.Label Label1 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Can Chi"
      WordWrap        =   -1  'True
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ty1"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Su73u"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Da62n"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ma4o"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Thi2n"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ty5"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ngo5"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Mu2i"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tha6n"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Da65u"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tua61t"
   End
   Begin LunarCalendar.Label LblChi 
      Height          =   495
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   7440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ho75i"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   1
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "A61t"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Bi1nh"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   3
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   16744576
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "D9inh"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   4
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ma65u"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   5
      Left            =   5040
      TabIndex        =   18
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ky3"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   6
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Canh"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   7
      Left            =   6600
      TabIndex        =   20
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ta6n"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   8
      Left            =   7440
      TabIndex        =   21
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   16744576
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nha6m"
   End
   Begin LunarCalendar.Label LblCan 
      Height          =   735
      Index           =   9
      Left            =   8280
      TabIndex        =   22
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Qu1y"
   End
End
Attribute VB_Name = "FrmLucHoaGiap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
