VERSION 5.00
Begin VB.UserControl UniComboBox 
   BackColor       =   &H80000005&
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   ScaleHeight     =   285
   ScaleWidth      =   1920
   ToolboxBitmap   =   "UniComboBox.ctx":0000
End
Attribute VB_Name = "UniComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ISubclass


' WinAPI:
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type SIZEAPI
   cX As Long
   cY As Long
End Type
'Private Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemId As Long
    ItemAction As Long
    ItemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageByref Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextExAsNull Lib "user32" Alias "DrawTextExW" (ByVal hdc As Long, ByVal lpsz As Long, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Long) As Long
'Private Declare Function DrawTextExAsNull Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, hWnd2 As Any, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function FindWindowExW Lib "user32" (ByVal hWnd1 As Long, hWnd2 As Long, ByVal lpsz1 As Long, lpsz2 As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthA Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As Long) As Long

'''''''''API dung de so sanh chuoi unicode
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function CompareString Lib "kernel32.dll" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
Const CSTR_LESS_THAN = 1
Const CSTR_EQUAL = 2
Const CSTR_GREATER_THAN = 3
Const LOCALE_SYSTEM_DEFAULT = &H400
Const LOCALE_USER_DEFAULT = &H800
Const NORM_IGNORECASE = &H1
Const NORM_IGNOREKANATYPE = &H10000
Const NORM_IGNORENONSPACE = &H2
Const NORM_IGNORESYMBOLS = &H4
Const NORM_IGNOREWIDTH = &H20000
Const SORT_STRINGSORT = &H1000


Private Const LF_FULLFACESIZE = 64
Private Type ENUMLOGFONTEX
    elfLogFont As LOGFONT
    elfFullName(LF_FULLFACESIZE - 1) As Byte
    elfStyle(LF_FACESIZE - 1) As Byte
    elfScript(LF_FACESIZE - 1) As Byte
End Type

Private Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ' Additional to TEXTMETRIC
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Private Type FONTSIGNATURE
   fsUsb(4) As Long
   fsCsb(2) As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type NEWTEXTMETRICEX
    ntmTm As NEWTEXTMETRIC
    ntmFontSig As FONTSIGNATURE
End Type
Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long

'/* EnumFonts Masks */
Private Const RASTER_FONTTYPE = 1&
Private Const DEVICE_FONTTYPE = 2&
Private Const TRUETYPE_FONTTYPE = 4&

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const GB2312_CHARSET = 134
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255
Private Const JOHAB_CHARSET = 130
Private Const HEBREW_CHARSET = 177
Private Const ARABIC_CHARSET = 178
Private Const GREEK_CHARSET = 161
Private Const TURKISH_CHARSET = 162
Private Const THAI_CHARSET = 222
Private Const EASTEUROPE_CHARSET = 238
Private Const RUSSIAN_CHARSET = 204

Private Const MAC_CHARSET = 77
Private Const BALTIC_CHARSET = 186

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Const WS_BORDER = &H800000
Private Const WS_TABSTOP = &H10000

Private Const GCL_HBRBACKGROUND = (-10)

Private Const GW_CHILD = 5
Private Const WM_SETFOCUS = &H7
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_DRAWITEM = &H2B
Private Const WM_SETFONT = &H30
Private Const WM_NOTIFY = &H4E
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_COMMAND = &H111
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_USER = &H400
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDOWN = &H201
Private Const GWL_ID As Long = -12
Private Const WM_CAPTURECHANGED As Long = &H215

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1

Private Const MA_NOACTIVATE = 3

Private Const BITSPIXEL = 12
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0

Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17 'win95 only

' Draw text flags:
Private Const DT_EDITCONTROL = &H2000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_TABSTOP = &H80
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_NOCLIP = &H100
Private Const DT_INTERNAL = &H1000
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EXPANDTABS = &H40
Private Const DT_CHARSTREAM = 4
Private Const DT_NOPREFIX = &H800
Private Const DT_CALCRECT = &H400

' Draw edge constants:
Private Const BF_LEFT = 1
Private Const BF_TOP = 2
Private Const BF_RIGHT = 4
Private Const BF_BOTTOM = 8
Private Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BF_MIDDLE = 2048
Private Const BDR_SUNKENINNER = 8
Private Const BDR_SUNKENOUTER = 2

' Combo box styles:
Private Const CBS_AUTOHSCROLL As Long = &H40&
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBS_HASSTRINGS = &H200&
Private Const CBS_DISABLENOSCROLL = &H800&
Private Const CBS_NOINTEGRALHEIGHT = &H400&
Private Const CBS_OWNERDRAWFIXED = &H10&
Private Const CBS_OWNERDRAWVARIABLE = &H20&
Private Const CBS_SIMPLE = &H1&
Private Const CBS_SORT = &H100&

' Combo box messages:
Private Const CB_ADDSTRING = &H143
Private Const CB_DELETESTRING = &H144
Private Const CB_DIR = &H145
Private Const CB_ERR = (-1)
Private Const CB_ERRSPACE = (-2)
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETDROPPEDCONTROLRECT = &H152
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_GETEDITSEL = &H140
Private Const CB_GETEXTENDEDUI = &H156
Private Const CB_GETITEMDATA = &H150
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_GETLOCALE = &H15A
Private Const CB_INSERTSTRING = &H14A
Private Const CB_LIMITTEXT = &H141
Private Const CB_MSGMAX = &H15B
Private Const CB_OKAY = 0
Private Const CB_RESETCONTENT = &H14B
Private Const CB_SELECTSTRING = &H14D
Private Const CB_SETCURSEL = &H14E
Private Const CB_SETEDITSEL = &H142
Private Const CB_SETEXTENDEDUI = &H155
Private Const CB_SETITEMDATA = &H151
Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_SETLOCALE = &H159
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160

' Combo box notifications:
Private Const CBN_CLOSEUP = 8
Private Const CBN_DBLCLK = 2
Private Const CBN_DROPDOWN = 7
Private Const CBN_EDITCHANGE = 5
Private Const CBN_EDITUPDATE = 6
Private Const CBN_KILLFOCUS = 4
Private Const CBN_SELCHANGE = 1
Private Const CBN_SELENDCANCEL = 10
Private Const CBN_SELENDOK = 9
Private Const CBN_SETFOCUS = 3

' Owner draw style types:
Private Const ODS_CHECKED = &H8
Private Const ODS_DISABLED = &H4
Private Const ODS_FOCUS = &H10
Private Const ODS_GRAYED = &H2
Private Const ODS_SELECTED = &H1
Private Const ODS_COMBOBOXEDIT = &H1000

' Owner draw action types:
Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_FOCUS = &H4
Private Const ODA_SELECT = &H2

Private Const CLR_NONE = -1

' Edit box:
Private Const EM_GETSEL = &HB0
Private Const EM_SETSEL = &HB1
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT

' CC API
Private Const H_MAX As Long = &HFFFF + 1
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Type NMHDR
   hwndFrom As Long
   idfrom As Long
   code As Long
End Type

Private Const CCM_FIRST              As Long = &H2000
Private Const CCM_SETUNICODEFORMAT   As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT   As Long = (CCM_FIRST + 6)
Private Const CCM_SETBKCOLOR         As Long = (CCM_FIRST + 1)        '// lParam is bkColor

' ImageList API:
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cX As Long, cY As Long) As Long

Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840

' CBEX API:
Private Const WC_COMBOBOXEX = "ComboBoxEx32"

Private Type COMBOBOXEXITEM
   mask As Long    ' CBEIF..
   iItem As Long
   pszText As Long ' String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   iOverlay As Long
   iIndent As Long
   lParam As Long
End Type

Private Const CBEIF_TEXT = &H1
Private Const CBEIF_IMAGE = &H2
Private Const CBEIF_SELECTEDIMAGE = &H4
Private Const CBEIF_OVERLAY = &H8
Private Const CBEIF_INDENT = &H10
Private Const CBEIF_LPARAM = &H20
Private Const CBEIF_DI_SETITEM = &H10000000

' Combo box extended messages:
Private Const CBEM_SETIMAGELIST = (WM_USER + 2)
Private Const CBEM_GETIMAGELIST = (WM_USER + 3)
Private Const CBEM_DELETEITEM = CB_DELETESTRING
Private Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL = (WM_USER + 7)
Private Const CBEM_SETEXSTYLE = (WM_USER + 8)
Private Const CBEM_GETEXSTYLE = (WM_USER + 9)
Private Const CBEM_HASEDITCHANGED = (WM_USER + 10)
Private Const CBEM_INSERTITEMA = (WM_USER + 1)
Private Const CBEM_INSERTITEMW = (WM_USER + 11)
Private Const CBEM_SETITEMA = (WM_USER + 5)
Private Const CBEM_SETITEMW = (WM_USER + 12)
Private Const CBEM_GETITEMA = (WM_USER + 4)
Private Const CBEM_GETITEMW = (WM_USER + 13)
Private Const CBEM_INSERTITEM = CBEM_INSERTITEMW 'CBEM_INSERTITEMA

' Combo box extended notifications:
Private Const CBEN_FIRST = (H_MAX - 800&)
Private Const CBEN_LAST = (H_MAX - 830&)
Private Const CBEN_GETDISPINFO = (CBEN_FIRST - 0)
Private Const CBEN_INSERTITEM = (CBEN_FIRST - 1)
Private Const CBEN_DELETEITEM = (CBEN_FIRST - 2)
Private Const CBEN_BEGINEDIT = (CBEN_FIRST - 4)
Private Const CBEN_ENDEDITA = (CBEN_FIRST - 5)
Private Const CBEN_ENDEDITW = (CBEN_FIRST - 6)
Private Const CBEN_ENDEDIT = CBEN_ENDEDITW 'CBEN_ENDEDITA

' Combo box extended styles:
Private Const CBES_EX_NOEDITIMAGE = &H1& ' no image to left of edit portion
Private Const CBES_EX_NOEDITIMAGEINDENT = &H2& ' edit box and dropdown box will not display images
Private Const CBES_EX_PATHWORDBREAKPROC = &H4& ' NT only. Edit box uses \ . and / as word delimiters
'#if (_WIN32_IE >= 0x0400)
Private Const CBES_EX_NOSIZELIMIT = &H8& ' Allow combo box ex vertical size < combo, clipped.
Private Const CBES_EX_CASESENSITIVE = &H10& ' case sensitive search

Private Const CBEMAXSTRLEN = 260
Private Type NMCBEENDEDIT
    hdr As NMHDR
    fChanged As Long
    iNewSelection As Long
    szText(0 To CBEMAXSTRLEN - 1) As Byte '// CBEMAXSTRLEN is 260
    iWhy As Integer
End Type
Private Type NMCBEENDEDITW
    hdr As NMHDR
    fChanged As Long
    iNewSelection As Long
    szText(0 To 518) As Byte '// CBEMAXSTRLEN is 260
    iWhy As Long
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, ByVal dwAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Enum EShellGetFileInfoConstants
   SHGFI_ICON = &H100                       ' // get icon
   SHGFI_DISPLAYNAME = &H200                ' // get display name
   SHGFI_TYPENAME = &H400                   ' // get type name
   SHGFI_ATTRIBUTES = &H800                 ' // get attributes
   SHGFI_ICONLOCATION = &H1000              ' // get icon location
   SHGFI_EXETYPE = &H2000                   ' // return exe type
   SHGFI_SYSICONINDEX = &H4000              ' // get system icon index
   SHGFI_LINKOVERLAY = &H8000               ' // put a link overlay on icon
   SHGFI_SELECTED = &H10000                 ' // show icon in selected state
   SHGFI_ATTR_SPECIFIED = &H20000           ' // get only specified attributes
   SHGFI_LARGEICON = &H0                    ' // get large icon
   SHGFI_SMALLICON = &H1                    ' // get small icon
   SHGFI_OPENICON = &H2                     ' // get open icon
   SHGFI_SHELLICONSIZE = &H4                ' // get shell size icon
   SHGFI_PIDL = &H8                         ' // pszPath is a pidl
   SHGFI_USEFILEATTRIBUTES = &H10           ' // use passed dwFileAttribute
End Enum
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Public Enum ECCXExtendedStyle
   eccxNoEditImage = CBES_EX_NOEDITIMAGE
   eccxNoImages = CBES_EX_NOEDITIMAGEINDENT
   eccxCaseSensitiveSearch = CBES_EX_CASESENSITIVE
End Enum
Public Enum EDriveType
   DRIVE_REMOVABLE = 2
   DRIVE_FIXED = 3
   DRIVE_REMOTE = 4
   DRIVE_CDROM = 5
   DRIVE_RAMDISK = 6
End Enum
Public Enum ECCXComboStyle
   eccxDropDownCombo
   eccxSimple
   eccxDropDownList
End Enum
' End edit reasons:
Public Enum ECCXEndEditReason
   CBENF_KILLFOCUS = 1
   CBENF_RETURN = 2
   CBENF_ESCAPE = 3
   CBENF_DROPDOWN = 4
End Enum
Public Enum ECCXDrawMode
    ' -- Owner draw styles --
    eccxDrawDefault              ' default comboboxex draw
    eccxDrawDefaultThenClient    ' default comboboxex draw then raise DrawItem event
    eccxDrawODCboList            ' ODCboLst style draw
    eccxDrawODCboListThenClient  ' ODCboLst style draw then raise DrawItem event
    eccxOwnerDraw                ' you do all drawing yourself
    
    ' -- Special styles --
    eccxColourPickerWithNames
    eccxColourPickerNoNames
    eccxSysColourPicker
    eccxDriveList
End Enum

' Column type enums
Public Enum ECCXColumnType
   eccxTextString = 0       ' The default - draw as text, sort as text
   eccxTextNumber = 1       ' Convert to number during sort
   eccxTextDateTime = 2     ' Convert to date for sort
   eccxImageListIcon = 4    ' Convert to icon index in image list & assume numeric during sort
End Enum

Private m_hWnd              As Long
Private m_hWndCbo           As Long
Private m_hWndEdit          As Long
Private m_hWndParent        As Long
Private m_hWndDropDown      As Long

Private m_bSubclass         As Boolean
Private m_hFnt              As Long
Private m_hFntOld           As Long
Private m_tLF               As LOGFONT
Private m_hUFnt             As Long
Private m_tULF              As LOGFONT
Private m_fnt               As StdFont
Private m_bInFocus          As Boolean
Private m_bEvents           As Boolean
Private m_bDesignTime       As Boolean
Private m_oBackColor        As OLE_COLOR
Private m_hBrBack           As Long
Private m_oForeColor        As OLE_COLOR

Private m_eStyle            As ECCXComboStyle
Private m_bSorted           As Boolean
Private m_bExtendedUI       As Boolean
Private m_eExStyle          As ECCXExtendedStyle
Private m_eClientDraw       As ECCXDrawMode
Private tCBItem             As COMBOBOXEXITEM

Private m_hIml              As Long
Private m_lIconSizeY        As Long
Private m_bEnabled          As Boolean
Private m_bRedraw           As Boolean
Private m_lWidth            As Long
Private m_lMaxLength        As Long
Private m_lNewIndex         As Long

Private m_iColCount         As Long
Private m_lColWidth()       As Long
Private m_eCoLType()        As ECCXColumnType

Private m_sDriveStrings     As String
Private m_cCbo              As UniComboBox
Private m_iType             As Long
Private m_iCharSet          As Long
Private bUnicode            As Boolean

' Auto complete mode for drop-down combo boxes:
Private m_bDoAutoComplete           As Boolean
Private m_bOnlyAutoCompleteItems    As Boolean
Private m_bDataIsSorted             As Boolean
Private m_IPAOHookStruct            As IPAOHookStructComboBox

Public Event AutoCompleteSelection(ByVal sItem As String, ByVal lIndex As Long)
Public Event BeginEdit(ByVal iIndex As Long)
Public Event Change()
Attribute Change.VB_MemberFlags = "200"
Public Event Click()
Public Event ListIndexChange()
Public Event CloseUp()
Public Event DblClick()
Public Event DrawItem(ByVal ItemIndex As Long, ByVal hdc As Long, ByVal bSelected As Boolean, ByVal bEnabled As Boolean, ByVal LeftPixels As Long, ByVal TopPixels As Long, ByVal RightPixels As Long, ByVal BottomPixels As Long, ByVal hFntOld As Long)
Public Event DropDown()
Public Event EndEdit(ByVal iIndex As Long, ByVal bChanged As Boolean, ByVal sText As String, ByVal eWHy As ECCXEndEditReason, ByVal iNewIndex As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event RequestDropDownResize(ByRef lLeft As Long, ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long, ByRef bCancel As Boolean)



Public Property Get AutoUnicode() As Boolean
    AutoUnicode = bUnicode
End Property

Public Property Let AutoUnicode(ByVal Auto_Uni As Boolean)
    bUnicode = Auto_Uni
    If Not (m_hWnd = 0) Then
       pCache True
       Clear
       pCache
    End If
    PropertyChanged "AutoUnicode"
End Property

Private Sub pDefaultDrawItem(ByVal hdc As Long, ByVal ItemId As Long, _
                             ByVal ItemAction As Long, ByVal ItemState As Long, _
                             ByVal left As Long, ByVal Top As Long, ByVal right As Long, ByVal bottom As Long)
    Dim tR As RECT, tIR As RECT, tTR As RECT
    Dim hPen As Long, hPenOld As Long
    Dim hBrush As Long, lCol As Long, hMem As Long
    Dim sItem As String, tP As POINTAPI
    Dim bSelected As Boolean, iColCount As Integer
    Dim lLeft As Long, bFocus As Boolean, lFocus As Long

        ' Debug.Print "DefaultDrawItem"
        lFocus = GetFocus()
        bFocus = ((lFocus = m_hWnd) Or (lFocus = m_hWndParent)) Or (lFocus = m_hWndCbo) Or (lFocus = m_hWndEdit)
            
        ' Determine the default draw mechanism:
        Select Case m_eClientDraw
            Case eccxColourPickerNoNames, eccxColourPickerWithNames, eccxSysColourPicker
            ' Do ColourPicker:
                pDrawColorPicker hdc, ItemId, ItemAction, ItemState, left, Top, right, bottom
            
            Case eccxDrawODCboList, eccxDrawODCboListThenClient, eccxDrawDefault
                With tR
                    .left = left
                    .Top = Top
                    .right = right
                    .bottom = bottom
                End With
                
                ' Debug.Print ItemId
                If (ItemId <> -1) Then
                    sItem = List(ItemId)
                Else
                    sItem = ""
                End If
            
                '' Debug.Print sItem, hdc, left, Right, tOp, Bottom
                If (ItemState And ODS_DISABLED) = ODS_DISABLED Then
                    lLeft = tR.left
                    If (ItemState And ODS_COMBOBOXEDIT) <> ODS_COMBOBOXEDIT Then
                        tR.left = tR.left + ItemIndent(ItemId)
                    End If
            
                    If (ItemId > -1) Then
                        If (ItemIcon(ItemId) > -1) Then
                            ImageList_DrawEx m_hIml, ItemIcon(ItemId), hdc, tR.left + 2, tR.Top, 0, 0, CLR_NONE, GetSysColor(vbWindowBackground And &H1F&), ILD_TRANSPARENT Or ILD_SELECTED
                            tR.left = tR.left + m_lIconSizeY + 4
                        End If
                    End If
                    If (ItemState And ODS_SELECTED) = ODS_SELECTED Then
                        lCol = GetSysColor(vbButtonFace And &H1F&)
                        SetBkColor hdc, lCol
                        lCol = GetSysColor(vbWindowBackground And &H1F&)
                        SetBkMode hdc, OPAQUE
                    Else
                        lCol = GetSysColor(vbButtonShadow And &H1F&)
                        SetBkMode hdc, TRANSPARENT
                    End If
                
                    tR.Top = tR.Top + 1
                    SetTextColor hdc, lCol
                
                    pDrawText hdc, ItemState, sItem, lLeft, DT_WORD_ELLIPSIS Or DT_SINGLELINE Or DT_LEFT, tR
                Else
                    SetBkMode hdc, OPAQUE
                    ' Set the forecolour to use for this draw:
                    ' Determine selection state:
                    bSelected = ((ItemState And ODS_SELECTED) = ODS_SELECTED)
                    If (bSelected) Then
                        ' Only draw selected in the combo when the
                        ' focus is on the control:
                        If (ItemState And ODS_COMBOBOXEDIT) = ODS_COMBOBOXEDIT Then
                            bSelected = False
                        End If
                    End If
                    ' Set the Text Colour of the DC to according to
                    ' the selection state:
                    If bSelected Then
                        ' Draw selected:
                        If m_eStyle <> eccxDropDownList Or bFocus Then
                            lCol = GetSysColor(vbHighlightText And &H1F&)
                            SetTextColor hdc, lCol
                            lCol = GetSysColor(vbHighlight And &H1F&)
                        Else
                            lCol = GetSysColor(vbWindowText And &H1F&)
                            SetTextColor hdc, lCol
                            OleTranslateColor m_oBackColor, 0, lCol 'GetSysColor(vbWindowBackground And &H1F&)
                        End If
                    Else
                        ' Draw normal:
                        OleTranslateColor UserControl.ForeColor, 0, lCol
                        SetTextColor hdc, lCol
                        ' Determine the back colour for this item:
                        OleTranslateColor m_oBackColor, 0, lCol
                    End If
                
                ' We only need to clear the background when
                ' the entire list box is being redrawn, or when
                ' the full-row select mode is on and the row is
                ' selected (this avoids some flicker):
'                    If (ItemAction = ODA_SELECT) Then
                        'hBrush = CreateSolidBrush(lCol)
                        'LSet tTR = tR
                        'FillRect hdc, tTR, hBrush
                        'DeleteObject hBrush
'                    End If
                
                    SetBkColor hdc, lCol
                    lLeft = tR.left
                        
                    ' Show the indent if this is not the edit box
                    ' portion of the combo box:
                    If (ItemState And ODS_COMBOBOXEDIT) <> ODS_COMBOBOXEDIT Then
                        tR.left = tR.left + ItemIndent(ItemId)
                    End If
                
'get text edit position
'                    GetWindowRect m_hWndEdit, tIR
                
                    ' If we have an icon, then draw it:
                    If (ItemIcon(ItemId) > -1) And sItem <> "" Then
                        ' Use the image list handle specified via the
                        ' ImageList property:
                        Select Case UserControl.Font.Size
                            Case Is <= 11: tIR.Top = tR.Top + UserControl.Font.Size - 8
                            Case Else: tIR.Top = tR.Top + UserControl.Font.Size - 10
                        End Select
                        ImageList_Draw m_hIml, ItemIcon(ItemId), hdc, tR.left + 2, tIR.Top, ILD_TRANSPARENT
'                        ImageList_Draw m_hIml, ItemIcon(ItemId), hdc, tR.Left + 2, tR.Top, ILD_TRANSPARENT
'                        ' Adjust draw position for the icon:
                        tR.left = tR.left + m_lIconSizeY + 4
                    End If
                
                ' Draw the text of the item:
                    tR.left = tR.left + 3:      tR.Top = tR.Top + 3
                    pDrawText hdc, ItemState, sItem, lLeft, DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_NOPREFIX Or DT_LEFT, tR
               End If
        End Select
End Sub

Private Sub pDrawColorPicker( _
        ByVal hdc As Long, _
        ByVal index As Long, _
        ByVal ItemAction As Long, _
        ByVal ItemState As Long, _
        ByVal LeftPixels As Long, ByVal TopPixels As Long, ByVal RightPixels As Long, ByVal BottomPixels As Long _
    )
Dim tR As RECT, hBrush As Long, tS As RECT
Dim bSelected As Boolean
Dim lCol As Long
    
   If (index <> -1) Then
      ' Debug.Print "DrawColorPicker"
    
      bSelected = ((ItemState And ODS_SELECTED) = ODS_SELECTED)

      SetBkMode hdc, TRANSPARENT
      
      tR.Top = TopPixels
      tR.bottom = BottomPixels
      tR.left = LeftPixels
      tR.right = RightPixels
      If (bSelected) Then
         hBrush = GetSysColorBrush(vbHighlight And &H1F&) 'CreateSolidBrush(gTranslateColor(vbHighlight))
         FillRect hdc, tR, hBrush
         DeleteObject hBrush
      Else
         If (ItemAction = ODA_SELECT) Then
            'hBrush = GetSysColorBrush(vbWindowBackground And &H1F&) 'CreateSolidBrush(gTranslateColor(vbWindowBackground))
            FillRect hdc, tR, m_hBrBack 'hBrush
            'DeleteObject hBrush
         End If
      End If
      
      'Debug.Print Index, hDC, bSelected, bEnabled, LeftPixels, TopPixels, RightPixels, BottomPixels
      
      tR.Top = TopPixels + 1
      tR.bottom = BottomPixels - 1
      tR.left = LeftPixels + 2
      If (m_eClientDraw = eccxColourPickerNoNames) Then
         tR.Top = tR.Top + 1
         tR.bottom = tR.bottom - 1
         tR.right = RightPixels - 2
      Else
         tR.right = tR.left + (tR.bottom - tR.Top)
      End If
      ' Draw sunken border:
      DrawEdge hdc, tR, BDR_SUNKENOUTER Or BDR_SUNKENINNER, (BF_RECT Or BF_MIDDLE)
      
      ' Draw the sample colour:
      OleTranslateColor ItemData(index), 0, lCol
      hBrush = CreateSolidBrush(lCol)
      LSet tS = tR
      tS.left = tS.left + 2
      tS.right = tS.right - 2
      tS.Top = tS.Top + 2
      tS.bottom = tS.bottom - 2
      FillRect hdc, tS, hBrush
      DeleteObject hBrush
      
      If (m_eClientDraw <> eccxColourPickerNoNames) Then
         ' Now write the caption
         If (bSelected) Then
            SetTextColor hdc, GetSysColor(vbHighlightText And &H1F&)
         Else
            SetTextColor hdc, GetSysColor(vbWindowText And &H1F&)
         End If
         tR.left = tR.right + 2
         tR.right = RightPixels
         DrawText hdc, StrPtr(List(index)), -1, tR, DT_LEFT Or DT_WORD_ELLIPSIS Or DT_SINGLELINE Or DT_NOPREFIX
      End If
   End If
    
End Sub
Private Sub pDrawText(ByVal hdc As Long, ByVal ItemState As Long, ByVal sItem As String, ByVal lLeft As Long, ByVal lAlign As Long, ByRef tR As RECT)
Dim tCR As RECT
Dim iColCount As Integer
Dim iCol As Integer
Dim sColVals() As String
   
   ' We potentially have > 1 column.  If this isn't the edit portion of a combo
   ' box, and we have specified that there are > 1 columns for the box,
   ' then draw according to the specified column widths.  Otherwise, use default
   ' drawing means.
   If (m_iColCount > 1) And (ItemState And ODS_COMBOBOXEDIT) <> ODS_COMBOBOXEDIT Then
      ' Split sItem according to vbTab:
      gSplitDelimitedString sItem, vbTab, sColVals(), iColCount
      ' Add attributes to truncate text and draw ellipsis (..) if too long
      lAlign = lAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING Or DT_NOPREFIX
      ' Set up rectangle for first column
      LSet tCR = tR
      tCR.right = lLeft + m_lColWidth(1)
      ' Always Draw the first item:
      If (m_eCoLType(1) = eccxImageListIcon) Then
         ImageList_Draw m_hIml, glCStr(sColVals(1), -1), hdc, tCR.left, tCR.Top - 2, ILD_TRANSPARENT
      Else
         DrawTextExAsNull hdc, StrPtr(sColVals(1)), -1, tCR, lAlign, 0
      End If
      For iCol = 2 To iColCount
         If (iCol > m_iColCount) Then
            ' Don't attempt to draw columns that we don't have:
            Exit For
         End If
         tCR.left = tCR.right + 1
         tCR.right = tCR.left + m_lColWidth(iCol)
         Select Case m_eCoLType(iCol)
         Case eccxImageListIcon
            ImageList_Draw m_hIml, glCStr(sColVals(iCol), -1), hdc, tCR.left, tCR.Top - 2, ILD_TRANSPARENT
         Case Else
            DrawTextExAsNull hdc, StrPtr(sColVals(iCol)), -1, tCR, lAlign, 0
         End Select
      Next iCol
   Else
      lAlign = DT_LEFT Or DT_NOPREFIX
      DrawTextExAsNull hdc, StrPtr(sItem), -1, tR, lAlign, 0
   End If
        
End Sub

Public Property Get AutoCompleteItemsAreSorted() As Boolean
   AutoCompleteItemsAreSorted = m_bDataIsSorted
End Property
Public Property Let AutoCompleteItemsAreSorted(ByVal bState As Boolean)
   m_bDataIsSorted = bState
   PropertyChanged "AutoCompleteItemsAreSorted"
End Property
Public Property Get AutoCompleteListItemsOnly() As Boolean
   AutoCompleteListItemsOnly = m_bOnlyAutoCompleteItems
End Property
Public Property Let AutoCompleteListItemsOnly(ByVal bState As Boolean)
   m_bOnlyAutoCompleteItems = bState
   PropertyChanged "AutoCompleteItemsListItemsOnly"
End Property

Public Property Get DoAutoComplete() As Boolean
   If (m_eStyle = eccxDropDownCombo) Or (m_eStyle = eccxSimple) Then
      DoAutoComplete = m_bDoAutoComplete
   Else
      'Err.Raise 383, App.EXEName & ".UniComboBox"
      DoAutoComplete = False
   End If
End Property
Public Property Let DoAutoComplete(ByVal bState As Boolean)
   m_bDoAutoComplete = bState
   PropertyChanged "DoAutoComplete"
End Property

Public Sub AutoCompleteKeyPress(ByRef iKeyAscii As Integer)
Dim sTotal As String, sLTotal As String, sUnSel As String, sLUnSel As String
Dim lLen As Long, iFound As Long, i As Long, lS As Long, lW As Long
Dim iStart As Long, iSelStart As Long, iSelLength As Long
Dim sText As String, str1 As String, str2 As String, hTL As Long

On Error GoTo ErrorHandler
   
   If (iKeyAscii = vbKeyReturn) Then
      If (ListIndex > -1) Then
        SelStart = 0
        SelLength = Len(List(ListIndex))
        RaiseEvent AutoCompleteSelection(List(ListIndex), ListIndex)
        Exit Sub
      End If
   ElseIf (iKeyAscii = vbKeyEscape) Then
      Exit Sub
   End If
   
   lS = SelStart
   lW = SelLength
   
   If (lS > 0) Then
      Call DetachEdit
      sUnSel = left$(pvGetWindowText(m_hWndEdit), lS)
      Call AttachEdit
   End If
   If (iKeyAscii = 8) Then
      If (Len(sUnSel) > 1) Then
         sTotal = left$(sUnSel, Len(sUnSel) - 1)
      Else
         sUnSel = ""
         iKeyAscii = 0
         Text = ""
         Exit Sub
      End If
   Else
      sTotal = sUnSel & ChrW$(iKeyAscii)
   End If
   
   ' try to match the the string entered:
   iFound = -1
   sLTotal = LCase$(sTotal)
   lLen = Len(sLTotal)
   
'str1, str2 la 2 chuoi Unicode can so sanh
    hTL = GetThreadLocale()
    str2 = LCase$(sLTotal)
    For i = 0 To ListCount - 1
        str1 = LCase$(left$(List(i), lLen))
'       If Left$(LCase$(List(i)), lLen) = LCase$(sLTotal) Then
'        If CompareString(hTL, NORM_IGNORECASE, StrPtr(str1), Len(str1), StrPtr(str2), Len(str2)) = CSTR_EQUAL Then
'            iFound = i
'            Exit For
'        End If
        If CompareString(hTL, NORM_IGNORECASE, StrPtr(str2), Len(str2), StrPtr(str1), Len(str1)) = CSTR_EQUAL Then
            iFound = i
            Exit For
        End If
    Next i
   
   If (iFound > -1) Then
      ListIndex = iFound
      iSelStart = Len(sTotal)
      iSelLength = Len(List(iFound)) - iSelStart
      'Debug.Print iSelStart, iSelLength
      SelStart = iSelStart
      SelLength = iSelLength
      'Debug.Print SelStart, SelLength
      iKeyAscii = 0
   Else
      If (m_bOnlyAutoCompleteItems) Then
         ' is there anything we can choose which has the same unmatched letters?
         iStart = ListIndex
         sLUnSel = LCase$(sUnSel)
         lLen = Len(sLUnSel)
         If (lLen > 0) Then
            If (m_bDataIsSorted) Then
               ' Its either the next one down or the first in the list:
               i = iStart + 1
               If StrComp(LCase$(left$(List(i), lLen)), sLUnSel) = 0 Then
                  iFound = i
               Else
                  For i = 0 To iStart - 1
                     If StrComp(LCase$(left$(List(i), lLen)), sLUnSel) = 0 Then
                        iFound = i
                        Exit For
                     End If
                  Next i
               End If
            Else
               ' it could be anything following list index, or anything preceeding it:
               For i = iStart + 1 To ListCount - 1
                  If StrComp(LCase$(left$(List(i), lLen)), sLUnSel) = 0 Then
                     iFound = i
                     Exit For
                  End If
               Next i
               If (iFound < 0) Then
                  For i = 0 To iStart - 1
                     If StrComp(LCase$(left$(List(i), lLen)), sLUnSel) = 0 Then
                        iFound = i
                        Exit For
                     End If
                  Next i
               End If
            End If
            If (iFound > -1) Then
               ListIndex = iFound
               SelStart = lLen
               SelLength = Len(List(iFound)) - SelStart + 1
            End If
         Else
             Beep
         End If
         iKeyAscii = 0
      Else
         Debug.Print "Not found, still works?"
         'SendMessageLong m_hWnd, CB_SETCURSEL, -1, 0
      End If
   End If
   Exit Sub
   
ErrorHandler:
   If (m_bOnlyAutoCompleteItems) Then
      iKeyAscii = 0
   End If
   Exit Sub
End Sub

Public Property Get ExtendedStyle(ByVal eStyle As ECCXExtendedStyle) As Boolean
   ExtendedStyle = ((m_eExStyle And eStyle) = eStyle)
End Property
Public Property Let ExtendedStyle(ByVal eStyle As ECCXExtendedStyle, ByVal bState As Boolean)
   If bState Then
      m_eExStyle = m_eExStyle Or eStyle
   Else
      m_eExStyle = m_eExStyle And Not eStyle
   End If
   If m_hWnd <> 0 Then
      SendMessageLong m_hWnd, CBEM_SETEXSTYLE, 0, m_eExStyle
   End If
End Property
Public Property Get DrawStyle() As ECCXDrawMode
   DrawStyle = m_eClientDraw
End Property
Public Property Let DrawStyle(ByVal eStyle As ECCXDrawMode)
   If eStyle <> m_eClientDraw Then
      m_eClientDraw = eStyle
      PropertyChanged "DrawStyle"
   End If
End Property

Public Property Get Font() As StdFont
   ' Get the control's default font:
   Set Font = UserControl.Font
End Property
Public Property Set Font(fntThis As StdFont)
Dim hFnt As Long
Dim tFnt As LOGFONT
Dim lH As Long
Dim tR As RECT

   ' Set the control's default font:
   Set UserControl.Font = fntThis
   ' Store a log font structure for this font:
   pOLEFontToLogFont fntThis, UserControl.hdc, tFnt
   ' Store old font handle:
   hFnt = m_hFnt
   ' Create a new version of the font:
   m_hFnt = CreateFontIndirect(tFnt)
   If (m_hWnd <> 0) Then
      ' Ensure the control has the correct font:
      SendMessageLong m_hWnd, WM_SETFONT, m_hFnt, 1
   End If
   ' Delete previous version, if we had one:
   If (hFnt <> 0) Then
      DeleteObject hFnt
   End If
   
   ' Make sure the User Control's height is correct:
   If m_eStyle <> eccxSimple Then
      lH = SendMessageLong(m_hWnd, CB_GETITEMHEIGHT, -1, 0)
      Debug.Print "Height;"; lH
      UserControl.Extender.Height = (lH + 6) * Screen.TwipsPerPixelY
   End If
   Set m_fnt = fntThis
   PropertyChanged "Font"
   
End Property
Private Property Get BackColor() As OLE_COLOR
   BackColor = m_oBackColor
End Property
Private Property Let BackColor(ByVal oColor As OLE_COLOR)
   m_oBackColor = oColor
Dim lC As Long
   If (m_hBrBack = 0) Then
      DeleteObject m_hBrBack
   End If
   OleTranslateColor oColor, 0, lC
   m_hBrBack = CreateSolidBrush(lC)
End Property
Private Property Get plDefaultItemHeight() As Long
Dim tR As RECT
Dim lHeight As Long
   DrawText UserControl.hdc, StrPtr("Xg"), -1, tR, DT_CALCRECT
   lHeight = (tR.bottom - tR.Top) + 2
   If (lHeight < m_lIconSizeY) Then
      lHeight = m_lIconSizeY
   End If
   plDefaultItemHeight = lHeight
End Property
Private Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        
        .lfCharSet = fntThis.Charset
    End With
End Sub
Friend Function TranslateAccelerator(lpMsg As Msg) As Long
   TranslateAccelerator = S_FALSE
   ' Here you can modify the response to the key down
   ' accelerator command using the values in lpMsg.  This
   ' can be used to capture Tabs, Returns, Arrows etc.
   ' Just process the message as required and return S_OK.
   If lpMsg.Message = WM_KEYDOWN Or lpMsg.Message = WM_KEYUP Then
   
      Dim bToEdit As Boolean
      Dim iKey As KeyCodeConstants
      Dim iSel As Long, iLen As Long
      Dim iShift As ShiftConstants
            
      iKey = lpMsg.wParam And &HFFFF&
      Select Case iKey
      Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn
         
         TranslateAccelerator = S_OK

         bToEdit = (GetFocus() = m_hWndEdit)
         If m_eStyle = eccxDropDownCombo Then
            If iKey = vbKeyHome Or iKey = vbKeyEnd Or iKey = vbKeyReturn Then
               If ComboIsDropped Then
                  If iKey = vbKeyHome Then
                     Debug.Print "Attempting to parse HOME"
                     iShift = piGetShiftState()
                     If (iShift And vbShiftMask) = vbShiftMask Then
                        iSel = SelStart
                        SelStart = 0
                        If iSel > 0 Then
                           SelLength = iSel + 1
                        End If
                     Else
                        SelStart = 0
                        SelLength = 0
                     End If
                     Exit Function
                  ElseIf iKey = vbKeyEnd Then
                     Debug.Print "Attempting to parse END"
                     iShift = piGetShiftState()
                     If (iShift And vbShiftMask) = vbShiftMask Then
                        iSel = SelStart
                        iLen = Len(Text)
                        If iLen - iSel >= 0 Then
                           SelLength = iLen - iSel
                        End If
                     Else
                        pSetSelStartEnd Len(Text), Len(Text)
                     End If
                     Exit Function
                  Else
                     Debug.Print "Forwarding"
                     bToEdit = True
                  End If
               End If
            End If
         End If
         If bToEdit Then
            SendMessageLong m_hWndEdit, lpMsg.Message, lpMsg.wParam, lpMsg.lParam
         Else
            SendMessageLong m_hWndCbo, lpMsg.Message, lpMsg.wParam, lpMsg.lParam
         End If
      End Select
   End If
End Function

Public Property Get Enabled() As Boolean
   Enabled = m_bEnabled
End Property
Public Property Let Enabled(ByVal bState As Boolean)
   If Not (m_bEnabled = bState) Then
      m_bEnabled = bState
      UserControl.Enabled = m_bEnabled
      EnableWindow UserControl.hwnd, Abs(m_bEnabled)
      If Not (m_hWnd = 0) Then
         EnableWindow m_hWnd, Abs(m_bEnabled)
      End If
      PropertyChanged "Enabled"
   End If
End Property

Public Property Get Style() As ECCXComboStyle
   Style = m_eStyle
End Property
Public Property Let Style(ByVal eStyle As ECCXComboStyle)
   If Not (m_eStyle = eStyle) Then
      m_eStyle = eStyle
      If Not (m_hWnd = 0) Then
         pCache True
         plCreate
         pCache
      End If
      UserControl_Resize
      PropertyChanged "Style"
   End If
End Property
Public Property Get Sorted() As Boolean
   Sorted = m_bSorted
End Property
Public Property Let Sorted(ByVal bState As Boolean)
   If m_bSorted <> bState Then
      m_bSorted = bState
      If Not (m_hWnd = 0) Then
         pCache True
         Clear
         pCache
      End If
      PropertyChanged "Sorted"
   End If
End Property
Private Sub pCache(Optional ByVal bState As Boolean)
Static s_tCBItem() As COMBOBOXEXITEM
Static s_iCount As Long
Static s_iListIndex As Long
Static s_sText As String
Dim i As Long
Dim sTemp As String
Dim sBuf As String

   If bState Then
      ' Cache:
      s_iCount = ListCount
      If s_iCount > 0 Then
         ReDim s_tCBItem(1 To s_iCount) As COMBOBOXEXITEM
         For i = 0 To s_iCount - 1
            With s_tCBItem(i + 1)
               .mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_OVERLAY Or CBEIF_INDENT Or CBEIF_LPARAM
               .iItem = i
               sTemp = String$(260, 0)
               sBuf = StrConv(sTemp, vbFromUnicode)
               .cchTextMax = LenB(sBuf)
               .pszText = StrPtr(sBuf)
            End With
            SendMessageLong m_hWnd, CBEM_GETITEMW, 0, VarPtr(s_tCBItem(i + 1))
         Next i
         s_iListIndex = ListIndex
      Else
         Erase s_tCBItem
      End If
      s_sText = Text
   Else
      ' Uncache:
      Redraw = False
      If s_iCount > 0 Then
         For i = 0 To s_iCount - 1
            With s_tCBItem(i + 1)
               AddItemAndData .pszText, .iImage, .iSelectedImage, .lParam, .iIndent
            End With
         Next i
         ListIndex = s_iListIndex
      End If
      If m_eStyle <> eccxDropDownList Then
         Text = s_sText
      End If
      Redraw = True
   End If
   
End Sub
Public Property Get DropDownWidth() As Long
   ' Get the width of the drop down portion of a combo box
   ' in pixels:
   DropDownWidth = m_lWidth
End Property
Public Property Let DropDownWidth(lWidth As Long)
Dim lR As Long
Dim lAWidth As Long
   ' Set the width of the drop down portion of a combo box
   ' in pixels:
   If Not (m_lWidth = lWidth) Then
      m_lWidth = lWidth
      If Not (m_hWnd = 0) Then
         If lWidth = -1 Then
            lWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
         End If
         ' The width of a combo box's drop down is set
         ' in dialog units which are basically the size
         ' of an average character in the system font:
         'lAWidth = lWidth \ plGetFontDialogUnits(m_hWnd)
         lR = SendMessageLong(m_hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)
      End If
      PropertyChanged "DropDownWidth"
   End If
End Property

Private Function plGetFontDialogUnits(ByVal hwnd As Long) As Long
Dim hFont As Long
Dim hFontOld As Long
Dim r As Long
Dim avgWidth As Long
Dim hdc As Long
Dim tmp As String
Dim sz As SIZEAPI
   
   'get the hdc to the main window
    hdc = GetDC(hwnd)
   
   'with the current font attributes, select the font
    hFont& = GetStockObject(ANSI_VAR_FONT)
    hFontOld& = SelectObject(hdc, hFont&)
   
   'get it's length, then calculate the average character width
    tmp$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    r& = GetTextExtentPoint32(hdc, tmp$, 52, sz)
    avgWidth& = (sz.cX \ 52)
   
   're-select the previous font & delete the hDc
    r& = SelectObject(hdc, hFontOld&)
    r& = DeleteObject(hFont&)
    r& = ReleaseDC(hwnd, hdc)
   
   'return the average character width
    plGetFontDialogUnits = avgWidth

End Function

Public Property Get ComboIsDropped() As Boolean
   ComboIsDropped = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
End Property
Public Sub ShowDropDown(ByVal bState As Boolean)
Dim wP As Long
Dim lR As Long
   ' In a combo box, show or hide the drop down portion:
   If Not (m_eStyle = eccxSimple) Then
      If Not (m_hWnd = 0) Then
         wP = -1 * bState
         lR = SendMessageLong(m_hWnd, CB_SHOWDROPDOWN, wP, 0)
      End If
   Else
      Err.Raise 383, App.EXEName & ".UniComboBox"
   End If
End Sub
Public Property Get ListCount() As Long
   ListCount = SendMessageLong(m_hWnd, CB_GETCOUNT, 0, 0)
End Property
Public Property Get ListIndex() As Long
    ListIndex = SendMessageLong(m_hWnd, CB_GETCURSEL, ByVal 0&, ByVal 0&)
End Property
Public Property Let ListIndex(ByVal lIndex As Long)
On Error GoTo err_index
Dim lR As Long
   lR = SendMessageLong(m_hWnd, CB_SETCURSEL, lIndex, 0)
   If lR = CB_ERR And lIndex <> -1 Then
      Err.Raise 381, App.EXEName & ".UniComboBox"
   Else
      RaiseEvent Click
   End If
err_index:
End Property
Public Property Get NewIndex() As Long
   NewIndex = m_lNewIndex
End Property
Private Sub pGetSelStartEnd(lStart As Long, lEnd As Long)
Dim lParam As Long
   ' Get the start and end of the selection in the edit
   ' box portion of a drop down combo box:
   If Not (m_hWnd = 0) Then
      lParam = SendMessageByref(m_hWndEdit, EM_GETSEL, lStart, lEnd)
   End If
End Sub
Private Sub pSetSelStartEnd(ByVal lStart As Long, ByVal lEnd As Long)
Dim lParam As Long
Dim lR As Long
   ' Set the start and end of the selection in the edit
   ' box portion of a drop down combo box:
   If Not (m_hWnd = 0) Then
      lStart = lStart And &H7FFF&
      lEnd = lEnd And &H7FFF&
      lR = SendMessageLong(m_hWndEdit, EM_SETSEL, lStart, lEnd)
      Debug.Print lEnd, lStart
   End If
End Sub

Property Get SelLength() As Long
Dim lStart As Long, lEnd As Long
   ' Return the length of the selected text in the edit
   ' box portion of a dropdown combo:
   If (m_eStyle = eccxDropDownList) Then
      If ListIndex > -1 Then
         SelLength = Len(List(ListIndex))
      Else
         SelLength = 0
      End If
   Else
      pGetSelStartEnd lStart, lEnd
      SelLength = lEnd - lStart
   End If
End Property
Property Let SelLength(ByVal lLength As Long)
Dim lStart As Long, lEnd As Long
   ' Set the length of the selected text in the edit
   ' box portion of a dropdown combo:
   If (m_eStyle <> eccxDropDownList) Then
      pGetSelStartEnd lStart, lEnd
      If (lEnd - lStart <> lLength) Then
         pSetSelStartEnd lStart, lStart + lLength
      End If
   Else
      Err.Raise 383, "UniComboBox." & App.EXEName
   End If
End Property
Property Get SelStart() As Long
Dim lStart As Long, lEnd As Long
   ' Return the start of the selected text in the edit
   ' box portion of a dropdown combo:
   If (m_eStyle <> eccxDropDownList) Then
      pGetSelStartEnd lStart, lEnd
      SelStart = lStart
   Else
      'Err.Raise 383, "UniComboBox." & App.EXEName
   End If
End Property
Property Let SelStart(ByVal lStart As Long)
Dim lOStart As Long, lEnd As Long
   ' Set the start of the selected text in the edit
   ' box portion of a dropdown combo:
   If (m_eStyle <> eccxDropDownList) Then
      pGetSelStartEnd lOStart, lEnd
      If (lStart <> lOStart) Then
         pSetSelStartEnd lStart, lEnd
      End If
   Else
      Err.Raise 383, "UniComboBox." & App.EXEName
   End If
End Property

Property Get SelText() As String
   ' Return the selected text from the edit
   ' box portion of a dropdown combo:
   If (m_eStyle = eccxDropDownList) Then
      Dim sText As String
      Dim lStart As Long, lEnd As Long
      
      pGetSelStartEnd lStart, lEnd
      sText = Text
      If (lEnd > 0) And Len(sText) > 0 Then
         If (lStart <= 0) Then
            lStart = 1
         End If
         lEnd = lEnd + 1
         If (lEnd > Len(sText)) Then lEnd = Len(sText)
         SelText = Mid$(sText, lStart, (lEnd - lStart))
      End If
   Else
      SelText = Text
   End If
End Property

'Property Let Text(ByVal sText As String)
'   ' Can only set the text in a drop down combo box:
'   If Not (m_eStyle = eccxDropDownList) Then
''      SendMessageLong m_hWndEdit, WM_SETTEXT, 0, StrPtr(sText)
'
'      SendMessageString m_hWnd, WM_SETTEXT, 0, sText & Chr$(0)
''      List(-1) = sText
'   Else
'      Err.Raise 383, "UniComboBox." & App.EXEName
'   End If
'End Property

Public Property Let Text(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : If there is an edit box, set the text.  Otherwise, search the list for an
'             item that matches sNew.  If found, set the listindex.  If not, raise an error.
'---------------------------------------------------------------------------------------
Dim lsAnsi As String
Dim liIndex As Long
    
    lsAnsi = StrConv(sNew & vbNullChar, vbFromUnicode)
    If m_hWnd Then
        If m_hWndEdit Then
            Call pvSetWindowText(sNew)
        Else
            liIndex = SendMessageLong(m_hWnd, CB_FINDSTRINGEXACT, 0, StrPtr(lsAnsi))
            If liIndex > -1 Then
                SendMessageLong m_hWnd, CB_SETCURSEL, liIndex, 0
            End If
        End If
    End If

End Property

'Property Get Text() As String
'Dim lR As Long
'Dim sText As String
'Dim iPos As Long
'   ' Returns either the text in the EditBox portion of a
'   ' drop down combo or the text of the (first) selected
'   ' list item:
'   If Not (m_hWnd = 0) Then
'      If Not (m_eStyle = eccxDropDownList) Then
''         Text = List(-1) ' --> ' This works correctly in IE4+ only
''         lR = SendMessageLong(m_hWndEdit, WM_GETTEXTLENGTH, 0, 0)
''         If (lR > 0) Then
''            sText = String$(lR + 1, Chr$(0))
'''            lR = SendMessageLong(m_hWndEdit, WM_GETTEXT, (lR + 1), StrPtr(sText))
''            lR = SendMessageString(m_hWndEdit, WM_GETTEXT, (lR + 1), sText)
'''            lR = SendMessage(m_hWndEdit, WM_GETTEXT, (lR + 1), sText)
''            If (lR > 0) Then
''               iPos = InStr(sText, vbNullChar)
''               If iPos <> 0 Then
''                  lR = iPos - 1
''               End If
''               Text = Left$(sText, lR)
''            End If
''         End If
'        'DetachEdit
'
'        Text = TextBox_GetText(m_hWndEdit)
'
'        'AttachEdit
'      Else
'         If (ListIndex > -1) Then
'            Text = List(ListIndex)
'         Else
'            Text = ""
'         End If
'      End If
'   End If
'End Property

Public Property Get Text() As String
    If m_hWndEdit Then
        Text = pvGetWindowText(m_hWndEdit)
    ElseIf m_hWnd Then
        Text = pItem_Text(SendMessage(m_hWnd, CB_GETCURSEL, 0, 0))
    End If
End Property

Private Property Get pItem_Text(ByVal iIndex As Long) As String
Dim a(260) As Byte
Dim lLen   As Long
Dim m_titem As COMBOBOXEXITEM
    If m_hWnd Then
        With m_titem
            .mask = CBEIF_TEXT
            .pszText = VarPtr(a(0))
            .cchTextMax = UBound(a)
            .iItem = iIndex
        End With
        lLen = SendMessage(m_hWnd, CBEM_GETITEMW, 0, m_titem)

        If lLen > 0 Then
'#If Unicode Then
            pItem_Text = a
'#Else
'            pItem_Text = Left$(StrConv(a(), vbUnicode), lLen)
'#End If
        Else
            pItem_Text = ""
        End If
    End If
    pItem_Text = pvStripNulls(pItem_Text)
'    pItem_Text = Left$(pItem_Text, InStr(pItem_Text, Chr$(0)) - 1)
End Property

Private Property Let pItem_Text(ByVal iIndex As Long, ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the text of a combo list item.
'---------------------------------------------------------------------------------------
Dim m_titem As COMBOBOXEXITEM
    If m_hWnd Then
        With m_titem
            .mask = CBEIF_TEXT
            .iItem = iIndex
'#If Unicode Then
            .pszText = StrPtr(sText)
'#Else
'            .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
'#End If
            .cchTextMax = Len(sText)
        End With
        SendMessage m_hWnd, CBEM_SETITEMW, 0, m_titem
    End If
End Property

Public Property Get MaxLength() As Long
   ' Same as MaxLength property of a Text control.  Only
   ' valid for drop down combo boxes:
   If Not (m_eStyle = eccxDropDownList) Then
      MaxLength = m_lMaxLength
   Else
      'Err.Raise 383, "UniComboBox." & App.EXEName
   End If
End Property
Public Property Let MaxLength(ByVal lLength As Long)
   ' Same as MaxLength property of a Text control.  Only
   ' valid for drop down combo boxes:
   If Not (m_eStyle = eccxDropDownCombo) Then
      ' Don't be silly:
      If (lLength > 30000&) Or (lLength <= 0) Then lLength = 30000&
      ' Set:
      m_lMaxLength = lLength
      SendMessageLong m_hWnd, CB_LIMITTEXT, lLength, 0
   Else
'      Err.Raise 383, "UniComboBox." & App.EXEName
   End If
End Property
Public Property Get Redraw() As Boolean
   Redraw = m_bRedraw
End Property
Public Property Let Redraw(ByVal bState As Boolean)
   If m_bRedraw <> bState Then
      m_bRedraw = bState
      If Not (m_hWnd = 0) Then
         SendMessageLong m_hWnd, WM_SETREDRAW, Abs(m_bRedraw), 0
      End If
      PropertyChanged "Redraw"
   End If
End Property
Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property
Public Property Get hWndComboEx() As Long
   hWndComboEx = m_hWnd
End Property
Public Property Get hWndCombo() As Long
   hWndCombo = m_hWndCbo
End Property
Public Property Get hWndEdit() As Long
   hWndEdit = m_hWndEdit
End Property
Public Sub Clear()
   SendMessageLong m_hWnd, CB_RESETCONTENT, 0, 0
   m_lNewIndex = -1
End Sub
Public Property Get ExtendedUI() As Boolean
   ExtendedUI = m_bExtendedUI
End Property
Public Property Let ExtendedUI(ByVal bState As Boolean)
   If m_bExtendedUI <> bState Then
      m_bExtendedUI = bState
      If Not (m_hWnd = 0) Then
         SendMessageLong m_hWnd, CB_SETEXTENDEDUI, Abs(bState), 0
      End If
      PropertyChanged "ExtendedUI"
   End If
End Property
Public Sub AddItem(ByVal sText As String)
   InsertItemAndData IIf(bUnicode, zToUnicode(sText), sText)
End Sub
Public Sub AddItemAndData(ByVal sText As String, Optional ByVal iIcon As Long = -1, _
                        Optional ByVal iIconSelected As Long = -1, Optional ByVal lItemData As Long = 0, _
                        Optional ByVal lIndent As Long = 0)
   InsertItemAndData IIf(bUnicode, zToUnicode(sText), sText), , iIcon, iIconSelected, lItemData, lIndent
End Sub
Public Sub InsertItem(ByVal sText As String, Optional ByVal lIndexBefore As Long = -1)
   InsertItemAndData sText, lIndexBefore
End Sub
Public Function InsertItemAndData(ByVal sText As String, Optional ByVal lIndexBefore As Long = -1, _
                            Optional ByVal iIcon As Long = -1, Optional ByVal iIconSelected As Long = -1, _
                            Optional ByVal lItemData As Long = 0, Optional ByVal lIndent As Long = 0) As Boolean
'Dim tCBItem As COMBOBOXEXITEM
Dim lR As Long
Dim i As Long
Dim iStart As Long
Dim iEnd As Long
Dim iComp As Long
Dim iRes As Long
Dim eCompare As VbCompareMethod
Static s_sLastText As String

   If m_bSorted Then
      ' We force the index to the appropriate point.
      ' Use a binary search...
      If ListCount > 0 Then
         If ExtendedStyle(eccxCaseSensitiveSearch) Then
            eCompare = vbBinaryCompare
         Else
            eCompare = vbTextCompare
         End If
         lIndexBefore = -1
         iEnd = ListCount - 1
         If iEnd > 0 Then
            Do While iEnd > iStart
               iComp = iStart + (iEnd - iStart) \ 2
               iRes = StrComp(sText, List(iComp), eCompare)
               If iRes = 0 Then
                  lIndexBefore = iComp
                  iStart = 0: iEnd = 0
               ElseIf iRes > 0 Then
                  iStart = iComp + 1
               Else
                  iEnd = iComp - 1
               End If
            Loop
         End If
            
         If lIndexBefore = -1 Then
            If iStart = iEnd Then
               If StrComp(sText, List(iEnd), eCompare) < 0 Then
                  lIndexBefore = iEnd
               Else
                  lIndexBefore = iEnd + 1
               End If
            Else
               If iEnd < iStart Then
                  If StrComp(sText, List(iStart), eCompare) < 0 Then
                     lIndexBefore = iStart
                  Else
                     Debug.Assert False
                  End If
               Else
                  Debug.Assert False
               End If
            End If
            If lIndexBefore >= ListCount Then
               lIndexBefore = -1
            End If
         End If
         
      End If
   End If

   With tCBItem
      .mask = CBEIF_TEXT Or CBEIF_INDENT _
               Or CBEIF_IMAGE Or CBEIF_LPARAM Or _
               CBEIF_SELECTEDIMAGE
      .pszText = StrPtr(sText)
      .cchTextMax = Len(sText)
      .iIndent = lIndent
      .iImage = iIcon
      .iSelectedImage = iIconSelected
      .lParam = lItemData
      .iItem = lIndexBefore
      .iOverlay = -1
   End With
   
    If lIndexBefore < -1 Then lIndexBefore = -1
'   m_lNewIndex = SendMessageW(m_hWnd, CBEM_INSERTITEMW, 0&, VarPtr(tCBItem))
    m_lNewIndex = SendMessage(m_hWnd, CBEM_INSERTITEMW, lIndexBefore, tCBItem)
    InsertItemAndData = CBool(m_lNewIndex > -1)
    If m_lNewIndex > -1 Then s_sLastText = sText
End Function
Public Function FindItemIndex(ByVal sToFind As String, Optional ByVal bExactMatch As Boolean = False) As Long
Dim lR As Long
Dim lFlag As Long
   ' Find the index of the item sToFind, optionally
   ' exact matching.  Return -1 if the item is not
   ' found.
   If Not (m_hWnd = 0) Then
      ' Set the message to send to the control:
      If (bExactMatch) Then
         lFlag = CB_FINDSTRINGEXACT
      Else
         lFlag = CB_FINDSTRING
      End If
      ' Find:
      lR = -1
'      lR = SendMessageString(m_hWnd, lFlag, 0, sToFind)
      lR = SendMessageLong(m_hWnd, lFlag, 0, StrPtr(sToFind))
      ' Return value:
      FindItemIndex = lR
   End If
End Function

'Public Property Get ItemIndent(ByVal lIndex As Long) As Long
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_INDENT
'   If GetItem(lIndex, tCBItem) Then
'      ItemIndent = tCBItem.iIndent
'   End If
'End Property
'
'Public Property Let ItemIndent(ByVal lIndex As Long, ByVal lIndent As Long)
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_INDENT
'   tCBItem.iIndent = lIndent
'   SetItem lIndex, tCBItem
'End Property

Public Property Get ItemIndent(ByVal index As Long) As Long
    ItemIndent = pItem_Info(index, CBEIF_INDENT)
End Property

Public Property Let ItemIndent(ByVal index As Long, ByVal lIndent As Long)
    pItem_Info(index, CBEIF_INDENT) = lIndent
End Property

'Public Property Get ItemIcon(ByVal lIndex As Long) As Long
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_IMAGE
'   If GetItem(lIndex, tCBItem) Then
'      ItemIcon = tCBItem.iImage
'   End If
'End Property

'Public Property Let ItemIcon(ByVal lIndex As Long, ByVal lIcon As Long)
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_IMAGE
'   tCBItem.iImage = lIcon
'   SetItem lIndex, tCBItem
'End Property

Public Property Get ItemIcon(ByVal index As Long) As Long
    ItemIcon = pItem_Info(index, CBEIF_IMAGE)
End Property

Public Property Let ItemIcon(ByVal index As Long, ByVal lIcon As Long)
    pItem_Info(index, CBEIF_IMAGE) = lIcon
End Property

'Public Property Get ItemData(ByVal lIndex As Long) As Long
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_LPARAM
'   If GetItem(lIndex, tCBItem) Then
'      ItemData = tCBItem.lParam
'   End If
'End Property

'Public Property Let ItemData(ByVal lIndex As Long, ByVal lData As Long)
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_LPARAM
'   tCBItem.lParam = lData
'   SetItem lIndex, tCBItem
'End Property
Public Property Get ItemData(ByVal index As Long) As Long
    ItemData = pItem_Info(index, CBEIF_LPARAM)
End Property

Public Property Let ItemData(ByVal index As Long, ByVal lData As Long)
    pItem_Info(index, CBEIF_LPARAM) = lData
End Property

'Public Property Get ItemIconSelected(ByVal lIndex As Long) As Long
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_SELECTEDIMAGE
'   If GetItem(lIndex, tCBItem) Then
'      ItemIconSelected = tCBItem.iSelectedImage
'   End If
'End Property
'
'Public Property Let ItemIconSelected(ByVal lIndex As Long, ByVal lIcon As Long)
'Dim tCBItem As COMBOBOXEXITEM
'   tCBItem.mask = CBEIF_SELECTEDIMAGE
'   tCBItem.iSelectedImage = lIcon
'   SetItem lIndex, tCBItem
'End Property

Public Property Get ItemIconSelected(ByVal index As Long) As Long
    ItemIconSelected = pItem_Info(index, CBEIF_SELECTEDIMAGE)
End Property

Public Property Let ItemIconSelected(ByVal index As Long, ByVal lIcon As Long)
    pItem_Info(index, CBEIF_SELECTEDIMAGE) = lIcon
End Property

Private Property Get pItem_Info(ByVal iIndex As Long, ByVal iMask As Long) As Long
'--------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get a 32 bit value in the COMBOBOXEXITEM structure.
'--------------------------------------------------------------------------------------
    If m_hWnd Then
        With tCBItem
            .mask = iMask
            .iItem = iIndex
            If SendMessage(m_hWnd, CBEM_GETITEMW, 0, tCBItem) Then
                If iMask = CBEIF_LPARAM Then
                    pItem_Info = .lParam
                ElseIf iMask = CBEIF_IMAGE Then
                    pItem_Info = .iImage
                ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                    pItem_Info = .iSelectedImage
                ElseIf iMask = CBEIF_INDENT Then
                    pItem_Info = .iIndent
                End If
            End If
        End With
    End If
End Property

Private Property Let pItem_Info(ByVal iIndex As Long, ByVal iMask As Long, ByVal iNew As Long)
'--------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set a 32 bit value in the COMBOBOXEXITEM structure.
'--------------------------------------------------------------------------------------
    If m_hWnd Then
        With tCBItem
            .mask = iMask
            .iItem = iIndex
            If iMask = CBEIF_LPARAM Then
                .lParam = iNew
            ElseIf iMask = CBEIF_IMAGE Then
                .iImage = iNew
            ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                .iSelectedImage = iNew
            ElseIf iMask = CBEIF_INDENT Then
                .iIndent = iNew
            End If
            SendMessage m_hWnd, CBEM_SETITEMW, 0, tCBItem
        End With
    End If
End Property

'Public Property Get List(ByVal lIndex As Long) As String
'Dim tCBItem As COMBOBOXEXITEM
'Dim sBuf As String
'Dim sTemp As String
'Dim iPos As Long
'   If lIndex = -1 Then
'      If m_eStyle = eccxDropDownCombo Then
'         List = Text
'      Else
'         List = ""
'      End If
'   Else
'      tCBItem.mask = CBEIF_TEXT
'      sTemp = String$(260, 0)
'      sBuf = StrConv(sTemp, vbFromUnicode)
'      tCBItem.cchTextMax = LenB(sBuf)
'      tCBItem.pszText = StrPtr(sBuf)
'      If GetItem(lIndex, tCBItem) Then
'         sTemp = sBuf   'StrPtr(tCBItem.pszText)
'         iPos = InStr(sTemp, vbNullChar)
'         If (iPos > 1) Then
'            List = Left$(sTemp, (iPos - 1))
'         Else
'            List = sTemp
'         End If
'      End If
'   End If

'End Property
'Public Property Let List(ByVal lIndex As Long, ByVal sItem As String)
'Dim tCBItem As COMBOBOXEXITEM
'
'    If bUnicode Then sItem = zToUnicode(sItem)
'    tCBItem.mask = CBEIF_TEXT
'    tCBItem.cchTextMax = LenB(sItem)
'    tCBItem.pszText = StrPtr(sItem)
'    SetItem lIndex, tCBItem
'End Property

Public Property Get List(ByVal iIndex As Long) As String
    List = pItem_Text(iIndex)
End Property

Public Property Let List(ByVal iIndex As Long, ByRef sNew As String)
    pItem_Text(iIndex) = sNew
End Property

'Private Function GetItem(ByVal lIndex As Long, ByRef tCBItem As COMBOBOXEXITEM) As Boolean
'Dim lR As Long
'   If InRange(lIndex) Then
'      tCBItem.iItem = lIndex
'      lR = SendMessageLong(m_hWnd, CBEM_GETITEMW, 0, VarPtr(tCBItem))
'   End If
'   If (lR = 0) And (lIndex <> -1) Then
'      Err.Raise 381, App.EXEName & ".UniComboBox"
'   Else
'      GetItem = True
'   End If
'End Function
'Private Function SetItem(ByVal lIndex As Long, ByRef tCBItem As COMBOBOXEXITEM) As Boolean
'On Error Resume Next
'Dim lR As Long
'   If InRange(lIndex) Then
'      tCBItem.iItem = lIndex
'   End If
'   lR = SendMessageLong(m_hWnd, CBEM_SETITEMW, 0, VarPtr(tCBItem))
'   If (lR = 0) Then
'      Err.Raise 381, App.EXEName & ".UniComboBox"
'   Else
'      SetItem = True
'   End If
'End Function
Private Property Get InRange(ByVal lIndex As Long) As Boolean
   InRange = (lIndex >= 0) And (lIndex < ListCount)
End Property
Public Sub RemoveItem(ByVal lIndex As Long)
   If SendMessageLong(m_hWnd, CBEM_DELETEITEM, lIndex, 0) = CB_ERR Then
      Err.Raise 381, App.EXEName & ".UniComboBox"
   End If
End Sub
Public Property Let ImageList(ByRef vThis As Variant)
Dim hIml As Long
Dim lX As Long

   ' Set the ImageList handle property either from a VB
   ' image list or directly:
   If VarType(vThis) = vbObject Then
       ' Assume VB ImageList control.  Note that unless
       ' some call has been made to an object within a
       ' VB ImageList the image list itself is not
       ' created.  Therefore hImageList returns error. So
       ' ensure that the ImageList has been initialised by
       ' drawing into nowhere:
       On Error Resume Next
       ' Get the image list initialised..
       vThis.ListImages(1).Draw 0, 0, 0, 1
       hIml = vThis.hImageList
       If (Err.Number <> 0) Then
         Err.Clear
         hIml = vThis.hIml
         If (Err.Number <> 0) Then
            hIml = 0
         End If
       End If
       On Error GoTo 0
   ElseIf VarType(vThis) = vbLong Then
       ' Assume ImageList handle:
       hIml = vThis
   Else
       Err.Raise vbObjectError + 1049, "vbalDriveCboEx." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
   End If
    
   ' If we have a valid image list, then associate it with the control:
   If (hIml <> 0) Then
      m_hIml = hIml
      ImageList_GetIconSize m_hIml, lX, m_lIconSizeY
      'Set the Imagelist for the ComboBox
      SendMessageLong m_hWnd, CBEM_SETIMAGELIST, 0, m_hIml
      Set Font = m_fnt
   End If
End Property

Private Function plCreate()
    Dim dwStyle As Long
    Dim lWidth As Long
    Dim lHeight As Long
   
        pDestroy
        
        dwStyle = WS_CHILD Or CBS_AUTOHSCROLL
        Select Case m_eStyle
            Case eccxSimple
                dwStyle = dwStyle Or CBS_SIMPLE
            Case eccxDropDownList
                dwStyle = dwStyle Or CBS_DROPDOWNLIST
            Case eccxDropDownCombo
                dwStyle = dwStyle Or CBS_DROPDOWN
            Case Else
                Debug.Assert False
                dwStyle = dwStyle Or CBS_DROPDOWN
        End Select
        lWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelX
        lHeight = (UserControl.ScaleHeight \ Screen.TwipsPerPixelX) * 8
        m_hWndParent = UserControl.hwnd
        m_hWnd = CreateWindowEX(0, WC_COMBOBOXEX, "", dwStyle, 0, 0, lWidth, lHeight, m_hWndParent, 0&, App.hInstance, 0&)
        
        If UserControl.Ambient.UserMode Then
            If m_hWnd Then
                SendMessageLongA m_hWnd, CCM_SETUNICODEFORMAT, 1&, 0&
                SendMessageLongA m_hWnd, CB_SETDROPPEDWIDTH, 0, 0
                SendMessageLongA m_hWnd, CB_LIMITTEXT, m_lMaxLength, 0
                SendMessageLongA m_hWnd, CB_SETEXTENDEDUI, -m_bExtendedUI, 0
                
                Call AddMsg 'start subclass
      
            'Set the Imagelist for the ComboBox
                If m_hIml <> 0 Then SendMessageLong m_hWnd, CBEM_SETIMAGELIST, 0, m_hIml
                
                SendMessageLong m_hWnd, CBEM_SETEXSTYLE, 0, m_eExStyle
   
            'stop subclass text edit
                If m_eStyle = eccxDropDownCombo Then Call DetachEdit
            End If
        End If
   
        SetParent m_hWnd, m_hWndParent
        MoveWindow m_hWnd, 0, 0, lWidth, lHeight, 1
        ShowWindow m_hWnd, SW_SHOWNORMAL
        EnableWindow m_hWnd, Abs(m_bEnabled)
'        SendMessageLong m_hWnd, WM_SETREDRAW, Abs(m_bRedraw), 0
        SendMessageLong m_hWnd, WM_SETREDRAW, -m_bRedraw, 0
End Function

Private Sub AttachEdit()
On Error GoTo Attach_Err
    If m_hWndEdit <> 0 Then
'        If m_eStyle = eccxDropDownCombo Then
            AttachMessage Me, m_hWndEdit, WM_SETFOCUS
            AttachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
            AttachMessage Me, m_hWndEdit, WM_KEYDOWN
            AttachMessage Me, m_hWndEdit, WM_CHAR
            AttachMessage Me, m_hWndEdit, WM_KEYUP
'        ElseIf m_eStyle = eccxSimple Then
'            AttachMessage Me, m_hWndEdit, WM_SETFOCUS
'            AttachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
'            AttachMessage Me, m_hWndEdit, WM_KEYDOWN
'            AttachMessage Me, m_hWndEdit, WM_CHAR
'            AttachMessage Me, m_hWndEdit, WM_KEYUP
'        End If
    End If
Attach_Err:
End Sub

Private Sub DetachEdit()
On Error GoTo Detach_Err
    If m_hWndEdit <> 0 Then
'        If m_eStyle = eccxDropDownCombo Then
            DetachMessage Me, m_hWndEdit, WM_SETFOCUS
            DetachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
            DetachMessage Me, m_hWndEdit, WM_KEYDOWN
            DetachMessage Me, m_hWndEdit, WM_CHAR
            DetachMessage Me, m_hWndEdit, WM_KEYUP
'        ElseIf m_eStyle = eccxSimple Then
'            DetachMessage Me, m_hWndEdit, WM_SETFOCUS
'            DetachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
'            DetachMessage Me, m_hWndEdit, WM_KEYDOWN
'            DetachMessage Me, m_hWndEdit, WM_CHAR
'            DetachMessage Me, m_hWndEdit, WM_KEYUP
'        End If
    End If
Detach_Err:
End Sub

Private Sub AddMsg()
    If m_bSubclass = False Then
        If m_hWnd Then
            AttachMessage Me, m_hWndParent, WM_COMMAND
            AttachMessage Me, m_hWndParent, WM_SETFOCUS
            AttachMessage Me, m_hWndParent, WM_NOTIFY
            AttachMessage Me, m_hWnd, WM_CTLCOLORLISTBOX
            AttachMessage Me, m_hWnd, WM_DRAWITEM
            
            m_hWndCbo = SendMessageLong(m_hWnd, CBEM_GETCOMBOCONTROL, 0, 0)
            AttachMessage Me, m_hWndCbo, WM_SETFOCUS
            AttachMessage Me, m_hWndCbo, WM_MOUSEACTIVATE
        
            Select Case m_eStyle
                Case eccxDropDownCombo:
                    m_hWndEdit = SendMessageLong(m_hWnd, CBEM_GETEDITCONTROL, 0, 0)
                    AttachMessage Me, m_hWndEdit, WM_SETFOCUS
                    AttachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
                    AttachMessage Me, m_hWndEdit, WM_KEYDOWN
                    AttachMessage Me, m_hWndEdit, WM_CHAR
                    AttachMessage Me, m_hWndEdit, WM_KEYUP
                    AttachMessage Me, m_hWndCbo, WM_KEYDOWN
                    AttachMessage Me, m_hWndCbo, WM_CHAR
                    AttachMessage Me, m_hWndCbo, WM_KEYUP
                    AttachMessage Me, m_hWndCbo, WM_CTLCOLOREDIT
                 
                Case eccxSimple:
                 ' **** PROBLEM **** - can't get hWnd...
'                    m_hWndEdit = FindWindowEx(m_hWnd, ByVal 0&, "EDIT", vbNullString)
                    m_hWndEdit = FindWindowEx(FindWindowEx(m_hWnd, ByVal 0&, "ComboBox", vbNullString), ByVal 0&, "Edit", vbNullString)
                    
                    If m_hWndEdit <> 0 Then
                        AttachMessage Me, m_hWndEdit, WM_SETFOCUS
                        AttachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
                        AttachMessage Me, m_hWndEdit, WM_KEYDOWN
                        AttachMessage Me, m_hWndEdit, WM_CHAR
                        AttachMessage Me, m_hWndEdit, WM_KEYUP
                        AttachMessage Me, m_hWndCbo, WM_CTLCOLOREDIT
                   End If
                    
                Case Else
                    AttachMessage Me, m_hWndCbo, WM_KEYDOWN
                    AttachMessage Me, m_hWndCbo, WM_CHAR
                    AttachMessage Me, m_hWndCbo, WM_KEYUP
                    
            End Select
            m_bSubclass = True
        End If
    End If
End Sub

Private Sub DelMsg()
   If m_bSubclass Then
      DetachMessage Me, m_hWndParent, WM_COMMAND
      DetachMessage Me, m_hWndParent, WM_SETFOCUS
      DetachMessage Me, m_hWndParent, WM_NOTIFY
      
      If m_hWnd <> 0 Then
         DetachMessage Me, m_hWnd, WM_DRAWITEM
         DetachMessage Me, m_hWnd, WM_CTLCOLORLISTBOX
         If m_hWndCbo <> 0 Then
            DetachMessage Me, m_hWndCbo, WM_SETFOCUS
            DetachMessage Me, m_hWndCbo, WM_MOUSEACTIVATE
            DetachMessage Me, m_hWndCbo, WM_KEYDOWN
            DetachMessage Me, m_hWndCbo, WM_CHAR
            DetachMessage Me, m_hWndCbo, WM_KEYUP
            
            If m_eStyle = eccxDropDownCombo Or m_eStyle = eccxSimple Then
               If m_hWndEdit <> 0 Then
                  DetachMessage Me, m_hWndEdit, WM_SETFOCUS
                  DetachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
                  DetachMessage Me, m_hWndEdit, WM_KEYDOWN
                  DetachMessage Me, m_hWndEdit, WM_CHAR
                  DetachMessage Me, m_hWndEdit, WM_KEYUP
                  DetachMessage Me, m_hWndCbo, WM_CTLCOLOREDIT
               End If
            Else
               DetachMessage Me, m_hWndCbo, WM_KEYDOWN
               DetachMessage Me, m_hWndCbo, WM_CHAR
               DetachMessage Me, m_hWndCbo, WM_KEYUP
            End If
         End If
      End If
      m_bSubclass = False
   End If
End Sub

Private Function pDestroy()
   Call DelMsg  'stop subclass
   
   If Not (m_hWnd = 0) Then
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      m_hWnd = 0
      m_hWndEdit = 0
   End If
   
   m_hWndParent = 0
   
   If Not (m_hFnt = 0) Then
      DeleteObject m_hFnt
   End If
      
   If Not (m_hBrBack = 0) Then
      DeleteObject m_hBrBack
      m_hBrBack = 0
   End If
End Function

Private Function piGetShiftState() As ShiftConstants
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
    iR = iR Or (-1 * pbKeyIsPressed(vbKeyShift))
    iR = iR Or (-2 * pbKeyIsPressed(vbKeyMenu))
    iR = iR Or (-4 * pbKeyIsPressed(vbKeyControl))
    piGetShiftState = iR
End Function
Private Function pbKeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        pbKeyIsPressed = True
    End If
End Function

Private Function plDrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDis As DRAWITEMSTRUCT
Dim bEnabled As Boolean
Dim bSelected As Boolean
Dim tLF As LOGFONT
Dim hMem As Long

   ' Debug.Print "OwnerDraw.."
   
   CopyMemory tDis, ByVal lParam, Len(tDis)
   
   ' Evaluate enabled/selected state of item:
   bEnabled = Not ((tDis.ItemState And ODS_DISABLED) = ODS_DISABLED)
   bSelected = ((tDis.ItemState And ODS_SELECTED) = ODS_SELECTED)
   If (bSelected) Then
       ' Only draw selected in the combo when the
       ' focus is on the control:
       If (tDis.ItemState And ODS_COMBOBOXEDIT) = ODS_COMBOBOXEDIT Then
           If (tDis.ItemState And ODS_FOCUS) <> ODS_FOCUS Then
               bSelected = False
           End If
       End If
   End If

   ' Ensure we have the correct font and colours selected:
   If (m_hFnt = 0) Then
       pOLEFontToLogFont UserControl.Font, UserControl.hdc, tLF
       m_hFnt = CreateFontIndirect(m_tLF)
   End If
   ' Get the item data for this item:
   
   If (tDis.ItemState And ODS_COMBOBOXEDIT) = ODS_COMBOBOXEDIT Then
       If Not (pbIsCurrentFont(m_tULF)) Then
           DeleteObject m_hFnt
           LSet m_tLF = m_tULF
           m_hFnt = CreateFontIndirect(m_tLF)
       End If
   End If
   m_hFntOld = SelectObject(tDis.hdc, m_hFnt)

   If m_eClientDraw <> eccxOwnerDraw Then
       ' Draw by default mechanism:
       pDefaultDrawItem tDis.hdc, tDis.ItemId, tDis.ItemAction, tDis.ItemState, _
           tDis.rcItem.left, tDis.rcItem.Top, tDis.rcItem.right, tDis.rcItem.bottom
   End If
   If m_eClientDraw <> eccxDrawDefault And m_eClientDraw <> eccxDrawODCboList Then
       ' Notify the client its time to draw:
       RaiseEvent DrawItem(tDis.ItemId, tDis.hdc, _
                           bSelected, bEnabled, _
                           tDis.rcItem.left, tDis.rcItem.Top, tDis.rcItem.right, tDis.rcItem.bottom, _
                           m_hFntOld)
   End If
   
   SelectObject tDis.hdc, m_hFntOld
   
   plDrawItem = 1
End Function
Private Function pbIsCurrentFont(tLF As LOGFONT) As Boolean
Dim sCurrentFace As String
Dim sItemFace As String
    If (tLF.lfFaceName(0) = 0) Then
        ' Default
        pbIsCurrentFont = True
    Else
        If (tLF.lfWeight = m_tLF.lfWeight) And (tLF.lfItalic = m_tLF.lfItalic) And (tLF.lfHeight = m_tLF.lfHeight) Then
            sCurrentFace = StrConv(tLF.lfFaceName, vbUnicode)
            sItemFace = StrConv(m_tLF.lfFaceName, vbUnicode)
            If (sCurrentFace = sItemFace) Then
                pbIsCurrentFont = True
            End If
        End If
    End If
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WM_MOUSEACTIVATE, WM_CHAR ', WM_KEYDOWN  ', WM_DRAWITEM
            ISubclass_MsgResponse = emrConsume
        Case Else
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNMH As NMHDR
Dim tNMHE As NMCBEENDEDIT
Dim tNMHEW As NMCBEENDEDITW
Dim tR As RECT
Dim tDis As DRAWITEMSTRUCT
Dim bCancel As Boolean
Dim sMsg As String
Dim iPos As Long
Dim iKeyCode As Integer
Dim sText As String
Static bAttach As Boolean

   Select Case iMsg
   Case WM_DRAWITEM
      If bAttach = True And m_eStyle = eccxDropDownCombo Then
         Call DetachEdit
         bAttach = False
      End If
      If m_eClientDraw = eccxDrawDefault Or m_eClientDraw = eccxDriveList Then
'      If m_eClientDraw = eccxDriveList Then
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         Exit Function
      End If
      ISubclass_WindowProc = plDrawItem(wParam, lParam)

   Case WM_NOTIFY
      CopyMemory tNMH, ByVal lParam, Len(tNMH)
      If tNMH.hwndFrom = m_hWnd Then
         Select Case tNMH.code
         Case CBEN_BEGINEDIT
'            If m_eStyle = eccxSimple Then
'                UserControl.Parent.Visible = False
'                Call DetachEdit
'                RaiseEvent BeginEdit(ListIndex)
'                Call AttachEdit
'                UserControl.Parent.Visible = True
'            Else
                RaiseEvent BeginEdit(ListIndex)
'                Dim sTam As String
'                If m_eStyle = eccxSimple And ListIndex > -1 Then
'                   Call DetachEdit
'                    sTam = List(SendMessageLong(m_hWnd, CB_GETCURSEL, ByVal 0&, ByVal 0&))
'                    If Trim$(sTam) <> "" Then SetWindowTextW m_hWndEdit, StrPtr(sTam)
'                    ISubclass_WindowProc = 0
'                    Call AttachEdit
'                End If

'            End If
         Case CBEN_DELETEITEM
            Debug.Print "CBEN_DELETEITEM"
            ' ... no need to intercept
         Case CBEN_INSERTITEM
            Debug.Print "CBEN_INSERTITEM"
            ' ... no need to intercept
         Case CBEN_ENDEDITW
            ' Debug.Print "EndEditW"
            CopyMemory tNMHEW, ByVal lParam, LenB(tNMHEW)
            sMsg = tNMHEW.szText
            iPos = InStr(sMsg, vbNullChar)
            If iPos > 1 Then
               sMsg = left$(sMsg, iPos - 1)
            ElseIf iPos = 1 Then
               sMsg = ""
            End If
            RaiseEvent EndEdit(ListIndex, (tNMHEW.fChanged <> 0), sMsg, tNMHEW.iWhy, tNMHEW.iNewSelection)
         Case CBEN_ENDEDITA
            ' Debug.Print "EndEditA"
            CopyMemory tNMHE, ByVal lParam, LenB(tNMHE)
            sMsg = StrConv(tNMHE.szText, vbUnicode)
            iPos = InStr(sMsg, vbNullChar)
            If iPos > 1 Then
               sMsg = left$(sMsg, iPos - 1)
            ElseIf iPos = 1 Then
               sMsg = ""
            End If
            RaiseEvent EndEdit(ListIndex, (tNMHE.fChanged <> 0), sMsg, tNMHE.iWhy, tNMHE.iNewSelection)
         End Select
      End If
    
   Case WM_CTLCOLORLISTBOX, WM_CTLCOLOREDIT
      ' This is the only way to get the handle of the
      ' list box portion of a combo box:
      If (iMsg = WM_CTLCOLORLISTBOX) Then
         If m_eStyle <> eccxSimple Then
            If (m_hWndDropDown = 0) Then
               m_hWndDropDown = lParam
               If (IsWindow(m_hWndDropDown)) Then
                  GetWindowRect m_hWndDropDown, tR
                  bCancel = False
                  RaiseEvent RequestDropDownResize(tR.left, tR.Top, tR.right, tR.bottom, bCancel)
                  If Not bCancel Then
                     MoveWindow m_hWndDropDown, tR.left, tR.Top, tR.right - tR.left, tR.bottom - tR.Top, 1
                  End If
               End If
               If m_hWndEdit <> 0 Then
                  SetFocusAPI m_hWndEdit
               End If
            End If
         ElseIf m_eStyle = eccxSimple Then
            ''
         End If
      End If
      Debug.Print "WM_CTLCOLOR", Hex(iMsg)
      ISubclass_WindowProc = m_hBrBack

   Case WM_COMMAND
      If lParam = m_hWnd Then
         ' Debug.Print "WM_COMMAND"
         Select Case (wParam \ &H10000) And &HFFFF&
            Case CBN_DBLCLK
               RaiseEvent DblClick
            Case CBN_DROPDOWN
               RaiseEvent DropDown
            Case CBN_CLOSEUP
               If bAttach = True And m_eStyle = eccxDropDownCombo Then
                   Call DetachEdit
                   bAttach = False
               End If
               m_hWndDropDown = 0
               RaiseEvent CloseUp
            Case CBN_SETFOCUS, CBN_KILLFOCUS
               ' Not required, handed by UserControl
'               Debug.Print "CBN_SETFOCUS"
            Case CBN_SELCHANGE
               RaiseEvent Change
               RaiseEvent ListIndexChange
               RaiseEvent Click
               
            Case CBN_EDITCHANGE
               If bAttach = False And m_eStyle = eccxDropDownCombo Then
                   Call AttachEdit
                   bAttach = True
               End If
               RaiseEvent Change
         End Select
      End If
   
   Case WM_KEYDOWN
      iKeyCode = (wParam And &HFF)
      RaiseEvent KeyDown(iKeyCode, piGetShiftState())
      If (iKeyCode = 0) Then
         ' consume
      Else
         If iKeyCode <> 0 Then
            wParam = wParam And Not &HFF&
            wParam = wParam Or (iKeyCode And &HFF&)
            If iKeyCode = 13 Then
                ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If
            If m_eStyle = eccxDropDownCombo And m_bDoAutoComplete Then
               If ComboIsDropped And iKeyCode = vbKeyReturn Then
                  sText = Text
                  ShowDropDown False
                  Text = sText
               End If
            End If
         End If
      End If
      
   Case WM_CHAR
      iKeyCode = (wParam And &HFF)
      If hwnd = m_hWndCbo Then
         If m_eStyle <> eccxDropDownList Then
            ' Forward the message on to the edit box:
            SendMessageLong m_hWndEdit, iMsg, wParam, lParam
            iKeyCode = 0
         End If
      End If
         
      If iKeyCode <> 0 Then
         RaiseEvent KeyPress(iKeyCode)
         If (iKeyCode = 0) Then
            ' consume:
         Else
            If (m_eStyle <> eccxDropDownList) Then
               If (m_bDoAutoComplete) Then
                  AutoCompleteKeyPress iKeyCode
                  Debug.Print iKeyCode
                  If (iKeyCode = vbKeyEscape) Then
                     ' consume:
                     Debug.Print "Escape"
                     ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
                     If ComboIsDropped Then
                        ShowDropDown False
                     End If
                     RaiseEvent AutoCompleteSelection(List(ListIndex), ListIndex)
                  End If
               End If
            End If
            wParam = wParam And Not &HFF&
            wParam = wParam Or (iKeyCode And &HFF&)
            If (iKeyCode <> 0) Then
               ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If
         End If
      End If
      
   Case WM_KEYUP
      ' Debug.Print "sending to ", hwnd
      iKeyCode = (wParam And &HFF)
      RaiseEvent KeyUp(iKeyCode, piGetShiftState())
      If (iKeyCode = 0) Then
         ' consume
      Else
         wParam = wParam And Not &HFF&
         wParam = wParam Or (iKeyCode And &HFF&)
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
            
   ' ------------------------------------------------------------------------------
   ' Implement focus.  Many many thanks to Mike Gainer for showing me this
   ' code.
   Case WM_SETFOCUS
      If Not m_bInFocus Then
         If IsWindowVisible(hwnd) Then
            If (m_hWndCbo = hwnd) Or (m_hWndEdit = hwnd) Or (m_hWnd = hwnd) Then
               ' The combo box itself
               Dim pOleObject                  As IOleObject
               Dim pOleInPlaceSite             As IOleInPlaceSite
               Dim pOleInPlaceFrame            As IOleInPlaceFrame
               Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
'               Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
               Dim PosRect                     As RECT
               Dim ClipRect                    As RECT
               Dim FrameInfo                   As OLEINPLACEFRAMEINFO
               Dim grfModifiers                As Long
               Dim AcceleratorMsg              As Msg

               'Get in-place frame and make sure it is set to our in-between
               'implementation of IOleInPlaceActiveObject in order to catch
               'TranslateAccelerator calls
               Set pOleObject = Me
               Set pOleInPlaceSite = pOleObject.GetClientSite
               pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
'               CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
               pOleInPlaceFrame.SetActiveObject m_IPAOHookStruct.ThisPointer, vbNullString
               If Not pOleInPlaceUIWindow Is Nothing Then
                  pOleInPlaceUIWindow.SetActiveObject m_IPAOHookStruct.ThisPointer, vbNullString
               End If
'               CopyMemory pOleInPlaceActiveObject, 0&, 4
               m_bInFocus = True
            Else
               ' The user control - forward focus to the
               ' Comboex control window:
               SetFocusAPI m_hWnd
            End If
         End If
      End If
      
   Case WM_MOUSEACTIVATE
      If Not m_bInFocus Then
         If GetFocus() <> m_hWndCbo And GetFocus() <> m_hWndEdit Then
            ' Click mouse down but miss the contained control; eat
            ' activate and setfocus to the the user control, this in
            ' turn focuses the contained Comboex
            SetFocusAPI UserControl.hwnd
            ISubclass_WindowProc = MA_NOACTIVATE
            Exit Function
         Else
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If
      Else
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
   ' End Implement focus.
   ' ------------------------------------------------------------------------------
   End Select
   
End Function

Private Sub UserControl_Initialize()
Dim iccex As tagInitCommonControlsEx
   debugmsg "UniComboBox:Initialize"
   
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
   ' Default conditions:
   m_bEnabled = True
   m_bRedraw = True
   
   ' Attach custom IOleInPlaceActiveObject interface
   Dim IPAO As IOleInPlaceActiveObject

   With m_IPAOHookStruct
      Set IPAO = Me
      CopyMemory .IPAOReal, IPAO, 4
      CopyMemory .TBEx, Me, 4
      .lpVTable = cboIPAOVTable
      .ThisPointer = VarPtr(m_IPAOHookStruct)
   End With
End Sub

Private Sub UserControl_InitProperties()
    bUnicode = True
    m_bDesignTime = Not (UserControl.Ambient.UserMode)
    UserControl.Extender.Width = ScaleX((TextWidth("A") + 6), vbPixels, vbContainerSize)
    plCreate
    Set Font = UserControl.Ambient.Font
    BackColor = vbWindowBackground
End Sub

Private Sub UserControl_LostFocus()
   Debug.Print "LostFocus"
   m_bInFocus = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    If Ambient.UserMode Then Call CheckLicensed 'khi chay chuong trinh thi kiem tra ban quyen
    Dim lH As Long
        
        bUnicode = PropBag.ReadProperty("AutoUnicode", True)
        m_bDesignTime = Not (UserControl.Ambient.UserMode)
        Debug.Print UserControl.Extender.Name
        
        Style = PropBag.ReadProperty("Style", eccxDropDownCombo)
        Enabled = PropBag.ReadProperty("Enabled", True)
        plCreate
        
        ExtendedUI = PropBag.ReadProperty("ExtendedUI", True)
        DropDownWidth = PropBag.ReadProperty("DropDownWidth", -1)
        AutoCompleteListItemsOnly = PropBag.ReadProperty("AutoCompleteListItemsOnly", False)
        AutoCompleteItemsAreSorted = PropBag.ReadProperty("AutoCompleteItemsAreSorted", False)
        DoAutoComplete = PropBag.ReadProperty("DoAutoComplete", False)
        DrawStyle = PropBag.ReadProperty("DrawStyle", eccxDrawDefault)
        Redraw = PropBag.ReadProperty("Redraw", True)
        BackColor = vbWindowBackground
        
'        AddItem "To Allow SetFont/Height"
        Dim iFnt As IFont, iFntCopy As IFont
        Set iFnt = UserControl.Font
        iFnt.Clone iFntCopy
        Set m_fnt = iFntCopy
        Set Font = PropBag.ReadProperty("Font", m_fnt)
        Clear
        
        If Not (m_bDesignTime) Then
            Select Case DrawStyle
                Case eccxDriveList
                    LoadDriveList Me, (m_lIconSizeY > 16)
                Case eccxSysColourPicker
                    LoadSysColorList Me
            End Select
        End If
        
        ' for VB6
        UserControl_Resize
        
        m_bEvents = True
End Sub

Private Sub UserControl_Resize()
Dim tR As RECT
Dim lHeight As Long

   If Not (m_hWnd = 0) Then
      If Not (m_eStyle = eccxSimple) Then
         If m_bDesignTime Then
            ' Make sure the User Control's height is correct:
            lHeight = SendMessageLong(m_hWnd, CB_GETITEMHEIGHT, -1, 0)
            UserControl.Extender.Height = (lHeight + 6) * Screen.TwipsPerPixelY
         End If
      End If
      
      GetClientRect UserControl.hwnd, tR
      MoveWindow m_hWnd, 0, 0, tR.right - tR.left, tR.bottom - tR.Top, 1
      If m_eStyle <> eccxSimple Then
         lHeight = tR.bottom - tR.Top + 2 + SendMessageLong(m_hWnd, CB_GETITEMHEIGHT, 0, 0) * 8
      Else
         lHeight = tR.bottom - tR.Top
      End If
      MoveWindow m_hWndCbo, 0, 0, tR.right - tR.left, lHeight, 1
   End If

End Sub

Private Sub UserControl_Terminate()
    ' Detach the custom IOleInPlaceActiveObject interface
   ' pointers.
   With m_IPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
   End With
   
   pDestroy
   
   debugmsg "UniComboBox:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoUnicode", bUnicode, True
    PropBag.WriteProperty "Style", Style, eccxDropDownCombo
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "ExtendedUI", ExtendedUI, True
    PropBag.WriteProperty "DropDownWidth", DropDownWidth, -1
    PropBag.WriteProperty "AutoCompleteListItemsOnly", AutoCompleteListItemsOnly, False
    PropBag.WriteProperty "AutoCompleteItemsAreSorted", AutoCompleteItemsAreSorted, False
    PropBag.WriteProperty "DoAutoComplete", DoAutoComplete, False
    PropBag.WriteProperty "DrawStyle", DrawStyle, eccxDrawDefault
    PropBag.WriteProperty "Redraw", Redraw, True
    'PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
End Sub

Private Sub gSplitDelimitedString(ByVal sString As String, ByVal sDelim As String, _
                                  ByRef sValues() As String, ByRef iCount As Integer)
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
Dim iPos As Integer
Dim iNextPos As Integer
Dim iDelimLen As Integer
    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub

Private Function glCStr(ByVal sThis As String, Optional ByVal lDefault As Long = 0) As Long
On Error Resume Next
    glCStr = CLng(sThis)
    If (Err.Number <> 0) Then
        glCStr = lDefault
    End If
End Function

Private Sub debugmsg(ByVal sMsg As String)
    #Const DEBUG_MSG = 0
    #If DEBUG_MSG = 1 Then
       MsgBox sMsg, vbInformation
    #Else
       Debug.Print sMsg
    #End If
End Sub

Private Sub LoadDriveList(ByVal cbo As UniComboBox, ByVal bLargeIcons As Boolean)
'// ==========================================================================
'// Load Items - collects all drive information and place it into the listbox,
'// return number of items added to the list: a negative value is an error;
'// ==========================================================================
Dim lAllDriveStrings As Long
Dim sDrive As String
Dim lR As Long
Dim dwIconSize As Long
Dim FileInfo As SHFILEINFO
Dim iPos As Long, iLastPos As Long
Dim iType As EDriveType
Dim hIml As Long
Dim dwFlags As Long
Dim lDefIndex As Long

   cbo.Clear
   cbo.Redraw = False
   
   '// allocate buffer for the drive strings: GetLogicalDriveStrings will tell
   '// me how much is needed (minus the trailing zero-byte)
   lAllDriveStrings = GetLogicalDriveStrings(0, ByVal 0&)

   m_sDriveStrings = String$(lAllDriveStrings + 1, 0) 'new _TCHAR[ lAllDriveStrings + sizeof( _T("")) ]; // + for trailer
   lR = GetLogicalDriveStrings(lAllDriveStrings, ByVal m_sDriveStrings)
   Debug.Assert lR = (lAllDriveStrings - 1)
  
   InitSystemImageList cbo, bLargeIcons
  
   '// now loop over each drive (string)
   If bLargeIcons Then
      dwIconSize = SHGFI_LARGEICON
   Else
      dwIconSize = SHGFI_SMALLICON
   End If
   
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, m_sDriveStrings, vbNullChar)
      
      If iPos <> 0 Then
         sDrive = Mid$(m_sDriveStrings, iLastPos, iPos - iLastPos)
         iLastPos = iPos + 1
      Else
         sDrive = Mid$(m_sDriveStrings, iLastPos)
      End If
      If Not sDrive = vbNullString Then
         lR = SHGetFileInfo(sDrive, FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or dwIconSize)
         If (lR = 0) Then  '// failure - which can be ignored
            Debug.Print "SHGetFileInfo failed, no more details available"
         Else
            '// insert icon and string into list box
            cbo.AddItemAndData FileInfo.szDisplayName, FileInfo.iIcon, FileInfo.iIcon, Asc(left$(sDrive, 1))
            If lDefIndex = 0 Then
               iType = GetDriveType(left$(sDrive, 2))
               If iType = 1 Or iType = DRIVE_FIXED Then
                  lDefIndex = cbo.NewIndex
               End If
            End If
         End If
         cbo.ListIndex = lDefIndex
      Else
         iPos = 0
      End If
   Loop While iPos <> 0
   cbo.Redraw = True
   
End Sub

Private Sub LoadSysColorList(ByRef cbo As UniComboBox)
      'assign system color names
   With cbo
      .Clear
      .Redraw = False
      .AddItemAndData "3DDKShadow", , , vb3DDKShadow
      .AddItemAndData "3DFace", , , vb3DFace
      .AddItemAndData "3DHighlight", , , vb3DHighlight
      .AddItemAndData "3DLight", , , vb3DLight
      .AddItemAndData "3DShadow", , , vb3DShadow
      .AddItemAndData "ActiveBorder", , , vbActiveBorder
      .AddItemAndData "ActiveTitleBar", , , vbActiveTitleBar
      .AddItemAndData "ApplicationWorkspace", , , vbApplicationWorkspace
      .AddItemAndData "ButtonFace", , , vbButtonFace
      .AddItemAndData "ButtonShadow", , , vbButtonShadow
      .AddItemAndData "ButtonText", , , vbButtonText
      .AddItemAndData "Desktop", , , vbDesktop
      .AddItemAndData "GrayText", , , vbGrayText
      .AddItemAndData "Highlight", , , vbHighlight
      .AddItemAndData "HighlightText", , , vbHighlightText
      .AddItemAndData "InactiveBorder", , , vbInactiveBorder
      .AddItemAndData "InactiveCaptionText", , , vbInactiveCaptionText
      .AddItemAndData "InactiveTitleBar", , , vbInactiveTitleBar
      .AddItemAndData "InfoBackground", , , vbInfoBackground
      .AddItemAndData "InfoText", , , vbInfoText
      .AddItemAndData "MenuBar", , , vbMenuBar
      .AddItemAndData "MenuText", , , vbMenuText
      .AddItemAndData "ScrollBars", , , vbScrollBars
      .AddItemAndData "TitleBarText", , , vbTitleBarText
      .AddItemAndData "WindowBackground", , , vbWindowBackground
      .AddItemAndData "WindowFrame", , , vbWindowFrame
      .AddItemAndData "WindowText", , , vbWindowText
      .ListIndex = 0
      .Redraw = True
   End With

End Sub

Private Sub InitSystemImageList(ByRef cbo As UniComboBox, ByVal bLargeIcons As Boolean)
Dim dwFlags As Long
Dim hIml As Long
Dim FileInfo As SHFILEINFO

   dwFlags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX
   If Not (bLargeIcons) Then
      dwFlags = dwFlags Or SHGFI_SMALLICON
   End If
   '// Load the image list - use an arbitrary file extension for the
   '// call to SHGetFileInfo (we don't want to touch the disk, so use
   '// FILE_ATTRIBUTE_NORMAL && SHGFI_USEFILEATTRIBUTES).
   hIml = SHGetFileInfo(".txt", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), dwFlags)
       
   ' MFC code sample says to do this, but this looks dubious to me.  Likely
   ' you will disrupt Explorer in Win9x...
   '// Make the background colour transparent, works better for lists etc.
   'ImageList_SetBkColor m_hIml, CLR_NONE
   
   cbo.ImageList = hIml

End Sub

Private Function pvGetTextLen(ByVal hwnd As Long) As Long
' Get length of the caption
    If IsWindowUnicode(hwnd) Then
        pvGetTextLen = GetWindowTextLengthW(hwnd)
    Else
        pvGetTextLen = GetWindowTextLengthA(hwnd)
    End If
End Function

Private Function pvStripNulls(ByVal sString As String) As String
Dim lPos As Long

    lPos = InStr(sString, vbNullChar)
    If (lPos = 1) Then
        pvStripNulls = vbNullString
    ElseIf (lPos > 1) Then
        pvStripNulls = left$(sString, lPos - 1)
        Exit Function
    End If
    pvStripNulls = sString
End Function

Private Function pvGetWindowText(ByVal hwnd As Long) As String
Dim lLen As Long
Dim sBuf As String

    lLen = 1 + pvGetTextLen(hwnd)
    If (lLen > 1) Then
        sBuf = String$(lLen, 0)
        If IsWindowUnicode(hwnd) Then
            GetWindowTextW hwnd, StrPtr(sBuf), lLen
        Else
            GetWindowTextA hwnd, sBuf, lLen
        End If
        pvGetWindowText = pvStripNulls(sBuf)
    Else
        pvGetWindowText = vbNullString
    End If
End Function

Private Sub pvSetWindowText(ByVal sText As String)
Dim lPtr As Long

    If IsWindowUnicode(m_hWndEdit) Then
        If Len(sText) = 0 Then
            SetWindowTextW m_hWndEdit, StrPtr(vbNullString)
            Exit Sub
        End If
        lPtr = StrPtr(sText)
        SetWindowTextW m_hWndEdit, lPtr
    Else
        If Len(sText) = 0 Then
            SetWindowTextA m_hWndEdit, vbNullString
            Exit Sub
        End If
        SetWindowTextA m_hWndEdit, sText
    End If
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next:   frmAbout.Show vbModal
End Sub
