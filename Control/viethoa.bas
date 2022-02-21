Attribute VB_Name = "viethoact"
'su dung module nay
'de hien thi tieng Viet trong chuong trinh
Public Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function MessageBox Lib "USER32.DLL" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar& Lib "User32" (ByVal hwnd&)
Private Declare Function GetMenuItemInfo Lib "User32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "User32" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetWindowText Lib "USER32.DLL" Alias "SetWindowTextW" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "User32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Const MIIM_DATA& = &H20
Private Const MIIM_TYPE& = &H10
Private Const MB_ICONINFORMATION& = &H40&
Private Const MF_BYPOSITION = &H400&

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


'''<<<
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "User32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Public Function cUni(str$) As String
    Dim ansi$, UNI$, i&, sTem$, sUni$, arrUNI() As String
    ansi = "a1 a2 a3 a4 a5 a6 a8 a61a62a63a64a65a81a82a83a84a85A1 A2 A3 A4 A5 A6 A8 A61A62A63A64A65A81A82A83A84A85e1 e2 e3 e4 e5 e6 e61e62e63e64e65E1 E2 E3 E4 E5 E6 E61E62E63E64E65i1 i2 i3 i4 i5 I1 I2 I3 I4 I5 o1 o2 o3 o4 o5 o6 o7 o61o62o63o64o65o71o72o73o74o75O1 O2 O3 O4 O5 O6 O7 O61O62O63O64O65O71O72O73O74O75u1 u2 u3 u4 u5 u7 u71u72u73u74u75U1 U2 U3 U4 U5 U7 U71U72U73U74U75y1 y2 y3 y4 y5 Y1 Y2 Y3 Y4 Y5 d9 D9 "
    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7,C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6,E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7,C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6,ED,EC,1EC9,129,1ECB,CD,CC,1EC8,128,1ECA,F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9,1EDB,1EDD,1EDF,1EE1,1EE3,D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8,1EDA,1EDC,1EDE,1EE0,1EE2,FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1,DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0,FD,1EF3,1EF7,1EF9,1EF5,DD,1EF2,1EF6,1EF8,1EF4,111,110"
    arrUNI = Split(UNI, ",")
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i + 1, 1)) = False Then
            sUni = sUni & Mid(str, i, 1)
        Else
            sTem = IIf(IsNumeric(Mid(str, i + 2, 1)), Mid(str, i, 3), Mid(str, i, 2))
            i = i + IIf(IsNumeric(Mid(str, i + 2, 1)), 2, 1)
            If InStr(ansi, sTem) > 0 Then sTem = ChrW("&h" & arrUNI(InStr(ansi, sTem) \ 3))
            sUni = sUni & sTem
        End If
    Next
    cUni = sUni
End Function
'''>>>
Public Function ABC2Uni(str$) As String
    Dim i&, arrUNI() As String, sUni$, ABC$, UNI$, tmp$
    Dim y As Long
    ABC = "¸µ¶·¹¨¾»¼½Æ©ÊÇÈÉËÐÌÎÏÑªÕÒÓÔÖÝ×ØÜÞãßáâä«èåæçé¬íêëìîóïñòô­øõö÷ùýúûüþ®¸µ¶·¹¡¾»¼½Æ¢ÊÇÈÉËÐÌÎÏÑ£ÕÒÓÔÖÝ×ØÜÞãßáâä¤èåæçé¥íêëìîóïñòô¦øõö÷ùýúûüþ§"
    UNI = "225,224,7843,227,7841,259,7855,7857,7859,7861,7863,226,7845,7847,7849,7851,7853,233,232,7867,7869,7865,234,7871,7873,7875,7877,7879,237,236,7881,297,7883,243,242,7887,245,7885,244,7889,7891,7893,7895,7897,417,7899,7901,7903,7905,7907,250,249,7911,361,7909,432,7913,7915,7917,7919,7921,253,7923,7927,7929,7925,273,225,224,7843,227,7841,258,7855,7857,7859,7861,7863,194,7845,7847,7849,7851,7853,233,232,7867,7869,7865,202,7871,7873,7875,7877,7879,237,236,7881,297,7883,243,242,7887,245,7885,212,7889,7891,7893,7895,7897,416,7899,7901,7903,7905,7907,250,249,7911,361,7909,431,7913,7915,7917,7919,7921,253,7923,7927,7929,7925,272"
    arrUNI = Split(UNI, ",")
    For i = 1 To Len(str$)
        tmp = Mid$(str$, i, 1)
        y = InStr(1, ABC, tmp)
        If y > 0 Then
            sUni = sUni & ChrW(arrUNI(y - 1))
        Else
            sUni = sUni & tmp
        End If
    Next
    ABC2Uni = sUni
End Function
Public Function Msg(ByVal Text As String, Optional ByVal Title As String = "Tho6ng ba1o", Optional Button As VbMsgBoxStyle = vbInformation, Optional ByVal FormhWnd As Long = 0) As VbMsgBoxResult
    Msg = MessageBox(FormhWnd, StrPtr(UNI(Text)), StrPtr(UNI(Title)), Button)
End Function
Public Function UNI(sText As String)
    Dim i As Integer, j As Integer
    Dim sCurChar As String, sPreChar As String, sPreTxt As String
    For j = 1 To 2
        For i = 2 To Len(sText)
            sCurChar = Mid(sText, i, 1)
            sPreTxt = left(sText, i - 2)
            sPreChar = Mid(sText, i - 1, 1)
            Select Case sCurChar
            Case "1"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE1) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EAF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EAE) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HE9) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HC9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EBF) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EBE) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&HED) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&HCD) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF3) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED0) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDB) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDA) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&HFA) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&HDA) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EE9) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EE8) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&HFD) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&HDD) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "2"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE0) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC0) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB0) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HE8) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HC8) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC0) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&HEC) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&HCC) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF2) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD2) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED2) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDD) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDC) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&HF9) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&HD9) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EEB) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEA) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF3) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF2) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "3"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H1EA3) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H1EA2) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA8) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB3) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB2) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EBB) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EBA) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC2) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H1EC9) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H1EC8) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1ECF) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1ECE) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDE) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1EE7) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1EE6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EED) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEC) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF7) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF6) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "4"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE3) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EAB) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EAA) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB5) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB4) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EBD) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EBC) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC4) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H129) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H128) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF5) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EE1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EE0) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H169) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H168) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EEF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEE) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF9) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF8) & Right(sText, Len(sText) - i)
                End Select
                
            Case "5"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H1EA1) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H1EA0) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EAD) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EAC) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB7) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB6) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EB9) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EB8) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC6) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H1ECB) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H1ECA) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1ECD) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1ECC) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED8) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EE3) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EE2) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1EE5) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1EE4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EF1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EF0) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF5) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF4) & Right(sText, Len(sText) - i)
                End Select
                
            Case "6"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE2) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC2) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HEA) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HCA) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF4) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD4) & Right(sText, Len(sText) - i)
                End Select
                
            Case "7"
                Select Case sPreChar
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1A1) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1A0) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1B0) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1AF) & Right(sText, Len(sText) - i)
                End Select
                
            Case "8"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H103) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H102) & Right(sText, Len(sText) - i)
                End Select
                
            Case "9"
                Select Case sPreChar
                    'd
                Case "d": sText = sPreTxt & ChrW$(&H111) & Right(sText, Len(sText) - i)
                Case "D": sText = sPreTxt & ChrW$(&H110) & Right(sText, Len(sText) - i)
                End Select
                
            End Select
        Next i
    Next j
    UNI = sText
End Function
Public Sub InitMenu(frm As Form)
    Dim hMenu&
    hMenu = GetMenu(frm.hwnd)
    VietnameseMenu hMenu
End Sub
Private Sub VietnameseMenu(ByVal hMenu As Long)
    Dim hSubMenu&, i%, nCnt%, sTmp$
    Dim MII As MENUITEMINFO
    
    nCnt = GetMenuItemCount(hMenu)
    If nCnt Then
        For i = 0 To nCnt - 1
            MII.cbSize = LenB(MII)
            MII.fMask = MIIM_TYPE Or MIIM_DATA
            MII.dwTypeData = String(&HFF, 0)
            MII.cch = Len(MII.dwTypeData)
            '           MII.hbmpChecked = MF_CHECKED Or MF_UNCHECKED
            
            'lay chuoi ten cua Menu
            GetMenuItemInfo hMenu, i, True, MII
            sTmp = left$(MII.dwTypeData, MII.cch)
            
            'Viet lai chuoi ten cua Menu va chuyen no sang Unicode
            sTmp = UNI(MII.dwTypeData)
            sTmp = StrConv(sTmp, vbUnicode) 'note
            MII.dwTypeData = sTmp
            SetMenuItemInfo hMenu, i, True, MII
            
            'lay Menu con cua mot MenuItem
            hSubMenu = GetSubMenu(hMenu, i)
            If hSubMenu Then
                VietnameseMenu hSubMenu
            End If
        Next
    End If
End Sub

Public Sub SetIcon(frm As Form, MenuNumber As Integer, SubMenuItemCount1 As Integer, Optional SubMenuItemCount2 As Integer, Optional SubMenuItemCount3 As Integer, Optional Icon As Picture, Optional isDefault As Boolean)
    On Error GoTo err
    Dim hMainMenu As Long, hSubMenu1 As Long, hSubMenu2 As Long, hSubMenu3 As Long
    
    
    MenuNumber = MenuNumber - 1
    SubMenuItemCount1 = SubMenuItemCount1 - 1
    SubMenuItemCount2 = SubMenuItemCount2 - 1
    SubMenuItemCount3 = SubMenuItemCount3 - 1
    
    hMainMenu = GetMenu(frm.hwnd)       'lay menu cua form
    
    
    If SubMenuItemCount1 >= 0 Then hSubMenu1 = GetSubMenu(hMainMenu, MenuNumber)        'lay menu con thu 1
    If SubMenuItemCount2 >= 0 Then hSubMenu2 = GetSubMenu(hSubMenu1, SubMenuItemCount1) 'lay menu con thu 2
    If SubMenuItemCount3 >= 0 Then hSubMenu3 = GetSubMenu(hSubMenu2, SubMenuItemCount2) 'lay menu con thu 3
    
    'neu chon Icon cho mot Menu khong ton tai trong Menu hien tai thi thoat khoi thuc tuc
    If (hSubMenu3 = 0 And SubMenuItemCount3 >= 0) Or (hSubMenu2 = 0 And SubMenuItemCount2 >= 0) Or (hSubMenu1 = 0 And SubMenuItemCount1 >= 0) Then Exit Sub
    
    'neu chon dat Icon cho menu con cap 3
    If hSubMenu3 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu3, SubMenuItemCount3, 1
        SetMenuItemBitmaps hSubMenu3, SubMenuItemCount3, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If
    
    'neu chon dat Icon cho menu con cap 2
    If hSubMenu2 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu2, SubMenuItemCount2, 1
        SetMenuItemBitmaps hSubMenu2, SubMenuItemCount2, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If
    
    'neu chon dat Icon cho menu con cap 1
    If hSubMenu1 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu1, SubMenuItemCount1, 1
        SetMenuItemBitmaps hSubMenu1, SubMenuItemCount1, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If
    
err:
    'loi xay ra khi chon menu can dat Icon ma khong dat icon
End Sub

'-----Lam Trong suot

Public Function isTransparent(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        isTransparent = True
    Else
        isTransparent = False
    End If
    If err Then
        isTransparent = False
    End If
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If err Then
        MakeTransparent = 2
    End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
    Dim Msg As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If err Then
        MakeOpaque = 2
    End If
End Function
Public Function ConvertToUnicode(sText As String)
    Dim i As Integer, j As Integer
    Dim sCurChar As String, sPreChar As String, sPreTxt As String
    For j = 1 To 2
        For i = 2 To Len(sText)
            sCurChar = Mid(sText, i, 1)
            sPreTxt = left(sText, i - 2)
            sPreChar = Mid(sText, i - 1, 1)
            Select Case sCurChar
            Case "1"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE1) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EAF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EAE) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HE9) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HC9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EBF) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EBE) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&HED) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&HCD) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF3) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HDA): sText = sPreTxt & ChrW$(&H1ED0) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDB) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDA) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&HFA) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&HDA) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EE9) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EE8) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&HFD) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&HDD) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "2"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE0) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC0) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB0) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HE8) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HC8) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC1) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC0) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&HEC) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&HCC) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF2) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD2) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HDA): sText = sPreTxt & ChrW$(&H1ED2) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDD) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDC) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&HF9) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&HD9) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EEB) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEA) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF3) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF2) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "3"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H1EA3) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H1EA2) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EA9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EA8) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB3) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB2) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EBB) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EBA) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC2) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H1EC9) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H1EC8) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1ECF) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1ECE) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HDA): sText = sPreTxt & ChrW$(&H1ED4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EDF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EDE) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1EE7) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1EE6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EED) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEC) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF7) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF6) & Right(sText, Len(sText) - i)
                    
                End Select
                
            Case "4"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE3) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC3) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EAB) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EAA) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB5) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB4) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EBD) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EBC) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC4) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H129) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H128) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF5) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD5) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HDA): sText = sPreTxt & ChrW$(&H1ED6) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EE1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EE0) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H169) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H168) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EEF) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EEE) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF9) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF8) & Right(sText, Len(sText) - i)
                End Select
                
            Case "5"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H1EA1) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H1EA0) & Right(sText, Len(sText) - i)
                Case ChrW$(&HE2): sText = sPreTxt & ChrW$(&H1EAD) & Right(sText, Len(sText) - i)
                Case ChrW$(&HC2): sText = sPreTxt & ChrW$(&H1EAC) & Right(sText, Len(sText) - i)
                Case ChrW$(&H103): sText = sPreTxt & ChrW$(&H1EB7) & Right(sText, Len(sText) - i)
                Case ChrW$(&H102): sText = sPreTxt & ChrW$(&H1EB6) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&H1EB9) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&H1EB8) & Right(sText, Len(sText) - i)
                Case ChrW$(&HEA): sText = sPreTxt & ChrW$(&H1EC7) & Right(sText, Len(sText) - i)
                Case ChrW$(&HCA): sText = sPreTxt & ChrW$(&H1EC6) & Right(sText, Len(sText) - i)
                    
                    'i
                Case "i": sText = sPreTxt & ChrW$(&H1ECB) & Right(sText, Len(sText) - i)
                Case "I": sText = sPreTxt & ChrW$(&H1ECA) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1ECD) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1ECC) & Right(sText, Len(sText) - i)
                Case ChrW$(&HF4): sText = sPreTxt & ChrW$(&H1ED9) & Right(sText, Len(sText) - i)
                Case ChrW$(&HD4): sText = sPreTxt & ChrW$(&H1ED8) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A1): sText = sPreTxt & ChrW$(&H1EE3) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1A0): sText = sPreTxt & ChrW$(&H1EE2) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1EE5) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1EE4) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1B0): sText = sPreTxt & ChrW$(&H1EF1) & Right(sText, Len(sText) - i)
                Case ChrW$(&H1AF): sText = sPreTxt & ChrW$(&H1EF0) & Right(sText, Len(sText) - i)
                    
                    'y
                Case "y": sText = sPreTxt & ChrW$(&H1EF5) & Right(sText, Len(sText) - i)
                Case "Y": sText = sPreTxt & ChrW$(&H1EF4) & Right(sText, Len(sText) - i)
                End Select
                
            Case "6"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&HE2) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&HC2) & Right(sText, Len(sText) - i)
                    
                    'e
                Case "e": sText = sPreTxt & ChrW$(&HEA) & Right(sText, Len(sText) - i)
                Case "E": sText = sPreTxt & ChrW$(&HCA) & Right(sText, Len(sText) - i)
                    
                    'o
                Case "o": sText = sPreTxt & ChrW$(&HF4) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&HD4) & Right(sText, Len(sText) - i)
                End Select
                
            Case "7"
                Select Case sPreChar
                    'o
                Case "o": sText = sPreTxt & ChrW$(&H1A1) & Right(sText, Len(sText) - i)
                Case "O": sText = sPreTxt & ChrW$(&H1A0) & Right(sText, Len(sText) - i)
                    
                    'u
                Case "u": sText = sPreTxt & ChrW$(&H1B0) & Right(sText, Len(sText) - i)
                Case "U": sText = sPreTxt & ChrW$(&H1AF) & Right(sText, Len(sText) - i)
                End Select
                
            Case "8"
                Select Case sPreChar
                    'a
                Case "a": sText = sPreTxt & ChrW$(&H103) & Right(sText, Len(sText) - i)
                Case "A": sText = sPreTxt & ChrW$(&H102) & Right(sText, Len(sText) - i)
                End Select
                
            Case "9"
                Select Case sPreChar
                    'd
                Case "d": sText = sPreTxt & ChrW$(&H111) & Right(sText, Len(sText) - i)
                Case "D": sText = sPreTxt & ChrW$(&H110) & Right(sText, Len(sText) - i)
                End Select
                
            End Select
        Next i
    Next j
    ConvertToUnicode = sText
End Function

Public Function ConvertToString(sText As String)
    Dim i As Integer
    Dim sChar As String, sTxtOut As String
    For i = 1 To Len(sText)
        sChar = Mid(sText, i, 1)
        Select Case sChar
            'a
        Case ChrW$(&HE1): sTxtOut = sTxtOut & "a1"
        Case ChrW$(&HC1): sTxtOut = sTxtOut & "A1"
            
        Case ChrW$(&HE0): sTxtOut = sTxtOut & "a2"
        Case ChrW$(&HC0): sTxtOut = sTxtOut & "A2"
            
        Case ChrW$(&H1EA3): sTxtOut = sTxtOut & "a3"
        Case ChrW$(&H1EA2): sTxtOut = sTxtOut & "A3"
            
        Case ChrW$(&HE3): sTxtOut = sTxtOut & "a4"
        Case ChrW$(&HC3): sTxtOut = sTxtOut & "A4"
            
        Case ChrW$(&H1EA1): sTxtOut = sTxtOut & "a5"
        Case ChrW$(&H1EA0): sTxtOut = sTxtOut & "A5"
            
        Case ChrW$(&HE2): sTxtOut = sTxtOut & "a6"
        Case ChrW$(&HC2): sTxtOut = sTxtOut & "A6"
            
        Case ChrW$(&H1EA5): sTxtOut = sTxtOut & "a61"
        Case ChrW$(&H1EA4): sTxtOut = sTxtOut & "A61"
            
        Case ChrW$(&H1EA7): sTxtOut = sTxtOut & "a62"
        Case ChrW$(&H1EA6): sTxtOut = sTxtOut & "A62"
            
        Case ChrW$(&H1EA9): sTxtOut = sTxtOut & "a63"
        Case ChrW$(&H1EA8): sTxtOut = sTxtOut & "A63"
            
        Case ChrW$(&H1EAB): sTxtOut = sTxtOut & "a64"
        Case ChrW$(&H1EAA): sTxtOut = sTxtOut & "A64"
            
        Case ChrW$(&H1EAD): sTxtOut = sTxtOut & "a65"
        Case ChrW$(&H1EAC): sTxtOut = sTxtOut & "A65"
            
        Case ChrW$(&H103): sTxtOut = sTxtOut & "a8"
        Case ChrW$(&H102): sTxtOut = sTxtOut & "A8"
            
        Case ChrW$(&H1EAF): sTxtOut = sTxtOut & "a81"
        Case ChrW$(&H1EAE): sTxtOut = sTxtOut & "A81"
            
        Case ChrW$(&H1EB1): sTxtOut = sTxtOut & "a82"
        Case ChrW$(&H1EB0): sTxtOut = sTxtOut & "A82"
            
        Case ChrW$(&H1EB3): sTxtOut = sTxtOut & "a83"
        Case ChrW$(&H1EB2): sTxtOut = sTxtOut & "A83"
            
        Case ChrW$(&H1EB5): sTxtOut = sTxtOut & "a84"
        Case ChrW$(&H1EB4): sTxtOut = sTxtOut & "A84"
            
        Case ChrW$(&H1EB7): sTxtOut = sTxtOut & "a85"
        Case ChrW$(&H1EB6): sTxtOut = sTxtOut & "A85"
            
            'e
        Case ChrW$(&HE9): sTxtOut = sTxtOut & "e1"
        Case ChrW$(&HC9): sTxtOut = sTxtOut & "E1"
            
        Case ChrW$(&HE8): sTxtOut = sTxtOut & "e2"
        Case ChrW$(&HC8): sTxtOut = sTxtOut & "E2"
            
        Case ChrW$(&H1EBB): sTxtOut = sTxtOut & "e3"
        Case ChrW$(&H1EBA): sTxtOut = sTxtOut & "E3"
            
        Case ChrW$(&H1EBD): sTxtOut = sTxtOut & "e4"
        Case ChrW$(&H1EBC): sTxtOut = sTxtOut & "E4"
            
        Case ChrW$(&H1EB9): sTxtOut = sTxtOut & "e5"
        Case ChrW$(&H1EB8): sTxtOut = sTxtOut & "E5"
            
        Case ChrW$(&HEA): sTxtOut = sTxtOut & "e6"
        Case ChrW$(&HCA): sTxtOut = sTxtOut & "E6"
            
        Case ChrW$(&H1EBF): sTxtOut = sTxtOut & "e61"
        Case ChrW$(&H1EBE): sTxtOut = sTxtOut & "E61"
            
        Case ChrW$(&H1EC1): sTxtOut = sTxtOut & "e62"
        Case ChrW$(&H1EC0): sTxtOut = sTxtOut & "E62"
            
        Case ChrW$(&H1EC3): sTxtOut = sTxtOut & "e63"
        Case ChrW$(&H1EC2): sTxtOut = sTxtOut & "E63"
            
        Case ChrW$(&H1EC5): sTxtOut = sTxtOut & "e64"
        Case ChrW$(&H1EC4): sTxtOut = sTxtOut & "E64"
            
        Case ChrW$(&H1EC7): sTxtOut = sTxtOut & "e65"
        Case ChrW$(&H1EC6): sTxtOut = sTxtOut & "E65"
            
            'i
        Case ChrW$(&HED): sTxtOut = sTxtOut & "i1"
        Case ChrW$(&HCD): sTxtOut = sTxtOut & "I1"
            
        Case ChrW$(&HEC): sTxtOut = sTxtOut & "i2"
        Case ChrW$(&HCC): sTxtOut = sTxtOut & "I2"
            
        Case ChrW$(&H1EC9): sTxtOut = sTxtOut & "i3"
        Case ChrW$(&H1EC8): sTxtOut = sTxtOut & "I3"
            
        Case ChrW$(&H129): sTxtOut = sTxtOut & "i4"
        Case ChrW$(&H128): sTxtOut = sTxtOut & "I4"
            
        Case ChrW$(&H1ECB): sTxtOut = sTxtOut & "i5"
        Case ChrW$(&H1ECA): sTxtOut = sTxtOut & "I5"
            
            'o
        Case ChrW$(&HF3): sTxtOut = sTxtOut & "o1"
        Case ChrW$(&HD3): sTxtOut = sTxtOut & "O1"
            
        Case ChrW$(&HF2): sTxtOut = sTxtOut & "o2"
        Case ChrW$(&HD2): sTxtOut = sTxtOut & "O2"
            
        Case ChrW$(&H1ECF): sTxtOut = sTxtOut & "o3"
        Case ChrW$(&H1ECE): sTxtOut = sTxtOut & "O3"
            
        Case ChrW$(&HF5): sTxtOut = sTxtOut & "o4"
        Case ChrW$(&HD5): sTxtOut = sTxtOut & "O4"
            
        Case ChrW$(&H1ECD): sTxtOut = sTxtOut & "o5"
        Case ChrW$(&H1ECC): sTxtOut = sTxtOut & "O5"
            
        Case ChrW$(&HF4): sTxtOut = sTxtOut & "o6"
        Case ChrW$(&HD4): sTxtOut = sTxtOut & "O6"
            
        Case ChrW$(&H1ED1): sTxtOut = sTxtOut & "o61"
        Case ChrW$(&H1ED0): sTxtOut = sTxtOut & "O61"
            
        Case ChrW$(&H1ED3): sTxtOut = sTxtOut & "o62"
        Case ChrW$(&H1ED2): sTxtOut = sTxtOut & "O62"
            
        Case ChrW$(&H1ED5): sTxtOut = sTxtOut & "o63"
        Case ChrW$(&H1ED4): sTxtOut = sTxtOut & "O63"
            
        Case ChrW$(&H1ED7): sTxtOut = sTxtOut & "o64"
        Case ChrW$(&H1ED6): sTxtOut = sTxtOut & "O64"
            
        Case ChrW$(&H1ED9): sTxtOut = sTxtOut & "o65"
        Case ChrW$(&H1ED8): sTxtOut = sTxtOut & "O65"
            
        Case ChrW$(&H1A1): sTxtOut = sTxtOut & "o7"
        Case ChrW$(&H1A0): sTxtOut = sTxtOut & "O7"
            
        Case ChrW$(&H1EDB): sTxtOut = sTxtOut & "o71"
        Case ChrW$(&H1EDA): sTxtOut = sTxtOut & "O71"
            
        Case ChrW$(&H1EDD): sTxtOut = sTxtOut & "o72"
        Case ChrW$(&H1EDC): sTxtOut = sTxtOut & "O72"
            
        Case ChrW$(&H1EDF): sTxtOut = sTxtOut & "o73"
        Case ChrW$(&H1EDE): sTxtOut = sTxtOut & "O73"
            
        Case ChrW$(&H1EE1): sTxtOut = sTxtOut & "o74"
        Case ChrW$(&H1EE0): sTxtOut = sTxtOut & "O74"
            
        Case ChrW$(&H1EE3): sTxtOut = sTxtOut & "o75"
        Case ChrW$(&H1EE2): sTxtOut = sTxtOut & "O75"
            
            'u
        Case ChrW$(&HFA): sTxtOut = sTxtOut & "u1"
        Case ChrW$(&HDA): sTxtOut = sTxtOut & "U1"
            
        Case ChrW$(&HF9): sTxtOut = sTxtOut & "u2"
        Case ChrW$(&HD9): sTxtOut = sTxtOut & "U2"
            
        Case ChrW$(&H1EE7): sTxtOut = sTxtOut & "u3"
        Case ChrW$(&H1EE6): sTxtOut = sTxtOut & "U3"
            
        Case ChrW$(&H169): sTxtOut = sTxtOut & "u4"
        Case ChrW$(&H168): sTxtOut = sTxtOut & "U4"
            
        Case ChrW$(&H1EE5): sTxtOut = sTxtOut & "u5"
        Case ChrW$(&H1EE4): sTxtOut = sTxtOut & "U5"
            
        Case ChrW$(&H1B0): sTxtOut = sTxtOut & "u7"
        Case ChrW$(&H1AF): sTxtOut = sTxtOut & "U7"
            
        Case ChrW$(&H1EE9): sTxtOut = sTxtOut & "u71"
        Case ChrW$(&H1EE8): sTxtOut = sTxtOut & "U71"
            
        Case ChrW$(&H1EEB): sTxtOut = sTxtOut & "u72"
        Case ChrW$(&H1EEA): sTxtOut = sTxtOut & "U72"
            
        Case ChrW$(&H1EED): sTxtOut = sTxtOut & "u73"
        Case ChrW$(&H1EEC): sTxtOut = sTxtOut & "U73"
            
        Case ChrW$(&H1EEF): sTxtOut = sTxtOut & "u74"
        Case ChrW$(&H1EEE): sTxtOut = sTxtOut & "U74"
            
        Case ChrW$(&H1EF1): sTxtOut = sTxtOut & "u75"
        Case ChrW$(&H1EF0): sTxtOut = sTxtOut & "U75"
            
            'y
        Case ChrW$(&HFD): sTxtOut = sTxtOut & "y1"
        Case ChrW$(&HDD): sTxtOut = sTxtOut & "Y1"
            
        Case ChrW$(&H1EF3): sTxtOut = sTxtOut & "y2"
        Case ChrW$(&H1EF2): sTxtOut = sTxtOut & "Y2"
            
        Case ChrW$(&H1EF7): sTxtOut = sTxtOut & "y3"
        Case ChrW$(&H1EF6): sTxtOut = sTxtOut & "Y3"
            
        Case ChrW$(&H1EF9): sTxtOut = sTxtOut & "y4"
        Case ChrW$(&H1EF8): sTxtOut = sTxtOut & "Y4"
            
        Case ChrW$(&H1EF5): sTxtOut = sTxtOut & "y5"
        Case ChrW$(&H1EF4): sTxtOut = sTxtOut & "Y5"
            
            'd
        Case ChrW$(&H111): sTxtOut = sTxtOut & "d9"
        Case ChrW$(&H110): sTxtOut = sTxtOut & "D9"
            
            'other
        Case Else: sTxtOut = sTxtOut & sChar
        End Select
    Next
    ConvertToString = sTxtOut
End Function
Public Function GetFileName(ByVal sPath As String) As String
    GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
    GetFolderPath = left(sPath, InStrRev(sPath, "\") - 1)
End Function
