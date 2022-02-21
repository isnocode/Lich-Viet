Attribute VB_Name = "ModNgayLe"
Option Explicit
Public Type TypeEvent
    Dong1 As String
    Dong2 As String
    Tacgia As String
    lColor As Long
End Type
Dim dD As Double, dM As Double, dy As Double
Public Function MyEvent(dD As Double, dM As Double, Optional dy As Double) As TypeEvent
    Dim aLuDate() As String
    MyEvent.lColor = RGB(214, 227, 188)
    If dy = 0 Then dy = Year(Date)
    With DuongLich2AmLich(dD, dM, dy, 7)
        Select Case .fdThang 'am lich
        Case 1
            Select Case .fdNgay
            Case 1
                MyEvent.Dong1 = "MO62NG 1 TE61T NGUYE6N D9A1N"
                MyEvent.Dong2 = "CHU1C MU72NG NA8M MO71I"
                MyEvent.lColor = vbRed
                Exit Function
            Case 2
                MyEvent.Dong1 = "MO62NG 2 TE61T NGUYE6N D9A1N"
                MyEvent.Dong2 = "CHU1C MU72NG NA8M MO71I"
                MyEvent.lColor = vbRed
                Exit Function
            Case 3
                MyEvent.Dong1 = "MO62NG 3 TE61T NGUYE6N D9A1N"
                MyEvent.Dong2 = "CHU1C MU72NG NA8M MO71I"
                MyEvent.lColor = vbRed
                Exit Function
            Case 15
                MyEvent.Dong1 = "TIE61T NGUYE6N TIE6U"
                MyEvent.Dong2 = "(Le64 ho65i lo62ng d9e2n)"
                MyEvent.lColor = vbRed
                Exit Function
            End Select
        Case 2
            Select Case .fdNgay
            Case 19
                MyEvent.Dong1 = "Nga2y sinh Quan A6m"
                MyEvent.Dong2 = ""
                Exit Function
            End Select
        Case 3
            Select Case .fdNgay
            Case 10
                MyEvent.Dong1 = "Nha81n ai d9i ngu7o75c ve62 xuo6i"
                MyEvent.Dong2 = "Nho71 nga2y gio64 to63 mo62ng 10 tha1ng 3"
                Exit Function
            End Select
        Case 4
        Case 5
            Select Case .fdNgay
            Case 5
                MyEvent.Dong1 = "TE61T D9OAN NGO5"
                MyEvent.Dong2 = "(Le64 ho65i d9ua thuye62n)"
                MyEvent.lColor = vbYellow
                Exit Function
            End Select
        Case 8
            If .fdNgay = 15 Then
                MyEvent.Dong1 = "TE61T TRUNG THU"
                MyEvent.Dong2 = "(Nga2y sum ho5p)"
                MyEvent.lColor = vbWhite
                Exit Function
            End If
        Case 9
            If .fdNgay = 9 Then
                MyEvent.Dong1 = "TE61T TRU2NG DU7O7NG"
                MyEvent.Dong2 = "(Nga2y leo nu1i va2 ha1i hoa)"
                MyEvent.lColor = RGB(255, 127, 223)
                Exit Function
            End If
        End Select
    End With
    '------------duong lich
    Select Case dM
    Case 1
        Select Case dD
        Case 1
            MyEvent.Dong1 = "TE61T DU7O7NG LI5CH"
            MyEvent.Dong2 = "CHU1C MU72NG NA8M MO71I"
            MyEvent.lColor = vbRed
        End Select
    Case 2
        Select Case dD
        Case 3
            MyEvent.Dong1 = "NGA2Y THA2NH LA65P D9A3NG"
            MyEvent.Dong2 = "CO65NG SA3N VIE65T NAM"
        Case 14
            MyEvent.Dong1 = "Le64 ti2nh nha6n"
            MyEvent.Dong2 = "Happy Valentine Day"
            MyEvent.lColor = RGB(191, 127, 255)
        End Select
    Case 3
        Select Case dD
        Case 8
            MyEvent.Dong1 = "NGA2Y QUO61C TE61 PHU5 NU74"
            MyEvent.Dong2 = "Tinh tha62n Hai Ba2 Tru7ng ba61t die65t"
        Case 26
            MyEvent.Dong1 = "Na2y tha2nh la65p"
            MyEvent.Dong2 = "D9oa2n Thanh Nie6n Co65ng Sa3n"
        End Select
    Case 4
        Select Case dD
        Case 1
            MyEvent.Dong1 = "NGA2Y QUO61C TE61 NO1I DO61I"
            MyEvent.Dong2 = ""
        Case 30
            MyEvent.Dong1 = "GIA3I PHO1NG D9A61T NU7O71C"
            MyEvent.Dong2 = "Gia3i pho1ng mie62n Nam Vie65t Nam"
        End Select
    Case 5
        Select Case dD
        Case 1
            MyEvent.Dong1 = "NGA2Y QUO61C TE61 LAO D9O65NG"
            MyEvent.Dong2 = ""
        End Select
    Case 6
        Select Case dD
        Case 1
            MyEvent.Dong1 = "NGA2Y QUO61C TE61 THIE61U NHI"
            MyEvent.Dong2 = ""
            MyEvent.lColor = RGB(254, 210, 76)
        End Select
    Case 9
        Select Case dD
        Case 2
            MyEvent.Dong1 = "QUO61C KHA1NH NU7O71C CHXHCN"
            MyEvent.Dong2 = "VIE65T NAM"
            MyEvent.lColor = vbRed
        End Select
    Case 10
        Select Case dD
        Case 20
            MyEvent.Dong1 = "NGA2Y PHU5 NU74 VIE65T NAM"
            MyEvent.Dong2 = ""
        End Select
    Case 11
        Select Case dD
        Case 20
            MyEvent.Dong1 = "NGA2Y NHA2 GIA1O VIE65T NAM"
            MyEvent.Dong2 = ""
        End Select
    Case 12
        Select Case dD
        Case 22
            MyEvent.Dong1 = "Ky3 nie65m " & Year(Date) - 1954 & " na8m tha2nh la65p"
            MyEvent.Dong2 = "QUA6N D9O65I NHA6N DA6N VIE65T NAM"
            MyEvent.lColor = RGB(0, 127, 31)
        Case 24
            MyEvent.Dong1 = "LE64 GIA1NG SINH"
            MyEvent.Dong2 = "Mu72ng gia1ng sinh an la2nh"
            MyEvent.lColor = vbWhite
        End Select
    End Select
    If MyEvent.Dong1 = "" Then
        MyEvent.Dong1 = "Kho6ng co1 vie65c gi2 kho1"
        MyEvent.Dong2 = "Chi3 so75 lo2ng kho6ng be62n"
        MyEvent.Tacgia = "Tu5c Ngu74"
        MyEvent.lColor = vbWhite
    End If
End Function
