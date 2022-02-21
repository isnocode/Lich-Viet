Attribute VB_Name = "mLunarDate"
'----------------------------------------------------------------------------------------
'* Copyright 2004 Ho Ngoc Duc [http://come.to/duc]. All Rights Reserved.<p>
'* Permission to use, copy, modify, and redistribute this software and its
'* documentation for personal, non-commercial use is hereby granted provided that
'* this copyright notice appears in all copies.
'----------------------------------------------------------------------------------------

Option Explicit

Public Type fThoiGian
    fdNgay As Double
    fdThang As Double
    fdNam As Double
End Type

Public Type ThongTinAmLich
    sThuTrongTuanVN As String
    fdNgayAL As fThoiGian
    dNgayOfLeap As Double
    dLeap As Double
    dLeapMonth As Double
    fdLeap2SonarFrom As fThoiGian
    fdLeap2SonarTo As fThoiGian
    dThangLenght As Double
    sDoiThangAL As String
    sNgayAmLich As String
    sThangAmLich As String
    sNamAmLich As String
    sGioAmLich As String
    sTietKhi As String
    sThuTrongTuanEng As String
    fdThoigianbatDauTietKhiDL As fThoiGian
    fdThoigianbatDauTietKhiAL As fThoiGian
End Type

Public Const PI = 3.14159265358979

Public ThienCan, DiaChi, VTenNgayTrongTuan, TietKhi

Private Sub TaoThongtinAmLich()
    ThienCan = Array("Gia1p", "A61t", "Bi1nh", "D9inh", "Ma65u", "Ky3", "Canh", "Ta6n", "Nha6m", "Quy1")
    DiaChi = Array("Ty1", "Su73u", "Da62n", "Ma4o", "Thi2n", "Ty5", "Ngo5", "Mu2i", "Tha6n", "Da65u", "Tua61t", "Ho75i")
    VTenNgayTrongTuan = Array("Chu3 nha65t", "Thu71 hai", "Thu71 ba", "Thu71 tu7", "Thu71 na8m", "Thu71 sa1u", "Thu71 ba3y")
    TietKhi = Array("Xua6n pha6n", "Thanh Minh", "Co61c vu4", "La65p ha5", _
    "Tie63u ma4n", "Mang chu3ng", "Ha5 chi1", "Tie63u thu73", _
    "D9a5i thu73", "La65p thu", "Xu73 thu73", "Ba5ch lo65", _
    "Thu pha6n", "Ha2n lo65", "Su7o7ng gia1ng", "La65p d9o6ng", _
    "Tie63u tuye61t", "D9a5i tuye61t", "D9o6ng chi1", "Tie63u ha2n", _
    "D9a5i ha2n", "La65p xua6n", "Vu4 thu3y", "Ki2nh tra65p")
End Sub

Private Function JuliusDays(dNgay As Double, dThang As Double, dNam As Double) As Double 'Done
    On Error GoTo Err_WrongDate
    
    Dim dD As Double, dM As Double, dy As Double
    
    dD = Int((14 - dThang) / 12)
    dy = dNam + 4800 - dD
    dM = dThang + 12 * dD - 3
    
    JuliusDays = dNgay + Int((153 * dM + 2) / 5) + 365 * dy + Int(dy / 4) - Int(dy / 100) + Int(dy / 400) - 32045
    
    If JuliusDays < 2299161 Then
        JuliusDays = dNgay + Int((153 * dM + 2) / 5) + 365 * dy + Int(dy / 4) - 32083
    End If
    Exit Function
    
Err_WrongDate:
    JuliusDays = -1
End Function

Private Function JuliusDays2Date(JDNumber As Double) As fThoiGian 'Done
    On Error GoTo Err_CantCalc
    
    Dim dA As Double, dB As Double, dC As Double, dD As Double
    Dim dE As Double, dM As Double
    
    With JuliusDays2Date
        If JDNumber < 2299161 Then
            dA = JDNumber
        Else
            dM = Int((JDNumber - 1867216.25) / 36524.25)
            dA = JDNumber + 1 + dM - Int(dM / 4)
        End If
        dB = dA + 1524
        dC = Int((dB - 122.1) / 365.25)
        dD = Int(365.25 * dC)
        dE = Int((dB - dD) / 30.6001)
        .fdNgay = Int(dB - dD - Int(30.6001 * dE))
        If dE < 14 Then
            .fdThang = dE - 1
        Else
            .fdThang = dE - 13
        End If
        If .fdThang < 3 Then
            .fdNam = dC - 4715
        Else
            .fdNam = dC - 4716
        End If
    End With
    Exit Function
    
Err_CantCalc:
    With JuliusDays2Date
        .fdNgay = -1
        .fdThang = -1
        .fdNam = -1
    End With
End Function

Private Function GetNewMoonDay(NMPos As Double, dTimeZone As Double) As Double 'Done
    'Return Julius Day at New Moon position from 01/01/1900
    Dim dT1 As Double, dT2 As Double, dT3 As Double
    Dim dDr As Double, dJD1 As Double, dM As Double
    Dim dMpr As Double, dF As Double, dC1 As Double
    Dim dDelta As Double, dJDNew As Double
    
    dT1 = NMPos / 1236.85
    dT2 = dT1 * dT1
    dT3 = dT2 * dT1
    
    dDr = PI / 180
    dJD1 = 2415020.75933 + 29.53058868 * NMPos + 0.0001178 * dT2 - 0.000000155 * dT3
    dJD1 = dJD1 + 0.00033 * Sin((166.56 + 132.87 * dT1 - 0.009173 * dT2) * dDr)
    dM = 359.2242 + 29.10535608 * NMPos - 0.0000333 * dT2 - 0.00000347 * dT3
    dMpr = 306.0253 + 385.81691806 * NMPos + 0.0107306 * dT2 + 0.00001236 * dT3
    dF = 21.2964 + 390.67050646 * NMPos - 0.0016528 * dT2 - 0.00000239 * dT3
    dC1 = (0.1734 - 0.000393 * dT1) * Sin(dM * dDr) + 0.0021 * Sin(2 * dDr * dM)
    dC1 = dC1 - 0.4068 * Sin(dMpr * dDr) + 0.0161 * Sin(dDr * 2 * dMpr)
    dC1 = dC1 - 0.0004 * Sin(dDr * 3 * dMpr)
    dC1 = dC1 + 0.0104 * Sin(dDr * 2 * dF) - 0.0051 * Sin(dDr * (dM + dMpr))
    dC1 = dC1 - 0.0074 * Sin(dDr * (dM - dMpr)) + 0.0004 * Sin(dDr * (2 * dF + dM))
    dC1 = dC1 - 0.0004 * Sin(dDr * (2 * dF - dM)) - 0.0006 * Sin(dDr * (2 * dF + dMpr))
    dC1 = dC1 + 0.001 * Sin(dDr * (2 * dF - dMpr)) + 0.0005 * Sin(dDr * (2 * dMpr + dM))
    
    If dT1 < -11 Then
        dDelta = 0.001 + 0.000839 * dT1 + 0.0002261 * dT2 - 0.00000845 * dT3 - 0.000000081 * dT1 * dT3
    Else
        dDelta = -0.000278 + 0.000265 * dT1 + 0.000262 * dT2
    End If
    
    dJDNew = dJD1 + dC1 - dDelta
    GetNewMoonDay = Int(dJDNew + 0.5 + dTimeZone / 24)
End Function

Private Function SunLongitude(dJD As Double, dTimeZone) As Double 'Done
    Dim dT1 As Double, dT2 As Double, dDr As Double
    Dim dM As Double, dL0 As Double, dDL As Double
    Dim dL As Double
    
    dT1 = (dJD - 2451545.5 - dTimeZone / 24) / 36525
    dT2 = dT1 * dT1
    dDr = PI / 180
    
    dM = 357.5291 + 35999.0503 * dT1 - 0.0001559 * dT2 - 0.00000048 * dT1 * dT2
    dL0 = 280.46645 + 36000.76983 * dT1 + 0.0003032 * dT2
    dDL = (1.9146 - 0.004817 * dT1 - 0.000014 * dT2) * Sin(dDr * dM)
    dDL = dDL + (0.019993 - 0.000101 * dT1) * Sin(dDr * 2 * dM) + 0.00029 * Sin(dDr * 3 * dM)
    dL = dL0 + dDL
    dL = dL * dDr
    dL = dL - PI * 2 * Int(dL / (PI * 2))
    
    SunLongitude = Int(dL / PI * 6)
End Function

Private Function SunLongitude2(dJD As Double) As Double 'Done
    Dim dT1 As Double, dT2 As Double, dDr As Double
    Dim dM As Double, dL0 As Double, dDL As Double
    Dim dL As Double
    
    dT1 = (dJD - 2451545) / 36525
    dT2 = dT1 * dT1
    dDr = PI / 180
    
    dM = 357.5291 + 35999.0503 * dT1 - 0.0001559 * dT2 - 0.00000048 * dT1 * dT2
    dL0 = 280.46645 + 36000.76983 * dT1 + 0.0003032 * dT2
    dDL = (1.9146 - 0.004817 * dT1 - 0.000014 * dT2) * Sin(dDr * dM)
    dDL = dDL + (0.019993 - 0.000101 * dT1) * Sin(dDr * 2 * dM) + 0.00029 * Sin(dDr * 3 * dM)
    dL = dL0 + dDL
    dL = dL * dDr
    dL = dL - PI * 2 * Int(dL / (PI * 2))
    
    SunLongitude2 = dL
End Function

Private Function GetSunLongitude(dJD As Double, dTimeZone As Double) As Double
    GetSunLongitude = Int(SunLongitude2(dJD - 0.5 - dTimeZone / 24) / PI * 12)
End Function

Private Function GetAmLichMonth11th(dNam As Double, dTimeZone As Double) As Double 'Done
    Dim dK As Double, dOff As Double, sunLong As Double
    Dim NM As Double
    
    dOff = JuliusDays(31, 12, dNam) - 2415021
    dK = Int(dOff / 29.530588853)
    NM = GetNewMoonDay(dK, dTimeZone)
    sunLong = SunLongitude(NM, dTimeZone)
    
    If sunLong >= 9 Then NM = GetNewMoonDay(dK - 1, dTimeZone)
    
    GetAmLichMonth11th = NM
End Function

Private Function GetLeapMonthOffset(dThang11th As Double, dTimeZone As Double) As Double 'Done
    Dim dK As Double, dLast As Double, dArc As Double, i As Integer
    
    dK = Int((dThang11th - 2415021.07699869) / 29.530588853 + 0.5)
    dLast = 0
    i = 1
    dArc = SunLongitude(GetNewMoonDay(dK + i, dTimeZone), dTimeZone)
    Do
        dLast = dArc
        i = i + 1
        dArc = SunLongitude(GetNewMoonDay(dK + i, dTimeZone), dTimeZone)
    Loop While (dArc <> dLast And i < 14)
    
    GetLeapMonthOffset = i - 1
End Function

Public Function DuongLich2AmLich(dNgay As Double, dThang As Double, dNam As Double, dTimeZone As Double) As fThoiGian
    
    Dim dK As Double, dNgayNum As Double, dThangStart As Double
    Dim dF11 As Double, dS11 As Double, dAmLichDay As Double
    Dim dAmLichMonth As Double, dAmLichYear As Double, dAmLichLeap As Double
    Dim dDiff As Double, dLeapMonthDiff As Double
    
    dNgayNum = JuliusDays(dNgay, dThang, dNam)
    dK = Int((dNgayNum - 2415024.07699869) / 29.530588853)
    dThangStart = GetNewMoonDay(dK + 1, dTimeZone)
    If dThangStart > dNgayNum Then dThangStart = GetNewMoonDay(dK, dTimeZone)
    dF11 = GetAmLichMonth11th(dNam, dTimeZone)
    dS11 = dF11
    If dF11 > dThangStart Then
        dAmLichYear = dNam
        dF11 = GetAmLichMonth11th(dNam - 1, dTimeZone)
    Else
        dAmLichYear = dNam + 1
        dS11 = GetAmLichMonth11th(dNam + 1, dTimeZone)
    End If
    dAmLichDay = dNgayNum - dThangStart + 1
    dDiff = Int((dThangStart - dF11) / 29)
    dAmLichLeap = 0
    dAmLichMonth = dDiff + 11
    If (dS11 - dF11) > 365 Then
        dLeapMonthDiff = GetLeapMonthOffset(dF11, dTimeZone)
        If dDiff >= dLeapMonthDiff Then
            dAmLichMonth = dDiff + 10
            If dDiff = dLeapMonthDiff Then dAmLichLeap = 1
        End If
    End If
    If dAmLichMonth > 12 Then dAmLichMonth = dAmLichMonth - 12
    If dAmLichMonth >= 11 And dDiff < 4 Then dAmLichYear = dAmLichYear - 1
    DuongLich2AmLich.fdNgay = dAmLichDay
    DuongLich2AmLich.fdThang = dAmLichMonth
    DuongLich2AmLich.fdNam = dAmLichYear
End Function

Public Function AmLich2DuongLich(dAmLichDay As Double, dAmLichMonth As Double, dAmLichYear As Double, dAmLichLeap As Double, dTimeZone As Double) As fThoiGian
    Dim dK As Double, dF11 As Double, dS11 As Double
    Dim dOff As Double, dLeapOff As Double, dLeapMonth As Double
    Dim dThangStart As Double
    If dAmLichMonth < 11 Then
        dF11 = GetAmLichMonth11th(dAmLichYear - 1, dTimeZone)
        dS11 = GetAmLichMonth11th(dAmLichYear, dTimeZone)
    Else
        dF11 = GetAmLichMonth11th(dAmLichYear, dTimeZone)
        dS11 = GetAmLichMonth11th(dAmLichYear + 1, dTimeZone)
    End If
    dOff = dAmLichMonth - 11
    If dOff < 0 Then dOff = dOff + 12
    If (dS11 - dF11) > 365 Then
        dLeapOff = GetLeapMonthOffset(dF11, dTimeZone)
        dLeapMonth = dLeapOff - 2
        If dLeapMonth < 0 Then dLeapMonth = dLeapMonth + 12
        If dAmLichLeap <> 0 And dAmLichMonth <> dLeapMonth Then
            AmLich2DuongLich.fdNgay = 0
            AmLich2DuongLich.fdThang = 0
            AmLich2DuongLich.fdNam = 0
        ElseIf dAmLichLeap <> 0 Or dAmLichMonth <> dLeapMonth Then
            dOff = dOff + 1
        End If
    End If
    dK = Int(0.5 + (dF11 - 2415021.07699869) / 29.530588853)
    dThangStart = GetNewMoonDay(dK + dOff, dTimeZone)
    With JuliusDays2Date(dThangStart + dAmLichDay - 1)
        AmLich2DuongLich.fdNgay = .fdNgay
        AmLich2DuongLich.fdThang = .fdThang
        AmLich2DuongLich.fdNam = .fdNam
    End With
End Function

'----------------------------------------------------------------------------------------
'* Test, Edit and Translate by vie87vn - www.caulacbovb.com
'----------------------------------------------------------------------------------------
Public Function GetThongTinAmLich(dNgay As Double, dThang As Double, dNam As Double, dTimeZone As Double) As ThongTinAmLich
    TaoThongtinAmLich
    Dim i As Integer, dJD As Double, dLLeap As Double
    Dim dLDay As Double, dLMonth As Double, dLYear As Double
    Dim dThangLen As Double, bLFullMonth As Boolean
    Dim dLM As Double, iPos As Integer, iCan As Integer, ichi As Integer
    
    iPos = ((JuliusDays(dNgay, dThang, dNam) + 1) Mod 7)
    
    With DuongLich2AmLich(dNgay, dThang, dNam, dTimeZone)
        dLDay = .fdNgay
        dLMonth = .fdThang
        dLYear = .fdNam
    End With
    
    dLM = GetLeapMonthOffset(GetAmLichMonth11th(dLYear, dTimeZone), dTimeZone)
    If dLM > 2 Then
        dLM = GetLeapMonthOffset(GetAmLichMonth11th(dLYear - 1, dTimeZone), dTimeZone)
    End If
    If dLM < 13 Then dLM = dLM - 2
    If dLM <= 0 Then dLM = dLM + 12
    dLLeap = IIf(dLM > 12, 0, 1)
    Dim dSDay As Double, dSMonth As Double, dSYear As Double
    dSDay = AmLich2DuongLich(1, dLMonth, dLYear, dLLeap, dTimeZone).fdNgay
    dSMonth = AmLich2DuongLich(1, dLMonth, dLYear, dLLeap, dTimeZone).fdThang
    dSYear = AmLich2DuongLich(1, dLMonth, dLYear, dLLeap, dTimeZone).fdNam
    dJD = JuliusDays(dSDay, dSMonth, dSYear)
    For i = 25 To 31
        If DuongLich2AmLich(JuliusDays2Date(dJD + i).fdNgay, JuliusDays2Date(dJD + i).fdThang, JuliusDays2Date(dJD + i).fdNam, dTimeZone).fdNgay = 1 Then
            Exit For
        End If
    Next i
    dThangLen = i
    
    With GetThongTinAmLich
        .sThuTrongTuanVN = VTenNgayTrongTuan(iPos)
        .fdNgayAL.fdNgay = dLDay
        .fdNgayAL.fdThang = dLMonth
        .fdNgayAL.fdNam = dLYear
        .dLeap = dLLeap
        .dLeapMonth = IIf(dLM > 12, 0, dLM)
        .dThangLenght = dThangLen
        iCan = (JuliusDays(dNgay, dThang, dNam) + 9) Mod 10
        ichi = (JuliusDays(dNgay, dThang, dNam) + 1) Mod 12
        .sNgayAmLich = ThienCan(iCan) & " " & DiaChi(ichi)
        iCan = (iCan * 2) Mod 10
        ichi = 0
        .sGioAmLich = ThienCan(iCan) & " " & DiaChi(ichi)
        iCan = (dLYear * 12 + dLMonth + 3) Mod 10
        ichi = IIf((dLMonth - 11) < 0, dLMonth + 1, dLMonth - 11)
        .sThangAmLich = ThienCan(iCan) & " " & DiaChi(ichi)
        iCan = (dLYear + 6) Mod 10
        ichi = (dLYear + 8) Mod 12
        .sNamAmLich = ThienCan(iCan) & " " & DiaChi(ichi)
        .dNgayOfLeap = 0
        If dLLeap = 1 Then
            .fdLeap2SonarFrom.fdNgay = AmLich2DuongLich(1, .dLeapMonth, dLYear, dLLeap, dTimeZone).fdNgay
            .fdLeap2SonarFrom.fdThang = AmLich2DuongLich(1, .dLeapMonth, dLYear, dLLeap, dTimeZone).fdThang
            .fdLeap2SonarFrom.fdNam = AmLich2DuongLich(1, .dLeapMonth, dLYear, dLLeap, dTimeZone).fdNam
            dJD = JuliusDays(.fdLeap2SonarFrom.fdNgay, .fdLeap2SonarFrom.fdThang, .fdLeap2SonarFrom.fdNam)
            For i = 25 To 31
                If DuongLich2AmLich(JuliusDays2Date(dJD + i).fdNgay, JuliusDays2Date(dJD + i).fdThang, JuliusDays2Date(dJD + i).fdNam, dTimeZone).fdNgay = 1 Then
                    Exit For
                End If
            Next i
            dJD = dJD + i - 1
            .fdLeap2SonarTo.fdNgay = JuliusDays2Date(dJD).fdNgay
            .fdLeap2SonarTo.fdThang = JuliusDays2Date(dJD).fdThang
            .fdLeap2SonarTo.fdNam = JuliusDays2Date(dJD).fdNam
            If dThang = .fdLeap2SonarFrom.fdThang Then
                If dNgay >= .fdLeap2SonarFrom.fdNgay Then .dNgayOfLeap = 1 Else .dNgayOfLeap = 0
            ElseIf dThang = .fdLeap2SonarTo.fdThang Then
                If dNgay <= .fdLeap2SonarTo.fdNgay Then .dNgayOfLeap = 1 Else .dNgayOfLeap = 0
            Else
                .dNgayOfLeap = 0
            End If
        Else
            .fdLeap2SonarFrom.fdNgay = 0
            .fdLeap2SonarFrom.fdThang = 0
            .fdLeap2SonarFrom.fdNam = 0
            .fdLeap2SonarTo.fdNgay = 0
            .fdLeap2SonarTo.fdThang = 0
            .fdLeap2SonarTo.fdNam = 0
        End If
        iPos = GetSunLongitude(JuliusDays(dNgay, dThang, dNam) + 1, dTimeZone)
        .sTietKhi = TietKhi(iPos)
        dJD = JuliusDays(dNgay, dThang, dNam) + 1
        For i = 0 To 20
            If GetSunLongitude(dJD - i, dTimeZone) <> iPos Then Exit For
        Next i
        .fdThoigianbatDauTietKhiDL.fdNgay = JuliusDays2Date(dJD - i).fdNgay
        .fdThoigianbatDauTietKhiDL.fdThang = JuliusDays2Date(dJD - i).fdThang
        .fdThoigianbatDauTietKhiDL.fdNam = JuliusDays2Date(dJD - i).fdNam
        .fdThoigianbatDauTietKhiAL.fdNgay = DuongLich2AmLich(.fdThoigianbatDauTietKhiDL.fdNgay, .fdThoigianbatDauTietKhiDL.fdThang, .fdThoigianbatDauTietKhiDL.fdNam, dTimeZone).fdNgay
        .fdThoigianbatDauTietKhiAL.fdThang = DuongLich2AmLich(.fdThoigianbatDauTietKhiDL.fdNgay, .fdThoigianbatDauTietKhiDL.fdThang, .fdThoigianbatDauTietKhiDL.fdNam, dTimeZone).fdThang
        .fdThoigianbatDauTietKhiAL.fdNam = DuongLich2AmLich(.fdThoigianbatDauTietKhiDL.fdNgay, .fdThoigianbatDauTietKhiDL.fdThang, .fdThoigianbatDauTietKhiDL.fdNam, dTimeZone).fdNam
        .sDoiThangAL = DoiThangAL(dLMonth)
        .sThuTrongTuanEng = GetDayName(.sThuTrongTuanVN)
    End With
End Function

Public Function DoiThangAL(dThang As Double) As String
    Select Case dThang
    Case 1: DoiThangAL = "Tha1ng gie6ng"
    Case 2: DoiThangAL = "Tha1ng hai"
    Case 3: DoiThangAL = "Tha1ng ba"
    Case 4: DoiThangAL = "Tha1ng tu7"
    Case 5: DoiThangAL = "Tha1ng na8m"
    Case 6: DoiThangAL = "Tha1ng sa1u"
    Case 7: DoiThangAL = "Tha1ng ba3y"
    Case 8: DoiThangAL = "Tha1ng ta1m"
    Case 9: DoiThangAL = "Tha1ng chi1n"
    Case 10: DoiThangAL = "Tha1ng mu7o72i"
    Case 11: DoiThangAL = "Tha1ng mu7o72i mo65t"
    Case 12: DoiThangAL = "Tha1ng cha5p"
    End Select
End Function
Public Function DoiquathangEnglish(ByVal Month As Integer) As String
    Select Case Month
    Case 1: DoiquathangEnglish = "January"
    Case 2: DoiquathangEnglish = "February"
    Case 3: DoiquathangEnglish = "Match"
    Case 4: DoiquathangEnglish = "April"
    Case 5: DoiquathangEnglish = "May"
    Case 6: DoiquathangEnglish = "June"
    Case 7: DoiquathangEnglish = "July"
    Case 8: DoiquathangEnglish = "August"
    Case 9: DoiquathangEnglish = "September"
    Case 10: DoiquathangEnglish = "October"
    Case 11: DoiquathangEnglish = "November"
    Case 12: DoiquathangEnglish = "December"
    End Select
End Function
Private Function doi_ra_JD(ByVal dD, ByVal mm, ByVal yy) As String
    Dim A, y, m, JD
    A = Int((14 - mm) / 12)
    y = yy + 4800 - A
    m = mm + 12 * A - 3
    JD = dD + Int((153 * m + 2) / 5) + 365 * y + Int(y / 4) - Int(y / 100) + Int(y / 400) - 32045
    If (JD < 2299161) Then JD = dD + Int((153 * m + 2) / 5) + 365 * y + Int(y / 4) - 32083
    doi_ra_JD = JD
End Function
Public Function Gio_Hoang_Dao(ByVal dD, ByVal mm, ByVal yy) As String
    Dim Ngay_JD1 As Double
    Ngay_JD1 = CDbl((doi_ra_JD(dD, mm, yy) + 1) Mod 12)
    Select Case Ngay_JD1
    Case 0: Gio_Hoang_Dao = "Ty1(23-1), Su73u(1-3), Ma4o(5-7), Ngo5(11-13), Tha6n(15-17), Da65u(17-19)"
    Case 1: Gio_Hoang_Dao = "Da62n(3-5), Ma4o(5-7), Ty5(9-11), Tha6n(15-17), Tua61t(19-21), Ho75i(21-23)"
    Case 2: Gio_Hoang_Dao = "Ty1(23-1), Su73u(1-3), Thi2n(7-9), Ty5(9-11), Mu2i(13-15), Tua61t(19-21)"
    Case 3: Gio_Hoang_Dao = "Ty1(23-1), Da62n(3-5), Ma4o(5-7), Ngo5(11-13), Mu2i(13-15), Da65u(17-19)"
    Case 4: Gio_Hoang_Dao = "Da62n(3-5), Thi2n(7-9), Ty5(9-11), Tha6n(15-17), Da65u(17-19), Ho75i(21-23)"
    Case 5: Gio_Hoang_Dao = "Su73u(1-3), Thi2n(7-9), Ngo5(11-13), Mu2i(13-15), Tua61t(19-21), Ho75i(21-23)"
    Case 6: Gio_Hoang_Dao = "Ty1(23-1), Su73u(1-3), Ma4o(5-7), Ngo5(11-13), Tha6n(15-17), Da65u(17-19)"
    Case 7: Gio_Hoang_Dao = "Da62n(3-5), Ma4o(5-7), Ty5(9-11), Tha6n(15-17), Tua61t(19-21), Ho75i(21-23)"
    Case 8: Gio_Hoang_Dao = "Ty1(23-1), Su73u(1-3), Thi2n(7-9), Ty5(9-11), Mu2i(13-15), Tua61t(19-21)"
    Case 9: Gio_Hoang_Dao = "Ty1(23-1), Da62n(3-5), Ma4o(5-7), Ngo5(11-13), Mu2i(13-15), Da65u(17-19)"
    Case 10: Gio_Hoang_Dao = "Da62n(3-5), Thi2n(7-9), Ty5(9-11), Tha6n(15-17), Da65u(17-19), Ho75i(21-23)"
    Case 11: Gio_Hoang_Dao = "Su73u(1-3), Thi2n(7-9), Ngo5(11-13), Mu2i(13-15), Tua61t(19-21), Ho75i(21-23)"
    End Select
End Function
Public Function sNgayHoangDao(dNgay As Double, dThang As Double, dNam As Double) As String
    Dim ichi As String
    ichi = DiaChi((JuliusDays(dNgay, dThang, dNam) + 1) Mod 12)
    Select Case dThang
    Case 1, 7
        If ichi = "Ty1" Or ichi = "Su73u" Or ichi = "Ty5" Or ichi = "Mu2i" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Ngo5" Or ichi = "Ma4o" Or ichi = "Ho75i" Or ichi = "Da65u" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    Case 2, 8
        If ichi = "Da62n" Or ichi = "Ma4o" Or ichi = "Da65u" Or ichi = "Mu2i" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Tha6n" Or ichi = "Ty5" Or ichi = "Ho75i" Or ichi = "Su73u" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    Case 3, 9
        If ichi = "Thi2n" Or ichi = "Da65u" Or ichi = "Ty5" Or ichi = "Ho75i" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Tua61t" Or ichi = "Ma4o" Or ichi = "Mu2i" Or ichi = "Su73u" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    Case 4, 10
        If ichi = "Ngo5" Or ichi = "Mu2i" Or ichi = "Ho75i" Or ichi = "Su73u" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Ty1" Or ichi = "Ma4o" Or ichi = "Ty5" Or ichi = "Da65u" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    Case 5, 11
        If ichi = "Tha6n" Or ichi = "Da65u" Or ichi = "Su73u" Or ichi = "Ma4o" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Da62n" Or ichi = "Mu2i" Or ichi = "Ho75i" Or ichi = "Ty5" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    Case 6, 12
        If ichi = "Tua61t" Or ichi = "Ho75i" Or ichi = "Ty5" Or ichi = "Ma4o" Then
            sNgayHoangDao = "Nga2y Hoa2ng D9a5o(to61t)"
        ElseIf ichi = "Thi2n" Or ichi = "Su73u" Or ichi = "Mu2i" Or ichi = "Da65u" Then
            sNgayHoangDao = "Nga2y Ha81c D9a5o(xa61u)"
        Else
            sNgayHoangDao = ""
        End If
    End Select
End Function
Public Function GetDayName(iDay As String) As String
    Select Case iDay
    Case "Thu71 hai": GetDayName = "Monday"
    Case "Thu71 ba": GetDayName = "Tuesday"
    Case "Thu71 tu7": GetDayName = "Wednesday"
    Case "Thu71 na8m": GetDayName = "Thursday"
    Case "Thu71 sa1u": GetDayName = "Friday"
    Case "Thu71 ba3y": GetDayName = "Saturday"
    Case "Chu3 nha65t": GetDayName = "Sunday"
    End Select
End Function
Public Function GetDaysOfMonth(iMonth As Double, iYear As Double) As Integer
    Select Case iMonth
    Case 1: GetDaysOfMonth = 31
    Case 2
        If IsDate("02/29/" & iYear) Then
            GetDaysOfMonth = 29
        Else
            GetDaysOfMonth = 28
        End If
    Case 3: GetDaysOfMonth = 31
    Case 4: GetDaysOfMonth = 30
    Case 5: GetDaysOfMonth = 31
    Case 6: GetDaysOfMonth = 30
    Case 7: GetDaysOfMonth = 31
    Case 8: GetDaysOfMonth = 31
    Case 9: GetDaysOfMonth = 30
    Case 10: GetDaysOfMonth = 31
    Case 11: GetDaysOfMonth = 30
    Case 12: GetDaysOfMonth = 31
    End Select
End Function
