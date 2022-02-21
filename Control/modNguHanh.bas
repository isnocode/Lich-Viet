Attribute VB_Name = "ModNguHanh"
Option Explicit

Public Type fNguHanhNapAm
    fdNguHanh As String
    fdTuoiXungKhac As String
    fdBatQuai As String
    fdAmDuong As String
End Type

Public Function LayNguHanhNapAm(NamAmLich As String) As fNguHanhNapAm
With LayNguHanhNapAm
Select Case NamAmLich
Case "Gia1p Ty1"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "A61t Su73u"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Bi1nh Da62n"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "D9inh Ma4o"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Ma65u Thi2n"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Ky3 Ty5"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Canh Ngo5"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Ta6n Mu2i"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Nha6m Tha6n"
.fdNguHanh = ""
.fdTuoiXungKhac = ""
.fdAmDuong = ""
Case "Quy1 Da65u"
'======10
Case "Gia1p Tua61t"
Case "A61t Ho75i"
Case "Bi1nh Ty1"
Case "D9inh Su73u"
Case "Ma65u Da62n"
Case "Ky3 Ma4o"
Case "Canh Thi2n"
Case "Ta6n Ty5"
Case "Nha6m Ngo5"
Case "Quy1 Mu2i"
'======20
Case "Gia1p Tha6n"
Case "A61t Da65u"
Case "Bi1nh Tua61t"
Case "D9inh Ho75i"
Case "Ma65u Ty1"
Case "Ky3 Su73u"
Case "Canh Da62n"
Case "Ta6n Ma4o"
Case "Nha6m Thi2n"
Case "Quy1 Ty5"
'======30
Case "Gia1p Ngo5"
Case "A61t Mu2i"
Case "Bi1nh Tha6n"
Case "D9inh Da65u"
Case "Ma65u Tua61t"
Case "Ky3 Ho75i"
Case "Canh Ty1"
Case "Ta6n Su73u"
Case "Nha6m Da62n"
Case "Quy1 Ma4o"
'======40
Case "Gia1p Thi2n"
Case "A61t Ty5"
Case "Bi1nh Ngo5"
Case "D9inh Mu2i"
Case "Ma65u Tha6n"
Case "Ky3 Da65u"
Case "Canh Tua61t"
Case "Ta6n Ho75i"
Case "Nha6m Ty1"
Case "Quy1 Su73u"
'======50
Case "Gia1p Da62n"
Case "A61t Ma4o"
Case "Bi1nh Thi2n"
Case "D9inh Ty5"
Case "Ma65u Ngo5"
Case "Ky3 Mu2i"
Case "Canh Tha6n"
Case "Ta6n Da65u"
Case "Nha6m Tua61t"
Case "Quy1 Ho75i"
'======60
End Select
End With
End Function


