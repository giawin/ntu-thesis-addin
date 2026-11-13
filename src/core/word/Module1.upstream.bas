Attribute VB_Name = "Module1"
Private mTTProfile As String
Private mTTSource As String
Private mTTDetected As String
Private mTTDetectedStrong As Boolean
Private mTTHeaderNote As String
Private mTTForceProfile As String
Private mTTRibbon As Object
Private mTTOvPid As String
Private mTTOvStr As String
Private mTTDaTach As String
Private mTTLuuTuBang As Boolean
Private mTTThuDongNgay As Boolean
Private mTTThuKhongHoi As Boolean
Private Const LIC_CAN_NHAP_MA As Boolean = False
Private Const TT_SETTING_APP As String = "OfficePartySuite"
Private Const TT_SETTING_SECTION As String = "TheThuc"
Private Const TT_DEFAULT_PROFILE As String = "DANG"
Public Function TT_ProfileCount() As Integer
TT_ProfileCount = 5
End Function
Public Function TT_ProfileIdAt(ByVal idx As Integer) As String
Select Case idx
Case 1: TT_ProfileIdAt = "DANG"
Case 2: TT_ProfileIdAt = "CHINH_QUYEN"
Case 3: TT_ProfileIdAt = "CONG_AN"
Case 4: TT_ProfileIdAt = "QUAN_DOI"
Case 5: TT_ProfileIdAt = "MTTQ"
Case Else: TT_ProfileIdAt = ""
End Select
End Function
Public Function TT_ProfileLabelAscii(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_ProfileLabelAscii = "Dang"
Case "CHINH_QUYEN": TT_ProfileLabelAscii = "Chinh quyen"
Case "CONG_AN": TT_ProfileLabelAscii = "Cong an"
Case "QUAN_DOI": TT_ProfileLabelAscii = "Quan doi"
Case "MTTQ": TT_ProfileLabelAscii = "MTTQ va to chuc CT-XH"
Case "AUTO": TT_ProfileLabelAscii = "Tu dong nhan dien"
Case Else: TT_ProfileLabelAscii = pid
End Select
End Function
Public Function TT_LegalBasisAscii(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_LegalBasisAscii = "Quy dinh so 399-QD/TW ngay 09/01/2026 cua Ban Bi thu; Huong dan so 05-HD/VPTW ngay 27/5/2026 cua Van phong Trung uong Dang"
Case "CHINH_QUYEN": TT_LegalBasisAscii = "Nghi dinh so 30/2020/ND-CP ngay 05/3/2020 cua Chinh phu ve cong tac van thu (Phu luc I: the thuc, ky thuat trinh bay van ban hanh chinh)"
Case "CONG_AN": TT_LegalBasisAscii = "Thong tu so 26/2019/TT-BCA ve cong tac van thu trong Cong an nhan dan (hop nhat tai Van ban hop nhat so 08/VBHN-BCA nam 2021); The thuc, ky thuat trinh bay theo Dieu 8, 9 Nghi dinh so 30/2020/ND-CP"
Case "QUAN_DOI": TT_LegalBasisAscii = "The thuc, ky thuat trinh bay theo Nghi dinh so 30/2020/ND-CP; Quy che cong tac van thu trong Bo Quoc phong (Thong tu so 68/2023/TT-BQP ngay 25/9/2023) - chua doi chieu toan van"
Case "MTTQ": TT_LegalBasisAscii = "Quyet dinh so 207/QD-MTTW-BTT ngay 29/8/2025 cua Ban Thuong truc Uy ban Trung uong MTTQ Viet Nam - chua doi chieu toan van; The thuc van ban cua tung to chuc chinh tri - xa hoi do Thuong truc to chuc do o Trung uong quy dinh"
Case Else: TT_LegalBasisAscii = ""
End Select
End Function
Private Function TT_RawN(ByVal pid As String, ByVal key As String) As Single
Select Case pid
Case "DANG"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13.5
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 14
Case "HDR_LEFT_MM": TT_RawN = 72
Case "HDR_RIGHT_MM": TT_RawN = 93
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 14
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 12
Case "TITLE_SIZE": TT_RawN = 15
Case "SUMMARY_SIZE": TT_RawN = 15
Case "BODY_SIZE": TT_RawN = 15
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 14
Case "NN_ITEM_SIZE": TT_RawN = 12
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "CHINH_QUYEN"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 13
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "CONG_AN"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 13
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "QUAN_DOI"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 13
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
Case "MTTQ"
Select Case key
Case "PAGE_TOP": TT_RawN = 20
Case "PAGE_BOTTOM": TT_RawN = 20
Case "PAGE_LEFT": TT_RawN = 30
Case "PAGE_RIGHT": TT_RawN = 15
Case "PAGE_HEADER": TT_RawN = 10
Case "PAGE_FOOTER": TT_RawN = 10
Case "PAGE_NUMBER_SIZE": TT_RawN = 13
Case "PARTY_MOTTO_SIZE": TT_RawN = 15
Case "QH_SIZE": TT_RawN = 13
Case "TN_SIZE": TT_RawN = 14
Case "ORGAN_SIZE": TT_RawN = 13
Case "HDR_LEFT_MM": TT_RawN = 64
Case "HDR_RIGHT_MM": TT_RawN = 101
Case "HDR_LEFT_MIN_MM": TT_RawN = 52
Case "DOC_NUMBER_SIZE": TT_RawN = 13
Case "DATE_SIZE": TT_RawN = 14
Case "CV_SUMMARY_SIZE": TT_RawN = 13
Case "TITLE_SIZE": TT_RawN = 14
Case "SUMMARY_SIZE": TT_RawN = 14
Case "BODY_SIZE": TT_RawN = 14
Case "INDENT_CM": TT_RawN = 1.27
Case "NN_LABEL_SIZE": TT_RawN = 12
Case "NN_ITEM_SIZE": TT_RawN = 11
Case "SIGN_SIZE": TT_RawN = 14
End Select
End Select
End Function
Private Function TT_RawB(ByVal pid As String, ByVal key As String) As Boolean
Select Case pid
Case "DANG"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = False
Case "CV_SUMMARY_ITALIC": TT_RawB = True
Case "LEGAL_BASIS_ITALIC": TT_RawB = False
Case "KINH_GUI_ITALIC": TT_RawB = True
Case "NN_LABEL_UNDERLINE": TT_RawB = True
Case "NN_LABEL_BOLD": TT_RawB = False
Case "NN_LABEL_ITALIC": TT_RawB = False
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = False
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = True
End Select
Case "CHINH_QUYEN"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "CONG_AN"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = True
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "QUAN_DOI"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = True
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = True
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = True
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
Case "MTTQ"
Select Case key
Case "AUTO_INSERT_MOTTO": TT_RawB = False
Case "ORGAN_UNDERLINE": TT_RawB = True
Case "CV_SUMMARY_ITALIC": TT_RawB = False
Case "LEGAL_BASIS_ITALIC": TT_RawB = False
Case "KINH_GUI_ITALIC": TT_RawB = False
Case "NN_LABEL_UNDERLINE": TT_RawB = False
Case "NN_LABEL_BOLD": TT_RawB = True
Case "NN_LABEL_ITALIC": TT_RawB = True
Case "ROLE_UNDER_AUTH_BOLD": TT_RawB = True
Case "RANK_BEFORE_NAME": TT_RawB = False
Case "CHECK_COLLECTIVE": TT_RawB = False
End Select
End Select
End Function
Private Function TT_RawS(ByVal pid As String, ByVal key As String) As String
Select Case pid
Case "DANG"
Select Case key
Case "MOTTO_KIND": TT_RawS = "DANG"
Case "RIGHT_ALIGN": TT_RawS = "RIGHT"
Case "DOC_NUMBER_STYLE": TT_RawS = "DANG"
Case "SIGN_PREFIX_STYLE": TT_RawS = "SLASH"
Case "SUMMARY_RULE": TT_RawS = "DASH"
Case "STATUS": TT_RawS = "official"
End Select
Case "CHINH_QUYEN"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "SUMMARY_RULE": TT_RawS = "LINE"
Case "STATUS": TT_RawS = "official"
End Select
Case "CONG_AN"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "SUMMARY_RULE": TT_RawS = "LINE"
Case "STATUS": TT_RawS = "official"
End Select
Case "QUAN_DOI"
Select Case key
Case "MOTTO_KIND": TT_RawS = "QUOC_HIEU"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "NHA_NUOC"
Case "SIGN_PREFIX_STYLE": TT_RawS = "DOT"
Case "SUMMARY_RULE": TT_RawS = "LINE"
Case "STATUS": TT_RawS = "needs_review"
End Select
Case "MTTQ"
Select Case key
Case "MOTTO_KIND": TT_RawS = "LINH_HOAT"
Case "RIGHT_ALIGN": TT_RawS = "CENTER"
Case "DOC_NUMBER_STYLE": TT_RawS = "LINH_HOAT"
Case "SIGN_PREFIX_STYLE": TT_RawS = "GIU_NGUYEN"
Case "SUMMARY_RULE": TT_RawS = "GIU_NGUYEN"
Case "STATUS": TT_RawS = "experimental"
End Select
End Select
End Function
Private Function TT_RawDocCodeTitle(ByVal pid As String, ByVal codeAsc As String) As String
Select Case pid
Case "DANG"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "CV": TT_RawDocCodeTitle = "CONG VAN"
Case "DA": TT_RawDocCodeTitle = "DU AN|DE AN"
Case "DL": TT_RawDocCodeTitle = "DIEU LE"
Case "GCN": TT_RawDocCodeTitle = "GIAY CHUNG NHAN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "KL": TT_RawDocCodeTitle = "KET LUAN"
Case "LKG": TT_RawDocCodeTitle = "LOI KEU GOI"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH|QUY DINH"
Case "QH": TT_RawDocCodeTitle = "QUY HOACH"
Case "TB": TT_RawDocCodeTitle = "THONG BAO|TUYEN BO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TT": TT_RawDocCodeTitle = "THONG TRI"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "CHINH_QUYEN"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "CONG_AN"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "QUAN_DOI"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
Case "MTTQ"
Select Case codeAsc
Case "BB": TT_RawDocCodeTitle = "BIEN BAN"
Case "BC": TT_RawDocCodeTitle = "BAO CAO"
Case "BGN": TT_RawDocCodeTitle = "BAN GHI NHO"
Case "BTT": TT_RawDocCodeTitle = "BAN THOA THUAN"
Case "CD": TT_RawDocCodeTitle = "CONG DIEN"
Case "CT": TT_RawDocCodeTitle = "CHI THI"
Case "CTR": TT_RawDocCodeTitle = "CHUONG TRINH"
Case "DA": TT_RawDocCodeTitle = "DE AN|DU AN"
Case "GGT": TT_RawDocCodeTitle = "GIAY GIOI THIEU"
Case "GM": TT_RawDocCodeTitle = "GIAY MOI"
Case "GNP": TT_RawDocCodeTitle = "GIAY NGHI PHEP"
Case "GUQ": TT_RawDocCodeTitle = "GIAY UY QUYEN"
Case "HD": TT_RawDocCodeTitle = "HUONG DAN|HOP DONG"
Case "KH": TT_RawDocCodeTitle = "KE HOACH"
Case "NQ": TT_RawDocCodeTitle = "NGHI QUYET"
Case "PA": TT_RawDocCodeTitle = "PHUONG AN"
Case "PB": TT_RawDocCodeTitle = "PHIEU BAO"
Case "PC": TT_RawDocCodeTitle = "PHIEU CHUYEN"
Case "PG": TT_RawDocCodeTitle = "PHIEU GUI"
Case "QC": TT_RawDocCodeTitle = "QUY CHE"
Case "QD": TT_RawDocCodeTitle = "QUYET DINH"
Case "QYD": TT_RawDocCodeTitle = "QUY DINH"
Case "SL": TT_RawDocCodeTitle = "BAN SAO LUC"
Case "SY": TT_RawDocCodeTitle = "BAN SAO Y"
Case "TB": TT_RawDocCodeTitle = "THONG BAO"
Case "TC": TT_RawDocCodeTitle = "THONG CAO"
Case "TRS": TT_RawDocCodeTitle = "BAN TRICH SAO"
Case "TTR": TT_RawDocCodeTitle = "TO TRINH"
End Select
End Select
End Function
Public Function TT_DetectFromHeaderText(ByVal sAsc As String, ByRef isStrong As Boolean) As String
Dim t As String: t = Replace(UCase(sAsc), " ", "")
isStrong = False
TT_DetectFromHeaderText = ""
If InStr(1, t, "DANGCONGSANVIETNAM") > 0 Then
isStrong = True
TT_DetectFromHeaderText = "DANG"
Exit Function
End If
Dim hasNational As Boolean: hasNational = False
If InStr(1, t, "CONGHOAXAHOICHUNGHIA") > 0 Or InStr(1, t, "DOCLAP-TUDO") > 0 Then hasNational = True
If Not hasNational Then
If InStr(1, t, "DANGUY") > 0 Or InStr(1, t, "DANGBO") > 0 Or InStr(1, t, "CHIBO") > 0 Or InStr(1, t, "TINHUY") > 0 Or InStr(1, t, "THANHUY") > 0 Or InStr(1, t, "BANTHUONGVU") > 0 Or InStr(1, t, "BANCHAPHANHDANGBO") > 0 Then
TT_DetectFromHeaderText = "DANG"
Exit Function
End If
End If
If InStr(1, t, "BOCONGAN") > 0 Or InStr(1, t, "CONGANTINH") > 0 Or InStr(1, t, "CONGANTHANHPHO") > 0 Or InStr(1, t, "CONGANXA") > 0 Or InStr(1, t, "CONGANPHUONG") > 0 Or InStr(1, t, "CONGANDACKHU") > 0 Or InStr(1, t, "CONGANNHANDAN") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "CONG_AN"
Exit Function
End If
If InStr(1, t, "BOQUOCPHONG") > 0 Or InStr(1, t, "QUANKHU") > 0 Or InStr(1, t, "QUANDOAN") > 0 Or InStr(1, t, "BOTULENH") > 0 Or InStr(1, t, "BOCHIHUYQUANSU") > 0 Or InStr(1, t, "BANCHIHUYQUANSU") > 0 Or InStr(1, t, "SUDOAN") > 0 Or InStr(1, t, "LUDOAN") > 0 Or InStr(1, t, "TRUNGDOAN") > 0 Or InStr(1, t, "BIENPHONG") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "QUAN_DOI"
Exit Function
End If
If InStr(1, t, "MATTRANTOQUOC") > 0 Or InStr(1, t, "UYBANMATTRAN") > 0 Or InStr(1, t, "LIENDOANLAODONG") > 0 Or InStr(1, t, "CONGDOAN") > 0 Or InStr(1, t, "HOINONGDAN") > 0 Or InStr(1, t, "HOILIENHIEPPHUNU") > 0 Or InStr(1, t, "HOICUUCHIENBINH") > 0 Or InStr(1, t, "DOANTNCS") > 0 Or InStr(1, t, "DOANTHANHNIEN") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "MTTQ"
Exit Function
End If
If InStr(1, t, "UYBANNHANDAN") > 0 Or InStr(1, t, "HOIDONGNHANDAN") > 0 Or InStr(1, t, "UBND") > 0 Or InStr(1, t, "HDND") > 0 Then
isStrong = hasNational
TT_DetectFromHeaderText = "CHINH_QUYEN"
Exit Function
End If
If hasNational Then
isStrong = True
TT_DetectFromHeaderText = "CHINH_QUYEN"
End If
End Function
Public Function TT_TextQuocHieu() As String
TT_TextQuocHieu = "C" & ChrW(7896) & "NG H" & ChrW(210) & "A X" & ChrW(195) & " H" & ChrW(7896) & "I CH" & ChrW(7910) & " NGH" & ChrW(296) & "A VI" & ChrW(7878) & "T NAM"
End Function
Public Function TT_TextTieuNgu() As String
TT_TextTieuNgu = ChrW(272) & ChrW(7897) & "c l" & ChrW(7853) & "p - T" & ChrW(7921) & " do - H" & ChrW(7841) & "nh ph" & ChrW(250) & "c"
End Function
Public Function TT_MauCount(ByVal grp As String) As Integer
Select Case grp
Case "DANG": TT_MauCount = 18
Case "NHA_NUOC": TT_MauCount = 20
Case "MTTQ": TT_MauCount = 12
Case Else: TT_MauCount = 0
End Select
End Function
Public Function TT_MauLabelAscii(ByVal grp As String, ByVal idx As Integer) As String
TT_MauLabelAscii = ""
Select Case grp
Case "DANG"
Select Case idx
Case 1: TT_MauLabelAscii = "Nghi quyet"
Case 2: TT_MauLabelAscii = "Quyet dinh"
Case 3: TT_MauLabelAscii = "Ke hoach"
Case 4: TT_MauLabelAscii = "Bao cao"
Case 5: TT_MauLabelAscii = "Thong bao"
Case 6: TT_MauLabelAscii = "To trinh"
Case 7: TT_MauLabelAscii = "Cong van"
Case 8: TT_MauLabelAscii = "Giay moi"
Case 9: TT_MauLabelAscii = "Bien ban"
Case 10: TT_MauLabelAscii = "Quy dinh"
Case 11: TT_MauLabelAscii = "Quy che"
Case 12: TT_MauLabelAscii = "Chi thi"
Case 13: TT_MauLabelAscii = "Ket luan"
Case 14: TT_MauLabelAscii = "Huong dan"
Case 15: TT_MauLabelAscii = "Chuong trinh"
Case 16: TT_MauLabelAscii = "Thong tri"
Case 17: TT_MauLabelAscii = "De an"
Case 18: TT_MauLabelAscii = "Giay gioi thieu"
End Select
Case "NHA_NUOC"
Select Case idx
Case 1: TT_MauLabelAscii = "Quyet dinh"
Case 2: TT_MauLabelAscii = "Ke hoach"
Case 3: TT_MauLabelAscii = "Bao cao"
Case 4: TT_MauLabelAscii = "Thong bao"
Case 5: TT_MauLabelAscii = "To trinh"
Case 6: TT_MauLabelAscii = "Cong van"
Case 7: TT_MauLabelAscii = "Giay moi"
Case 8: TT_MauLabelAscii = "Bien ban"
Case 9: TT_MauLabelAscii = "Quyet dinh ban hanh, phe duyet van ban"
Case 10: TT_MauLabelAscii = "Quy dinh kem theo quyet dinh"
Case 11: TT_MauLabelAscii = "Quy che kem theo quyet dinh"
Case 12: TT_MauLabelAscii = "Huong dan"
Case 13: TT_MauLabelAscii = "Chuong trinh"
Case 14: TT_MauLabelAscii = "Quyet dinh uy quyen"
Case 15: TT_MauLabelAscii = "Giay gioi thieu"
Case 16: TT_MauLabelAscii = "Cong dien"
Case 17: TT_MauLabelAscii = "Phieu chuyen"
Case 18: TT_MauLabelAscii = "Nghi quyet (ca biet)"
Case 19: TT_MauLabelAscii = "Chi thi"
Case 20: TT_MauLabelAscii = "Giay nghi phep"
End Select
Case "MTTQ"
Select Case idx
Case 1: TT_MauLabelAscii = "Ke hoach"
Case 2: TT_MauLabelAscii = "Bao cao"
Case 3: TT_MauLabelAscii = "Thong bao"
Case 4: TT_MauLabelAscii = "Cong van"
Case 5: TT_MauLabelAscii = "Giay moi"
Case 6: TT_MauLabelAscii = "To trinh"
Case 7: TT_MauLabelAscii = "Quyet dinh"
Case 8: TT_MauLabelAscii = "Chuong trinh phoi hop"
Case 9: TT_MauLabelAscii = "Huong dan"
Case 10: TT_MauLabelAscii = "Nghi quyet hoi nghi"
Case 11: TT_MauLabelAscii = "Bien ban"
Case 12: TT_MauLabelAscii = "Giay gioi thieu"
End Select
End Select
End Function
Public Function TT_MauData(ByVal grp As String, ByVal idx As Integer) As String
TT_MauData = ""
Select Case grp
Case "DANG"
Select Case idx
Case 1: TT_MauData = TT_Mau_DANG_1()
Case 2: TT_MauData = TT_Mau_DANG_2()
Case 3: TT_MauData = TT_Mau_DANG_3()
Case 4: TT_MauData = TT_Mau_DANG_4()
Case 5: TT_MauData = TT_Mau_DANG_5()
Case 6: TT_MauData = TT_Mau_DANG_6()
Case 7: TT_MauData = TT_Mau_DANG_7()
Case 8: TT_MauData = TT_Mau_DANG_8()
Case 9: TT_MauData = TT_Mau_DANG_9()
Case 10: TT_MauData = TT_Mau_DANG_10()
Case 11: TT_MauData = TT_Mau_DANG_11()
Case 12: TT_MauData = TT_Mau_DANG_12()
Case 13: TT_MauData = TT_Mau_DANG_13()
Case 14: TT_MauData = TT_Mau_DANG_14()
Case 15: TT_MauData = TT_Mau_DANG_15()
Case 16: TT_MauData = TT_Mau_DANG_16()
Case 17: TT_MauData = TT_Mau_DANG_17()
Case 18: TT_MauData = TT_Mau_DANG_18()
End Select
Case "NHA_NUOC"
Select Case idx
Case 1: TT_MauData = TT_Mau_NHA_NUOC_1()
Case 2: TT_MauData = TT_Mau_NHA_NUOC_2()
Case 3: TT_MauData = TT_Mau_NHA_NUOC_3()
Case 4: TT_MauData = TT_Mau_NHA_NUOC_4()
Case 5: TT_MauData = TT_Mau_NHA_NUOC_5()
Case 6: TT_MauData = TT_Mau_NHA_NUOC_6()
Case 7: TT_MauData = TT_Mau_NHA_NUOC_7()
Case 8: TT_MauData = TT_Mau_NHA_NUOC_8()
Case 9: TT_MauData = TT_Mau_NHA_NUOC_9()
Case 10: TT_MauData = TT_Mau_NHA_NUOC_10()
Case 11: TT_MauData = TT_Mau_NHA_NUOC_11()
Case 12: TT_MauData = TT_Mau_NHA_NUOC_12()
Case 13: TT_MauData = TT_Mau_NHA_NUOC_13()
Case 14: TT_MauData = TT_Mau_NHA_NUOC_14()
Case 15: TT_MauData = TT_Mau_NHA_NUOC_15()
Case 16: TT_MauData = TT_Mau_NHA_NUOC_16()
Case 17: TT_MauData = TT_Mau_NHA_NUOC_17()
Case 18: TT_MauData = TT_Mau_NHA_NUOC_18()
Case 19: TT_MauData = TT_Mau_NHA_NUOC_19()
Case 20: TT_MauData = TT_Mau_NHA_NUOC_20()
End Select
Case "MTTQ"
Select Case idx
Case 1: TT_MauData = TT_Mau_MTTQ_1()
Case 2: TT_MauData = TT_Mau_MTTQ_2()
Case 3: TT_MauData = TT_Mau_MTTQ_3()
Case 4: TT_MauData = TT_Mau_MTTQ_4()
Case 5: TT_MauData = TT_Mau_MTTQ_5()
Case 6: TT_MauData = TT_Mau_MTTQ_6()
Case 7: TT_MauData = TT_Mau_MTTQ_7()
Case 8: TT_MauData = TT_Mau_MTTQ_8()
Case 9: TT_MauData = TT_Mau_MTTQ_9()
Case 10: TT_MauData = TT_Mau_MTTQ_10()
Case 11: TT_MauData = TT_Mau_MTTQ_11()
Case 12: TT_MauData = TT_Mau_MTTQ_12()
End Select
End Select
End Function
Private Function TT_Mau_DANG_1() As String
Dim s As String
s = ""
s = s & "CODE|" & "NQ" & vbLf
s = s & "L|" & "NGH" & ChrW(7882) & " QUY" & ChrW(7870) & "T" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[Ph" & ChrW(7847) & "n m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u: c" & ChrW(259) & "n c" & ChrW(7913) & ", b" & ChrW(7889) & "i c" & ChrW(7843) & "nh ban h" & ChrW(224) & "nh ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & "]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [M" & ChrW(7908) & "C TI" & ChrW(202) & "U, NHI" & ChrW(7878) & "M V" & ChrW(7908) & ", GI" & ChrW(7842) & "I PH" & ChrW(193) & "P]" & vbLf
s = s & "L|" & "1. [M" & ChrW(7909) & "c ti" & ChrW(234) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", gi" & ChrW(7843) & "i ph" & ChrW(225) & "p]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_1 = s
End Function
Private Function TT_Mau_DANG_2() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " vi" & ChrW(7879) & "c ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [" & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " " & ChrW(272) & ChrW(7843) & "ng; quy ch" & ChrW(7871) & " l" & ChrW(224) & "m vi" & ChrW(7879) & "c; v" & ChrW(259) & "n b" & ChrW(7843) & "n li" & ChrW(234) & "n quan];" & vbLf
s = s & "L|" & "- X" & ChrW(233) & "t " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_2 = s
End Function
Private Function TT_Mau_DANG_3() As String
Dim s As String
s = ""
s = s & "CODE|" & "KH" & vbLf
s = s & "L|" & "K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung k" & ChrW(7871) & " ho" & ChrW(7841) & "ch]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(259) & "n c" & ChrW(7913) & ", m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u k" & ChrW(7871) & " ho" & ChrW(7841) & "ch.]" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "1. M" & ChrW(7909) & "c " & ChrW(273) & ChrW(237) & "ch" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. Y" & ChrW(234) & "u c" & ChrW(7847) & "u" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_3 = s
End Function
Private Function TT_Mau_DANG_4() As String
Dim s As String
s = ""
s = s & "CODE|" & "BC" & vbLf
s = s & "L|" & "B" & ChrW(193) & "O C" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung b" & ChrW(225) & "o c" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[M" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u b" & ChrW(225) & "o c" & ChrW(225) & "o.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & " TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [H" & ChrW(7840) & "N CH" & ChrW(7870) & ", NGUY" & ChrW(202) & "N NH" & ChrW(194) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. [PH" & ChrW(431) & ChrW(416) & "NG H" & ChrW(431) & ChrW(7898) & "NG, NHI" & ChrW(7878) & "M V" & ChrW(7908) & " TH" & ChrW(7900) & "I GIAN T" & ChrW(7898) & "I]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "IV. [KI" & ChrW(7870) & "N NGH" & ChrW(7882) & ", " & ChrW(272) & ChrW(7872) & " XU" & ChrW(7844) & "T]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_4 = s
End Function
Private Function TT_Mau_DANG_5() As String
Dim s As String
s = ""
s = s & "CODE|" & "TB" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG B" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_5 = s
End Function
Private Function TT_Mau_DANG_6() As String
Dim s As String
s = ""
s = s & "CODE|" & "TTr" & vbLf
s = s & "L|" & "T" & ChrW(7900) & " TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n t" & ChrW(7901) & " tr" & ChrW(236) & "nh]." & vbLf
s = s & "L|" & "[L" & ChrW(253) & " do, c" & ChrW(259) & "n c" & ChrW(7913) & " tr" & ChrW(236) & "nh.]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & ".]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c] xem x" & ChrW(233) & "t, quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_6 = s
End Function
Private Function TT_Mau_DANG_7() As String
Dim s As String
s = ""
s = s & "CODE|" & "CV" & vbLf
s = s & "CV|1" & vbLf
s = s & "SUM|" & "V/v [tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n c" & ChrW(244) & "ng v" & ChrW(259) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n.]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_7 = s
End Function
Private Function TT_Mau_DANG_8() As String
Dim s As String
s = ""
s = s & "CODE|" & "GM" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y M" & ChrW(7900) & "I" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: d" & ChrW(7921) & " h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng k" & ChrW(237) & "nh m" & ChrW(7901) & "i: [" & ChrW(272) & ChrW(7841) & "i bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7901) & "i]." & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [n" & ChrW(7897) & "i dung]." & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " chu" & ChrW(7849) & "n b" & ChrW(7883) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u, li" & ChrW(234) & "n h" & ChrW(7879) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ").]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_8 = s
End Function
Private Function TT_Mau_DANG_9() As String
Dim s As String
s = ""
s = s & "CODE|" & "BB" & vbLf
s = s & "L|" & "BI" & ChrW(202) & "N B" & ChrW(7842) & "N" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "Th" & ChrW(224) & "nh ph" & ChrW(7847) & "n tham d" & ChrW(7921) & ": [th" & ChrW(224) & "nh ph" & ChrW(7847) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & ": [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(432) & " k" & ChrW(253) & " (ng" & ChrW(432) & ChrW(7901) & "i ghi bi" & ChrW(234) & "n b" & ChrW(7843) & "n): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [theo di" & ChrW(7877) & "n bi" & ChrW(7871) & "n cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "Cu" & ChrW(7897) & "c h" & ChrW(7885) & "p k" & ChrW(7871) & "t th" & ChrW(250) & "c v" & ChrW(224) & "o [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " TR" & ChrW(204) & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
s = s & "LR|" & "TH" & ChrW(431) & " K" & ChrW(221) & vbLf
s = s & "LH|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_9 = s
End Function
Private Function TT_Mau_DANG_10() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " " & ChrW(272) & ChrW(7843) & "ng;" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [quy ch" & ChrW(7871) & " l" & ChrW(224) & "m vi" & ChrW(7879) & "c, v" & ChrW(259) & "n b" & ChrW(7843) & "n li" & ChrW(234) & "n quan]." & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " [c" & ChrW(7845) & "p " & ChrW(7911) & "y] quy " & ChrW(273) & ChrW(7883) & "nh [...] nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng I" & vbLf
s = s & "L|" & "QUY " & ChrW(272) & ChrW(7882) & "NH CHUNG" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. Ph" & ChrW(7841) & "m vi " & ChrW(273) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh, " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng II" & vbLf
s = s & "L|" & "[N" & ChrW(7896) & "I DUNG QUY " & ChrW(272) & ChrW(7882) & "NH]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [T" & ChrW(234) & "n " & ChrW(273) & "i" & ChrW(7873) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng III" & vbLf
s = s & "L|" & "T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. Tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 4. Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c thi h" & ChrW(224) & "nh" & vbLf
s = s & "L|" & "Quy " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y c" & ChrW(243) & " hi" & ChrW(7879) & "u l" & ChrW(7921) & "c t" & ChrW(7915) & " ng" & ChrW(224) & "y k" & ChrW(253) & "." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_10 = s
End Function
Private Function TT_Mau_DANG_11() As String
Dim s As String
s = ""
s = s & "CODE|" & "QC" & vbLf
s = s & "L|" & "QUY CH" & ChrW(7870) & vbLf
s = s & "L|" & "[l" & ChrW(224) & "m vi" & ChrW(7879) & "c c" & ChrW(7911) & "a ... nhi" & ChrW(7879) & "m k" & ChrW(7923) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " " & ChrW(272) & ChrW(7843) & "ng;" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n li" & ChrW(234) & "n quan]." & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng I" & vbLf
s = s & "L|" & "NH" & ChrW(7918) & "NG QUY " & ChrW(272) & ChrW(7882) & "NH CHUNG" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [Ph" & ChrW(7841) & "m vi, " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng II" & vbLf
s = s & "L|" & "[TR" & ChrW(193) & "CH NHI" & ChrW(7878) & "M, QUY" & ChrW(7872) & "N H" & ChrW(7840) & "N, CH" & ChrW(7870) & " " & ChrW(272) & ChrW(7896) & " L" & ChrW(192) & "M VI" & ChrW(7878) & "C]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [T" & ChrW(234) & "n " & ChrW(273) & "i" & ChrW(7873) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng III" & vbLf
s = s & "L|" & ChrW(272) & "I" & ChrW(7872) & "U KHO" & ChrW(7842) & "N THI H" & ChrW(192) & "NH" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c, tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_11 = s
End Function
Private Function TT_Mau_DANG_12() As String
Dim s As String
s = ""
s = s & "CODE|" & "CT" & vbLf
s = s & "L|" & "CH" & ChrW(7880) & " TH" & ChrW(7882) & vbLf
s = s & "L|" & "[c" & ChrW(7911) & "a Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " ... v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(225) & "nh gi" & ChrW(225) & " t" & ChrW(236) & "nh h" & ChrW(236) & "nh, n" & ChrW(234) & "u y" & ChrW(234) & "u c" & ChrW(7847) & "u c" & ChrW(7847) & "n ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7841) & "o.]" & vbLf
s = s & "L|" & "[Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " ... y" & ChrW(234) & "u c" & ChrW(7847) & "u c" & ChrW(225) & "c c" & ChrW(7845) & "p " & ChrW(7911) & "y, t" & ChrW(7893) & " ch" & ChrW(7913) & "c " & ChrW(273) & ChrW(7843) & "ng th" & ChrW(7921) & "c hi" & ChrW(7879) & "n t" & ChrW(7889) & "t c" & ChrW(225) & "c nhi" & ChrW(7879) & "m v" & ChrW(7909) & " sau:]" & vbLf
s = s & "L|" & "1. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", gi" & ChrW(7843) & "i ph" & ChrW(225) & "p.]" & vbLf
s = s & "L|" & "2. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", gi" & ChrW(7843) & "i ph" & ChrW(225) & "p.]" & vbLf
s = s & "L|" & "3. [T" & ChrW(7893) & " ch" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "L|" & "[Ch" & ChrW(7881) & " th" & ChrW(7883) & " n" & ChrW(224) & "y " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ph" & ChrW(7893) & " bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7871) & "n ...]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_12 = s
End Function
Private Function TT_Mau_DANG_13() As String
Dim s As String
s = ""
s = s & "CODE|" & "KL" & vbLf
s = s & "L|" & "K" & ChrW(7870) & "T LU" & ChrW(7852) & "N" & vbLf
s = s & "L|" & "[c" & ChrW(7911) & "a Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " ... v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[Ng" & ChrW(224) & "y ... th" & ChrW(225) & "ng ... n" & ChrW(259) & "m ..., Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " ... h" & ChrW(7885) & "p, nghe ... b" & ChrW(225) & "o c" & ChrW(225) & "o v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "[Sau khi th" & ChrW(7843) & "o lu" & ChrW(7853) & "n, Ban Th" & ChrW(432) & ChrW(7901) & "ng v" & ChrW(7909) & " ... k" & ChrW(7871) & "t lu" & ChrW(7853) & "n nh" & ChrW(432) & " sau:]" & vbLf
s = s & "L|" & "1. [N" & ChrW(7897) & "i dung k" & ChrW(7871) & "t lu" & ChrW(7853) & "n.]" & vbLf
s = s & "L|" & "2. [N" & ChrW(7897) & "i dung k" & ChrW(7871) & "t lu" & ChrW(7853) & "n.]" & vbLf
s = s & "L|" & "3. [T" & ChrW(7893) & " ch" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_13 = s
End Function
Private Function TT_Mau_DANG_14() As String
Dim s As String
s = ""
s = s & "CODE|" & "HD" & vbLf
s = s & "L|" & "H" & ChrW(431) & ChrW(7898) & "NG D" & ChrW(7850) & "N" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "- C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n [...] nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "I. [M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [N" & ChrW(7896) & "I DUNG H" & ChrW(431) & ChrW(7898) & "NG D" & ChrW(7850) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_14 = s
End Function
Private Function TT_Mau_DANG_15() As String
Dim s As String
s = ""
s = s & "CODE|" & "CTr" & vbLf
s = s & "L|" & "CH" & ChrW(431) & ChrW(416) & "NG TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "[h" & ChrW(224) & "nh " & ChrW(273) & ChrW(7897) & "ng / c" & ChrW(244) & "ng t" & ChrW(225) & "c ... n" & ChrW(259) & "m ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(259) & "n c" & ChrW(7913) & ", m" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh.]" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_15 = s
End Function
Private Function TT_Mau_DANG_16() As String
Dim s As String
s = ""
s = s & "CODE|" & "TT" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG TRI" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng tri: ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7841) & "o, gi" & ChrW(7843) & "i th" & ChrW(237) & "ch, h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n th" & ChrW(7921) & "c hi" & ChrW(7879) & "n ...]" & vbLf
s = s & "L|" & "1. [N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. [N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "3. [T" & ChrW(7893) & " ch" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_16 = s
End Function
Private Function TT_Mau_DANG_17() As String
Dim s As String
s = ""
s = s & "CODE|" & ChrW(272) & "A" & vbLf
s = s & "L|" & ChrW(272) & ChrW(7872) & " " & ChrW(193) & "N" & vbLf
s = s & "L|" & "[v" & ChrW(7873) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "I. S" & ChrW(7920) & " C" & ChrW(7846) & "N THI" & ChrW(7870) & "T, C" & ChrW(258) & "N C" & ChrW(7912) & " X" & ChrW(194) & "Y D" & ChrW(7920) & "NG " & ChrW(272) & ChrW(7872) & " " & ChrW(193) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. M" & ChrW(7908) & "C TI" & ChrW(202) & "U, NHI" & ChrW(7878) & "M V" & ChrW(7908) & ", GI" & ChrW(7842) & "I PH" & ChrW(193) & "P" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_17 = s
End Function
Private Function TT_Mau_DANG_18() As String
Dim s As String
s = ""
s = s & "CODE|" & "GGT" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y GI" & ChrW(7898) & "I THI" & ChrW(7878) & "U" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u:" & vbLf
s = s & "L|" & ChrW(272) & ChrW(7891) & "ng ch" & ChrW(237) & ": [h" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7913) & "c v" & ChrW(7909) & ": [ch" & ChrW(7913) & "c v" & ChrW(7909) & ", " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " c" & ChrW(244) & "ng t" & ChrW(225) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7917) & " " & ChrW(273) & ChrW(7871) & "n: [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c " & ChrW(273) & ChrW(432) & ChrW(7907) & "c gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u t" & ChrW(7899) & "i l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c: [n" & ChrW(7897) & "i dung l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c] gi" & ChrW(250) & "p " & ChrW(273) & ChrW(7905) & " " & ChrW(273) & ChrW(7891) & "ng ch" & ChrW(237) & " c" & ChrW(243) & " t" & ChrW(234) & "n tr" & ChrW(234) & "n ho" & ChrW(224) & "n th" & ChrW(224) & "nh nhi" & ChrW(7879) & "m v" & ChrW(7909) & "." & vbLf
s = s & "L|" & "Gi" & ChrW(7845) & "y n" & ChrW(224) & "y c" & ChrW(243) & " gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(273) & ChrW(7871) & "n h" & ChrW(7871) & "t ng" & ChrW(224) & "y [...]." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u V" & ChrW(259) & "n ph" & ChrW(242) & "ng [c" & ChrW(7845) & "p " & ChrW(7911) & "y]." & vbLf
s = s & "A|" & "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & vbLf
s = s & "R|" & "[B" & ChrW(205) & " TH" & ChrW(431) & " / PH" & ChrW(211) & " B" & ChrW(205) & " TH" & ChrW(431) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_DANG_18 = s
End Function
Private Function TT_Mau_NHA_NUOC_1() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[TH" & ChrW(7848) & "M QUY" & ChrW(7872) & "N BAN H" & ChrW(192) & "NH, V" & ChrW(205) & " D" & ChrW(7908) & ": CH" & ChrW(7910) & " T" & ChrW(7882) & "CH " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh th" & ChrW(7849) & "m quy" & ChrW(7873) & "n, ch" & ChrW(7913) & "c n" & ChrW(259) & "ng, nhi" & ChrW(7879) & "m v" & ChrW(7909) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(7897) & "i dung, c" & ChrW(417) & " s" & ChrW(7903) & " ban h" & ChrW(224) & "nh];" & vbLf
s = s & "L|" & "Theo " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c thi h" & ChrW(224) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_1 = s
End Function
Private Function TT_Mau_NHA_NUOC_2() As String
Dim s As String
s = ""
s = s & "CODE|" & "KH" & vbLf
s = s & "L|" & "K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] ban h" & ChrW(224) & "nh K" & ChrW(7871) & " ho" & ChrW(7841) & "ch v" & ChrW(7873) & " vi" & ChrW(7879) & "c [...] v" & ChrW(7899) & "i c" & ChrW(225) & "c n" & ChrW(7897) & "i dung nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "1. M" & ChrW(7909) & "c " & ChrW(273) & ChrW(237) & "ch" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. Y" & ChrW(234) & "u c" & ChrW(7847) & "u" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_2 = s
End Function
Private Function TT_Mau_NHA_NUOC_3() As String
Dim s As String
s = ""
s = s & "CODE|" & "BC" & vbLf
s = s & "L|" & "B" & ChrW(193) & "O C" & ChrW(193) & "O" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[M" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u b" & ChrW(225) & "o c" & ChrW(225) & "o.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & " TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [H" & ChrW(7840) & "N CH" & ChrW(7870) & ", NGUY" & ChrW(202) & "N NH" & ChrW(194) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. [PH" & ChrW(431) & ChrW(416) & "NG H" & ChrW(431) & ChrW(7898) & "NG, NHI" & ChrW(7878) & "M V" & ChrW(7908) & " TH" & ChrW(7900) & "I GIAN T" & ChrW(7898) & "I]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "IV. [KI" & ChrW(7870) & "N NGH" & ChrW(7882) & ", " & ChrW(272) & ChrW(7872) & " XU" & ChrW(7844) & "T]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_3 = s
End Function
Private Function TT_Mau_NHA_NUOC_4() As String
Dim s As String
s = ""
s = s & "CODE|" & "TB" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG B" & ChrW(193) & "O" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_4 = s
End Function
Private Function TT_Mau_NHA_NUOC_5() As String
Dim s As String
s = ""
s = s & "CODE|" & "TTr" & vbLf
s = s & "L|" & "T" & ChrW(7900) & " TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(236) & "nh [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c, c" & ChrW(225) & " nh" & ChrW(226) & "n nh" & ChrW(7853) & "n t" & ChrW(7901) & " tr" & ChrW(236) & "nh] v" & ChrW(7873) & " vi" & ChrW(7879) & "c [...], v" & ChrW(7899) & "i c" & ChrW(225) & "c n" & ChrW(7897) & "i dung nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung t" & ChrW(7901) & " tr" & ChrW(236) & "nh.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_5 = s
End Function
Private Function TT_Mau_NHA_NUOC_6() As String
Dim s As String
s = ""
s = s & "CODE|" & vbLf
s = s & "CV|1" & vbLf
s = s & "SUM|" & "V/v [tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c, c" & ChrW(225) & " nh" & ChrW(226) & "n nh" & ChrW(7853) & "n c" & ChrW(244) & "ng v" & ChrW(259) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n.]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_6 = s
End Function
Private Function TT_Mau_NHA_NUOC_7() As String
Dim s As String
s = ""
s = s & "CODE|" & "GM" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y M" & ChrW(7900) & "I" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng k" & ChrW(237) & "nh m" & ChrW(7901) & "i: [" & ChrW(272) & ChrW(7841) & "i bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7901) & "i]." & vbLf
s = s & "L|" & "T" & ChrW(7899) & "i d" & ChrW(7921) & ": [cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", h" & ChrW(7897) & "i th" & ChrW(7843) & "o, l" & ChrW(7877) & " ...]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & ": [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & " ng" & ChrW(432) & ChrW(7901) & "i ch" & ChrW(7911) & " tr" & ChrW(236) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "[Nh" & ChrW(7919) & "ng " & ChrW(273) & "i" & ChrW(7873) & "u c" & ChrW(7847) & "n l" & ChrW(432) & "u " & ChrW(253) & ": chu" & ChrW(7849) & "n b" & ChrW(7883) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u, n" & ChrW(7897) & "i dung ph" & ChrW(225) & "t bi" & ChrW(7875) & "u, trang ph" & ChrW(7909) & "c (n" & ChrW(7871) & "u c" & ChrW(243) & ").]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_7 = s
End Function
Private Function TT_Mau_NHA_NUOC_8() As String
Dim s As String
s = ""
s = s & "CODE|" & "BB" & vbLf
s = s & "ND|1" & vbLf
s = s & "L|" & "BI" & ChrW(202) & "N B" & ChrW(7842) & "N" & vbLf
s = s & "L|" & "[T" & ChrW(234) & "n cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", h" & ChrW(7897) & "i th" & ChrW(7843) & "o ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "Th" & ChrW(224) & "nh ph" & ChrW(7847) & "n tham d" & ChrW(7921) & ": [th" & ChrW(224) & "nh ph" & ChrW(7847) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & " (ch" & ChrW(7911) & " t" & ChrW(7885) & "a): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(432) & " k" & ChrW(253) & " (ng" & ChrW(432) & ChrW(7901) & "i ghi bi" & ChrW(234) & "n b" & ChrW(7843) & "n): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung (theo di" & ChrW(7877) & "n bi" & ChrW(7871) & "n cu" & ChrW(7897) & "c h" & ChrW(7885) & "p, h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", h" & ChrW(7897) & "i th" & ChrW(7843) & "o ...):" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "K" & ChrW(7871) & "t qu" & ChrW(7843) & " bi" & ChrW(7875) & "u quy" & ChrW(7871) & "t, th" & ChrW(244) & "ng qua: [...] (n" & ChrW(7871) & "u c" & ChrW(243) & ")." & vbLf
s = s & "L|" & "Cu" & ChrW(7897) & "c h" & ChrW(7885) & "p (h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", h" & ChrW(7897) & "i th" & ChrW(7843) & "o ...) k" & ChrW(7871) & "t th" & ChrW(250) & "c v" & ChrW(224) & "o [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]./." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, H" & ChrW(7891) & " s" & ChrW(417) & "." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " T" & ChrW(7884) & "A" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
s = s & "LR|" & "TH" & ChrW(431) & " K" & ChrW(221) & vbLf
s = s & "LH|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_8 = s
End Function
Private Function TT_Mau_NHA_NUOC_9() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "Ban h" & ChrW(224) & "nh (Ph" & ChrW(234) & " duy" & ChrW(7879) & "t) [t" & ChrW(234) & "n v" & ChrW(259) & "n b" & ChrW(7843) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ban h" & ChrW(224) & "nh, ph" & ChrW(234) & " duy" & ChrW(7879) & "t]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[TH" & ChrW(7848) & "M QUY" & ChrW(7872) & "N BAN H" & ChrW(192) & "NH, V" & ChrW(205) & " D" & ChrW(7908) & ": " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N ... / CH" & ChrW(7910) & " T" & ChrW(7882) & "CH " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh th" & ChrW(7849) & "m quy" & ChrW(7873) & "n, ch" & ChrW(7913) & "c n" & ChrW(259) & "ng, nhi" & ChrW(7879) & "m v" & ChrW(7909) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(7897) & "i dung, c" & ChrW(417) & " s" & ChrW(7903) & " ban h" & ChrW(224) & "nh];" & vbLf
s = s & "L|" & "Theo " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. Ban h" & ChrW(224) & "nh (Ph" & ChrW(234) & " duy" & ChrW(7879) & "t) k" & ChrW(232) & "m theo Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y [t" & ChrW(234) & "n v" & ChrW(259) & "n b" & ChrW(7843) & "n]." & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c thi h" & ChrW(224) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y./." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_9 = s
End Function
Private Function TT_Mau_NHA_NUOC_10() As String
Dim s As String
s = ""
s = s & "CODE|" & vbLf
s = s & "ND|1" & vbLf
s = s & "NS|1" & vbLf
s = s & "NK|1" & vbLf
s = s & "L|" & "QUY " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung quy " & ChrW(273) & ChrW(7883) & "nh]" & vbLf
s = s & "L|" & "(K" & ChrW(232) & "m theo Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh s" & ChrW(7889) & "    /Q" & ChrW(272) & "-{VIETTAT} ng" & ChrW(224) & "y    th" & ChrW(225) & "ng    n" & ChrW(259) & "m     c" & ChrW(7911) & "a [c" & ChrW(417) & " quan ban h" & ChrW(224) & "nh quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh])" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng I" & vbLf
s = s & "L|" & "QUY " & ChrW(272) & ChrW(7882) & "NH CHUNG" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. Ph" & ChrW(7841) & "m vi " & ChrW(273) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. " & ChrW(272) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng II" & vbLf
s = s & "L|" & "[N" & ChrW(7896) & "I DUNG QUY " & ChrW(272) & ChrW(7882) & "NH]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [T" & ChrW(234) & "n " & ChrW(273) & "i" & ChrW(7873) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng III" & vbLf
s = s & "L|" & "T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 4. Tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung]./." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & vbLf
s = s & "H|" & vbLf
TT_Mau_NHA_NUOC_10 = s
End Function
Private Function TT_Mau_NHA_NUOC_11() As String
Dim s As String
s = ""
s = s & "CODE|" & vbLf
s = s & "ND|1" & vbLf
s = s & "NS|1" & vbLf
s = s & "NK|1" & vbLf
s = s & "L|" & "QUY CH" & ChrW(7870) & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung quy ch" & ChrW(7871) & "]" & vbLf
s = s & "L|" & "(K" & ChrW(232) & "m theo Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh s" & ChrW(7889) & "    /Q" & ChrW(272) & "-{VIETTAT} ng" & ChrW(224) & "y    th" & ChrW(225) & "ng    n" & ChrW(259) & "m     c" & ChrW(7911) & "a [c" & ChrW(417) & " quan ban h" & ChrW(224) & "nh quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh])" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng I" & vbLf
s = s & "L|" & "QUY " & ChrW(272) & ChrW(7882) & "NH CHUNG" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. Ph" & ChrW(7841) & "m vi " & ChrW(273) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh, " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng II" & vbLf
s = s & "L|" & "[N" & ChrW(7896) & "I DUNG QUY CH" & ChrW(7870) & "]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [T" & ChrW(234) & "n " & ChrW(273) & "i" & ChrW(7873) & "u]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "Ch" & ChrW(432) & ChrW(417) & "ng III" & vbLf
s = s & "L|" & "T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. Tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung]./." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & vbLf
s = s & "H|" & vbLf
TT_Mau_NHA_NUOC_11 = s
End Function
Private Function TT_Mau_NHA_NUOC_12() As String
Dim s As String
s = ""
s = s & "CODE|" & "HD" & vbLf
s = s & "L|" & "H" & ChrW(431) & ChrW(7898) & "NG D" & ChrW(7850) & "N" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n v" & ChrW(7873) & " vi" & ChrW(7879) & "c [...], nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_12 = s
End Function
Private Function TT_Mau_NHA_NUOC_13() As String
Dim s As String
s = ""
s = s & "CODE|" & "CTr" & vbLf
s = s & "L|" & "CH" & ChrW(431) & ChrW(416) & "NG TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ")];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ")]." & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] ban h" & ChrW(224) & "nh Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh v" & ChrW(7873) & " vi" & ChrW(7879) & "c [...] v" & ChrW(7899) & "i c" & ChrW(225) & "c n" & ChrW(7897) & "i dung nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_13 = s
End Function
Private Function TT_Mau_NHA_NUOC_14() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c " & ChrW(7911) & "y quy" & ChrW(7873) & "n cho [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c, c" & ChrW(225) & " nh" & ChrW(226) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(7911) & "y quy" & ChrW(7873) & "n]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[TH" & ChrW(7848) & "M QUY" & ChrW(7872) & "N BAN H" & ChrW(192) & "NH, V" & ChrW(205) & " D" & ChrW(7908) & ": " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " Lu" & ChrW(7853) & "t T" & ChrW(7893) & " ch" & ChrW(7913) & "c ch" & ChrW(237) & "nh quy" & ChrW(7873) & "n " & ChrW(273) & ChrW(7883) & "a ph" & ChrW(432) & ChrW(417) & "ng ng" & ChrW(224) & "y [...] th" & ChrW(225) & "ng [...] n" & ChrW(259) & "m [...];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. " & ChrW(7910) & "y quy" & ChrW(7873) & "n cho [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c, c" & ChrW(225) & " nh" & ChrW(226) & "n] th" & ChrW(7921) & "c hi" & ChrW(7879) & "n [n" & ChrW(7897) & "i dung, ph" & ChrW(7841) & "m vi " & ChrW(7911) & "y quy" & ChrW(7873) & "n]." & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. Th" & ChrW(7901) & "i h" & ChrW(7841) & "n " & ChrW(7911) & "y quy" & ChrW(7873) & "n: [...]; c" & ChrW(225) & "ch th" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n v" & ChrW(224) & " c" & ChrW(225) & "c " & ChrW(273) & "i" & ChrW(7873) & "u ki" & ChrW(7879) & "n c" & ChrW(7847) & "n thi" & ChrW(7871) & "t " & ChrW(273) & ChrW(7875) & " th" & ChrW(7921) & "c hi" & ChrW(7879) & "n nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", quy" & ChrW(7873) & "n h" & ChrW(7841) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(7911) & "y quy" & ChrW(7873) & "n: [...]." & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ", c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y./." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_14 = s
End Function
Private Function TT_Mau_NHA_NUOC_15() As String
Dim s As String
s = ""
s = s & "CODE|" & "GGT" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y GI" & ChrW(7898) & "I THI" & ChrW(7878) & "U" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u:" & vbLf
s = s & "L|" & ChrW(212) & "ng (b" & ChrW(224) & "): [h" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n ng" & ChrW(432) & ChrW(7901) & "i " & ChrW(273) & ChrW(432) & ChrW(7907) & "c gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u]." & vbLf
s = s & "L|" & "Ch" & ChrW(7913) & "c v" & ChrW(7909) & ": [ch" & ChrW(7913) & "c v" & ChrW(7909) & ", c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " c" & ChrW(244) & "ng t" & ChrW(225) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7917) & " " & ChrW(273) & ChrW(7871) & "n: [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c " & ChrW(273) & ChrW(432) & ChrW(7907) & "c gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u t" & ChrW(7899) & "i l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c: [n" & ChrW(7897) & "i dung l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " Qu" & ChrW(253) & " c" & ChrW(417) & " quan t" & ChrW(7841) & "o " & ChrW(273) & "i" & ChrW(7873) & "u ki" & ChrW(7879) & "n " & ChrW(273) & ChrW(7875) & " " & ChrW(244) & "ng (b" & ChrW(224) & ") c" & ChrW(243) & " t" & ChrW(234) & "n " & ChrW(7903) & " tr" & ChrW(234) & "n ho" & ChrW(224) & "n th" & ChrW(224) & "nh nhi" & ChrW(7879) & "m v" & ChrW(7909) & "." & vbLf
s = s & "L|" & "Gi" & ChrW(7845) & "y n" & ChrW(224) & "y c" & ChrW(243) & " gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(273) & ChrW(7871) & "n h" & ChrW(7871) & "t ng" & ChrW(224) & "y [...]./." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_15 = s
End Function
Private Function TT_Mau_NHA_NUOC_16() As String
Dim s As String
s = ""
s = s & "CODE|" & "C" & ChrW(272) & vbLf
s = s & "L|" & "C" & ChrW(212) & "NG " & ChrW(272) & "I" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung " & ChrW(273) & "i" & ChrW(7879) & "n]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "CH" & ChrW(7910) & " T" & ChrW(7882) & "CH " & ChrW(7910) & "Y BAN NH" & ChrW(194) & "N D" & ChrW(194) & "N [T" & ChrW(202) & "N X" & ChrW(195) & ", PH" & ChrW(431) & ChrW(7900) & "NG, " & ChrW(272) & ChrW(7862) & "C KHU] " & ChrW(273) & "i" & ChrW(7879) & "n:" & vbLf
s = s & "L|" & "- [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n " & ChrW(273) & "i" & ChrW(7879) & "n];" & vbLf
s = s & "L|" & "- [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n " & ChrW(273) & "i" & ChrW(7879) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung " & ChrW(273) & "i" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " T" & ChrW(7882) & "CH" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_16 = s
End Function
Private Function TT_Mau_NHA_NUOC_17() As String
Dim s As String
s = ""
s = s & "CODE|" & "PC" & vbLf
s = s & "L|" & "PHI" & ChrW(7870) & "U CHUY" & ChrW(7874) & "N" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung phi" & ChrW(7871) & "u chuy" & ChrW(7875) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_17 = s
End Function
Private Function TT_Mau_NHA_NUOC_18() As String
Dim s As String
s = ""
s = s & "CODE|" & "NQ" & vbLf
s = s & "L|" & "NGH" & ChrW(7882) & " QUY" & ChrW(7870) & "T" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[T" & ChrW(202) & "N T" & ChrW(7852) & "P TH" & ChrW(7874) & " BAN H" & ChrW(192) & "NH, V" & ChrW(205) & " D" & ChrW(7908) & ": H" & ChrW(7896) & "I " & ChrW(272) & ChrW(7890) & "NG NH" & ChrW(194) & "N D" & ChrW(194) & "N ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n quy " & ChrW(273) & ChrW(7883) & "nh th" & ChrW(7849) & "m quy" & ChrW(7873) & "n];" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "X" & ChrW(233) & "t [t" & ChrW(7901) & " tr" & ChrW(236) & "nh, " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " ...]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T NGH" & ChrW(7882) & ":" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [N" & ChrW(7897) & "i dung ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c, tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh.]" & vbLf
s = s & "L|" & "Ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t n" & ChrW(224) & "y " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c [t" & ChrW(7853) & "p th" & ChrW(7875) & " ban h" & ChrW(224) & "nh] th" & ChrW(244) & "ng qua ng" & ChrW(224) & "y [...] th" & ChrW(225) & "ng [...] n" & ChrW(259) & "m [...]./." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "TM. [T" & ChrW(202) & "N T" & ChrW(7852) & "P TH" & ChrW(7874) & "]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_18 = s
End Function
Private Function TT_Mau_NHA_NUOC_19() As String
Dim s As String
s = ""
s = s & "CODE|" & "CT" & vbLf
s = s & "L|" & "CH" & ChrW(7880) & " TH" & ChrW(7882) & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(225) & "nh gi" & ChrW(225) & " t" & ChrW(236) & "nh h" & ChrW(236) & "nh, n" & ChrW(234) & "u y" & ChrW(234) & "u c" & ChrW(7847) & "u c" & ChrW(7847) & "n ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7841) & "o.]" & vbLf
s = s & "L|" & "[Th" & ChrW(7849) & "m quy" & ChrW(7873) & "n ban h" & ChrW(224) & "nh] y" & ChrW(234) & "u c" & ChrW(7847) & "u:" & vbLf
s = s & "L|" & "1. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & " giao cho c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ".]" & vbLf
s = s & "L|" & "2. [Nhi" & ChrW(7879) & "m v" & ChrW(7909) & " giao cho c" & ChrW(417) & " quan, " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & ".]" & vbLf
s = s & "L|" & "3. [T" & ChrW(7893) & " ch" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT, [" & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " so" & ChrW(7841) & "n th" & ChrW(7843) & "o]." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_19 = s
End Function
Private Function TT_Mau_NHA_NUOC_20() As String
Dim s As String
s = ""
s = s & "CODE|" & "GNP" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y NGH" & ChrW(7880) & " PH" & ChrW(201) & "P" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "X" & ChrW(233) & "t " & ChrW(272) & ChrW(417) & "n " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " ngh" & ChrW(7881) & " ph" & ChrW(233) & "p ng" & ChrW(224) & "y [...] th" & ChrW(225) & "ng [...] n" & ChrW(259) & "m [...] c" & ChrW(7911) & "a " & ChrW(244) & "ng (b" & ChrW(224) & ") [h" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]," & vbLf
s = s & "L|" & "[Quy" & ChrW(7873) & "n h" & ChrW(7841) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & " ng" & ChrW(432) & ChrW(7901) & "i k" & ChrW(253) & "] c" & ChrW(7845) & "p cho:" & vbLf
s = s & "L|" & ChrW(212) & "ng (b" & ChrW(224) & "): [h" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7913) & "c v" & ChrW(7909) & ": [ch" & ChrW(7913) & "c v" & ChrW(7909) & ", " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " c" & ChrW(244) & "ng t" & ChrW(225) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(432) & ChrW(7907) & "c ngh" & ChrW(7881) & " ph" & ChrW(233) & "p trong th" & ChrW(7901) & "i gian [s" & ChrW(7889) & " ng" & ChrW(224) & "y], k" & ChrW(7875) & " t" & ChrW(7915) & " ng" & ChrW(224) & "y [...] " & ChrW(273) & ChrW(7871) & "n h" & ChrW(7871) & "t ng" & ChrW(224) & "y [...]." & vbLf
s = s & "L|" & "T" & ChrW(7841) & "i: [n" & ChrW(417) & "i ngh" & ChrW(7881) & " ph" & ChrW(233) & "p]./." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "[TM. / KT. / TL. ...]" & vbLf
s = s & "R|" & "[CH" & ChrW(7912) & "C V" & ChrW(7908) & " NG" & ChrW(431) & ChrW(7900) & "I K" & ChrW(221) & "]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_NHA_NUOC_20 = s
End Function
Private Function TT_Mau_MTTQ_1() As String
Dim s As String
s = ""
s = s & "CODE|" & "KH" & vbLf
s = s & "L|" & "K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung k" & ChrW(7871) & " ho" & ChrW(7841) & "ch]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " M" & ChrW(7863) & "t tr" & ChrW(7853) & "n T" & ChrW(7893) & " qu" & ChrW(7889) & "c Vi" & ChrW(7879) & "t Nam;" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng tr" & ChrW(7921) & "c " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam [c" & ChrW(7845) & "p ...] x" & ChrW(226) & "y d" & ChrW(7921) & "ng K" & ChrW(7871) & " ho" & ChrW(7841) & "ch [...] nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "1. M" & ChrW(7909) & "c " & ChrW(273) & ChrW(237) & "ch" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "2. Y" & ChrW(234) & "u c" & ChrW(7847) & "u" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[Ph" & ChrW(226) & "n c" & ChrW(244) & "ng nhi" & ChrW(7879) & "m v" & ChrW(7909) & ", th" & ChrW(7901) & "i gian th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_1 = s
End Function
Private Function TT_Mau_MTTQ_2() As String
Dim s As String
s = ""
s = s & "CODE|" & "BC" & vbLf
s = s & "L|" & "B" & ChrW(193) & "O C" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung b" & ChrW(225) & "o c" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[M" & ChrW(7903) & " " & ChrW(273) & ChrW(7847) & "u b" & ChrW(225) & "o c" & ChrW(225) & "o.]" & vbLf
s = s & "L|" & "I. [T" & ChrW(204) & "NH H" & ChrW(204) & "NH, K" & ChrW(7870) & "T QU" & ChrW(7842) & " TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. [H" & ChrW(7840) & "N CH" & ChrW(7870) & ", NGUY" & ChrW(202) & "N NH" & ChrW(194) & "N]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. [PH" & ChrW(431) & ChrW(416) & "NG H" & ChrW(431) & ChrW(7898) & "NG, NHI" & ChrW(7878) & "M V" & ChrW(7908) & " TH" & ChrW(7900) & "I GIAN T" & ChrW(7898) & "I]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "IV. [KI" & ChrW(7870) & "N NGH" & ChrW(7882) & ", " & ChrW(272) & ChrW(7872) & " XU" & ChrW(7844) & "T]" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_2 = s
End Function
Private Function TT_Mau_MTTQ_3() As String
Dim s As String
s = ""
s = s & "CODE|" & "TB" & vbLf
s = s & "L|" & "TH" & ChrW(212) & "NG B" & ChrW(193) & "O" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung th" & ChrW(244) & "ng b" & ChrW(225) & "o.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_3 = s
End Function
Private Function TT_Mau_MTTQ_4() As String
Dim s As String
s = ""
s = s & "CODE|" & vbLf
s = s & "CV|1" & vbLf
s = s & "SUM|" & "V/v [tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n]" & vbLf
s = s & "L|" & "K" & ChrW(237) & "nh g" & ChrW(7917) & "i: [C" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c, c" & ChrW(225) & " nh" & ChrW(226) & "n nh" & ChrW(7853) & "n c" & ChrW(244) & "ng v" & ChrW(259) & "n]." & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung c" & ChrW(244) & "ng v" & ChrW(259) & "n.]" & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " tr" & ChrW(234) & "n;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_4 = s
End Function
Private Function TT_Mau_MTTQ_5() As String
Dim s As String
s = ""
s = s & "CODE|" & "GM" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y M" & ChrW(7900) & "I" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u: d" & ChrW(7921) & " h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng tr" & ChrW(7921) & "c " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam [c" & ChrW(7845) & "p ...] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng k" & ChrW(237) & "nh m" & ChrW(7901) & "i: [" & ChrW(272) & ChrW(7841) & "i bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7901) & "i]." & vbLf
s = s & "L|" & "T" & ChrW(7899) & "i d" & ChrW(7921) & ": [h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]." & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "[" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " chu" & ChrW(7849) & "n b" & ChrW(7883) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u, li" & ChrW(234) & "n h" & ChrW(7879) & " (n" & ChrW(7871) & "u c" & ChrW(243) & ").]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_5 = s
End Function
Private Function TT_Mau_MTTQ_6() As String
Dim s As String
s = ""
s = s & "CODE|" & "TTr" & vbLf
s = s & "L|" & "T" & ChrW(7900) & " TR" & ChrW(204) & "NH" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung t" & ChrW(7901) & " tr" & ChrW(236) & "nh]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " M" & ChrW(7863) & "t tr" & ChrW(7853) & "n T" & ChrW(7893) & " qu" & ChrW(7889) & "c Vi" & ChrW(7879) & "t Nam;" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng tr" & ChrW(7921) & "c " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam [c" & ChrW(7845) & "p ...] tr" & ChrW(236) & "nh [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c nh" & ChrW(7853) & "n t" & ChrW(7901) & " tr" & ChrW(236) & "nh] v" & ChrW(7873) & " vi" & ChrW(7879) & "c [...], nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung t" & ChrW(7901) & " tr" & ChrW(236) & "nh.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_6 = s
End Function
Private Function TT_Mau_MTTQ_7() As String
Dim s As String
s = ""
s = s & "CODE|" & "Q" & ChrW(272) & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH" & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c [...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C " & ChrW(7910) & "Y BAN M" & ChrW(7862) & "T TR" & ChrW(7852) & "N T" & ChrW(7892) & " QU" & ChrW(7888) & "C VI" & ChrW(7878) & "T NAM [C" & ChrW(7844) & "P ...]" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " M" & ChrW(7863) & "t tr" & ChrW(7853) & "n T" & ChrW(7893) & " qu" & ChrW(7889) & "c Vi" & ChrW(7879) & "t Nam;" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "];" & vbLf
s = s & "L|" & "X" & ChrW(233) & "t " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & " c" & ChrW(7911) & "a [b" & ChrW(7897) & " ph" & ChrW(7853) & "n, c" & ChrW(225) & " nh" & ChrW(226) & "n " & ChrW(273) & ChrW(7873) & " ngh" & ChrW(7883) & "]." & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T " & ChrW(272) & ChrW(7882) & "NH:" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 1. [N" & ChrW(7897) & "i dung quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 2. [Hi" & ChrW(7879) & "u l" & ChrW(7921) & "c thi h" & ChrW(224) & "nh.]" & vbLf
s = s & "L|" & ChrW(272) & "i" & ChrW(7873) & "u 3. [C" & ChrW(225) & "c b" & ChrW(7897) & " ph" & ChrW(7853) & "n, c" & ChrW(225) & " nh" & ChrW(226) & "n] ch" & ChrW(7883) & "u tr" & ChrW(225) & "ch nhi" & ChrW(7879) & "m thi h" & ChrW(224) & "nh Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh n" & ChrW(224) & "y./." & vbLf
s = s & "N|" & "- Nh" & ChrW(432) & " " & ChrW(272) & "i" & ChrW(7873) & "u 3;" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_7 = s
End Function
Private Function TT_Mau_MTTQ_8() As String
Dim s As String
s = ""
s = s & "CODE|" & "CTr" & vbLf
s = s & "L|" & "CH" & ChrW(431) & ChrW(416) & "NG TR" & ChrW(204) & "NH PH" & ChrW(7888) & "I H" & ChrW(7906) & "P" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " M" & ChrW(7863) & "t tr" & ChrW(7853) & "n T" & ChrW(7893) & " qu" & ChrW(7889) & "c Vi" & ChrW(7879) & "t Nam;" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "[C" & ChrW(225) & "c b" & ChrW(234) & "n] th" & ChrW(7889) & "ng nh" & ChrW(7845) & "t Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh ph" & ChrW(7889) & "i h" & ChrW(7907) & "p [...] nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "I. M" & ChrW(7908) & "C " & ChrW(272) & ChrW(205) & "CH, Y" & ChrW(202) & "U C" & ChrW(7846) & "U" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "II. N" & ChrW(7896) & "I DUNG PH" & ChrW(7888) & "I H" & ChrW(7906) & "P" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "L|" & "III. T" & ChrW(7892) & " CH" & ChrW(7912) & "C TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_8 = s
End Function
Private Function TT_Mau_MTTQ_9() As String
Dim s As String
s = ""
s = s & "CODE|" & "HD" & vbLf
s = s & "L|" & "H" & ChrW(431) & ChrW(7898) & "NG D" & ChrW(7850) & "N" & vbLf
s = s & "L|" & "[Tr" & ChrW(237) & "ch y" & ChrW(7871) & "u n" & ChrW(7897) & "i dung h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " " & ChrW(272) & "i" & ChrW(7873) & "u l" & ChrW(7879) & " M" & ChrW(7863) & "t tr" & ChrW(7853) & "n T" & ChrW(7893) & " qu" & ChrW(7889) & "c Vi" & ChrW(7879) & "t Nam;" & vbLf
s = s & "L|" & "C" & ChrW(259) & "n c" & ChrW(7913) & " [v" & ChrW(259) & "n b" & ChrW(7843) & "n l" & ChrW(224) & "m c" & ChrW(259) & "n c" & ChrW(7913) & "]." & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng tr" & ChrW(7921) & "c " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam [c" & ChrW(7845) & "p ...] h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n [...] nh" & ChrW(432) & " sau:" & vbLf
s = s & "L|" & "[N" & ChrW(7897) & "i dung h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_9 = s
End Function
Private Function TT_Mau_MTTQ_10() As String
Dim s As String
s = ""
s = s & "CODE|" & "NQ" & vbLf
s = s & "L|" & "NGH" & ChrW(7882) & " QUY" & ChrW(7870) & "T" & vbLf
s = s & "L|" & "[H" & ChrW(7897) & "i ngh" & ChrW(7883) & " " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam ... l" & ChrW(7847) & "n th" & ChrW(7913) & " ..., nhi" & ChrW(7879) & "m k" & ChrW(7923) & " ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "[Ng" & ChrW(224) & "y ... th" & ChrW(225) & "ng ... n" & ChrW(259) & "m ..., t" & ChrW(7841) & "i ..., " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam ... t" & ChrW(7893) & " ch" & ChrW(7913) & "c H" & ChrW(7897) & "i ngh" & ChrW(7883) & " l" & ChrW(7847) & "n th" & ChrW(7913) & " ...]" & vbLf
s = s & "L|" & "QUY" & ChrW(7870) & "T NGH" & ChrW(7882) & ":" & vbLf
s = s & "L|" & "1. [N" & ChrW(7897) & "i dung ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & "2. [N" & ChrW(7897) & "i dung ngh" & ChrW(7883) & " quy" & ChrW(7871) & "t.]" & vbLf
s = s & "L|" & "3. [T" & ChrW(7893) & " ch" & ChrW(7913) & "c th" & ChrW(7921) & "c hi" & ChrW(7879) & "n.]" & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. " & ChrW(7910) & "Y BAN M" & ChrW(7862) & "T TR" & ChrW(7852) & "N T" & ChrW(7892) & " QU" & ChrW(7888) & "C VI" & ChrW(7878) & "T NAM [C" & ChrW(7844) & "P ...]" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " T" & ChrW(7882) & "CH" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_10 = s
End Function
Private Function TT_Mau_MTTQ_11() As String
Dim s As String
s = ""
s = s & "CODE|" & "BB" & vbLf
s = s & "ND|1" & vbLf
s = s & "L|" & "BI" & ChrW(202) & "N B" & ChrW(7842) & "N" & vbLf
s = s & "L|" & "[H" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p ...]" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Th" & ChrW(7901) & "i gian b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m: [" & ChrW(273) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m]." & vbLf
s = s & "L|" & "Th" & ChrW(224) & "nh ph" & ChrW(7847) & "n tham d" & ChrW(7921) & ": [th" & ChrW(224) & "nh ph" & ChrW(7847) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7911) & " tr" & ChrW(236) & ": [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "Th" & ChrW(432) & " k" & ChrW(253) & " (ng" & ChrW(432) & ChrW(7901) & "i ghi bi" & ChrW(234) & "n b" & ChrW(7843) & "n): [h" & ChrW(7885) & " t" & ChrW(234) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & "N" & ChrW(7897) & "i dung: [theo di" & ChrW(7877) & "n bi" & ChrW(7871) & "n h" & ChrW(7897) & "i ngh" & ChrW(7883) & ", cu" & ChrW(7897) & "c h" & ChrW(7885) & "p]." & vbLf
s = s & "L|" & "H" & ChrW(7897) & "i ngh" & ChrW(7883) & " (cu" & ChrW(7897) & "c h" & ChrW(7885) & "p) k" & ChrW(7871) & "t th" & ChrW(250) & "c v" & ChrW(224) & "o [gi" & ChrW(7901) & ", ng" & ChrW(224) & "y, th" & ChrW(225) & "ng, n" & ChrW(259) & "m]./." & vbLf
s = s & "A|" & vbLf
s = s & "R|" & "CH" & ChrW(7910) & " TR" & ChrW(204) & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
s = s & "LR|" & "TH" & ChrW(431) & " K" & ChrW(221) & vbLf
s = s & "LH|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_11 = s
End Function
Private Function TT_Mau_MTTQ_12() As String
Dim s As String
s = ""
s = s & "CODE|" & "GGT" & vbLf
s = s & "L|" & "GI" & ChrW(7844) & "Y GI" & ChrW(7898) & "I THI" & ChrW(7878) & "U" & vbLf
s = s & "L|" & "-----" & vbLf
s = s & "L|" & "Ban Th" & ChrW(432) & ChrW(7901) & "ng tr" & ChrW(7921) & "c " & ChrW(7910) & "y ban MTTQ Vi" & ChrW(7879) & "t Nam [c" & ChrW(7845) & "p ...] tr" & ChrW(226) & "n tr" & ChrW(7885) & "ng gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u:" & vbLf
s = s & "L|" & ChrW(212) & "ng (b" & ChrW(224) & "): [h" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]." & vbLf
s = s & "L|" & "Ch" & ChrW(7913) & "c v" & ChrW(7909) & ": [ch" & ChrW(7913) & "c v" & ChrW(7909) & "]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7917) & " " & ChrW(273) & ChrW(7871) & "n: [c" & ChrW(417) & " quan, t" & ChrW(7893) & " ch" & ChrW(7913) & "c " & ChrW(273) & ChrW(432) & ChrW(7907) & "c gi" & ChrW(7899) & "i thi" & ChrW(7879) & "u t" & ChrW(7899) & "i l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & "V" & ChrW(7873) & " vi" & ChrW(7879) & "c: [n" & ChrW(7897) & "i dung l" & ChrW(224) & "m vi" & ChrW(7879) & "c]." & vbLf
s = s & "L|" & ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " Qu" & ChrW(253) & " c" & ChrW(417) & " quan t" & ChrW(7841) & "o " & ChrW(273) & "i" & ChrW(7873) & "u ki" & ChrW(7879) & "n " & ChrW(273) & ChrW(7875) & " " & ChrW(244) & "ng (b" & ChrW(224) & ") c" & ChrW(243) & " t" & ChrW(234) & "n " & ChrW(7903) & " tr" & ChrW(234) & "n ho" & ChrW(224) & "n th" & ChrW(224) & "nh nhi" & ChrW(7879) & "m v" & ChrW(7909) & "." & vbLf
s = s & "L|" & "Gi" & ChrW(7845) & "y n" & ChrW(224) & "y c" & ChrW(243) & " gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(273) & ChrW(7871) & "n h" & ChrW(7871) & "t ng" & ChrW(224) & "y [...]./." & vbLf
s = s & "N|" & "- [N" & ChrW(417) & "i nh" & ChrW(7853) & "n];" & vbLf
s = s & "N|" & "- L" & ChrW(432) & "u: VT." & vbLf
s = s & "A|" & "TM. BAN TH" & ChrW(431) & ChrW(7900) & "NG TR" & ChrW(7920) & "C" & vbLf
s = s & "R|" & "[CH" & ChrW(7910) & " T" & ChrW(7882) & "CH / PH" & ChrW(211) & " CH" & ChrW(7910) & " T" & ChrW(7882) & "CH]" & vbLf
s = s & "H|" & "[H" & ChrW(7885) & " v" & ChrW(224) & " t" & ChrW(234) & "n]" & vbLf
TT_Mau_MTTQ_12 = s
End Function
Public Function TT_TextTieuNguDang() As String
TT_TextTieuNguDang = ChrW(272) & ChrW(7842) & "NG C" & ChrW(7896) & "NG S" & ChrW(7842) & "N VI" & ChrW(7878) & "T NAM"
End Function
Public Function TT_GetSetting() As String
Dim v As String
On Error Resume Next
v = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "DonVi", "AUTO")
On Error GoTo 0
If Len(v) = 0 Then v = "AUTO"
If v <> "AUTO" And TT_IndexOfProfile(v) = 0 Then v = "AUTO"
TT_GetSetting = v
End Function
Public Sub TT_SaveSetting(ByVal pid As String)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonVi", pid
On Error GoTo 0
End Sub
Private Function TT_IndexOfProfile(ByVal pid As String) As Integer
Dim i As Integer
TT_IndexOfProfile = 0
For i = 1 To TT_ProfileCount()
If TT_ProfileIdAt(i) = pid Then
TT_IndexOfProfile = i
Exit Function
End If
Next i
End Function
Public Sub RibbonDonVi_GetSelectedIndex(control As Object, ByRef returnedVal)
Dim v As String
v = TT_GetSetting()
If v = "AUTO" Then
returnedVal = 0
Else
returnedVal = TT_IndexOfProfile(v)
End If
End Sub
Public Sub RibbonDonVi_OnAction(control As Object, selectedId As String, selectedIndex As Integer)
If selectedIndex <= 0 Then
Call TT_SaveSetting("AUTO")
Else
Call TT_SaveSetting(TT_ProfileIdAt(selectedIndex))
End If
Call TT_RibbonRefresh
Call TT_CanhBaoDonViLechKhoi
End Sub
Private Sub TT_CanhBaoDonViLechKhoi()
On Error Resume Next
If TT_GetSetting() = "AUTO" Then Exit Sub
Dim grp As String, ten As String, lech As Boolean
grp = TT_UnitGroupCurrent()
ten = " " & UCase(ToAscii(TT_UnitGet(grp, "COQUAN") & " " & TT_UnitGet(grp, "CAPTREN"))) & " "
If Len(Trim(ten)) = 0 Then Exit Sub
If grp = "DANG" Then
lech = (InStr(1, ten, "UY BAN NHAN DAN") > 0 Or InStr(1, ten, " UBND ") > 0 Or InStr(1, ten, "HOI DONG NHAN DAN") > 0)
Else
lech = (InStr(1, ten, " DANG UY ") > 0 Or InStr(1, ten, " DANG BO ") > 0 Or InStr(1, ten, " CHI BO ") > 0 Or _
InStr(1, ten, " TINH UY ") > 0 Or InStr(1, ten, " THANH UY ") > 0)
End If
If lech Then
MsgBox "Don vi ban hanh dang chon cua khoi nay la: " & ToAscii(TT_UnitLabelAt(grp, TT_UnitSlot(grp))) & vbCrLf & vbCrLf & _
"Ten nay co ve thuoc he thong khac voi khoi vua chon (" & TT_ProfileLabelAscii(TT_GetSetting()) & ")." & vbCrLf & _
"Moi khoi luu danh sach don vi rieng. Hay bam 'Sua thong tin' de nhap dung co quan cua khoi nay, hoac 'Them don vi'.", vbInformation, "Don vi ban hanh"
End If
On Error GoTo 0
End Sub
Public Sub RibbonOnLoad(ribbon As Object)
Set mTTRibbon = ribbon
Call CN_HenKiemTraKhiMo
End Sub
Private Sub TT_RibbonRefresh()
On Error Resume Next
If Not mTTRibbon Is Nothing Then mTTRibbon.Invalidate
On Error GoTo 0
End Sub
Public Sub RibbonChonKhoi(control As Object)
Call ChonKhoiDonVi
End Sub
Public Sub ChonKhoiDonVi()
Dim msg As String
Dim i As Integer
msg = "Chon khoi the thuc van ban (go so roi bam OK):" & vbCrLf & vbCrLf & "0. Tu dong nhan dien theo dong tieu ngu / co quan ban hanh"
For i = 1 To TT_ProfileCount()
msg = msg & vbCrLf & i & ". " & TT_ProfileLabelAscii(TT_ProfileIdAt(i))
Next i
msg = msg & vbCrLf & vbCrLf & "Van ban cua cap uy, to chuc dang trong Cong an, Quan doi: chon khoi Dang."
Dim cur As String
Dim curIdx As Integer
cur = TT_GetSetting()
If cur = "AUTO" Then
curIdx = 0
Else
curIdx = TT_IndexOfProfile(cur)
End If
Dim ans As String
ans = InputBox(msg, "Khoi the thuc van ban", CStr(curIdx))
If Len(Trim(ans)) = 0 Then Exit Sub
If Not IsNumeric(ans) Then Exit Sub
Dim n As Integer
n = Val(ans)
If n = 0 Then
Call TT_SaveSetting("AUTO")
ElseIf n >= 1 And n <= TT_ProfileCount() Then
Call TT_SaveSetting(TT_ProfileIdAt(n))
Else
Exit Sub
End If
MsgBox "Da chon khoi: " & TT_ProfileLabelAscii(TT_GetSetting()), vbInformation, "Khoi the thuc van ban"
End Sub
Private Function TT_HeaderAreaAscii(ByVal doc As Document) As String
Dim buf As String
Dim nonEmpty As Integer
Dim p As Paragraph
On Error Resume Next
For Each p In doc.Paragraphs
Dim sRaw As String
sRaw = CleanParaText(p.Range.text)
If Len(sRaw) > 0 Then
Dim sA As String
sA = Trim(ToAscii(sRaw))
If nonEmpty >= 3 And Not p.Range.Information(wdWithInTable) Then
If IsPartyDocTitle(sA) Or IsKinhGuiStart(sRaw, sA) Or IsStrictBodyOrHeadingStart(sRaw, sA) Then Exit For
End If
buf = buf & " " & sA
nonEmpty = nonEmpty + 1
If nonEmpty >= 14 Then Exit For
End If
Next p
On Error GoTo 0
TT_HeaderAreaAscii = buf
End Function
Public Function TT_Resolve(ByVal doc As Document) As String
Dim setting As String
Dim strong As Boolean
setting = TT_GetSetting()
If Len(mTTForceProfile) > 0 Then setting = mTTForceProfile
mTTHeaderNote = ""
mTTDetected = TT_DetectFromHeaderText(TT_HeaderAreaAscii(doc), strong)
mTTDetectedStrong = strong
If setting = "AUTO" Then
If Len(mTTDetected) > 0 Then
mTTProfile = mTTDetected
mTTSource = "TU_NHAN_DIEN"
Else
mTTProfile = TT_DEFAULT_PROFILE
mTTSource = "MAC_DINH"
End If
Else
mTTProfile = setting
mTTSource = "CHON"
End If
TT_Resolve = mTTProfile
End Function
Public Function TT_Profile() As String
If Len(mTTProfile) = 0 Then mTTProfile = TT_DEFAULT_PROFILE
TT_Profile = mTTProfile
End Function
Public Function TT_N(ByVal key As String) As Single
Dim ov As Single
TT_N = TT_RawN(TT_Profile(), key)
If TT_OvGet(key, ov) Then TT_N = ov
End Function
Public Function TT_OptCount(ByVal grp As String) As Integer
If grp = "DANG" Then
TT_OptCount = 9
Else
TT_OptCount = 20
End If
End Function
Public Function TT_OptKey(ByVal grp As String, ByVal i As Integer) As String
TT_OptKey = ""
If grp <> "DANG" Then
Select Case i
Case 1: TT_OptKey = "BODY_SIZE"
Case 2: TT_OptKey = "TITLE_SIZE"
Case 3: TT_OptKey = "SUMMARY_SIZE"
Case 4: TT_OptKey = "QH_SIZE"
Case 5: TT_OptKey = "TN_SIZE"
Case 6: TT_OptKey = "ORGAN_SIZE"
Case 7: TT_OptKey = "DATE_SIZE"
Case 8: TT_OptKey = "CV_SUMMARY_SIZE"
Case 9: TT_OptKey = "SIGN_SIZE"
Case 10: TT_OptKey = "PAGE_NUMBER_SIZE"
Case 11: TT_OptKey = "INDENT_CM"
Case 12: TT_OptKey = "PAGE_TOP"
Case 13: TT_OptKey = "PAGE_BOTTOM"
Case 14: TT_OptKey = "PAGE_LEFT"
Case 15: TT_OptKey = "PAGE_RIGHT"
Case 16: TT_OptKey = "LINE_SPACING"
Case 17: TT_OptKey = "PARA_BEFORE"
Case 18: TT_OptKey = "PARA_AFTER"
Case 19: TT_OptKey = "TEN_DAI"
Case 20: TT_OptKey = "GOM_CHU_LE"
End Select
ElseIf grp = "DANG" Then
Select Case i
Case 1: TT_OptKey = "BODY_SIZE"
Case 2: TT_OptKey = "TITLE_SIZE"
Case 3: TT_OptKey = "SUMMARY_SIZE"
Case 4: TT_OptKey = "INDENT_CM"
Case 5: TT_OptKey = "LINE_SPACING"
Case 6: TT_OptKey = "PARA_BEFORE"
Case 7: TT_OptKey = "PARA_AFTER"
Case 8: TT_OptKey = "TEN_DAI"
Case 9: TT_OptKey = "GOM_CHU_LE"
End Select
End If
End Function
Public Function TT_OptLabel(ByVal grp As String, ByVal i As Integer) As String
TT_OptLabel = ""
If grp <> "DANG" Then
Select Case i
Case 1: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " n" & ChrW(7897) & "i dung"
Case 2: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " t" & ChrW(234) & "n lo" & ChrW(7841) & "i v" & ChrW(259) & "n b" & ChrW(7843) & "n"
Case 3: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " tr" & ChrW(237) & "ch y" & ChrW(7871) & "u"
Case 4: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " Qu" & ChrW(7889) & "c hi" & ChrW(7879) & "u"
Case 5: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " Ti" & ChrW(234) & "u ng" & ChrW(7919)
Case 6: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " t" & ChrW(234) & "n c" & ChrW(417) & " quan"
Case 7: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " " & ChrW(273) & ChrW(7883) & "a danh, ng" & ChrW(224) & "y th" & ChrW(225) & "ng"
Case 8: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " tr" & ChrW(237) & "ch y" & ChrW(7871) & "u c" & ChrW(244) & "ng v" & ChrW(259) & "n (V/v)"
Case 9: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " quy" & ChrW(7873) & "n h" & ChrW(7841) & "n, ch" & ChrW(7913) & "c v" & ChrW(7909) & ", h" & ChrW(7885) & " t" & ChrW(234) & "n ng" & ChrW(432) & ChrW(7901) & "i k" & ChrW(253)
Case 10: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " s" & ChrW(7889) & " trang"
Case 11: TT_OptLabel = "Th" & ChrW(7909) & "t " & ChrW(273) & ChrW(7847) & "u d" & ChrW(242) & "ng (cm)"
Case 12: TT_OptLabel = "L" & ChrW(7873) & " tr" & ChrW(234) & "n (mm)"
Case 13: TT_OptLabel = "L" & ChrW(7873) & " d" & ChrW(432) & ChrW(7899) & "i (mm)"
Case 14: TT_OptLabel = "L" & ChrW(7873) & " tr" & ChrW(225) & "i (mm)"
Case 15: TT_OptLabel = "L" & ChrW(7873) & " ph" & ChrW(7843) & "i (mm)"
Case 16: TT_OptLabel = "Gi" & ChrW(227) & "n d" & ChrW(242) & "ng n" & ChrW(7897) & "i dung"
Case 17: TT_OptLabel = "Gi" & ChrW(227) & "n " & ChrW(273) & "o" & ChrW(7841) & "n - ph" & ChrW(237) & "a tr" & ChrW(234) & "n (pt)"
Case 18: TT_OptLabel = "Gi" & ChrW(227) & "n " & ChrW(273) & "o" & ChrW(7841) & "n - ph" & ChrW(237) & "a d" & ChrW(432) & ChrW(7899) & "i (pt)"
Case 19: TT_OptLabel = "T" & ChrW(234) & "n c" & ChrW(417) & " quan d" & ChrW(224) & "i b" & ChrW(7883) & " r" & ChrW(7899) & "t d" & ChrW(242) & "ng"
Case 20: TT_OptLabel = "Gom ch" & ChrW(7919) & " l" & ChrW(7867) & " cu" & ChrW(7889) & "i " & ChrW(273) & "o" & ChrW(7841) & "n"
End Select
ElseIf grp = "DANG" Then
Select Case i
Case 1: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " n" & ChrW(7897) & "i dung"
Case 2: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " t" & ChrW(234) & "n lo" & ChrW(7841) & "i v" & ChrW(259) & "n b" & ChrW(7843) & "n"
Case 3: TT_OptLabel = "C" & ChrW(7905) & " ch" & ChrW(7919) & " tr" & ChrW(237) & "ch y" & ChrW(7871) & "u"
Case 4: TT_OptLabel = "Th" & ChrW(7909) & "t " & ChrW(273) & ChrW(7847) & "u d" & ChrW(242) & "ng (cm)"
Case 5: TT_OptLabel = "Gi" & ChrW(227) & "n d" & ChrW(242) & "ng n" & ChrW(7897) & "i dung"
Case 6: TT_OptLabel = "Gi" & ChrW(227) & "n " & ChrW(273) & "o" & ChrW(7841) & "n - ph" & ChrW(237) & "a tr" & ChrW(234) & "n (pt)"
Case 7: TT_OptLabel = "Gi" & ChrW(227) & "n " & ChrW(273) & "o" & ChrW(7841) & "n - ph" & ChrW(237) & "a d" & ChrW(432) & ChrW(7899) & "i (pt)"
Case 8: TT_OptLabel = "T" & ChrW(234) & "n c" & ChrW(417) & " quan d" & ChrW(224) & "i b" & ChrW(7883) & " r" & ChrW(7899) & "t d" & ChrW(242) & "ng"
Case 9: TT_OptLabel = "Gom ch" & ChrW(7919) & " l" & ChrW(7867) & " cu" & ChrW(7889) & "i " & ChrW(273) & "o" & ChrW(7841) & "n"
End Select
End If
End Function
Public Function TT_OptMin(ByVal grp As String, ByVal i As Integer) As Single
TT_OptMin = 0
If grp <> "DANG" Then
Select Case i
Case 1: TT_OptMin = 13
Case 2: TT_OptMin = 13
Case 3: TT_OptMin = 13
Case 4: TT_OptMin = 12
Case 5: TT_OptMin = 13
Case 6: TT_OptMin = 12
Case 7: TT_OptMin = 13
Case 8: TT_OptMin = 12
Case 9: TT_OptMin = 13
Case 10: TT_OptMin = 13
Case 11: TT_OptMin = 1
Case 12: TT_OptMin = 20
Case 13: TT_OptMin = 20
Case 14: TT_OptMin = 30
Case 15: TT_OptMin = 15
Case 16: TT_OptMin = 1
Case 17: TT_OptMin = 0
Case 18: TT_OptMin = 0
Case 19: TT_OptMin = 1
Case 20: TT_OptMin = 0
End Select
ElseIf grp = "DANG" Then
Select Case i
Case 1: TT_OptMin = 14
Case 2: TT_OptMin = 15
Case 3: TT_OptMin = 14
Case 4: TT_OptMin = 1
Case 5: TT_OptMin = 18
Case 6: TT_OptMin = 0
Case 7: TT_OptMin = 0
Case 8: TT_OptMin = 1
Case 9: TT_OptMin = 0
End Select
End If
End Function
Public Function TT_OptMax(ByVal grp As String, ByVal i As Integer) As Single
TT_OptMax = 0
If grp <> "DANG" Then
Select Case i
Case 1: TT_OptMax = 14
Case 2: TT_OptMax = 14
Case 3: TT_OptMax = 14
Case 4: TT_OptMax = 13
Case 5: TT_OptMax = 14
Case 6: TT_OptMax = 13
Case 7: TT_OptMax = 14
Case 8: TT_OptMax = 13
Case 9: TT_OptMax = 14
Case 10: TT_OptMax = 14
Case 11: TT_OptMax = 1.27
Case 12: TT_OptMax = 25
Case 13: TT_OptMax = 25
Case 14: TT_OptMax = 35
Case 15: TT_OptMax = 20
Case 16: TT_OptMax = 1.5
Case 17: TT_OptMax = 12
Case 18: TT_OptMax = 12
Case 19: TT_OptMax = 2
Case 20: TT_OptMax = 1
End Select
ElseIf grp = "DANG" Then
Select Case i
Case 1: TT_OptMax = 15
Case 2: TT_OptMax = 16
Case 3: TT_OptMax = 15
Case 4: TT_OptMax = 1.27
Case 5: TT_OptMax = 22
Case 6: TT_OptMax = 12
Case 7: TT_OptMax = 12
Case 8: TT_OptMax = 2
Case 9: TT_OptMax = 1
End Select
End If
End Function
Public Function TT_OptNote(ByVal grp As String, ByVal i As Integer) As String
TT_OptNote = ""
If grp <> "DANG" Then
Select Case i
Case 1: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 2: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 3: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 4: TT_OptNote = "12 " & ChrW(8211) & " 13 (N" & ChrW(272) & " 30/2020)"
Case 5: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 6: TT_OptNote = "12 " & ChrW(8211) & " 13 (N" & ChrW(272) & " 30/2020)"
Case 7: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 8: TT_OptNote = "12 " & ChrW(8211) & " 13 (N" & ChrW(272) & " 30/2020)"
Case 9: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 10: TT_OptNote = "13 " & ChrW(8211) & " 14 (N" & ChrW(272) & " 30/2020)"
Case 11: TT_OptNote = "1 ho" & ChrW(7863) & "c 1,27 cm (N" & ChrW(272) & " 30/2020)"
Case 12: TT_OptNote = "20 " & ChrW(8211) & " 25 mm (N" & ChrW(272) & " 30/2020)"
Case 13: TT_OptNote = "20 " & ChrW(8211) & " 25 mm (N" & ChrW(272) & " 30/2020)"
Case 14: TT_OptNote = "30 " & ChrW(8211) & " 35 mm (N" & ChrW(272) & " 30/2020)"
Case 15: TT_OptNote = "15 " & ChrW(8211) & " 20 mm (N" & ChrW(272) & " 30/2020)"
Case 16: TT_OptNote = "1 " & ChrW(8211) & " 1,5 d" & ChrW(242) & "ng (N" & ChrW(272) & " 30/2020). " & ChrW(272) & ChrW(7875) & " tr" & ChrW(7889) & "ng = d" & ChrW(242) & "ng " & ChrW(273) & ChrW(417) & "n"
Case 17: TT_OptNote = "0 " & ChrW(8211) & " 12. Quy " & ChrW(273) & ChrW(7883) & "nh: c" & ChrW(225) & "c " & ChrW(273) & "o" & ChrW(7841) & "n c" & ChrW(225) & "ch nhau t" & ChrW(7889) & "i thi" & ChrW(7875) & "u 6pt"
Case 18: TT_OptNote = "0 " & ChrW(8211) & " 12. " & ChrW(272) & ChrW(7863) & "t 0 " & ChrW(273) & ChrW(7875) & " v" & ChrW(259) & "n b" & ChrW(7843) & "n g" & ChrW(7885) & "n h" & ChrW(417) & "n"
Case 19: TT_OptNote = "1 = h" & ChrW(7841) & " c" & ChrW(7905) & " ch" & ChrW(7919) & " (khuy" & ChrW(234) & "n d" & ChrW(249) & "ng, gi" & ChrW(7919) & " " & ChrW(273) & ChrW(250) & "ng l" & ChrW(7873) & "); 2 = l" & ChrW(7845) & "n l" & ChrW(7873) & " tr" & ChrW(225) & "i (gi" & ChrW(7919) & " c" & ChrW(7905) & " ch" & ChrW(7919) & ")"
Case 20: TT_OptNote = "0 = kh" & ChrW(244) & "ng; 1 = c" & ChrW(243) & ": d" & ChrW(242) & "ng cu" & ChrW(7889) & "i ch" & ChrW(7881) & " c" & ChrW(243) & " 1 t" & ChrW(7915) & " th" & ChrW(236) & " co nh" & ChrW(7865) & " ch" & ChrW(7919) & " " & ChrW(273) & ChrW(7875) & " k" & ChrW(233) & "o l" & ChrW(234) & "n d" & ChrW(242) & "ng tr" & ChrW(234) & "n"
End Select
ElseIf grp = "DANG" Then
Select Case i
Case 1: TT_OptNote = "14 " & ChrW(8211) & " 15 (HD 05-HD/VPTW)"
Case 2: TT_OptNote = "15 " & ChrW(8211) & " 16 (HD 05-HD/VPTW)"
Case 3: TT_OptNote = "14 " & ChrW(8211) & " 15 (HD 05-HD/VPTW)"
Case 4: TT_OptNote = "kho" & ChrW(7843) & "ng 1 cm (HD 05-HD/VPTW); 1 " & ChrW(8211) & " 1,27"
Case 5: TT_OptNote = "18 " & ChrW(8211) & " 22 pt, ki" & ChrW(7875) & "u Exactly (HD 05-HD/VPTW). " & ChrW(272) & ChrW(7875) & " tr" & ChrW(7889) & "ng = d" & ChrW(242) & "ng " & ChrW(273) & ChrW(417) & "n nh" & ChrW(432) & " tr" & ChrW(432) & ChrW(7899) & "c gi" & ChrW(7901)
Case 6: TT_OptNote = "0 " & ChrW(8211) & " 12. Quy " & ChrW(273) & ChrW(7883) & "nh: c" & ChrW(225) & "c " & ChrW(273) & "o" & ChrW(7841) & "n c" & ChrW(225) & "ch nhau t" & ChrW(7889) & "i thi" & ChrW(7875) & "u 6pt"
Case 7: TT_OptNote = "0 " & ChrW(8211) & " 12. " & ChrW(272) & ChrW(7863) & "t 0 " & ChrW(273) & ChrW(7875) & " v" & ChrW(259) & "n b" & ChrW(7843) & "n g" & ChrW(7885) & "n h" & ChrW(417) & "n"
Case 8: TT_OptNote = "1 = h" & ChrW(7841) & " c" & ChrW(7905) & " ch" & ChrW(7919) & " (khuy" & ChrW(234) & "n d" & ChrW(249) & "ng, gi" & ChrW(7919) & " " & ChrW(273) & ChrW(250) & "ng l" & ChrW(7873) & "); 2 = l" & ChrW(7845) & "n l" & ChrW(7873) & " tr" & ChrW(225) & "i (gi" & ChrW(7919) & " c" & ChrW(7905) & " ch" & ChrW(7919) & ")"
Case 9: TT_OptNote = "0 = kh" & ChrW(244) & "ng; 1 = c" & ChrW(243) & ": d" & ChrW(242) & "ng cu" & ChrW(7889) & "i ch" & ChrW(7881) & " c" & ChrW(243) & " 1 t" & ChrW(7915) & " th" & ChrW(236) & " co nh" & ChrW(7865) & " ch" & ChrW(7919) & " " & ChrW(273) & ChrW(7875) & " k" & ChrW(233) & "o l" & ChrW(234) & "n d" & ChrW(242) & "ng tr" & ChrW(234) & "n"
End Select
End If
End Function
Public Function TT_OptDefault(ByVal pid As String, ByVal key As String) As Single
Select Case key
Case "LINE_SPACING": TT_OptDefault = 0
Case "PARA_BEFORE": TT_OptDefault = 6
Case "PARA_AFTER": TT_OptDefault = 6
Case "TEN_DAI"
If GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "HA_CO") = "LAN_LE" Then
TT_OptDefault = 2
Else
TT_OptDefault = 1
End If
Case "GOM_CHU_LE"
If GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "GomChuLe", "0") = "1" Then
TT_OptDefault = 1
Else
TT_OptDefault = 0
End If
Case Else: TT_OptDefault = TT_RawN(pid, key)
End Select
End Function
Private Sub TT_OvLoad(ByVal pid As String)
Dim grp As String, i As Integer, k As String, v As String, s As String, muc As String
grp = TT_TemplateGroupOf(pid)
muc = TT_UnitSection(pid, TT_UnitSlot(pid))
s = "|"
On Error Resume Next
For i = 1 To TT_OptCount(grp)
k = TT_OptKey(grp, i)
v = ""
v = GetSetting(TT_SETTING_APP, muc, "TC_" & k, "")
If Len(v) > 0 Then
If Val(v) >= TT_OptMin(grp, i) - 0.001 And Val(v) <= TT_OptMax(grp, i) + 0.001 Then s = s & k & "=" & v & "|"
End If
Next i
On Error GoTo 0
mTTOvStr = s
mTTOvPid = pid
End Sub
Public Function TT_OvGet(ByVal key As String, ByRef giaTri As Single) As Boolean
TT_OvGet = False
Dim pid As String, vt As Long, ket As Long
pid = TT_Profile()
If mTTOvPid <> pid Or Len(mTTOvStr) = 0 Then Call TT_OvLoad(pid)
vt = InStr(1, mTTOvStr, "|" & key & "=")
If vt = 0 Then Exit Function
vt = vt + Len(key) + 2
ket = InStr(vt, mTTOvStr, "|")
If ket = 0 Then Exit Function
giaTri = Val(Mid(mTTOvStr, vt, ket - vt))
TT_OvGet = True
End Function
Public Function TT_Opt(ByVal key As String) As Single
Dim v As Single
If TT_OvGet(key, v) Then
TT_Opt = v
Else
TT_Opt = TT_OptDefault(TT_Profile(), key)
End If
End Function
Public Function TT_GianTruoc() As Single
TT_GianTruoc = TT_Opt("PARA_BEFORE")
End Function
Public Function TT_GianSau() As Single
TT_GianSau = TT_Opt("PARA_AFTER")
End Function
Private Function TT_LaDongTrong(ByVal p As Paragraph) As Boolean
TT_LaDongTrong = False
If p Is Nothing Then Exit Function
On Error GoTo Thoat
If p.Range.Information(wdWithInTable) Then Exit Function
If Len(Replace(Replace(p.Range.text, vbCr, ""), " ", "")) > 0 Then Exit Function
If p.Range.InlineShapes.Count > 0 Then Exit Function
If p.Range.ShapeRange.Count > 0 Then Exit Function
TT_LaDongTrong = True
Thoat:
End Function
Private Sub TT_DatDongTrong(ByVal p As Paragraph)
On Error Resume Next
With p.Format
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
End With
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = False
On Error GoTo 0
End Sub
Private Sub TT_DongTrongQuanhTieuDe(ByVal pTen As Paragraph, ByVal pKe As Paragraph)
On Error Resume Next
Dim pv As Paragraph, pn As Paragraph
Set pv = Nothing
Set pv = pTen.Previous
If Not pv Is Nothing Then
If TT_LaDongTrong(pv) Then
Call TT_DatDongTrong(pv)
ElseIf InStr(1, pv.Range.text, Chr(12)) = 0 Then
pTen.Range.InsertParagraphBefore
Call TT_DatDongTrong(pTen.Previous)
End If
End If
If Not pKe Is Nothing Then
Set pn = Nothing
Set pn = pKe.Next
If Not pn Is Nothing Then
If TT_LaDongTrong(pn) Then
Call TT_DatDongTrong(pn)
Else
pKe.Range.InsertParagraphAfter
Call TT_DatDongTrong(pKe.Next)
End If
End If
End If
On Error GoTo 0
End Sub
Private Function TT_NeoKeTrichYeu(ByVal pKe As Paragraph) As Paragraph
Set TT_NeoKeTrichYeu = Nothing
On Error GoTo Thoat
If pKe Is Nothing Then Exit Function
If Not IsSeparatorLine(CleanParaText(pKe.Range.text)) Then Exit Function
If pKe.Range.Information(wdWithInTable) Then Exit Function
If pKe.Range.InlineShapes.Count > 0 Then Exit Function
If pKe.Range.ShapeRange.Count > 0 Then Exit Function
Dim pv As Paragraph, buoc As Integer
buoc = 0
Set pv = pKe.Previous
Do While Not pv Is Nothing And buoc < 3
If pv.Range.Information(wdWithInTable) Then Exit Function
If Len(CleanParaText(pv.Range.text)) > 0 Then Exit Do
buoc = buoc + 1
Set pv = pv.Previous
Loop
If pv Is Nothing Then Exit Function
If Len(CleanParaText(pv.Range.text)) = 0 Then Exit Function
If InStr(1, pv.Range.text, Chr(12)) > 0 Then Exit Function
pKe.Range.Delete
Set TT_NeoKeTrichYeu = pv
Thoat:
End Function
Private Sub TT_VeKeDuoiTrichYeu(ByVal doc As Document, ByVal p As Paragraph, ByVal thuTu As Long)
On Error Resume Next
If p Is Nothing Then Exit Sub
Dim vungChu As Single
vungChu = 0
With p.Range.Sections(1).PageSetup
vungChu = .PageWidth - .LeftMargin - .RightMargin
End With
If vungChu <= 0 Or vungChu > 1200 Then vungChu = MillimetersToPoints(165)
Dim soDong As Long
soDong = 1
soDong = p.Range.ComputeStatistics(wdStatisticLines)
If soDong < 1 Then soDong = 1
Dim rongChu As Single, rongDoan As Single, pq As Paragraph, demDoan As Integer
rongChu = 0
demDoan = 0
Set pq = p
Do While Not pq Is Nothing And demDoan < 6
If pq.Range.Information(wdWithInTable) Then Exit Do
If Len(CleanParaText(pq.Range.text)) = 0 Then Exit Do
If pq.Format.Alignment <> wdAlignParagraphCenter Then Exit Do
If pq.Range.ComputeStatistics(wdStatisticLines) > 1 Then
rongDoan = vungChu
Else
rongDoan = TT_EstimateTextWidth(pq)
End If
If rongDoan > rongChu Then rongChu = rongDoan
demDoan = demDoan + 1
Set pq = pq.Previous
Loop
If rongChu <= 0 Then rongChu = TT_EstimateTextWidth(p)
If soDong > 1 Or rongChu <= 0 Or rongChu > vungChu Then rongChu = vungChu
Dim dai As Single
dai = rongChu * 0.4
If dai < CentimetersToPoints(1.5) Then dai = CentimetersToPoints(1.5)
Dim sz As Single
sz = p.Range.Font.Size
If sz <= 0 Or sz > 40 Then sz = 14
Dim trai As Single, tren As Single
trai = (vungChu - dai) / 2
If trai < 0 Then trai = 0
tren = soDong * sz * 1.17 + 6
Call TT_ChenKe(doc, p, "LineTrichYeu" & CStr(thuTu), trai, tren, dai, False)
On Error GoTo 0
End Sub
Private Sub TT_KhoiPhucKeTrichYeu(ByVal doc As Document)
On Error Resume Next
If doc.Shapes.Count = 0 Then Exit Sub
Dim i As Long, shp As Shape, pNeo As Paragraph, ten As String, viTri As Long
For i = doc.Shapes.Count To 1 Step -1
Set shp = Nothing
Set shp = doc.Shapes(i)
If Not shp Is Nothing Then
ten = ""
ten = shp.Name
If ten Like "LineTrichYeu*" Then
Set pNeo = Nothing
Set pNeo = shp.Anchor.Paragraphs(1)
shp.Delete
If Not pNeo Is Nothing Then
If Not pNeo.Range.Information(wdWithInTable) Then
viTri = pNeo.Range.End - 1
doc.Range(viTri, viTri).InsertAfter vbCr & "-----"
End If
End If
End If
End If
Next i
On Error GoTo 0
End Sub
Public Sub TT_ApGianDong(ByVal p As Paragraph, ByVal giaTri As Single, ByVal laDang As Boolean)
On Error Resume Next
If giaTri <= 0 Then
p.Format.LineSpacingRule = wdLineSpaceSingle
ElseIf laDang Then
p.Format.LineSpacingRule = wdLineSpaceExactly
p.Format.LineSpacing = giaTri
ElseIf giaTri <= 1.001 Then
p.Format.LineSpacingRule = wdLineSpaceSingle
Else
p.Format.LineSpacingRule = wdLineSpaceMultiple
p.Format.LineSpacing = LinesToPoints(giaTri)
End If
On Error GoTo 0
End Sub
Public Function TT_GomChuLe() As Boolean
TT_GomChuLe = (TT_Opt("GOM_CHU_LE") >= 0.5)
End Function
Public Sub RibbonGomChuLe_GetPressed(control As Object, ByRef returnedVal)
returnedVal = TT_GomChuLe()
End Sub
Public Sub RibbonGomChuLe_OnAction(control As Object, pressed As Boolean)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "GomChuLe", IIf(pressed, "1", "0")
On Error GoTo 0
End Sub
Private Function TT_SoHienThi(ByVal v As Single) As String
Dim s As String
s = Trim(Str(Round(v, 2)))
If Left(s, 1) = "." Then s = "0" & s
TT_SoHienThi = Replace(s, ".", ",")
End Function
Public Function TT_OptPidCurrent() As String
TT_OptPidCurrent = TT_UnitGroupCurrent()
End Function
Private Function TT_TenKhoiCoDau(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_TenKhoiCoDau = "Kh" & ChrW(7889) & "i " & ChrW(272) & ChrW(7843) & "ng"
Case "CHINH_QUYEN": TT_TenKhoiCoDau = "Kh" & ChrW(7889) & "i Ch" & ChrW(237) & "nh quy" & ChrW(7873) & "n"
Case "CONG_AN": TT_TenKhoiCoDau = "Kh" & ChrW(7889) & "i C" & ChrW(244) & "ng an"
Case "QUAN_DOI": TT_TenKhoiCoDau = "Kh" & ChrW(7889) & "i Qu" & ChrW(226) & "n " & ChrW(273) & ChrW(7897) & "i"
Case "MTTQ": TT_TenKhoiCoDau = "Kh" & ChrW(7889) & "i MTTQ v" & ChrW(224) & " t" & ChrW(7893) & " ch" & ChrW(7913) & "c CT-XH"
Case Else: TT_TenKhoiCoDau = TT_ProfileLabelAscii(pid)
End Select
End Function
Public Sub TT_MoFormTuyChon(ByVal pid As String)
On Error Resume Next
Dim d As Document, grp As String, n As Integer, i As Integer, tb As Table, r As Range
Dim huongDan As String, v As String, nhanLuu As String, pLuu As Paragraph
Set d = TT_FormDoc()
If Not (d Is Nothing) Then d.Close 0
Set d = Nothing
Set d = Documents.Add
If d Is Nothing Then Exit Sub
Dim slot As Integer, muc As String
slot = TT_UnitSlot(pid)
muc = TT_UnitSection(pid, slot)
d.Variables.Add "TT_FormTuyChon", pid & "|" & CStr(slot)
grp = TT_TemplateGroupOf(pid)
n = TT_OptCount(grp)
With d.PageSetup
.TopMargin = 28
.BottomMargin = 28
.LeftMargin = 34
.RightMargin = 34
End With
Set r = d.Range(0, 0)
r.InsertAfter "T" & ChrW(217) & "Y CH" & ChrW(7884) & "N TH" & ChrW(7874) & " TH" & ChrW(7912) & "C" & " - " & TT_TenKhoiCoDau(pid) & vbCr & TT_UnitLabelAt(pid, slot) & vbCr
huongDan = "S" & ChrW(7917) & "a c" & ChrW(7897) & "t " & ChrW(8220) & "Gi" & ChrW(225) & " tr" & ChrW(7883) & ChrW(8221) & " trong kho" & ChrW(7843) & "ng cho ph" & ChrW(233) & "p r" & ChrW(7891) & "i"
huongDan = huongDan & " b" & ChrW(7845) & "m n" & ChrW(250) & "t " & ChrW(8220) & "L" & ChrW(432) & "u b" & ChrW(7843) & "ng nh" & ChrW(7853) & "p" & ChrW(8221) & " tr" & ChrW(234) & "n thanh Ribbon. " & ChrW(272) & ChrW(7875)
huongDan = huongDan & " tr" & ChrW(7889) & "ng m" & ChrW(7897) & "t " & ChrW(244) & " = d" & ChrW(249) & "ng m" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh c" & ChrW(7911) & "a kh" & ChrW(7889) & "i. S" & ChrW(7889) & " in"
huongDan = huongDan & " " & ChrW(273) & ChrW(7853) & "m l" & ChrW(224) & " gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " " & ChrW(273) & ChrW(227) & " t" & ChrW(7921) & " " & ChrW(273) & ChrW(7863) & "t. L" & ChrW(432) & "u " & ChrW(253) & " c" & ChrW(7911) & "a N" & ChrW(272)
huongDan = huongDan & " 30: c" & ChrW(7905) & " ch" & ChrW(7919) & " trong c" & ChrW(249) & "ng m" & ChrW(7897) & "t v" & ChrW(259) & "n b" & ChrW(7843) & "n t" & ChrW(259) & "ng hay"
huongDan = huongDan & " gi" & ChrW(7843) & "m ph" & ChrW(7843) & "i th" & ChrW(7889) & "ng nh" & ChrW(7845) & "t. " & ChrW(272) & ChrW(226) & "y ch" & ChrW(7881) & " l" & ChrW(224) & " b" & ChrW(7843) & "ng nh" & ChrW(7853) & "p"
huongDan = huongDan & " t" & ChrW(7841) & "m: KH" & ChrW(212) & "NG c" & ChrW(7847) & "n l" & ChrW(432) & "u file n" & ChrW(224) & "y."
r.InsertAfter huongDan & vbCr
d.Paragraphs(1).Range.Font.Bold = True
d.Paragraphs(1).Range.Font.Size = 13
d.Paragraphs(2).Range.Font.Bold = True
d.Paragraphs(2).Range.Font.Size = 12
d.Paragraphs(3).Range.Font.Italic = True
d.Paragraphs(3).Range.Font.Size = 11
d.Content.Font.Name = "Times New Roman"
d.Content.ParagraphFormat.SpaceAfter = 4
Set r = d.Range(d.Content.End - 1, d.Content.End - 1)
Set tb = d.Tables.Add(r, n + 1, 3)
tb.Borders.Enable = True
tb.Range.Font.Name = "Times New Roman"
tb.Range.Font.Size = 12
tb.Range.Font.Bold = False
tb.Range.Font.Italic = False
tb.Range.ParagraphFormat.SpaceAfter = 0
tb.AllowAutoFit = False
tb.Columns(1).Width = 200
tb.Columns(2).Width = 70
tb.Columns(3).Width = 240
tb.Cell(1, 1).Range.text = "Th" & ChrW(244) & "ng s" & ChrW(7889)
tb.Cell(1, 2).Range.text = "Gi" & ChrW(225) & " tr" & ChrW(7883)
tb.Cell(1, 3).Range.text = ChrW(272) & ChrW(432) & ChrW(7907) & "c ph" & ChrW(233) & "p"
tb.Rows(1).Range.Font.Bold = True
tb.Rows(1).Shading.BackgroundPatternColor = RGB(217, 217, 217)
For i = 1 To n
tb.Cell(i + 1, 1).Range.text = TT_OptLabel(grp, i)
tb.Cell(i + 1, 1).Shading.BackgroundPatternColor = RGB(242, 242, 242)
v = ""
v = GetSetting(TT_SETTING_APP, muc, "TC_" & TT_OptKey(grp, i), "")
If Len(v) > 0 Then
tb.Cell(i + 1, 2).Range.text = TT_SoHienThi(Val(v))
tb.Cell(i + 1, 2).Range.Font.Bold = True
ElseIf TT_OptKey(grp, i) <> "LINE_SPACING" Then
tb.Cell(i + 1, 2).Range.text = TT_SoHienThi(TT_OptDefault(pid, TT_OptKey(grp, i)))
End If
tb.Cell(i + 1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
tb.Cell(i + 1, 3).Range.text = TT_OptNote(grp, i)
tb.Cell(i + 1, 3).Range.Font.Italic = True
tb.Cell(i + 1, 3).Range.Font.Size = 11
Next i
d.ActiveWindow.WindowState = wdWindowStateNormal
d.ActiveWindow.Caption = "B" & ChrW(7843) & "ng nh" & ChrW(7853) & "p t" & ChrW(7841) & "m - kh" & ChrW(244) & "ng c" & ChrW(7847) & "n l" & ChrW(432) & "u file"
d.ActiveWindow.Width = 760
d.ActiveWindow.Height = 760
d.ActiveWindow.View.Type = wdPrintView
d.ActiveWindow.View.Zoom.Percentage = 100
tb.Cell(2, 2).Range.Select
Selection.Collapse wdCollapseStart
d.Saved = True
Call TT_RibbonRefresh
DoEvents
If Not (mTTRibbon Is Nothing) Then mTTRibbon.ActivateTab "tabVanBanDang"
On Error GoTo 0
End Sub
Public Function TT_LuuFormTuyChon(ByVal d As Document, ByVal showMsg As Boolean) As Boolean
TT_LuuFormTuyChon = False
On Error Resume Next
Dim pid As String, grp As String, n As Integer, i As Integer, t As String, tb As Table, v As Single
Dim loi As String, k As String, macDinh As Single
Dim slot As Integer, muc As String, vtBar As Long
pid = d.Variables("TT_FormTuyChon").value
vtBar = InStr(1, pid, "|")
slot = 1
If vtBar > 0 Then
slot = Val(Mid(pid, vtBar + 1))
pid = Left(pid, vtBar - 1)
End If
muc = TT_UnitSection(pid, slot)
grp = TT_TemplateGroupOf(pid)
n = TT_OptCount(grp)
If d.Tables.Count = 0 Then Exit Function
Set tb = d.Tables(1)
If tb.Rows.Count < n + 1 Then
If showMsg Then MsgBox "Bang tuy chon da bi sua mat dong. Hay dong cua so nay va mo lai bang tuy chon.", vbExclamation, "Tuy chon the thuc"
Exit Function
End If
For i = 1 To n
k = TT_OptKey(grp, i)
t = tb.Cell(i + 1, 2).Range.text
t = Replace(Replace(Replace(Replace(t, Chr(7), ""), vbCr, ""), Chr(11), ""), " ", "")
t = Replace(t, ",", ".")
If Len(t) = 0 Then
DeleteSetting TT_SETTING_APP, muc, "TC_" & k
Else
v = Val(t)
If t Like "*[!0-9.]*" Or t Like "*.*.*" Or v < TT_OptMin(grp, i) - 0.001 Or v > TT_OptMax(grp, i) + 0.001 Then
loi = loi & vbCrLf & " - " & ToAscii(TT_OptLabel(grp, i)) & ": '" & t & "' (cho phep " & TT_SoHienThi(TT_OptMin(grp, i)) & " - " & TT_SoHienThi(TT_OptMax(grp, i)) & ")"
Else
macDinh = TT_OptDefault(pid, k)
If Abs(v - macDinh) < 0.001 And k <> "LINE_SPACING" Then
DeleteSetting TT_SETTING_APP, muc, "TC_" & k
Else
SaveSetting TT_SETTING_APP, muc, "TC_" & k, Trim(Str(v))
End If
End If
End If
Next i
mTTOvPid = ""
mTTOvStr = ""
If Len(loi) > 0 Then
If showMsg Then MsgBox "Cac dong sau co gia tri ngoai khoang quy dinh cho phep nen CHUA duoc luu (cac dong khac da luu):" & loi & vbCrLf & vbCrLf & "Hay sua lai roi bam Luu lan nua.", vbExclamation, "Tuy chon the thuc"
Exit Function
End If
Call TT_DongBangNhap(d)
Call TT_RibbonRefresh
TT_LuuFormTuyChon = True
If showMsg Then MsgBox "Da luu tuy chon the thuc cho don vi: " & ToAscii(TT_UnitLabelAt(pid, slot)) & " (khoi " & TT_ProfileLabelAscii(pid) & ")." & vbCrLf & "Tuy chon ap dung tu lan bam Chuan hoa tiep theo." & vbCrLf & "Co the dong cua so bang nhap (khong can luu file).", vbInformation, "Tuy chon the thuc"
On Error GoTo 0
End Function
Public Sub RibbonTuyChon_Mo(control As Object)
Call TT_MoFormTuyChon(TT_OptPidCurrent())
End Sub
Private Sub TT_GomChuLeCuoiDoan(ByVal doc As Document)
On Error Resume Next
Dim p As Paragraph, t As String, doDai As Long, vt As Long, viTriDau As Long
Dim rA As Range, rB As Range, y1 As Single, y2 As Single, k As Integer, daGom As Boolean, sp As Single
For Each p In doc.Paragraphs
t = p.Range.text
doDai = Len(t) - 1
If doDai >= 90 Then
If p.Format.Alignment = wdAlignParagraphJustify Then
If Not p.Range.Information(wdWithInTable) Then
sp = 0
sp = p.Range.Font.Spacing
If sp < 0 And sp >= -0.5 Then p.Range.Font.Spacing = 0
vt = InStrRev(t, " ", doDai)
If vt > 20 And doDai - vt <= 12 And p.Range.End - p.Range.Start = Len(t) Then
viTriDau = p.Range.Start
Set rA = doc.Range(viTriDau + vt - 2, viTriDau + vt - 2)
Set rB = doc.Range(viTriDau + vt, viTriDau + vt)
y1 = -1: y2 = -1
y1 = rA.Information(wdVerticalPositionRelativeToPage)
y2 = rB.Information(wdVerticalPositionRelativeToPage)
If y1 >= 0 And y2 >= 0 And Abs(y1 - y2) > 1 Then
daGom = False
For k = 1 To 4
p.Range.Font.Spacing = -0.1 * k
y1 = -1: y2 = -1
y1 = rA.Information(wdVerticalPositionRelativeToPage)
y2 = rB.Information(wdVerticalPositionRelativeToPage)
If y1 >= 0 And y2 >= 0 And Abs(y1 - y2) <= 1 Then
daGom = True
Exit For
End If
Next k
If Not daGom Then p.Range.Font.Spacing = 0
End If
End If
End If
End If
End If
Next p
On Error GoTo 0
End Sub
Public Function TT_B(ByVal key As String) As Boolean
TT_B = TT_RawB(TT_Profile(), key)
End Function
Public Function TT_S(ByVal key As String) As String
TT_S = TT_RawS(TT_Profile(), key)
End Function
Public Function TT_IsNational() As Boolean
TT_IsNational = (TT_S("MOTTO_KIND") <> "DANG")
End Function
Public Function TT_RightColumnAlignment() As Long
If TT_S("RIGHT_ALIGN") = "CENTER" Then
TT_RightColumnAlignment = wdAlignParagraphCenter
Else
TT_RightColumnAlignment = wdAlignParagraphRight
End If
End Function
Public Function TT_DocCodeTitle(ByVal codeAsc As String) As String
TT_DocCodeTitle = TT_RawDocCodeTitle(TT_Profile(), UCase(Trim(codeAsc)))
End Function
Public Function TT_ReportHeader() As String
Dim s As String
s = "Khoi the thuc: " & TT_ProfileLabelAscii(TT_Profile())
Select Case mTTSource
Case "TU_NHAN_DIEN": s = s & " (tu dong nhan dien)"
Case "MAC_DINH": s = s & " (khong nhan dien duoc, dung khoi mac dinh)"
Case Else: s = s & " (nguoi dung chon)"
End Select
s = s & vbCrLf & "Can cu: " & TT_LegalBasisAscii(TT_Profile())
Select Case TT_S("STATUS")
Case "experimental": s = s & vbCrLf & "Luu y: bo the thuc khoi nay dang THU NGHIEM, chi kiem tra noi long."
Case "needs_review": s = s & vbCrLf & "Luu y: bo the thuc khoi nay CAN DOI CHIEU THEM van ban goc."
End Select
TT_ReportHeader = s
End Function
Public Function TT_MismatchWarning() As String
TT_MismatchWarning = ""
If mTTSource <> "CHON" Or Not mTTDetectedStrong Or Len(mTTDetected) = 0 Then Exit Function
If mTTDetected = mTTProfile Then Exit Function
If (TT_RawS(mTTDetected, "MOTTO_KIND") = "DANG") <> (TT_RawS(mTTProfile, "MOTTO_KIND") = "DANG") Then
TT_MismatchWarning = "[Khoi the thuc] Dang chon khoi '" & TT_ProfileLabelAscii(mTTProfile) & _
"' nhung dong tieu ngu tren van ban thuoc khoi '" & TT_ProfileLabelAscii(mTTDetected) & _
"'. Kiem tra lai muc Khoi tren Ribbon (hoac chon Tu dong nhan dien)."
End If
End Function
Private Function TT_EstimateTextWidth(ByVal p As Paragraph) As Single
Dim t As String
t = CleanParaText(p.Range.text)
Dim sz As Single
sz = p.Range.Font.Size
If sz <= 0 Or sz > 40 Then sz = 13
If InStr(1, Replace(ToAscii(t), " ", ""), "DOCLAP-TUDO-HANHPHUC") > 0 Then
TT_EstimateTextWidth = CentimetersToPoints(5.95) * sz / 14
ElseIf UCase(t) = t Then
TT_EstimateTextWidth = Len(t) * sz * 0.62
Else
TT_EstimateTextWidth = Len(t) * sz * 0.48
End If
End Function
Private Function TT_ParaLines(ByVal p As Paragraph) As Long
TT_ParaLines = 1
If p Is Nothing Then Exit Function
On Error Resume Next
Dim n As Long, r As Range
n = 0
n = p.Range.ComputeStatistics(wdStatisticLines)
If n < 1 Then
Set r = p.Range.Duplicate
r.MoveEnd wdCharacter, -1
n = r.ComputeStatistics(wdStatisticLines)
End If
If n < 1 Then
Dim y1 As Single, y2 As Single
Set r = p.Range.Duplicate
r.MoveEnd wdCharacter, -1
y1 = r.Characters(1).Information(wdVerticalPositionRelativeToPage)
y2 = r.Characters(r.Characters.Count).Information(wdVerticalPositionRelativeToPage)
If y2 - y1 > p.Range.Font.Size * 0.6 Then n = 2 Else n = 1
End If
If n < 1 Then n = 1
TT_ParaLines = n
End Function
Private Function TT_OrganWraps(ByVal organParas As Collection) As Boolean
TT_OrganWraps = False
Dim o As Variant
On Error Resume Next
For Each o In organParas
If TT_ParaLines(o) > 1 Then
TT_OrganWraps = True
Exit Function
End If
Next o
End Function
Private Sub TT_WidenOrganCol(ByVal tbl As Table, ByVal organParas As Collection, ByVal pKey As Paragraph, ByVal pDate As Paragraph, ByVal totalW As Single, ByRef rightW As Single)
Dim k As Integer
On Error Resume Next
For k = 1 To 25
If Not TT_OrganWraps(organParas) Then Exit Sub
rightW = rightW - MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
If TT_ParaLines(pKey) > 1 Or TT_ParaLines(pDate) > 1 Then
rightW = rightW + MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Exit Sub
End If
Next k
End Sub
Private Function TT_TenDaiMode() As String
If TT_Opt("TEN_DAI") >= 1.5 Then
TT_TenDaiMode = "LAN_LE"
Else
TT_TenDaiMode = "HA_CO"
End If
End Function
Public Sub RibbonTenDai_GetSelectedIndex(control As Object, ByRef returnedVal)
If TT_TenDaiMode() = "LAN_LE" Then returnedVal = 1 Else returnedVal = 0
End Sub
Public Sub RibbonTenDai_OnAction(control As Object, selectedId As String, selectedIndex As Integer)
On Error Resume Next
If selectedIndex = 1 Then
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "LAN_LE"
Else
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "TenDai", "HA_CO"
End If
End Sub
Private Sub TT_LanLeTrai(ByVal tbl As Table, ByVal organParas As Collection, ByRef totalW As Single, ByVal rightW As Single, ByRef lanW As Single)
Dim k As Integer, baseIndent As Single
On Error Resume Next
baseIndent = tbl.Rows.LeftIndent + lanW
For k = 1 To 8
If Not TT_OrganWraps(organParas) Then Exit Sub
If lanW + MillimetersToPoints(2) > MillimetersToPoints(15) + 0.5 Then Exit Sub
lanW = lanW + MillimetersToPoints(2)
totalW = totalW + MillimetersToPoints(2)
tbl.Rows.LeftIndent = baseIndent - lanW
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Call TT_DebugLog("  Lan le: " & Format(PointsToMillimeters(lanW), "0") & "mm, indent " & Format(PointsToMillimeters(tbl.Rows.LeftIndent), "0.0") & "mm, cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0") & ", o " & Format(PointsToMillimeters(tbl.Cell(1, 1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Cell(1, 2).Width), "0"))
Next k
End Sub
Private Sub TT_NarrowForRight(ByVal tbl As Table, ByRef organParas As Collection, ByVal pRight As Paragraph, ByVal totalW As Single, ByRef rightW As Single)
Dim k As Integer
On Error Resume Next
For k = 1 To 15
If TT_ParaLines(pRight) <= 1 Then Exit Sub
If totalW - rightW - MillimetersToPoints(2) < MillimetersToPoints(52) Then Exit Sub
rightW = rightW + MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
If TT_OrganWraps(organParas) Then Set organParas = TT_SplitLongOrgans(organParas, False)
If TT_OrganWraps(organParas) Then
rightW = rightW - MillimetersToPoints(2)
tbl.Columns(1).Width = totalW - rightW
tbl.Columns(2).Width = rightW
Exit Sub
End If
Next k
End Sub
Private Function TT_SplitLongOrgans(ByVal organParas As Collection, ByVal standardOnly As Boolean) As Collection
Dim outC As New Collection
Dim o As Variant, p As Paragraph
On Error Resume Next
For Each o In organParas
Set p = o
outC.Add p
If TT_ParaLines(p) > 1 Then
Dim t As String, a As String, pos As Long, raw As String, off As Long
raw = p.Range.text
t = CleanParaText(raw)
a = UCase(ToAscii(t))
pos = 0
If Len(a) = Len(t) Then
If a Like "UY BAN NHAN DAN *" Then
pos = 16
ElseIf a Like "HOI DONG NHAN DAN *" Then
pos = 18
ElseIf a Like "BO CHI HUY QUAN SU *" Then
pos = 19
ElseIf Not standardOnly Then
Dim kw As Variant
For Each kw In Array(" THANH PHO ", " TINH ", " DAC KHU ", " PHUONG ", " XA ")
pos = InStr(2, a, kw)
If pos > 0 Then Exit For
Next kw
End If
End If
If pos > 1 Then
off = InStr(1, raw, t) - 1
If off < 0 Then off = 0
Dim r As Range
Set r = p.Range.Duplicate
r.SetRange p.Range.Start + off + pos - 1, p.Range.Start + off + pos
If r.text = " " Then
Dim pStart As Long, p1 As Paragraph
pStart = p.Range.Start
r.text = vbCr
Set p1 = p.Range.Document.Range(pStart, pStart).Paragraphs(1)
outC.Remove outC.Count
outC.Add p1
outC.Add p1.Next
End If
End If
End If
Next o
Set TT_SplitLongOrgans = outC
End Function
Private Sub TT_InDamTenCoQuan(ByVal organParas As Collection)
On Error Resume Next
Dim n As Long
n = organParas.Count
If n = 0 Then Exit Sub
Dim nhom() As Long
ReDim nhom(1 To n)
Dim i As Long, soNhom As Long
soNhom = 1
nhom(1) = 1
For i = 2 To n
If TT_DongNoiTiepTen(CleanParaText(organParas(i - 1).Range.text), CleanParaText(organParas(i).Range.text)) Then
nhom(i) = soNhom
Else
soNhom = soNhom + 1
nhom(i) = soNhom
End If
Next i
For i = 1 To n
organParas(i).Range.Font.Bold = (soNhom = 1 Or nhom(i) > 1)
Next i
On Error GoTo 0
End Sub
Private Function TT_DongNoiTiepTen(ByVal truoc As String, ByVal sau As String) As Boolean
TT_DongNoiTiepTen = False
Dim a As String, b As String
a = Trim(ToAscii(truoc))
b = Trim(ToAscii(sau))
If Len(a) = 0 Or Len(b) = 0 Then Exit Function
If Left(Trim(sau), 1) = "[" Or (Left(Trim(truoc), 1) = "[" And Right(Trim(truoc), 1) = "]") Then Exit Function
Dim ds As Variant, k As Integer
ds = Array("DANG UY", "DANG BO", "CHI BO", "UY BAN", "HOI DONG", "SO ", "BAN ", "PHONG ", "VAN PHONG", "CONG AN", _
"BO CHI HUY", "BO TU LENH", "TINH UY", "THANH UY", "HUYEN UY", "QUAN UY", "THI UY", "DOAN ", "HOI ", "TRUONG ", _
"TRUNG TAM", "CHI CUC", "CUC ", "TONG CUC", "VU ", "MAT TRAN", "THANH TRA", "KHO BAC", "BENH VIEN", "TRAM ", _
"VIEN ", "HOC VIEN", "LIEN DOAN", "CONG DOAN", "QUY ", "NGAN HANG", "BAO ", "DAI ", "TOA AN", "DOI ", _
"TIEU BAN", "CHI NHANH", "CONG TY", "BO ", "KHOI ", "LIEN MINH", "HOP TAC XA", "DON VI", "TIEU DOAN", "TRUNG DOAN", _
"SU DOAN", "LU DOAN", "DAI DOI", "TRUNG DOI", "CUM ", "KHU ", "TO ")
For k = LBound(ds) To UBound(ds)
If Left(b, Len(ds(k))) = ds(k) Then Exit Function
Next k
TT_DongNoiTiepTen = True
End Function
Private Function TT_ChenKe(ByVal doc As Document, ByVal p As Paragraph, ByVal tenHinh As String, ByVal trai As Single, _
ByVal tren As Single, ByVal dai As Single, ByVal trongBang As Boolean) As Shape
Set TT_ChenKe = Nothing
On Error Resume Next
Dim ke As Shape
Set ke = Nothing
Dim tenTam As String, soThem As Long, daCo As Boolean, shK As Shape
tenTam = tenHinh
soThem = 1
Do
daCo = False
For Each shK In doc.Shapes
If shK.Name = tenTam Then daCo = True: Exit For
Next shK
If Not daCo Then Exit Do
soThem = soThem + 1
tenTam = tenHinh & "_" & CStr(soThem)
Loop While soThem < 50
Dim soDoanTruoc As Long
soDoanTruoc = doc.Paragraphs.Count
Dim xml As String
xml = "<?xml version=""1.0"" standalone=""yes""?>" & _
"<w:wordDocument xmlns:w=""http://schemas.microsoft.com/office/word/2003/wordml"" xmlns:v=""urn:schemas-microsoft-com:vml"" " & _
"xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:w10=""urn:schemas-microsoft-com:office:word"">" & _
"<w:body><w:p><w:r><w:pict><v:line id=""" & tenTam & """ " & _
"style=""position:absolute;z-index:1;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;" & _
"mso-position-vertical:absolute;mso-position-vertical-relative:text"" " & _
"from=""" & TT_SoPt(trai) & "," & TT_SoPt(tren) & """ to=""" & TT_SoPt(trai + dai) & "," & TT_SoPt(tren) & """ " & _
"strokecolor=""black"" strokeweight="".75pt""" & IIf(trongBang, " o:allowincell=""t""", "") & ">" & _
"<w10:anchorlock/></v:line></w:pict></w:r></w:p></w:body></w:wordDocument>"
Dim r As Range
Set r = p.Range.Duplicate
r.Collapse wdCollapseStart
Err.Clear
r.InsertXML xml
If Err.Number = 0 Then
If doc.Paragraphs.Count > soDoanTruoc Then
Dim pThua As Paragraph
Set pThua = p.Previous
If Not pThua Is Nothing Then
If Len(CleanParaText(pThua.Range.text)) = 0 And pThua.Range.ShapeRange.Count > 0 Then
pThua.Range.Characters(pThua.Range.Characters.Count).Delete
End If
End If
End If
Dim sH As Shape
For Each sH In doc.Shapes
If sH.Name = tenTam Then
If sH.Anchor.Start >= p.Range.Start - 1 And sH.Anchor.Start <= p.Range.End Then
Set ke = sH
Else
sH.Delete
End If
Exit For
End If
Next sH
End If
Err.Clear
If Not ke Is Nothing Then
With ke
.LockAnchor = True
If trongBang Then .LayoutInCell = True
.Line.Weight = 0.75
.Line.ForeColor.RGB = RGB(0, 0, 0)
.Line.DashStyle = 1
.WrapFormat.Type = 3
End With
Else
Set ke = doc.Shapes.AddLine(trai, tren, trai + dai, tren, p.Range)
If ke Is Nothing Then Exit Function
With ke
.Name = tenTam
.LockAnchor = True
If trongBang Then .LayoutInCell = True
.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
.Left = trai
.Top = tren
.Line.Weight = 0.75
.Line.ForeColor.RGB = RGB(0, 0, 0)
.Line.DashStyle = 1
.WrapFormat.Type = 3
End With
End If
Set TT_ChenKe = ke
On Error GoTo 0
End Function
Private Sub TT_DoiDauDauDong(ByVal doc As Document)
On Error Resume Next
Dim p As Paragraph
For Each p In doc.Paragraphs
Dim t As String
t = p.Range.text
If Len(t) >= 3 Then
Dim ma As Long
ma = AscW(Left(t, 1)) And &HFFFF&
Dim thay As String
thay = ""
Select Case ma
Case &H2022&, &H25CF&, &H25AA&, &H25A0&, &H2219&, &HB7&, &H2023&, &H2043&, &H25BA&, &H25B6&, _
&H27A2&, &H27A4&, &H2756&, &H2713&, &H2714&, &H2611&, &H2610&, &H25C6&, &H25C7&, &H2666&, _
&HF0B7&, &HF0A7&, &HF076&, &HF0D8&, &HF0FC&, &HF06E&, &HF0A8&
thay = "-"
Case &H2013&, &H2014&, &H2212&
If Mid(t, 2, 1) = " " Or Mid(t, 2, 1) = vbTab Then thay = "-"
Case &H25E6&, &H25CB&, &H25A1&, &HF06F&
thay = "+"
Case &H2A&
If Mid(t, 2, 1) = " " Then thay = "-"
End Select
If Len(thay) > 0 Then
Dim k As Long, ch As String
k = 2
Do While k < Len(t)
ch = Mid(t, k, 1)
If ch = " " Or ch = vbTab Or ch = ChrW(160) Or ch = ChrW(8194) Or ch = ChrW(8195) Or ch = ChrW(8201) Then
k = k + 1
Else
Exit Do
End If
Loop
If k < Len(t) And Asc(Mid(t, k, 1)) <> 13 Then
Dim r As Range
Set r = p.Range.Duplicate
r.SetRange p.Range.Start, p.Range.Start + k - 1
r.text = thay & " "
r.Font.Name = "Times New Roman"
End If
End If
End If
Next p
On Error GoTo 0
End Sub
Private Function TT_SoPt(ByVal v As Single) As String
TT_SoPt = Replace(Trim(Str(Round(v, 2))), ",", ".") & "pt"
End Function
Private Sub TT_AddCenteredLineUnder(ByVal doc As Document, ByVal p As Paragraph, ByVal shapeName As String, ByVal ratio As Single, ByVal fallbackWidthPt As Single)
On Error Resume Next
Dim c As Cell
Set c = Nothing
Set c = p.Range.Cells(1)
If c Is Nothing Then Exit Sub
Dim lPad As Single, rPad As Single
lPad = 5.4: rPad = 5.4
lPad = c.LeftPadding
rPad = c.RightPadding
Dim textAreaW As Single
textAreaW = c.Width - lPad - rPad
If textAreaW <= 0 Or textAreaW > 600 Then textAreaW = 240
Dim textW As Single
textW = TT_EstimateTextWidth(p)
If textW <= 0 Then textW = fallbackWidthPt
If textW > textAreaW Then textW = textAreaW
Dim lineW As Single
lineW = textW * ratio
If lineW < 28 Then lineW = 28
Dim nLines As Long
nLines = 1
nLines = p.Range.ComputeStatistics(wdStatisticLines)
If nLines < 1 Then nLines = 1
Dim lPos As Single, tPos As Single
lPos = (textAreaW - lineW) / 2
If lPos < 0 Then lPos = 0
tPos = nLines * p.Range.Font.Size * 1.17 + 4
Call TT_ChenKe(doc, p, shapeName, lPos, tPos, lineW, True)
On Error GoTo 0
End Sub
Private Function TT_FindSignerName(ByVal doc As Document) As String
TT_FindSignerName = ""
On Error Resume Next
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.text)
If IsFooterSignatureTable(tbl, tAsc) Then
Dim c As Cell
For Each c In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(c.Range.text))
If InStr(1, cAsc, "NOI NHAN") = 0 And HasSignatureKeywordInCell(cAsc) Then
Dim pS As Paragraph
Dim lastText As String
lastText = ""
For Each pS In c.Range.Paragraphs
If Len(CleanParaText(pS.Range.text)) > 0 Then lastText = CleanParaText(pS.Range.text)
Next pS
If Len(lastText) > 0 Then
If Not IsSignatureRoleOrAuth(Trim(ToAscii(lastText))) Then
TT_FindSignerName = lastText
End If
End If
Exit Function
End If
Next c
End If
Next tbl
On Error GoTo 0
End Function
Private Function TT_HasRankPrefix(ByVal nameAsc As String) As Boolean
Dim t As String
t = Trim(UCase(nameAsc))
TT_HasRankPrefix = (t Like "DAI TUONG *" Or t Like "THUONG TUONG *" Or t Like "TRUNG TUONG *" Or t Like "THIEU TUONG *" Or _
t Like "DAI TA *" Or t Like "THUONG TA *" Or t Like "TRUNG TA *" Or t Like "THIEU TA *" Or _
t Like "DAI UY *" Or t Like "THUONG UY *" Or t Like "TRUNG UY *" Or t Like "THIEU UY *" Or _
t Like "DO DOC *" Or t Like "PHO DO DOC *" Or t Like "CHUAN DO DOC *")
End Function
Private Sub TT_DebugLog(ByVal msg As String)
On Error Resume Next
Dim baseDir As String
baseDir = Environ("HOME")
If Len(baseDir) = 0 Then Exit Sub
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim f As Integer
f = FreeFile
Open baseDir & "/thethuc_debug.log" For Append As #f
Print #f, Format(Now, "hh:nn:ss") & " [" & Format(Timer, "0.00") & "] " & msg
Close #f
On Error GoTo 0
End Sub
Private Function TT_EncU(ByVal s As String) As String
Dim i As Long, w As Long, r As String
r = ""
For i = 1 To Len(s)
w = AscW(Mid$(s, i, 1))
If w < 0 Then w = w + 65536
r = r & Right$("0000" & Hex$(w), 4)
Next i
TT_EncU = r
End Function
Private Function TT_DecU(ByVal h As String) As String
Dim i As Long, r As String
r = ""
On Error GoTo DecFail
For i = 1 To Len(h) - 3 Step 4
r = r & ChrW(CLng("&H" & Mid$(h, i, 4)))
Next i
TT_DecU = r
Exit Function
DecFail:
TT_DecU = r
End Function
Private Function TT_UnitMax() As Integer
TT_UnitMax = 9
End Function
Private Function TT_UnitSection(ByVal grp As String, ByVal n As Integer) As String
If n <= 1 Then
TT_UnitSection = "DonVi_" & grp
Else
TT_UnitSection = "DonVi_" & grp & "_" & CStr(n)
End If
End Function
Private Sub TT_UnitTach(ByVal grp As String)
If InStr(1, mTTDaTach, "|" & grp & "|") > 0 Then Exit Sub
mTTDaTach = mTTDaTach & "|" & grp & "|"
If grp = "DANG" Or TT_IndexOfProfile(grp) = 0 Then Exit Sub
On Error Resume Next
If GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "DaTachDonVi_" & grp, "") = "1" Then Exit Sub
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "DaTachDonVi_" & grp, "1"
If Len(GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & grp, "")) > 0 Then Exit Sub
If Len(GetSetting(TT_SETTING_APP, "DonVi_" & grp, "COQUAN", "")) > 0 Then Exit Sub
Dim soCu As Integer, n As Integer, i As Integer, v As String
soCu = Val(GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_NHA_NUOC", "1"))
If soCu < 1 Then soCu = 1
If soCu > TT_UnitMax() Then soCu = TT_UnitMax()
For n = 1 To soCu
For i = 1 To 7
v = ""
v = GetSetting(TT_SETTING_APP, TT_UnitSection("NHA_NUOC", n), TT_UnitFieldAt(i), "")
If Len(v) > 0 Then SaveSetting TT_SETTING_APP, TT_UnitSection(grp, n), TT_UnitFieldAt(i), v
Next i
Next n
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & grp, CStr(soCu)
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & grp, GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_NHA_NUOC", "1")
On Error GoTo 0
End Sub
Public Function TT_UnitCount(ByVal grp As String) As Integer
Dim v As Integer
Call TT_UnitTach(grp)
On Error Resume Next
v = Val(GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & grp, "1"))
On Error GoTo 0
If v < 1 Then v = 1
If v > TT_UnitMax() Then v = TT_UnitMax()
TT_UnitCount = v
End Function
Public Function TT_UnitSlot(ByVal grp As String) As Integer
Dim v As Integer
Call TT_UnitTach(grp)
On Error Resume Next
v = Val(GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & grp, "1"))
On Error GoTo 0
If v < 1 Then v = 1
If v > TT_UnitCount(grp) Then v = TT_UnitCount(grp)
TT_UnitSlot = v
End Function
Public Function TT_UnitGetAt(ByVal grp As String, ByVal n As Integer, ByVal field As String) As String
Dim v As String
Call TT_UnitTach(grp)
On Error Resume Next
v = GetSetting(TT_SETTING_APP, TT_UnitSection(grp, n), field, "")
On Error GoTo 0
TT_UnitGetAt = TT_DecU(v)
End Function
Public Sub TT_UnitSetAt(ByVal grp As String, ByVal n As Integer, ByVal field As String, ByVal value As String)
Call TT_UnitTach(grp)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_UnitSection(grp, n), field, TT_EncU(Trim(value))
On Error GoTo 0
End Sub
Public Function TT_UnitGet(ByVal grp As String, ByVal field As String) As String
TT_UnitGet = TT_UnitGetAt(grp, TT_UnitSlot(grp), field)
End Function
Public Sub TT_UnitSet(ByVal grp As String, ByVal field As String, ByVal value As String)
Call TT_UnitSetAt(grp, TT_UnitSlot(grp), field, value)
End Sub
Private Function TT_UnitFieldAt(ByVal i As Integer) As String
Select Case i
Case 1: TT_UnitFieldAt = "CAPTREN"
Case 2: TT_UnitFieldAt = "COQUAN"
Case 3: TT_UnitFieldAt = "VIETTAT"
Case 4: TT_UnitFieldAt = "DIADANH"
Case 5: TT_UnitFieldAt = "THAMQUYEN"
Case 6: TT_UnitFieldAt = "CHUCVU"
Case 7: TT_UnitFieldAt = "HOTEN"
Case Else: TT_UnitFieldAt = ""
End Select
End Function
Public Function TT_UnitLabelAt(ByVal grp As String, ByVal n As Integer) As String
Dim ten As String, vt As String, nhan As String
ten = Trim(Replace(TT_UnitGetAt(grp, n, "COQUAN"), "|", " "))
Do While InStr(1, ten, "  ") > 0
ten = Replace(ten, "  ", " ")
Loop
vt = Trim(TT_UnitGetAt(grp, n, "VIETTAT"))
If Len(ten) = 0 And Len(vt) = 0 Then
nhan = ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & " " & CStr(n) & " (ch" & ChrW(432) & "a nh" & ChrW(7853) & "p)"
ElseIf Len(vt) = 0 Then
nhan = ten
ElseIf Len(ten) = 0 Then
nhan = vt
Else
nhan = vt & " - " & ten
End If
If Len(nhan) > 60 Then nhan = Left(nhan, 58) & "..."
TT_UnitLabelAt = nhan
End Function
Public Sub TT_UnitSelect(ByVal grp As String, ByVal n As Integer)
If n < 1 Then n = 1
If n > TT_UnitCount(grp) Then n = TT_UnitCount(grp)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & grp, CStr(n)
On Error GoTo 0
mTTOvPid = ""
mTTOvStr = ""
End Sub
Private Sub TT_UnitCopySlot(ByVal grp As String, ByVal tu As Integer, ByVal den As Integer)
Dim i As Integer, k As String, v As String, nhom As String
For i = 1 To 7
Call TT_UnitSetAt(grp, den, TT_UnitFieldAt(i), TT_UnitGetAt(grp, tu, TT_UnitFieldAt(i)))
Next i
nhom = TT_TemplateGroupOf(grp)
On Error Resume Next
For i = 1 To TT_OptCount(nhom)
k = "TC_" & TT_OptKey(nhom, i)
v = ""
v = GetSetting(TT_SETTING_APP, TT_UnitSection(grp, tu), k, "")
If Len(v) > 0 Then
SaveSetting TT_SETTING_APP, TT_UnitSection(grp, den), k, v
Else
DeleteSetting TT_SETTING_APP, TT_UnitSection(grp, den), k
End If
Next i
On Error GoTo 0
mTTOvPid = ""
mTTOvStr = ""
End Sub
Public Function TT_UnitAdd(ByVal grp As String, ByVal tenCoQuan As String) As Integer
TT_UnitAdd = 0
Dim n As Integer, cu As Integer
n = TT_UnitCount(grp)
If n >= TT_UnitMax() Then Exit Function
cu = TT_UnitSlot(grp)
n = n + 1
Call TT_UnitCopySlot(grp, cu, n)
Call TT_UnitSetAt(grp, n, "COQUAN", tenCoQuan)
Call TT_UnitSetAt(grp, n, "VIETTAT", "")
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & grp, CStr(n)
On Error GoTo 0
Call TT_UnitSelect(grp, n)
TT_UnitAdd = n
End Function
Public Function TT_UnitDelete(ByVal grp As String, ByVal n As Integer) As Boolean
TT_UnitDelete = False
Dim soDV As Integer, k As Integer
soDV = TT_UnitCount(grp)
If soDV <= 1 Or n < 1 Or n > soDV Then Exit Function
For k = n To soDV - 1
Call TT_UnitCopySlot(grp, k + 1, k)
Next k
On Error Resume Next
DeleteSetting TT_SETTING_APP, TT_UnitSection(grp, soDV)
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & grp, CStr(soDV - 1)
On Error GoTo 0
k = TT_UnitSlot(grp)
If k > n Then k = k - 1
Call TT_UnitSelect(grp, k)
TT_UnitDelete = True
End Function
Public Sub TT_KiemThuNhieuDonVi()
Dim baseDir As String
baseDir = Environ("HOME")
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim g As String, kq As String, i As Integer
g = "KIEMTHU"
On Error Resume Next
For i = 1 To 9
DeleteSetting TT_SETTING_APP, TT_UnitSection(g, i)
Next i
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & g
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & g
On Error GoTo 0
Call TT_UnitSet(g, "COQUAN", "DANG UY XA A")
Call TT_UnitSet(g, "VIETTAT", "DU")
Call TT_UnitSet(g, "DIADANH", "Xa A")
kq = "1) dem=" & TT_UnitCount(g) & " chon=" & TT_UnitSlot(g) & " nhan=[" & ToAscii(TT_UnitLabelAt(g, 1)) & "]"
i = TT_UnitAdd(g, "VAN PHONG")
Call TT_UnitSet(g, "VIETTAT", "VPDU")
kq = kq & " | 2) them=" & i & " dem=" & TT_UnitCount(g) & " chon=" & TT_UnitSlot(g) & " coquan=" & TT_UnitGet(g, "COQUAN") & " diadanh(chep)=" & TT_UnitGet(g, "DIADANH")
i = TT_UnitAdd(g, "UY BAN KIEM TRA")
Call TT_UnitSet(g, "VIETTAT", "UBKT")
Call TT_UnitSelect(g, 1)
kq = kq & " | 3) dem=" & TT_UnitCount(g) & " chon=" & TT_UnitSlot(g) & " coquan=" & TT_UnitGet(g, "COQUAN") & " [2]=" & TT_UnitGetAt(g, 2, "VIETTAT") & " [3]=" & TT_UnitGetAt(g, 3, "VIETTAT")
Call TT_UnitSelect(g, 3)
kq = kq & " | 4) xoa2=" & TT_UnitDelete(g, 2) & " dem=" & TT_UnitCount(g) & " chon=" & TT_UnitSlot(g) & " coquan=" & TT_UnitGet(g, "COQUAN") & " [2]=" & TT_UnitGetAt(g, 2, "VIETTAT") & " [3]=[" & TT_UnitGetAt(g, 3, "VIETTAT") & "]"
kq = kq & " | 5) xoa1=" & TT_UnitDelete(g, 1) & " xoaCuoi=" & TT_UnitDelete(g, 1) & " dem=" & TT_UnitCount(g) & " coquan=" & TT_UnitGet(g, "COQUAN")
kq = kq & " | 6) hienHome=" & TT_HienHome() & " nhomThat=" & TT_UnitGroupCurrent() & " demThat=" & TT_UnitCount(TT_UnitGroupCurrent()) & " ribbon=" & CStr(Not (mTTRibbon Is Nothing))
On Error Resume Next
For i = 1 To 9
DeleteSetting TT_SETTING_APP, TT_UnitSection(g, i)
Next i
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & g
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & g
If Not (mTTRibbon Is Nothing) Then mTTRibbon.ActivateTab "tabVanBanDang"
On Error GoTo 0
Call TT_DebugLog("KIEM THU NHIEU DON VI: " & kq)
End Sub
Public Sub TT_KiemThuFormMo()
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Call TT_UnitSetAt("KIEMTHU", 1, "CAPTREN", ChrW(272) & ChrW(7842) & "NG B" & ChrW(7896) & " X" & ChrW(195) & " A")
Call TT_UnitSetAt("KIEMTHU", 1, "COQUAN", ChrW(272) & ChrW(7842) & "NG " & ChrW(7910) & "Y X" & ChrW(195) & " A")
Call TT_UnitSetAt("KIEMTHU", 1, "VIETTAT", ChrW(272) & "U")
Call TT_UnitSetAt("KIEMTHU", 1, "THAMQUYEN", "T/M BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & "|B" & ChrW(205) & " TH" & ChrW(431))
Call TT_MoFormDonVi("KIEMTHU", 1)
End Sub
Public Sub TT_KiemThuFormLuu()
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim d As Document, kq As String, ok As Boolean, i As Integer
Set d = TT_FormDoc()
If Not (d Is Nothing) Then
d.Tables(1).Cell(2, 2).Range.text = "V" & ChrW(258) & "N PH" & ChrW(210) & "NG" & vbCr & ChrW(272) & ChrW(7842) & "NG " & ChrW(7910) & "Y"
d.Tables(1).Cell(3, 2).Range.text = "VP" & ChrW(272) & "U"
End If
ok = TT_LuuFormDonVi(False)
kq = "luu=" & ok & " | COQUAN=[" & ToAscii(TT_UnitGetAt("KIEMTHU", 1, "COQUAN")) & "] | VIETTAT ma ky tu 3=" & AscW(Mid(TT_UnitGetAt("KIEMTHU", 1, "VIETTAT"), 3, 1)) & _
" | THAMQUYEN=[" & ToAscii(TT_UnitGetAt("KIEMTHU", 1, "THAMQUYEN")) & "] | CAPTREN=[" & ToAscii(TT_UnitGetAt("KIEMTHU", 1, "CAPTREN")) & "] | formConMo=" & CStr(Not (TT_FormDoc() Is Nothing))
On Error Resume Next
For i = 1 To 9
DeleteSetting TT_SETTING_APP, TT_UnitSection("KIEMTHU", i)
Next i
On Error GoTo 0
Call TT_DebugLog("KIEM THU FORM: " & kq)
End Sub
Public Sub TT_KiemThuDat(ByVal loai As String, ByVal ten As String, ByVal giaTri As String)
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
On Error Resume Next
Select Case loai
Case "KHOI"
Call TT_DebugLog("KIEM THU: khoi cu=" & TT_GetSetting() & " -> " & giaTri)
Call TT_SaveSetting(giaTri)
Case "GOM"
Call TT_DebugLog("KIEM THU: gom chu le cu=" & TT_GomChuLe() & " -> " & giaTri)
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "GomChuLe", giaTri
Case "CHON"
Call TT_UnitSelect(ten, Val(giaTri))
Case "THEMDV"
Call TT_DebugLog("KIEM THU: them don vi " & ten & " -> so " & TT_UnitAdd(ten, giaTri))
Case "XOADV"
Call TT_DebugLog("KIEM THU: xoa don vi " & ten & " so " & giaTri & " -> " & TT_UnitDelete(ten, Val(giaTri)))
Case "IN"
Call TT_KiemThuInDonVi
Case "XOACAIDAT"
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, ten
Case "MAULUU"
Dim dMau2 As Document
Set dMau2 = TT_TaoMauTheoTag(ten & "|" & giaTri, False)
If Not (dMau2 Is Nothing) Then
dMau2.SaveAs2 Environ("HOME") & "/tt_mau_" & ten & "_" & giaTri & ".docx", 16
dMau2.Close 0
End If
Case "MAU"
Dim dMau As Document, tMau As String
Set dMau = TT_TaoMauTheoTag(ten & "|" & giaTri, False)
If Not (dMau Is Nothing) Then
If dMau.Tables.Count > 0 Then tMau = dMau.Tables(1).Range.text
Call TT_DebugLog("KIEM THU MAU " & ten & "|" & giaTri & ": " & Replace(Replace(ToAscii(Left(tMau, 160)), vbCr, " / "), Chr(7), " "))
dMau.Close 0
Else
Call TT_DebugLog("KIEM THU MAU " & ten & "|" & giaTri & ": khong tao duoc")
End If
Case Else
If Len(giaTri) = 0 Then
DeleteSetting TT_SETTING_APP, TT_UnitSection(loai, TT_UnitSlot(loai)), "TC_" & ten
Else
SaveSetting TT_SETTING_APP, TT_UnitSection(loai, TT_UnitSlot(loai)), "TC_" & ten, giaTri
End If
End Select
mTTOvPid = ""
mTTOvStr = ""
Call TT_RibbonRefresh
On Error GoTo 0
End Sub
Public Sub TT_KiemThuThemDonVi()
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim g As String, i As Integer, kq As String, d As Document, ok As Boolean
g = "KIEMTHU"
On Error Resume Next
For i = 1 To 9
DeleteSetting TT_SETTING_APP, TT_UnitSection(g, i)
Next i
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & g
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & g
Call TT_UnitSetAt(g, 1, "COQUAN", "DANG UY XA A")
Call TT_UnitSetAt(g, 1, "DIADANH", "Xa A")
Call TT_MoFormDonVi(g, 0)
kq = "mo bang moi: so=" & TT_UnitCount(g)
Set d = TT_FormDoc()
d.Close 0
kq = kq & " | dong khong luu: so=" & TT_UnitCount(g)
Call TT_MoFormDonVi(g, 0)
ok = TT_LuuFormDonVi(False)
kq = kq & " | luu khi chua co ten=" & ok & " so=" & TT_UnitCount(g)
Set d = TT_FormDoc()
d.Tables(1).Cell(2, 2).Range.text = "VAN PHONG"
d.Tables(1).Cell(3, 2).Range.text = "VPDU"
ok = TT_LuuFormDonVi(False)
kq = kq & " | luu co ten=" & ok & " so=" & TT_UnitCount(g) & " chon=" & TT_UnitSlot(g) & " [2]=" & ToAscii(TT_UnitLabelAt(g, 2)) & " diadanh[2]=" & TT_UnitGetAt(g, 2, "DIADANH") & " bangConMo=" & CStr(Not (TT_FormDoc() Is Nothing))
For i = 1 To 9
DeleteSetting TT_SETTING_APP, TT_UnitSection(g, i)
Next i
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "SoDonVi_" & g
DeleteSetting TT_SETTING_APP, TT_SETTING_SECTION, "DonViChon_" & g
On Error GoTo 0
Call TT_DebugLog("KIEM THU THEM DON VI: " & kq)
End Sub
Public Sub TT_KiemThuBamDupLuu()
Dim h As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
On Error Resume Next
mTTThuDongNgay = (Len(Dir(h & "/thethuc_debug.dongngay")) > 0)
mTTThuKhongHoi = True
Call TT_UnitSetAt("KIEMTHU", 1, "COQUAN", "DANG UY XA A")
Call TT_MoFormDonVi("KIEMTHU", 1)
Dim d As Document
Set d = TT_FormDoc()
Call TT_DebugLog("BAM DUP LUU: bat dau, dongNgay=" & mTTThuDongNgay & " soTruong=" & d.Fields.Count)
d.Tables(1).Cell(3, 2).Range.text = "DU"
d.Fields(1).DoClick
Call TT_DebugLog("BAM DUP LUU: xong, VIETTAT=" & TT_UnitGetAt("KIEMTHU", 1, "VIETTAT") & " soVanBan=" & Documents.Count & " bangConMo=" & CStr(Not (TT_FormDoc() Is Nothing)))
mTTThuDongNgay = False
mTTThuKhongHoi = False
DeleteSetting TT_SETTING_APP, TT_UnitSection("KIEMTHU", 1)
On Error GoTo 0
End Sub
Public Sub TT_KiemThuInDonVi()
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim g As Variant, i As Integer, kq As String
For Each g In Array("DANG", "CHINH_QUYEN", "CONG_AN", "QUAN_DOI", "MTTQ")
kq = kq & " || " & g & ": so=" & TT_UnitCount(CStr(g)) & " chon=" & TT_UnitSlot(CStr(g))
For i = 1 To TT_UnitCount(CStr(g))
kq = kq & " [" & i & "]=" & ToAscii(TT_UnitLabelAt(CStr(g), i))
Next i
Next g
Call TT_DebugLog("DON VI: khoi=" & TT_GetSetting() & kq)
End Sub
Public Sub TT_KiemThuLenh()
Dim h As String, f As Integer, dong As String, a() As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
If Len(Dir(h & "/thethuc_debug.cmd")) = 0 Then Exit Sub
On Error Resume Next
f = FreeFile
Open h & "/thethuc_debug.cmd" For Input As #f
Do While Not EOF(f)
Line Input #f, dong
dong = Trim(Replace(dong, vbCr, ""))
If Len(dong) > 0 Then
a = Split(dong & "||", "|")
If a(0) = "FORM" Then
Call TT_MoFormTuyChon(a(1))
Else
Call TT_KiemThuDat(a(0), a(1), a(2))
End If
End If
Loop
Close #f
Kill h & "/thethuc_debug.cmd"
On Error GoTo 0
End Sub
Public Sub TT_KiemThuFormTuyChonMo(ByVal pid As String)
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Call TT_MoFormTuyChon(pid)
End Sub
Public Sub TT_KiemThuFormTuyChonLuu()
If Len(Dir(Environ("HOME") & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim d As Document, pid As String, ok As Boolean, kq As String
Set d = TT_FormDoc()
If d Is Nothing Then
Call TT_DebugLog("KIEM THU FORM TUY CHON: khong co bang")
Exit Sub
End If
On Error Resume Next
pid = d.Variables("TT_FormTuyChon").value
pid = Left(pid, InStr(1, pid & "|", "|") - 1)
d.Tables(1).Cell(2, 2).Range.text = "14"
d.Tables(1).Cell(3, 2).Range.text = "99"
ok = TT_LuuFormTuyChon(d, False)
kq = "pid=" & pid & " | luu(co 1 dong sai)=" & ok & " | BODY_SIZE=" & GetSetting(TT_SETTING_APP, TT_UnitSection(pid, TT_UnitSlot(pid)), "TC_BODY_SIZE", "(trong)") & _
" | TITLE_SIZE=" & GetSetting(TT_SETTING_APP, TT_UnitSection(pid, TT_UnitSlot(pid)), "TC_TITLE_SIZE", "(trong)") & " | bangConMo=" & CStr(Not (TT_FormDoc() Is Nothing))
Set d = TT_FormDoc()
d.Tables(1).Cell(3, 2).Range.text = ""
d.Tables(1).Cell(2, 2).Range.text = ""
ok = TT_LuuFormTuyChon(d, False)
kq = kq & " | luu lai(de trong)=" & ok & " | BODY_SIZE=" & GetSetting(TT_SETTING_APP, TT_UnitSection(pid, TT_UnitSlot(pid)), "TC_BODY_SIZE", "(trong)") & " | bangConMo=" & CStr(Not (TT_FormDoc() Is Nothing))
On Error GoTo 0
Call TT_DebugLog("KIEM THU FORM TUY CHON: " & kq)
End Sub
Public Sub RibbonDVBH_MenuLabel(control As Object, ByRef returnedVal)
Dim g As String, nhan As String
g = TT_UnitGroupCurrent()
nhan = TT_UnitLabelAt(g, TT_UnitSlot(g))
If Len(nhan) > 34 Then nhan = Left(nhan, 32) & "..."
returnedVal = ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & ": " & nhan
End Sub
Public Sub RibbonDVBH_ItemLabel(control As Object, ByRef returnedVal)
Dim g As String, n As Integer
g = TT_UnitGroupCurrent()
n = Val(control.tag)
If n = TT_UnitSlot(g) Then
returnedVal = ChrW(10003) & "  " & TT_UnitLabelAt(g, n)
Else
returnedVal = "     " & TT_UnitLabelAt(g, n)
End If
End Sub
Public Sub RibbonDVBH_ItemVisible(control As Object, ByRef returnedVal)
returnedVal = (Val(control.tag) <= TT_UnitCount(TT_UnitGroupCurrent()))
End Sub
Public Sub RibbonDVBH_ItemPick(control As Object)
Call TT_UnitSelect(TT_UnitGroupCurrent(), Val(control.tag))
Call TT_RibbonRefresh
End Sub
Public Sub RibbonDVBH_GetItemCount(control As Object, ByRef returnedVal)
returnedVal = TT_UnitCount(TT_UnitGroupCurrent())
End Sub
Public Sub RibbonDVBH_GetItemLabel(control As Object, index As Variant, ByRef returnedVal)
returnedVal = TT_UnitLabelAt(TT_UnitGroupCurrent(), CInt(index) + 1)
End Sub
Public Sub RibbonDVBH_GetItemID(control As Object, index As Variant, ByRef returnedVal)
returnedVal = control.ID & "_" & CStr(CInt(index) + 1)
End Sub
Public Sub RibbonDVBH_GetSelectedIndex(control As Object, ByRef returnedVal)
returnedVal = TT_UnitSlot(TT_UnitGroupCurrent()) - 1
End Sub
Public Sub RibbonDVBH_OnAction(control As Object, selectedId As String, selectedIndex As Integer)
Call TT_UnitSelect(TT_UnitGroupCurrent(), selectedIndex + 1)
Call TT_RibbonRefresh
End Sub
Public Sub RibbonDVBH_Them(control As Object)
Dim grp As String
grp = TT_UnitGroupCurrent()
If TT_UnitCount(grp) >= TT_UnitMax() Then
MsgBox "Da du " & TT_UnitMax() & " don vi cho khoi nay. Hay xoa bot don vi khong dung truoc khi them.", vbExclamation, "Don vi ban hanh"
Exit Sub
End If
Call TT_MoFormDonVi(grp, 0)
End Sub
Public Sub RibbonDVBH_Xoa(control As Object)
Dim grp As String, n As Integer
grp = TT_UnitGroupCurrent()
If TT_UnitCount(grp) <= 1 Then
MsgBox "Chi con 1 don vi, khong xoa duoc. Muon doi thong tin thi sua truc tiep cac o.", vbInformation, "Don vi ban hanh"
Exit Sub
End If
n = TT_UnitSlot(grp)
If MsgBox("Xoa don vi dang chon (so " & n & ": " & ToAscii(TT_UnitLabelAt(grp, n)) & ")?", vbYesNo + vbQuestion, "Don vi ban hanh") <> vbYes Then Exit Sub
Call TT_UnitDelete(grp, n)
Call TT_RibbonRefresh
End Sub
Private Function TT_FormNhan(ByVal i As Integer) As String
Select Case i
Case 1: TT_FormNhan = "C" & ChrW(7845) & "p tr" & ChrW(234) & "n"
Case 2: TT_FormNhan = "C" & ChrW(417) & " quan ban h" & ChrW(224) & "nh"
Case 3: TT_FormNhan = "Ch" & ChrW(7919) & " vi" & ChrW(7871) & "t t" & ChrW(7855) & "t"
Case 4: TT_FormNhan = ChrW(272) & ChrW(7883) & "a danh"
Case 5: TT_FormNhan = "Th" & ChrW(7849) & "m quy" & ChrW(7873) & "n k" & ChrW(253)
Case 6: TT_FormNhan = "Ch" & ChrW(7913) & "c v" & ChrW(7909) & " ng" & ChrW(432) & ChrW(7901) & "i k" & ChrW(253)
Case 7: TT_FormNhan = "H" & ChrW(7885) & " t" & ChrW(234) & "n ng" & ChrW(432) & ChrW(7901) & "i k" & ChrW(253)
Case Else: TT_FormNhan = ""
End Select
End Function
Private Function TT_FormDoc() As Document
Set TT_FormDoc = Nothing
On Error Resume Next
Dim d As Document, v As String
For Each d In Documents
v = ""
v = d.Variables("TT_FormDonVi").value
If Len(v) = 0 Then v = d.Variables("TT_FormTuyChon").value
If Len(v) > 0 Then
Set TT_FormDoc = d
Exit For
End If
Next d
On Error GoTo 0
End Function
Public Sub TT_MoFormDonVi(ByVal grp As String, ByVal n As Integer)
On Error Resume Next
Dim d As Document
Set d = TT_FormDoc()
If Not (d Is Nothing) Then d.Close 0
Set d = Nothing
Set d = Documents.Add
If d Is Nothing Then Exit Sub
d.Variables.Add "TT_FormDonVi", grp & "|" & CStr(n)
With d.PageSetup
.TopMargin = 28
.BottomMargin = 28
.LeftMargin = 34
.RightMargin = 34
End With
Dim r As Range, tb As Table, i As Integer, tenKhoi As String, huongDan As String
tenKhoi = TT_TenKhoiCoDau(grp)
Set r = d.Range(0, 0)
If n = 0 Then
r.InsertAfter "TH" & ChrW(202) & "M " & ChrW(272) & ChrW(416) & "N V" & ChrW(7882) & " BAN H" & ChrW(192) & "NH M" & ChrW(7898) & "I" & " - " & tenKhoi & vbCr
Else
r.InsertAfter "TH" & ChrW(212) & "NG TIN " & ChrW(272) & ChrW(416) & "N V" & ChrW(7882) & " BAN H" & ChrW(192) & "NH" & " - " & tenKhoi & " - " & ChrW(273) & ChrW(417) & "n v" & ChrW(7883) & " s" & ChrW(7889) & " " & CStr(n) & vbCr
End If
huongDan = ChrW(272) & "i" & ChrW(7873) & "n v" & ChrW(224) & "o c" & ChrW(7897) & "t b" & ChrW(234) & "n ph" & ChrW(7843) & "i. Xong th" & ChrW(236) & " b" & ChrW(7845) & "m n" & ChrW(250) & "t " & ChrW(8220) & "L" & ChrW(432) & "u"
huongDan = huongDan & " b" & ChrW(7843) & "ng nh" & ChrW(7853) & "p" & ChrW(8221) & " tr" & ChrW(234) & "n thanh Ribbon (tab Home ho" & ChrW(7863) & "c"
huongDan = huongDan & " tab TH" & ChrW(7874) & " TH" & ChrW(7912) & "C V" & ChrW(258) & "N B" & ChrW(7842) & "N). T" & ChrW(234) & "n d" & ChrW(224) & "i mu" & ChrW(7889) & "n xu" & ChrW(7889) & "ng"
huongDan = huongDan & " d" & ChrW(242) & "ng th" & ChrW(236) & " b" & ChrW(7845) & "m Enter ngay trong " & ChrW(244) & ". " & ChrW(272) & ChrW(226) & "y ch" & ChrW(7881) & " l" & ChrW(224)
huongDan = huongDan & " b" & ChrW(7843) & "ng nh" & ChrW(7853) & "p t" & ChrW(7841) & "m: KH" & ChrW(212) & "NG c" & ChrW(7847) & "n l" & ChrW(432) & "u file n" & ChrW(224) & "y (kh" & ChrW(244) & "ng"
huongDan = huongDan & " b" & ChrW(7845) & "m Ctrl+S). L" & ChrW(432) & "u xong c" & ChrW(243) & " th" & ChrW(7875) & " " & ChrW(273) & ChrW(243) & "ng c" & ChrW(7917) & "a s" & ChrW(7893) & " n" & ChrW(224) & "y."
r.InsertAfter huongDan & vbCr
d.Paragraphs(1).Range.Font.Bold = True
d.Paragraphs(1).Range.Font.Size = 13
d.Paragraphs(2).Range.Font.Italic = True
d.Paragraphs(2).Range.Font.Size = 11
d.Content.Font.Name = "Times New Roman"
d.Content.ParagraphFormat.SpaceAfter = 4
Set r = d.Range(d.Content.End - 1, d.Content.End - 1)
Set tb = d.Tables.Add(r, 7, 2)
tb.Borders.Enable = True
tb.Range.Font.Name = "Times New Roman"
tb.Range.Font.Size = 13
tb.Range.Font.Bold = False
tb.Range.Font.Italic = False
tb.Range.ParagraphFormat.SpaceAfter = 0
tb.AllowAutoFit = False
tb.Columns(1).Width = 135
tb.Columns(2).Width = 300
For i = 1 To 7
tb.Cell(i, 1).Range.text = TT_FormNhan(i)
tb.Cell(i, 1).Range.Font.Bold = True
tb.Cell(i, 1).Shading.BackgroundPatternColor = RGB(235, 235, 235)
If n = 0 Then
If TT_UnitFieldAt(i) <> "COQUAN" And TT_UnitFieldAt(i) <> "VIETTAT" Then
tb.Cell(i, 2).Range.text = Replace(TT_UnitGetAt(grp, TT_UnitSlot(grp), TT_UnitFieldAt(i)), "|", vbCr)
End If
Else
tb.Cell(i, 2).Range.text = Replace(TT_UnitGetAt(grp, n, TT_UnitFieldAt(i)), "|", vbCr)
End If
tb.Cell(i, 2).TopPadding = 3
tb.Cell(i, 2).BottomPadding = 3
Next i
Set r = d.Range(d.Content.End - 1, d.Content.End - 1)
huongDan = "G" & ChrW(7907) & "i " & ChrW(253) & ": C" & ChrW(7845) & "p tr" & ChrW(234) & "n = " & ChrW(272) & ChrW(7843) & "ng b" & ChrW(7897) & " c" & ChrW(7845) & "p tr" & ChrW(234) & "n ho" & ChrW(7863) & "c c" & ChrW(417)
huongDan = huongDan & " quan ch" & ChrW(7911) & " qu" & ChrW(7843) & "n. Ch" & ChrW(7919) & " vi" & ChrW(7871) & "t t" & ChrW(7855) & "t d" & ChrW(249) & "ng trong s" & ChrW(7889) & ", k" & ChrW(253)
huongDan = huongDan & " hi" & ChrW(7879) & "u (VD: " & ChrW(272) & "U, VP" & ChrW(272) & "U, UBND). Th" & ChrW(7849) & "m quy" & ChrW(7873) & "n k" & ChrW(253) & " " & ChrW(273) & ChrW(7875)
huongDan = huongDan & " tr" & ChrW(7889) & "ng n" & ChrW(7871) & "u k" & ChrW(253) & " tr" & ChrW(7921) & "c ti" & ChrW(7871) & "p (VD c" & ChrW(243) & " th" & ChrW(7849) & "m quy" & ChrW(7873) & "n: T/M"
huongDan = huongDan & " BAN TH" & ChrW(431) & ChrW(7900) & "NG V" & ChrW(7908) & ", KT. CH" & ChrW(7910) & " T" & ChrW(7882) & "CH)."
r.InsertAfter huongDan
d.Paragraphs(d.Paragraphs.Count).Range.Font.Size = 10
d.Paragraphs(d.Paragraphs.Count).Range.Font.Italic = True
d.Paragraphs(d.Paragraphs.Count).Range.Font.Bold = False
d.ActiveWindow.WindowState = wdWindowStateNormal
d.ActiveWindow.Caption = "B" & ChrW(7843) & "ng nh" & ChrW(7853) & "p t" & ChrW(7841) & "m - kh" & ChrW(244) & "ng c" & ChrW(7847) & "n l" & ChrW(432) & "u file"
d.ActiveWindow.Width = 700
d.ActiveWindow.Height = 560
d.ActiveWindow.View.Type = wdPrintView
d.ActiveWindow.View.Zoom.Percentage = 100
If n = 0 Then
tb.Cell(2, 2).Range.Select
Else
tb.Cell(1, 2).Range.Select
End If
Selection.Collapse wdCollapseStart
d.Saved = True
Call TT_RibbonRefresh
DoEvents
If Not (mTTRibbon Is Nothing) Then mTTRibbon.ActivateTab "tabVanBanDang"
On Error GoTo 0
End Sub
Public Function TT_LuuFormDonVi(ByVal showMsg As Boolean) As Boolean
TT_LuuFormDonVi = False
On Error Resume Next
Dim d As Document, v As String, bar As Long, grp As String, n As Integer, i As Integer, t As String, tb As Table
Set d = TT_FormDoc()
If Not (d Is Nothing) Then
v = ""
v = d.Variables("TT_FormTuyChon").value
If Len(v) > 0 Then
TT_LuuFormDonVi = TT_LuuFormTuyChon(d, showMsg)
Exit Function
End If
End If
If d Is Nothing Then
If showMsg Then MsgBox "Chua mo bang thong tin don vi. Bam nut 'Sua thong tin' truoc, dien xong roi bam 'Luu don vi'.", vbInformation, "Don vi ban hanh"
Exit Function
End If
v = d.Variables("TT_FormDonVi").value
bar = InStr(1, v, "|")
If bar = 0 Or d.Tables.Count = 0 Then Exit Function
grp = Left(v, bar - 1)
n = Val(Mid(v, bar + 1))
Set tb = d.Tables(1)
If tb.Rows.Count < 7 Then
If showMsg Then MsgBox "Bang thong tin don vi da bi sua mat dong. Hay dong cua so nay va bam 'Sua thong tin' lai.", vbExclamation, "Don vi ban hanh"
Exit Function
End If
If n = 0 Then
t = Replace(Replace(Replace(tb.Cell(2, 2).Range.text & tb.Cell(3, 2).Range.text, Chr(7), ""), vbCr, ""), " ", "")
If Len(t) = 0 Then
If showMsg Then MsgBox "Chua nhap ten co quan ban hanh (dong 2). Hay dien roi bam Luu, hoac dong cua so nay neu khong muon them don vi.", vbExclamation, "Don vi ban hanh"
Exit Function
End If
n = TT_UnitAdd(grp, "")
If n = 0 Then
If showMsg Then MsgBox "Da du so don vi cho khoi nay, khong them duoc.", vbExclamation, "Don vi ban hanh"
Exit Function
End If
End If
For i = 1 To 7
t = tb.Cell(i, 2).Range.text
t = Replace(Replace(t, Chr(7), ""), Chr(11), vbCr)
Do While Len(t) > 0 And (Right(t, 1) = vbCr Or Right(t, 1) = " ")
t = Left(t, Len(t) - 1)
Loop
Do While InStr(1, t, vbCr & vbCr) > 0
t = Replace(t, vbCr & vbCr, vbCr)
Loop
t = Replace(t, vbCr, " | ")
Call TT_UnitSetAt(grp, n, TT_UnitFieldAt(i), Trim(t))
Next i
Call TT_DongBangNhap(d)
Call TT_RibbonRefresh
TT_LuuFormDonVi = True
If showMsg Then MsgBox "Da luu thong tin don vi so " & n & "." & vbCrLf & "Co the dong cua so bang nhap (khong can luu file).", vbInformation, "Don vi ban hanh"
On Error GoTo 0
End Function
Public Sub RibbonDVBH_Sua(control As Object)
Dim grp As String
grp = TT_UnitGroupCurrent()
Call TT_MoFormDonVi(grp, TT_UnitSlot(grp))
End Sub
Public Sub RibbonDVBH_Luu(control As Object)
Call TT_LuuFormDonVi(True)
End Sub
Public Sub TT_LuuDonViTuBang()
mTTLuuTuBang = True
Call TT_LuuFormDonVi(Not mTTThuKhongHoi)
mTTLuuTuBang = False
End Sub
Private Sub TT_DongBangNhap(ByVal d As Document)
On Error Resume Next
If mTTThuDongNgay Then
d.Close 0
Exit Sub
End If
d.Variables("TT_FormDonVi").Delete
d.Variables("TT_FormTuyChon").Delete
d.Variables.Add "TT_FormDaLuu", "1"
d.Saved = True
Application.OnTime Now + TimeSerial(0, 0, 1), "TT_DongBangNhapDaLuu"
On Error GoTo 0
End Sub
Public Sub TT_DongBangNhapDaLuu()
On Error Resume Next
Dim d As Document, v As String, i As Long
For i = Documents.Count To 1 Step -1
Set d = Documents(i)
v = ""
v = d.Variables("TT_FormDaLuu").value
If v = "1" Then d.Close 0
Next i
On Error GoTo 0
End Sub
Public Sub RibbonDVBH_LuuVisible(control As Object, ByRef returnedVal)
returnedVal = Not (TT_FormDoc() Is Nothing)
End Sub
Public Function TT_HienHome() As Boolean
Dim v As String
On Error Resume Next
v = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "HienHome", "1")
On Error GoTo 0
TT_HienHome = (v <> "0")
End Function
Public Sub RibbonHienHome_GetPressed(control As Object, ByRef returnedVal)
returnedVal = TT_HienHome()
End Sub
Public Sub RibbonHienHome_OnAction(control As Object, pressed As Boolean)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "HienHome", IIf(pressed, "1", "0")
On Error GoTo 0
Call TT_RibbonRefresh
End Sub
Public Sub RibbonHome_GetVisible(control As Object, ByRef returnedVal)
returnedVal = TT_HienHome()
End Sub
Public Function TT_UnitGroupCurrent() As String
Dim pid As String, g As String
pid = TT_GetSetting()
If pid <> "AUTO" Then
If TT_IndexOfProfile(pid) > 0 Then
TT_UnitGroupCurrent = pid
Exit Function
End If
End If
On Error Resume Next
g = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", "DANG")
On Error GoTo 0
If g = "NHA_NUOC" Then g = "CHINH_QUYEN"
If TT_IndexOfProfile(g) = 0 Then g = "DANG"
TT_UnitGroupCurrent = g
End Function
Public Sub RibbonUnit_GetText(control As Object, ByRef returnedVal)
returnedVal = TT_UnitGet(TT_UnitGroupCurrent(), control.tag)
End Sub
Public Sub RibbonUnit_OnChange(control As Object, text As String)
Call TT_UnitSet(TT_UnitGroupCurrent(), control.tag, text)
End Sub
Public Sub RibbonUnit_GetLabel(control As Object, ByRef returnedVal)
If TT_UnitGroupCurrent() = "DANG" Then
returnedVal = "(" & ChrW(273) & "ang s" & ChrW(7917) & "a: kh" & ChrW(7889) & "i " & ChrW(272) & ChrW(7843) & "ng)"
Else
returnedVal = "(" & ChrW(273) & "ang s" & ChrW(7917) & "a: kh" & ChrW(7889) & "i Nh" & ChrW(224) & " n" & ChrW(432) & ChrW(7899) & "c)"
End If
End Sub
Private Sub TT_ReadSignerInfo(ByVal src As Document, ByRef auth As String, ByRef role As String, ByRef signer As String)
auth = "": role = "": signer = ""
On Error GoTo SigFail
Dim tbl As Table
For Each tbl In src.Tables
If IsFooterSignatureTable(tbl, ToAscii(tbl.Range.text)) Then
Dim c As Cell
For Each c In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(c.Range.text))
If InStr(1, cAsc, "NOI NHAN") = 0 And HasSignatureKeywordInCell(cAsc) Then
Dim p As Paragraph
For Each p In c.Range.Paragraphs
Dim t As String
t = CleanParaText(p.Range.text)
If Len(t) > 0 Then
Dim a As String
a = Trim(ToAscii(t))
If (a Like "T/M *" Or a Like "TM.*" Or a Like "K/T *" Or a Like "KT.*" Or a Like "T/L *" Or a Like "TL.*" Or a Like "TUQ.*" Or a Like "Q. *" Or a Like "Q.[A-Z]*") And Len(auth) = 0 And Len(role) = 0 Then
auth = t
ElseIf IsSignatureRoleOrAuth(a) And Len(role) = 0 Then
role = t
ElseIf Not IsSignatureRoleOrAuth(a) Then
signer = t
End If
End If
Next p
Exit Sub
End If
Next c
End If
Next tbl
Exit Sub
SigFail:
End Sub
Public Sub RibbonUnit_LayTuVanBan(control As Object)
Call TT_UnitLayTuVanBan(True)
End Sub
Public Function TT_UnitLayTuVanBan(ByVal showMsg As Boolean) As Boolean
TT_UnitLayTuVanBan = False
If Documents.Count = 0 Then Exit Function
Dim src As Document
Set src = ActiveDocument
Dim strong As Boolean, pid As String, grp As String
pid = TT_DetectFromHeaderText(TT_HeaderAreaAscii(src), strong)
If Len(pid) = 0 Then pid = TT_GetSetting()
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Then
If showMsg Then MsgBox "Khong nhan dien duoc khoi cua van ban dang mo (can co bang tieu de da chuan hoa).", vbExclamation, "Thong tin don vi"
Exit Function
End If
Dim organLines As String, orgAbbr As String, placeName As String
Call TT_ReadOrganInfo(src, grp, organLines, orgAbbr, placeName)
If Len(organLines) = 0 Then
If showMsg Then MsgBox "Van ban dang mo chua co bang tieu de (Ten co quan, So ky hieu). Hay Chuan hoa the thuc truoc.", vbExclamation, "Thong tin don vi"
Exit Function
End If
Dim parts() As String
parts = Split(organLines, vbCr)
Dim capTren As String, coQuan As String
coQuan = parts(UBound(parts))
Dim k As Long
For k = LBound(parts) To UBound(parts) - 1
If Left(parts(k), 1) <> "[" Then
If Len(capTren) > 0 Then capTren = capTren & " | "
capTren = capTren & parts(k)
End If
Next k
Call TT_UnitSet(pid, "CAPTREN", capTren)
Call TT_UnitSet(pid, "COQUAN", coQuan)
If Left(orgAbbr, 1) <> "[" Then Call TT_UnitSet(pid, "VIETTAT", orgAbbr)
If Left(placeName, 1) <> "[" Then Call TT_UnitSet(pid, "DIADANH", placeName)
Dim auth As String, role As String, signer As String
Call TT_ReadSignerInfo(src, auth, role, signer)
If Len(auth) > 0 Then Call TT_UnitSet(pid, "THAMQUYEN", auth)
If Len(role) > 0 Then Call TT_UnitSet(pid, "CHUCVU", role)
If Len(signer) > 0 And Left(signer, 1) <> "[" Then Call TT_UnitSet(pid, "HOTEN", signer)
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", pid
On Error GoTo 0
Call TT_RibbonRefresh
TT_UnitLayTuVanBan = True
If showMsg Then MsgBox "Da luu thong tin don vi (khoi " & TT_ProfileLabelAscii(pid) & ") tu van ban dang mo. Xem va sua o nhom 'Thong tin don vi' tren Ribbon.", vbInformation, "Thong tin don vi"
End Function
Public Sub RibbonMau_GetVisible(control As Object, ByRef returnedVal)
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Then
returnedVal = True
Else
returnedVal = (TT_TemplateGroupOf(pid) = control.tag)
End If
End Sub
Public Sub RibbonMau_GetLabel(control As Object, ByRef returnedVal)
If TT_GetSetting() = "AUTO" Then
If control.tag = "DANG" Then
returnedVal = "M" & ChrW(7851) & "u kh" & ChrW(7889) & "i " & ChrW(272) & ChrW(7843) & "ng"
ElseIf control.tag = "MTTQ" Then
returnedVal = "M" & ChrW(7851) & "u kh" & ChrW(7889) & "i MTTQ"
Else
returnedVal = "M" & ChrW(7851) & "u kh" & ChrW(7889) & "i Nh" & ChrW(224) & " n" & ChrW(432) & ChrW(7899) & "c"
End If
Else
returnedVal = "T" & ChrW(7841) & "o v" & ChrW(259) & "n b" & ChrW(7843) & "n m" & ChrW(7851) & "u"
End If
End Sub
Public Sub RibbonMoDonVi(control As Object)
On Error GoTo NoActivate
If mTTRibbon Is Nothing Then GoTo NoActivate
mTTRibbon.ActivateTab "tabVanBanDang"
Exit Sub
NoActivate:
MsgBox "Thong tin don vi nam o tab 'THE THUC VAN BAN' tren thanh Ribbon (nhom 'Thong Tin Don Vi').", vbInformation, "Thong tin don vi"
End Sub
Public Sub RibbonTaoMauItem(control As Object)
Call TT_TaoMauTheoTag(control.tag, True)
End Sub
Public Function TT_TaoMauTheoTag(ByVal tagVal As String, ByVal showMsg As Boolean) As Document
Set TT_TaoMauTheoTag = Nothing
If showMsg Then
If Not CheckLicenseActivation() Then Exit Function
End If
Dim bar As Long
bar = InStr(1, tagVal, "|")
If bar = 0 Then Exit Function
Dim grp As String, nM As Integer
grp = Left(tagVal, bar - 1)
nM = Val(Mid(tagVal, bar + 1))
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Or TT_TemplateGroupOf(pid) <> grp Then
If grp = "DANG" Then
pid = "DANG"
ElseIf TT_IndexOfProfile(grp) > 0 Then
pid = grp
Else
pid = "CHINH_QUYEN"
End If
End If
On Error Resume Next
SaveSetting TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", pid
On Error GoTo 0
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Dim khoaDV As String
khoaDV = TT_UnitGroupCurrent()
If Not (TT_TemplateGroupOf(khoaDV) = grp Or (TT_TemplateGroupOf(khoaDV) = "" And grp = "NHA_NUOC" And khoaDV <> "DANG")) Then khoaDV = pid
Set TT_TaoMauTheoTag = TT_CreateFromTemplate(pid, nM, srcDoc, showMsg, khoaDV)
Call TT_RibbonRefresh
End Function
Public Sub RibbonTaoMau(control As Object)
Call TaoVanBanMau
End Sub
Private Function TT_TemplateGroupOf(ByVal pid As String) As String
If Len(pid) > 0 And pid <> "AUTO" Then
If TT_MauCount(pid) > 0 Then
TT_TemplateGroupOf = pid
Exit Function
End If
End If
Select Case TT_RawS(pid, "MOTTO_KIND")
Case "DANG": TT_TemplateGroupOf = "DANG"
Case "QUOC_HIEU": TT_TemplateGroupOf = "NHA_NUOC"
Case Else: TT_TemplateGroupOf = ""
End Select
End Function
Private Function TT_AskNumber(ByVal msg As String, ByVal title As String, ByVal defaultVal As Integer, ByVal maxVal As Integer) As Integer
TT_AskNumber = 0
Dim ans As String
ans = InputBox(msg, title, CStr(defaultVal))
If Len(Trim(ans)) = 0 Then Exit Function
If Not IsNumeric(ans) Then Exit Function
Dim n As Integer
n = Val(ans)
If n >= 1 And n <= maxVal Then TT_AskNumber = n
End Function
Private Sub TT_ReadOrganInfo(ByVal src As Document, ByVal grp As String, ByRef organLines As String, ByRef orgAbbr As String, ByRef placeName As String)
organLines = "": orgAbbr = "": placeName = ""
On Error GoTo ReadFail
If src Is Nothing Then Exit Sub
If src.Tables.Count = 0 Then Exit Sub
Dim tbl As Table
Set tbl = src.Tables(1)
Dim tAsc As String
tAsc = ToAscii(tbl.Range.text)
Dim strong As Boolean
Dim srcPid As String
srcPid = TT_DetectFromHeaderText(tAsc, strong)
If Len(srcPid) = 0 Then Exit Sub
If TT_TemplateGroupOf(srcPid) <> grp Then Exit Sub
Dim c As Cell
Set c = tbl.Range.Cells(1)
Dim p As Paragraph
For Each p In c.Range.Paragraphs
Dim t As String
t = CleanParaText(p.Range.text)
Dim a As String
a = Trim(ToAscii(t))
If Len(a) > 0 Then
If IsDocNumberLineAsc(a) Then
Dim slashPos As Long
slashPos = InStr(1, t, "/")
If slashPos > 0 Then
Dim tail As String
tail = Trim(Mid(t, slashPos + 1))
If DocNumberStyleOf(t) = "NHA_NUOC" Then
Dim dashPos As Long
dashPos = InStr(1, tail, "-")
If dashPos > 0 Then
If Len(TT_RawDocCodeTitle(srcPid, ToAscii(Left(tail, dashPos - 1)))) > 0 Then
orgAbbr = Trim(Mid(tail, dashPos + 1))
Else
orgAbbr = Trim(Left(tail, dashPos - 1))
End If
Else
orgAbbr = tail
End If
Else
orgAbbr = tail
End If
End If
Exit For
ElseIf a = "*" Or a = "***" Or IsSeparatorLine(t) Then
ElseIf a Like "V/V*" Then
Exit For
Else
Call AppendLineToBuffer(organLines, t)
End If
End If
Next p
If grp = "DANG" And Len(organLines) > 0 And InStr(1, organLines, vbCr) = 0 Then
Dim oneAsc As String
oneAsc = Trim(ToAscii(organLines))
If oneAsc Like "DANG UY XA *" Or oneAsc Like "DANG UY PHUONG *" Or oneAsc Like "DANG UY DAC KHU *" Or oneAsc Like "CHI BO *" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & organLines
End If
End If
For Each p In tbl.Range.Paragraphs
Dim tD As String
tD = CleanParaText(p.Range.text)
Dim aD As String
aD = ToAscii(tD)
If InStr(1, aD, "NGAY") > 0 And InStr(1, aD, "THANG") > 0 Then
Dim commaPos As Long
commaPos = InStr(1, tD, ",")
If commaPos > 1 Then placeName = Trim(Left(tD, commaPos - 1))
Exit For
End If
Next p
Exit Sub
ReadFail:
organLines = "": orgAbbr = "": placeName = ""
End Sub
Private Function TT_PartySuperiorPlaceholder() As String
TT_PartySuperiorPlaceholder = "[" & ChrW(272) & ChrW(7842) & "NG B" & ChrW(7896) & " C" & ChrW(7844) & "P TR" & ChrW(202) & "N TR" & ChrW(7920) & "C TI" & ChrW(7870) & "P]"
End Function
Private Function TT_HighlightPlaceholders(ByVal d As Document) As Long
Dim cnt As Long
cnt = 0
On Error Resume Next
Dim p As Paragraph
For Each p In d.Paragraphs
Dim t As String
t = p.Range.text
Dim pos As Long
pos = InStr(1, t, "[")
Dim guard As Integer
guard = 0
Do While pos > 0 And guard < 50
guard = guard + 1
Dim pe As Long
pe = InStr(pos + 1, t, "]")
If pe = 0 Then Exit Do
d.Range(p.Range.Start + pos - 1, p.Range.Start + pe).HighlightColorIndex = wdYellow
cnt = cnt + 1
pos = InStr(pe + 1, t, "[")
Loop
Next p
On Error GoTo 0
TT_HighlightPlaceholders = cnt
End Function
Public Sub TaoVanBanMau()
If Not CheckLicenseActivation() Then Exit Sub
Dim pid As String
pid = TT_GetSetting()
If pid = "AUTO" Then
Dim msgK As String
Dim i As Integer
msgK = "Tao van ban mau cho khoi nao? (go so roi bam OK)" & vbCrLf
For i = 1 To TT_ProfileCount()
msgK = msgK & vbCrLf & i & ". " & TT_ProfileLabelAscii(TT_ProfileIdAt(i))
Next i
Dim nK As Integer
nK = TT_AskNumber(msgK, "Tao van ban mau", 1, TT_ProfileCount())
If nK = 0 Then Exit Sub
pid = TT_ProfileIdAt(nK)
End If
Dim grp As String
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Or TT_MauCount(grp) = 0 Then
MsgBox "Khoi '" & TT_ProfileLabelAscii(pid) & "' chua co mau van ban (bo the thuc dang thu nghiem).", vbInformation, "Tao van ban mau"
Exit Sub
End If
Dim msgM As String
msgM = "Chon mau van ban - khoi " & TT_ProfileLabelAscii(pid) & " (go so roi bam OK):" & vbCrLf
For i = 1 To TT_MauCount(grp)
msgM = msgM & vbCrLf & i & ". " & TT_MauLabelAscii(grp, i)
Next i
Dim nM As Integer
nM = TT_AskNumber(msgM, "Tao van ban mau", 1, TT_MauCount(grp))
If nM = 0 Then Exit Sub
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Call TT_CreateFromTemplate(pid, nM, srcDoc, True)
End Sub
Public Function TT_CreateFromTemplate(ByVal pid As String, ByVal nM As Integer, ByVal srcDoc As Document, ByVal showMsg As Boolean, Optional ByVal khoaDonVi As String = "") As Document
Set TT_CreateFromTemplate = Nothing
If Len(khoaDonVi) = 0 Then khoaDonVi = pid
Dim grp As String
grp = TT_TemplateGroupOf(pid)
If Len(grp) = 0 Or nM < 1 Or nM > TT_MauCount(grp) Then Exit Function
Dim organLines As String, orgAbbr As String, placeName As String
Dim infoSource As String
infoSource = ""
Dim uCoQuan As String
uCoQuan = TT_UnitGet(khoaDonVi, "COQUAN")
If Len(uCoQuan) > 0 Then
Dim uCapTren As String
uCapTren = TT_UnitGet(khoaDonVi, "CAPTREN")
If Len(uCapTren) > 0 Then
organLines = Replace(Replace(uCapTren, " | ", vbCr), "|", vbCr) & vbCr & Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
ElseIf grp = "DANG" Then
organLines = Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
Dim uAsc As String
uAsc = Trim(ToAscii(Replace(organLines, vbCr, " ")))
If uAsc Like "DANG UY XA *" Or uAsc Like "DANG UY PHUONG *" Or uAsc Like "DANG UY DAC KHU *" Or uAsc Like "CHI BO *" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & organLines
End If
Else
organLines = Replace(Replace(uCoQuan, " | ", vbCr), "|", vbCr)
End If
orgAbbr = TT_UnitGet(khoaDonVi, "VIETTAT")
placeName = TT_UnitGet(khoaDonVi, "DIADANH")
infoSource = "DONVI"
Else
Call TT_ReadOrganInfo(srcDoc, grp, organLines, orgAbbr, placeName)
If Len(organLines) > 0 Then infoSource = "VANBAN"
End If
Dim tookInfo As Boolean
tookInfo = (Len(infoSource) > 0)
Application.ScreenUpdating = False
Dim d As Document
Set d = Documents.Add
Call TT_BuildTemplateDoc(d, grp, TT_MauData(grp, nM), organLines, orgAbbr, placeName, _
TT_UnitGet(khoaDonVi, "THAMQUYEN"), TT_UnitGet(khoaDonVi, "CHUCVU"), TT_UnitGet(khoaDonVi, "HOTEN"), pid)
d.Activate
mTTForceProfile = pid
Call ChuanHoa_VanBanDang(True)
mTTForceProfile = ""
On Error Resume Next
d.Variables("TheThuc_BaoCao").Delete
On Error GoTo 0
Dim nPh As Long
nPh = TT_HighlightPlaceholders(d)
Application.ScreenUpdating = True
On Error Resume Next
Dim pSel As Paragraph
For Each pSel In d.Paragraphs
If Not pSel.Range.Information(wdWithInTable) Then
Dim posB As Long
posB = InStr(1, pSel.Range.text, "[")
If posB > 0 Then
Dim posE As Long
posE = InStr(posB, pSel.Range.text, "]")
If posE > posB Then d.Range(pSel.Range.Start + posB - 1, pSel.Range.Start + posE).Select
Exit For
End If
End If
Next pSel
On Error GoTo 0
Set TT_CreateFromTemplate = d
If Not showMsg Then Exit Function
Dim info As String
If infoSource = "DONVI" Then
info = "Ten co quan, ky hieu, dia danh, nguoi ky lay tu Thong tin don vi da luu (sua tren Ribbon)."
ElseIf infoSource = "VANBAN" Then
info = "Ten co quan, ky hieu, dia danh lay tu van ban dang mo."
Else
info = "Chua nhap Thong tin don vi tren Ribbon: ten co quan, ky hieu, dia danh de trong cho dien."
End If
MsgBox "Da tao mau: " & TT_MauLabelAscii(grp, nM) & " (khoi " & TT_ProfileLabelAscii(pid) & ")." & vbCrLf & _
info & vbCrLf & "Con " & nPh & " cho can dien trong dau [ ] (to vang). So va ngay thang de trong cho van thu.", _
vbInformation, "Tao van ban mau"
End Function
Public Sub TT_KiemThuDonVi()
Dim baseDir As String
baseDir = Environ("HOME")
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
If Documents.Count = 0 Then Exit Sub
Dim src As Document
Set src = ActiveDocument
Dim ok As Boolean
ok = TT_UnitLayTuVanBan(False)
Dim grp As String
grp = GetSetting(TT_SETTING_APP, TT_SETTING_SECTION, "NhomGanNhat", "DANG")
Call TT_DebugLog("LayTuVanBan=" & ok & " nhom=" & grp & " | CAPTREN=" & ToAscii(TT_UnitGet(grp, "CAPTREN")) & " | COQUAN=" & ToAscii(TT_UnitGet(grp, "COQUAN")) & _
" | VIETTAT=" & ToAscii(TT_UnitGet(grp, "VIETTAT")) & " | DIADANH=" & ToAscii(TT_UnitGet(grp, "DIADANH")) & _
" | THAMQUYEN=" & ToAscii(TT_UnitGet(grp, "THAMQUYEN")) & " | CHUCVU=" & ToAscii(TT_UnitGet(grp, "CHUCVU")) & " | HOTEN=" & ToAscii(TT_UnitGet(grp, "HOTEN")))
src.Close SaveChanges:=False
Dim idxCV As Integer
Dim i As Integer
For i = 1 To TT_MauCount(grp)
If TT_MauLabelAscii(grp, i) = "Cong van" Then idxCV = i
Next i
Dim d As Document
Set d = TT_TaoMauTheoTag(grp & "|" & idxCV, False)
If Not d Is Nothing Then
On Error Resume Next
d.SaveAs FileName:=baseDir & "/Documents/thethuc_test/out/donvi_" & grp & ".docx", FileFormat:=wdFormatXMLDocument
d.Close SaveChanges:=False
On Error GoTo 0
End If
End Sub
Public Sub TT_KiemThuTatCaMau()
Dim baseDir As String
baseDir = Environ("HOME")
If Len(Dir(baseDir & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim outDir As String
outDir = baseDir & "/Documents/thethuc_test/out/"
Dim srcDoc As Document
Set srcDoc = Nothing
On Error Resume Next
If Documents.Count > 0 Then Set srcDoc = ActiveDocument
On Error GoTo 0
Dim tag As String
If srcDoc Is Nothing Then
tag = "trong"
Else
tag = "nguon"
End If
Dim pids(1 To 3) As String
pids(1) = "DANG": pids(2) = "CHINH_QUYEN": pids(3) = "CONG_AN"
Dim ip As Integer, im As Integer
For ip = 1 To 3
For im = 1 To TT_MauCount(TT_TemplateGroupOf(pids(ip)))
Dim d As Document
Set d = TT_CreateFromTemplate(pids(ip), im, srcDoc, False)
If Not d Is Nothing Then
On Error Resume Next
d.Variables.Add "TheThuc_BaoCao", "Mau " & pids(ip) & " " & im
d.SaveAs FileName:=outDir & "mau_" & tag & "_" & pids(ip) & "_" & Format(im, "00") & ".docx", FileFormat:=wdFormatXMLDocument
d.Close SaveChanges:=False
On Error GoTo 0
Call TT_DebugLog("Tao mau " & pids(ip) & " " & im & " xong")
End If
Next im
Next ip
If Not srcDoc Is Nothing Then srcDoc.Activate
End Sub
Private Sub TT_BuildTemplateDoc(ByVal d As Document, ByVal grp As String, ByVal data As String, _
ByVal organLines As String, ByVal orgAbbr As String, ByVal placeName As String, _
ByVal uAuth As String, ByVal uRole As String, ByVal uName As String, _
Optional ByVal pid As String = "")
On Error Resume Next
Dim code As String, isCV As Boolean, summary As String
Dim bodyText As String, recipText As String
Dim auth As String, role As String, signer As String
Dim lRole As String, lName As String, hasLeftSigner As Boolean, noDate As Boolean
Dim noNumber As Boolean, noSigner As Boolean
isCV = False: hasLeftSigner = False: noDate = False: noNumber = False: noSigner = False
Dim items() As String
items = Split(data, vbLf)
Dim k As Long
For k = LBound(items) To UBound(items)
Dim it As String
it = items(k)
Dim bar As Long
bar = InStr(1, it, "|")
If bar > 0 Then
Dim tg As String, v As String
tg = Left(it, bar - 1)
v = Mid(it, bar + 1)
Select Case tg
Case "CODE": code = v
Case "CV": isCV = True
Case "ND": noDate = True
Case "NS": noNumber = True
Case "NK": noSigner = True
Case "SUM": summary = v
Case "L": Call AppendLineToBuffer(bodyText, v)
Case "N": Call AppendLineToBuffer(recipText, v)
Case "A": auth = v
Case "R": role = v
Case "H": signer = v
Case "LR": lRole = v: hasLeftSigner = True
Case "LH": lName = v
End Select
End If
Next k
If Not hasLeftSigner Then
If Len(uRole) > 0 Or Len(uName) > 0 Then
auth = uAuth
ElseIf Len(uAuth) > 0 Then
auth = uAuth
End If
If Len(uRole) > 0 Then role = uRole
If Len(uName) > 0 Then signer = uName
End If
If Len(pid) > 0 And Left(signer, 1) = "[" Then
If TT_RawB(pid, "RANK_BEFORE_NAME") Then signer = "[C" & ChrW(7845) & "p b" & ChrW(7853) & "c h" & ChrW(224) & "m] " & signer
End If
If hasLeftSigner And Len(pid) > 0 And Left(lName, 1) = "[" Then
If TT_RawB(pid, "RANK_BEFORE_NAME") Then lName = "[C" & ChrW(7845) & "p b" & ChrW(7853) & "c h" & ChrW(224) & "m] " & lName
End If
If Len(organLines) = 0 Then
If grp = "DANG" Then
organLines = TT_PartySuperiorPlaceholder() & vbCr & _
"[" & ChrW(272) & ChrW(7842) & "NG " & ChrW(7910) & "Y / CHI B" & ChrW(7896) & " BAN H" & ChrW(192) & "NH]"
ElseIf pid = "CONG_AN" Or pid = "QUAN_DOI" Or pid = "MTTQ" Then
Select Case pid
Case "CONG_AN": organLines = "[C" & ChrW(212) & "NG AN T" & ChrW(7880) & "NH ...]" & vbCr & "[C" & ChrW(212) & "NG AN X" & ChrW(195) & " ...]"
Case "QUAN_DOI": organLines = "[B" & ChrW(7896) & " CH" & ChrW(7880) & " HUY QU" & ChrW(194) & "N S" & ChrW(7920) & " T" & ChrW(7880) & "NH ...]" & vbCr & "[BAN CH" & ChrW(7880) & " HUY QU" & ChrW(194) & "N S" & ChrW(7920) & " X" & ChrW(195) & " ...]"
Case "MTTQ": organLines = "[" & ChrW(7910) & "Y BAN MTTQ VI" & ChrW(7878) & "T NAM T" & ChrW(7880) & "NH ...]" & vbCr & "[" & ChrW(7910) & "Y BAN MTTQ VI" & ChrW(7878) & "T NAM X" & ChrW(195) & " ...]"
End Select
Else
organLines = "[T" & ChrW(202) & "N C" & ChrW(416) & " QUAN C" & ChrW(7844) & "P TR" & ChrW(202) & "N (N" & ChrW(7870) & "U C" & ChrW(211) & ")]" & vbCr & _
"[T" & ChrW(202) & "N C" & ChrW(416) & " QUAN BAN H" & ChrW(192) & "NH]"
End If
End If
If Len(orgAbbr) = 0 Then orgAbbr = "[VI" & ChrW(7870) & "T T" & ChrW(7854) & "T]"
If Len(placeName) = 0 Then placeName = "[" & ChrW(272) & ChrW(7883) & "a danh]"
bodyText = Replace(bodyText, "{VIETTAT}", orgAbbr)
Dim soLine As String
If grp = "DANG" Then
soLine = "S" & ChrW(7889) & "   -" & code & "/" & orgAbbr
ElseIf isCV Then
soLine = "S" & ChrW(7889) & ":   /" & orgAbbr & "-[" & ChrW(272) & "VST]"
Else
soLine = "S" & ChrW(7889) & ":   /" & code & "-" & orgAbbr
End If
Dim leftText As String
leftText = organLines
If grp = "DANG" Then Call AppendLineToBuffer(leftText, "*")
If Not noNumber Then Call AppendLineToBuffer(leftText, soLine)
If isCV And Len(summary) > 0 Then Call AppendLineToBuffer(leftText, summary)
Dim rightText As String
If grp = "DANG" Then
rightText = TT_TextTieuNguDang()
Else
rightText = TT_TextQuocHieu() & vbCr & TT_TextTieuNgu()
End If
If Not noDate Then Call AppendLineToBuffer(rightText, placeName & ", ng" & ChrW(224) & "y    th" & ChrW(225) & "ng    n" & ChrW(259) & "m " & Year(Date))
Dim tH As Table
Set tH = d.Tables.Add(d.Range(0, 0), 1, 2)
tH.Borders.Enable = False
tH.Cell(1, 1).Range.text = leftText
tH.Cell(1, 2).Range.text = rightText
Dim rB As Range
Set rB = d.Paragraphs(d.Paragraphs.Count).Range
rB.InsertBefore bodyText & vbCr
If noSigner Then GoTo XongChuKy
Dim leftF As String, rightF As String
If hasLeftSigner Then
leftF = lRole & vbCr & vbCr & vbCr & vbCr & vbCr & lName
ElseIf Len(recipText) > 0 Then
leftF = "N" & ChrW(417) & "i nh" & ChrW(7853) & "n:" & vbCr & recipText
End If
If Len(auth) > 0 Then rightF = Replace(Replace(auth, " | ", vbCr), "|", vbCr) & vbCr
rightF = rightF & Replace(Replace(role, " | ", vbCr), "|", vbCr) & vbCr & vbCr & vbCr & vbCr & vbCr & signer
Dim rF As Range
Set rF = d.Paragraphs(d.Paragraphs.Count).Range
Dim tF As Table
Set tF = d.Tables.Add(rF, 1, 2)
tF.Borders.Enable = False
tF.Cell(1, 1).Range.text = leftF
tF.Cell(1, 2).Range.text = rightF
If hasLeftSigner And Len(recipText) > 0 Then
Dim rN As Range
Set rN = d.Range(d.Content.End - 1, d.Content.End - 1)
rN.InsertAfter vbCr & "N" & ChrW(417) & "i nh" & ChrW(7853) & "n:" & vbCr & recipText
End If
XongChuKy:
On Error GoTo 0
End Sub
Public Sub RibbonChuanHoa(control As Object)
Call ChuanHoa_VanBanDang(False)
End Sub
Public Sub ChuanHoa_VanBanDang_Silent()
Call ChuanHoa_VanBanDang(True)
End Sub
Sub ChuanHoa_VanBanDang(Optional ByVal silentMode As Boolean = False)
If Documents.Count = 0 Then Exit Sub
If Not silentMode Then
Call CN_TuKiemTraHangTuan(True)
If Not CheckLicenseActivation() Then Exit Sub
End If
Application.ScreenUpdating = False
On Error Resume Next
Dim cuChinhTa As Boolean, cuNguPhap As Boolean
cuChinhTa = Application.Options.CheckSpellingAsYouType
cuNguPhap = Application.Options.CheckGrammarAsYouType
If cuChinhTa Then Application.Options.CheckSpellingAsYouType = False
If cuNguPhap Then Application.Options.CheckGrammarAsYouType = False
Dim doc As Document
Set doc = ActiveDocument
mTTOvPid = ""
mTTOvStr = ""
Dim docPhanSau As Document
Set docPhanSau = Nothing
Set docPhanSau = TT_TachVanBanTiepTheo(doc)
If Not (docPhanSau Is Nothing) Then
docPhanSau.Activate
Call ChuanHoa_VanBanDang(True)
doc.Activate
Application.ScreenUpdating = False
On Error Resume Next
End If
Call TT_DebugLog("BAT DAU: " & doc.Name & " | doan van: " & doc.Paragraphs.Count)
Call TT_Resolve(doc)
Call TT_DebugLog("Khoi: " & TT_Profile())
On Error Resume Next
doc.ConvertNumbersToText
Call TT_DebugLog("0.0 ConvertNumbersToText xong")
Call TT_DoiDauDauDong(doc)
Call TT_DebugLog("0.01 Doi dau dau dong xong")
Call TT_GoBangBoc(doc)
Call TT_GoHopVanBanDauTrang(doc)
On Error Resume Next
Call TT_KhoiPhucKeTrichYeu(doc)
Call RemoveHeaderFloatingShapes(doc, Nothing)
Call TT_DebugLog("0.1 RemoveHeaderFloatingShapes xong")
Call SafeReplaceSoftBreaksInRange(doc.Content)
Call TT_DebugLog("0.2 SafeReplaceSoftBreaks xong")
Call TT_ChuanKinhGui(doc)
Call AutoCreateHeaderTable(doc)
Call TT_DebugLog("0.3a AutoCreateHeaderTable xong")
Call SplitGluedHeaderAndTitleLines(doc)
Call TT_DebugLog("0.3b SplitGlued xong")
Call AutoCreateKinhGuiTable(doc)
Call TT_DebugLog("0.3c AutoCreateKinhGuiTable xong")
Call AutoCreateFooterTable(doc)
Call TT_DebugLog("0.3d AutoCreateFooterTable xong")
Dim pCur As Paragraph, pPrev As Paragraph, demXoa As Long
Set pCur = Nothing
Set pCur = doc.Paragraphs.Last
Do While Not (pCur Is Nothing)
demXoa = demXoa + 1
If demXoa > 200000 Then Exit Do
Set pPrev = Nothing
Set pPrev = pCur.Previous
If pPrev Is Nothing Then Exit Do
If Len(CleanParaText(pCur.Range.text)) = 0 Then
If Len(CleanParaText(pPrev.Range.text)) = 0 Then
If Not pCur.Range.Information(wdWithInTable) And Not pPrev.Range.Information(wdWithInTable) Then
On Error Resume Next
pCur.Range.Delete
On Error Resume Next
End If
End If
End If
Set pCur = pPrev
Loop
Call TT_DebugLog("BUOC 1 bat dau")
Dim sec As Section
For Each sec In doc.Sections
On Error Resume Next
With sec.PageSetup
.PaperSize = wdPaperA4
.TopMargin = MillimetersToPoints(TT_N("PAGE_TOP"))
.BottomMargin = MillimetersToPoints(TT_N("PAGE_BOTTOM"))
.LeftMargin = MillimetersToPoints(TT_N("PAGE_LEFT"))
.RightMargin = MillimetersToPoints(TT_N("PAGE_RIGHT"))
.HeaderDistance = MillimetersToPoints(TT_N("PAGE_HEADER"))
.FooterDistance = MillimetersToPoints(TT_N("PAGE_FOOTER"))
.MirrorMargins = False
.DifferentFirstPageHeaderFooter = True
End With
On Error Resume Next
Next sec
Call TT_DebugLog("BUOC 2 bat dau")
For Each sec In doc.Sections
On Error Resume Next
Dim hdr As HeaderFooter
Set hdr = sec.Headers(wdHeaderFooterPrimary)
hdr.Range.Delete
hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
hdr.Range.Font.Name = "Times New Roman"
hdr.Range.Font.Size = TT_N("PAGE_NUMBER_SIZE")
hdr.Range.Font.Bold = False
hdr.Range.Font.Italic = False
hdr.Range.Font.ColorIndex = wdBlack
hdr.Range.Fields.Add Range:=hdr.Range, Type:=wdFieldPage
hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
sec.Headers(wdHeaderFooterFirstPage).Range.Delete
On Error Resume Next
Next sec
Call TT_DebugLog("BUOC 3 bat dau")
Dim tFit As Table
For Each tFit In doc.Tables
On Error Resume Next
tFit.AutoFitBehavior 2
tFit.AllowAutoFit = True
On Error Resume Next
Next tFit
Dim tbl As Table
Dim tblCounter As Integer
tblCounter = 0
Dim c As Cell
Dim pCell As Paragraph
Dim pDCS As Paragraph
Dim cRight As Cell
For Each tbl In doc.Tables
tblCounter = tblCounter + 1
Dim tAsc As String
tAsc = ToAscii(tbl.Range.text)
Dim isHeaderTbl As Boolean
isHeaderTbl = False
Dim hdrDetectStrong As Boolean
If tblCounter = 1 And (InStr(1, tAsc, "DANG CONG SAN") > 0 Or InStr(1, tAsc, "VIET NAM") > 0 Or InStr(1, tAsc, "DANG BO") > 0 Or InStr(1, tAsc, "DANG UY") > 0 Or InStr(1, tAsc, "CHI BO") > 0 Or InStr(1, tAsc, "TINH UY") > 0 Or InStr(1, tAsc, "THANH UY") > 0 Or InStr(1, tAsc, "HUYEN UY") > 0) Then
isHeaderTbl = True
ElseIf tblCounter = 1 And tbl.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(tAsc, hdrDetectStrong)) > 0 And InStr(1, tAsc, "NOI NHAN") = 0 Then
isHeaderTbl = True
ElseIf InStr(1, tAsc, "DANG CONG SAN VIET NAM") > 0 Or InStr(1, tAsc, "CONG HOA XA HOI CHU NGHIA") > 0 Then
isHeaderTbl = True
End If
If isHeaderTbl Then
Set pDCS = Nothing
Set cRight = Nothing
Dim pQuocHieu As Paragraph: Set pQuocHieu = Nothing
Dim pTieuNgu As Paragraph: Set pTieuNgu = Nothing
Dim pNgayThangHdr As Paragraph: Set pNgayThangHdr = Nothing
Dim pLastOrgan As Paragraph: Set pLastOrgan = Nothing
Dim organParas As Collection: Set organParas = New Collection
Call RemoveHeaderFloatingShapes(doc, tbl)
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim totalCells As Integer
totalCells = tbl.Range.Cells.Count
Dim cellIdx As Integer
cellIdx = 1
Dim leftColCqIdx As Integer
leftColCqIdx = 1
Dim leftColInTrichYeu As Boolean
leftColInTrichYeu = False
Dim leftColSawSo As Boolean
leftColSawSo = False
Dim sawKinhGuiRowInHeaderTbl As Boolean
sawKinhGuiRowInHeaderTbl = False
For Each c In tbl.Range.Cells
Dim cellTextRaw As String
cellTextRaw = CleanParaText(c.Range.text)
If Len(cellTextRaw) = 0 Then GoTo NextCellHeader
Dim cAsc As String
cAsc = ToAscii(c.Range.text)
If HasKinhGuiOrTrinh(cAsc) Then
sawKinhGuiRowInHeaderTbl = True
GoTo NextCellHeader
End If
If sawKinhGuiRowInHeaderTbl Or Left(cellTextRaw, 1) = "-" Then
GoTo NextCellHeader
End If
Dim isRightCell As Boolean
isRightCell = False
If InStr(1, cAsc, "DANG CONG SAN") > 0 Or InStr(1, cAsc, "DOC LAP - TU DO") > 0 Or InStr(1, cAsc, "CONG HOA XA HOI") > 0 Then
isRightCell = True
Set cRight = c
ElseIf InStr(1, cAsc, "NGAY") > 0 And InStr(1, cAsc, "THANG") > 0 And InStr(1, cAsc, "NAM") > 0 And InStr(1, cAsc, "V/V") = 0 Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
ElseIf InStr(1, cAsc, "DANG BO") > 0 Or InStr(1, cAsc, "DANG UY") > 0 Or InStr(1, cAsc, "CHI BO") > 0 Or _
InStr(1, cAsc, "TINH UY") > 0 Or InStr(1, cAsc, "THANH UY") > 0 Or InStr(1, cAsc, "HUYEN UY") > 0 Or _
InStr(1, cAsc, "SO ") > 0 Or InStr(1, cAsc, "SO:") > 0 Or InStr(1, cAsc, "V/V") > 0 Or InStr(1, cAsc, "VE VIEC") > 0 Then
isRightCell = False
Else
On Error Resume Next
If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
ElseIf cellIdx = totalCells And totalCells = 2 Then
isRightCell = True
If cRight Is Nothing Then Set cRight = c
End If
On Error Resume Next
End If
On Error Resume Next
c.Borders.Enable = False
c.Shading.BackgroundPatternColor = wdColorAutomatic
c.Shading.Texture = wdTextureNone
Call SafeReplaceSoftBreaksInRange(c.Range)
On Error Resume Next
Call RemoveEmptyParagraphsInCell(doc, c)
For Each pCell In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pCell)
Dim pAsc As String
pAsc = Trim(ToAscii(pCell.Range.text))
If Len(pAsc) = 0 Then GoTo SkipEmptyLineHeader
With pCell.Format
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
If isRightCell Then
pCell.Format.Alignment = TT_RightColumnAlignment()
If InStr(1, pAsc, "CONG HOA XA HOI") > 0 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Size = TT_N("QH_SIZE")
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.Alignment = wdAlignParagraphCenter
Set pQuocHieu = pCell
ElseIf InStr(1, Replace(pAsc, " ", ""), "DOCLAP-TUDO") > 0 Then
If CleanParaText(pCell.Range.text) <> TT_TextTieuNgu() And InStr(1, Replace(pAsc, " ", ""), "DOCLAP-TUDO-HANHPHUC") > 0 Then
Call SafeSetParagraphText(doc, pCell, TT_TextTieuNgu())
End If
pCell.Range.Font.Size = TT_N("TN_SIZE")
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Format.SpaceAfter = 8
Set pTieuNgu = pCell
ElseIf InStr(1, pAsc, "DANG CONG SAN") > 0 Or InStr(1, pAsc, "VIET NAM") > 0 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Size = 15
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
pCell.Format.SpaceAfter = 4
Set pDCS = pCell
ElseIf InStr(1, pAsc, "NGAY") > 0 And InStr(1, pAsc, "THANG") > 0 Then
Set pNgayThangHdr = pCell
Dim origNgay As String, fixNgay As String
origNgay = CleanParaText(pCell.Range.text)
fixNgay = ChuanHoa_NgayThang(origNgay)
If fixNgay <> origNgay And Len(fixNgay) > 0 Then
Call SafeSetParagraphText(doc, pCell, fixNgay)
End If
pCell.Format.Alignment = TT_RightColumnAlignment()
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("DATE_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = True
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
ElseIf pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
Call SafeSetParagraphText(doc, pCell, "")
ElseIf TT_S("MOTTO_KIND") = "LINH_HOAT" Then
pCell.Range.Font.Size = 14
Else
pCell.Range.Font.Size = 14
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
End If
Else
If leftColSawSo Then
If pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = 10
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
Else
leftColInTrichYeu = True
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
End If
ElseIf pAsc Like "V/V*" Or pAsc Like "V.V*" Or pAsc Like "VE VIEC*" Or InStr(1, pAsc, "V/V") > 0 Or InStr(1, pAsc, "V.V") > 0 Then
leftColInTrichYeu = True
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
ElseIf IsDocNumberLineAsc(pAsc) Then
leftColSawSo = True
Dim origSo As String, fixSo As String
origSo = CleanParaText(pCell.Range.text)
fixSo = ChuanHoa_SoKyHieu(origSo)
If fixSo <> origSo And Len(fixSo) > 0 Then
Call SafeSetParagraphText(doc, pCell, fixSo)
End If
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Name = "Times New Roman"
pCell.Range.Font.Size = TT_N("DOC_NUMBER_SIZE")
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
pCell.Range.Font.ColorIndex = wdBlack
pCell.Range.Font.Underline = wdUnderlineNone
pCell.Format.SpaceBefore = 0
pCell.Format.SpaceAfter = 0
pCell.Format.LineSpacingRule = wdLineSpaceSingle
ElseIf pAsc = "*" Or pAsc = "***" Or IsSeparatorLine(pAsc) Then
If TT_B("ORGAN_UNDERLINE") Then
Call SafeSetParagraphText(doc, pCell, "")
Else
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = 14
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
End If
Else
Set pLastOrgan = pCell
organParas.Add pCell
pCell.Format.Alignment = wdAlignParagraphCenter
pCell.Range.Font.Size = TT_N("ORGAN_SIZE")
If leftColCqIdx = 1 Then
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Bold = False
pCell.Range.Font.Italic = False
Else
pCell.Range.Case = wdUpperCase
pCell.Range.Font.Bold = True
pCell.Range.Font.Italic = False
End If
leftColCqIdx = leftColCqIdx + 1
End If
End If
SkipEmptyLineHeader:
Next pCell
NextCellHeader:
cellIdx = cellIdx + 1
Next c
If TT_IsNational() And leftColCqIdx = 3 Then
On Error Resume Next
Dim orgFirstP As Paragraph
Set orgFirstP = tbl.Range.Cells(1).Range.Paragraphs(1)
Dim orgFirstAsc As String
orgFirstAsc = Trim(ToAscii(CleanParaText(orgFirstP.Range.text)))
If orgFirstAsc = "UY BAN NHAN DAN" Or orgFirstAsc = "HOI DONG NHAN DAN" Then orgFirstP.Range.Font.Bold = True
On Error Resume Next
End If
If leftColCqIdx = 2 Then
On Error Resume Next
Dim firstLeftP As Paragraph
Set firstLeftP = tbl.Range.Cells(1).Range.Paragraphs(1)
firstLeftP.Range.Font.Bold = True
On Error Resume Next
End If
Call TT_InDamTenCoQuan(organParas)
If TT_IsNational() Then
On Error Resume Next
If tbl.Columns.Count = 2 Then
Dim hdrTotalW As Single
hdrTotalW = 0
With tbl.Range.Sections(1).PageSetup
hdrTotalW = .PageWidth - .LeftMargin - .RightMargin
End With
If hdrTotalW < 300 Or hdrTotalW > 800 Then hdrTotalW = MillimetersToPoints(165)
Dim hdrRightW As Single
hdrRightW = MillimetersToPoints(TT_N("HDR_RIGHT_MM"))
Dim hdrMinLeftW As Single
hdrMinLeftW = MillimetersToPoints(TT_N("HDR_LEFT_MIN_MM"))
If hdrTotalW - hdrRightW < hdrMinLeftW Then hdrRightW = hdrTotalW - hdrMinLeftW
tbl.AllowAutoFit = False
If tbl.Rows.LeftIndent < -MillimetersToPoints(3) Then tbl.Rows.LeftIndent = 0
Dim hdrLanW As Single: hdrLanW = 0
tbl.Columns(1).Width = hdrTotalW - hdrRightW
tbl.Columns(2).Width = hdrRightW
If Not (pQuocHieu Is Nothing) Then
Dim shiftStep As Integer
Dim qhLines As Long
For shiftStep = 1 To 40
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines <= 1 Then Exit For
If hdrTotalW - hdrRightW - MillimetersToPoints(2) < hdrMinLeftW Then Exit For
hdrRightW = hdrRightW + MillimetersToPoints(2)
tbl.Columns(1).Width = hdrTotalW - hdrRightW
tbl.Columns(2).Width = hdrRightW
Next shiftStep
Call TT_DebugLog("  Cot Hdr: trai " & Format(PointsToMillimeters(hdrTotalW - hdrRightW), "0") & "mm, phai " & Format(PointsToMillimeters(hdrRightW), "0") & "mm, so lan doi " & (shiftStep - 1))
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines > 1 Then
pQuocHieu.Range.Font.Size = 12
If Not (pTieuNgu Is Nothing) Then pTieuNgu.Range.Font.Size = 13
If Not (pNgayThangHdr Is Nothing) Then pNgayThangHdr.Range.Font.Size = 13
Call TT_DebugLog("  Quoc hieu van xuong dong -> giam dong bo 12/13/13")
qhLines = 1
qhLines = pQuocHieu.Range.ComputeStatistics(wdStatisticLines)
If qhLines > 1 Then
mTTHeaderNote = "[Quoc hieu] Quoc hieu van bi xuong dong du da noi rong cot va giam co chu 12/13/13. Kiem tra lai do rong bang tieu de (Quoc hieu phai nam tren 1 dong)."
End If
End If
If TT_OrganWraps(organParas) Then Set organParas = TT_SplitLongOrgans(organParas, True): If organParas.Count > 0 Then Set pLastOrgan = organParas(organParas.Count)
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
Call TT_DebugLog("  Truoc buoc ten dai (" & TT_TenDaiMode() & "): cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0") & ", rot dong " & CStr(TT_OrganWraps(organParas)) & ", QH dong " & TT_ParaLines(pQuocHieu))
If TT_TenDaiMode() = "LAN_LE" Then
Call TT_LanLeTrai(tbl, organParas, hdrTotalW, hdrRightW, hdrLanW)
Else
If TT_OrganWraps(organParas) And pQuocHieu.Range.Font.Size > 12 Then
pQuocHieu.Range.Font.Size = 12
If Not (pTieuNgu Is Nothing) Then pTieuNgu.Range.Font.Size = 13
If Not (pNgayThangHdr Is Nothing) Then pNgayThangHdr.Range.Font.Size = 13
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
If TT_OrganWraps(organParas) Then
Dim oPara As Variant
For Each oPara In organParas
If oPara.Range.Font.Size > 12 Then oPara.Range.Font.Size = 12
Next oPara
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
End If
If TT_OrganWraps(organParas) Then
Set organParas = TT_SplitLongOrgans(organParas, False)
If organParas.Count > 0 Then Set pLastOrgan = organParas(organParas.Count)
Call TT_WidenOrganCol(tbl, organParas, pQuocHieu, pNgayThangHdr, hdrTotalW, hdrRightW)
End If
Call TT_DebugLog("  Noi cot ten co quan: trai " & Format(PointsToMillimeters(hdrTotalW - hdrRightW), "0") & "mm")
End If
End If
If TT_B("ORGAN_UNDERLINE") And Not (pLastOrgan Is Nothing) Then
If pLastOrgan.Format.SpaceAfter < 6 Then pLastOrgan.Format.SpaceAfter = 6
End If
On Error Resume Next
ElseIf tbl.Columns.Count = 2 And Not (pDCS Is Nothing) Then
On Error Resume Next
Dim dTotalW As Single, dRightW As Single
Dim pDateDang As Paragraph, pR As Paragraph, pRAsc As String
Set pDateDang = Nothing
For Each pR In tbl.Cell(1, 2).Range.Paragraphs
pRAsc = UCase(ToAscii(CleanParaText(pR.Range.text)))
If InStr(1, pRAsc, "NGAY") > 0 And InStr(1, pRAsc, "NAM") > 0 Then
Set pDateDang = pR
Exit For
End If
Next pR
Dim dLanW As Single: dLanW = 0
If tbl.Rows.LeftIndent < -MillimetersToPoints(3) Then
tbl.Rows.LeftIndent = 0
With tbl.Range.Sections(1).PageSetup
tbl.Columns(1).Width = .PageWidth - .LeftMargin - .RightMargin - tbl.Columns(2).Width
End With
End If
dTotalW = tbl.Columns(1).Width + tbl.Columns(2).Width
Call TT_DebugLog("  Khoi Dang - co dong ngay thang: " & CStr(Not (pDateDang Is Nothing)) & ", so dong " & TT_ParaLines(pDateDang) & ", ten co quan rot dong: " & CStr(TT_OrganWraps(organParas)) & ", cot " & Format(PointsToMillimeters(tbl.Columns(1).Width), "0") & "/" & Format(PointsToMillimeters(tbl.Columns(2).Width), "0"))
dRightW = tbl.Columns(2).Width
If TT_OrganWraps(organParas) Or TT_ParaLines(pDateDang) > 1 Then tbl.AllowAutoFit = False
If TT_OrganWraps(organParas) Then
Call TT_WidenOrganCol(tbl, organParas, pDCS, pDateDang, dTotalW, dRightW)
If TT_TenDaiMode() = "LAN_LE" Then Call TT_LanLeTrai(tbl, organParas, dTotalW, dRightW, dLanW)
If TT_OrganWraps(organParas) Then
Set organParas = TT_SplitLongOrgans(organParas, False)
Call TT_WidenOrganCol(tbl, organParas, pDCS, pDateDang, dTotalW, dRightW)
End If
End If
Call TT_NarrowForRight(tbl, organParas, pDateDang, dTotalW, dRightW)
Call TT_DebugLog("  Khoi Dang - cot trai " & Format(PointsToMillimeters(dTotalW - dRightW), "0") & "mm")
On Error Resume Next
End If
If Not (pTieuNgu Is Nothing) Then
Call TT_AddCenteredLineUnder(doc, pTieuNgu, "LineTieuNguQuocHieu", 1, CentimetersToPoints(6))
End If
If TT_B("ORGAN_UNDERLINE") And Not (pLastOrgan Is Nothing) Then
Call TT_AddCenteredLineUnder(doc, pLastOrgan, "LineCoQuanBanHanh", 0.4, CentimetersToPoints(5))
End If
If Not (pDCS Is Nothing) And pTieuNgu Is Nothing Then
pDCS.Range.Font.Underline = wdUnderlineNone
On Error Resume Next
pDCS.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
On Error Resume Next
Dim cRightCell As Cell
Set cRightCell = Nothing
On Error Resume Next
Set cRightCell = pDCS.Range.Cells(1)
On Error Resume Next
If cRightCell Is Nothing Then Set cRightCell = cRight
Dim rPad As Single: rPad = 5.4
Dim cellW As Single: cellW = 248
If Not (cRightCell Is Nothing) Then
cellW = cRightCell.Width
On Error Resume Next
rPad = cRightCell.RightPadding
On Error Resume Next
End If
Dim lineWidth As Single: lineWidth = CentimetersToPoints(7.2)
Dim lPos As Single: lPos = cellW - lineWidth - rPad - 5.5
If lPos < 0 Then lPos = 0
On Error Resume Next
Call TT_ChenKe(doc, pDCS, "LineTieuNguDang", lPos, 16.5, lineWidth, True)
On Error Resume Next
End If
Exit For
End If
Next tbl
Call TT_DebugLog("BUOC 3.2 bat dau")
For Each tbl In doc.Tables
Dim tAscKG As String
tAscKG = ToAscii(tbl.Range.text)
If IsKinhGuiTable(tbl, tAscKG) Then
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim cKG As Cell
Dim seenKinhGuiCell As Boolean
seenKinhGuiCell = False
For Each cKG In tbl.Range.Cells
Dim cellTxtKG As String
cellTxtKG = CleanParaText(cKG.Range.text)
If Len(cellTxtKG) > 0 Then
Dim cAscKG As String
cAscKG = Trim(ToAscii(cellTxtKG))
If InStr(1, cAscKG, "DANG CONG SAN") > 0 Or InStr(1, cAscKG, "CONG HOA XA HOI") > 0 Or _
InStr(1, Replace(cAscKG, " ", ""), "DOCLAP-TUDO") > 0 Or _
(InStr(1, cAscKG, "NGAY") > 0 And InStr(1, cAscKG, "THANG") > 0 And InStr(1, cAscKG, "NAM") > 0) Or _
IsDocNumberLineAsc(cAscKG) Then
GoTo NextCellKG
End If
If HasKinhGuiOrTrinh(cAscKG) Then
seenKinhGuiCell = True
On Error Resume Next
cKG.Borders.Enable = False
cKG.Shading.BackgroundPatternColor = wdColorAutomatic
cKG.Shading.Texture = wdTextureNone
cKG.VerticalAlignment = wdCellAlignVerticalTop
On Error Resume Next
cKG.Range.Font.Bold = False
cKG.Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
Dim pKGLeft As Paragraph
For Each pKGLeft In cKG.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pKGLeft)
Dim pAscLKG As String
pAscLKG = Trim(ToAscii(pKGLeft.Range.text))
If Len(pAscLKG) > 0 Then
Call SafeSetParagraphText(doc, pKGLeft, KinhGuiPrefixFor(pAscLKG))
With pKGLeft.Format
.Alignment = wdAlignParagraphRight
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pKGLeft.Range.Font.Name = "Times New Roman"
pKGLeft.Range.Font.Size = 14
pKGLeft.Range.Font.Bold = False
pKGLeft.Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
pKGLeft.Range.Font.ColorIndex = wdBlack
pKGLeft.Range.Font.Underline = wdUnderlineNone
End If
Next pKGLeft
ElseIf seenKinhGuiCell Then
On Error Resume Next
cKG.Borders.Enable = False
cKG.Shading.BackgroundPatternColor = wdColorAutomatic
cKG.Shading.Texture = wdTextureNone
cKG.VerticalAlignment = wdCellAlignVerticalTop
Call SafeReplaceSoftBreaksInRange(cKG.Range)
On Error Resume Next
cKG.Range.Font.Bold = False
cKG.Range.Font.Italic = False
Dim lastLineKG As Integer: lastLineKG = 0
Dim counterKG As Integer: counterKG = 1
Dim tPkg As Paragraph
For Each tPkg In cKG.Range.Paragraphs
If Len(CleanParaText(tPkg.Range.text)) > 0 Then lastLineKG = counterKG
counterKG = counterKG + 1
Next tPkg
Dim pKGRight As Paragraph
Dim lineKGIdx As Integer: lineKGIdx = 1
For Each pKGRight In cKG.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pKGRight)
If Len(CleanParaText(pKGRight.Range.text)) > 0 Then
If lastLineKG > 1 Then
Dim fChr As String
fChr = pKGRight.Range.Characters(1).text
If fChr = ChrW(8211) Or fChr = ChrW(8212) Or fChr = ChrW(8210) Or fChr = ChrW(8722) Or _
fChr = ChrW(8226) Or fChr = ChrW(8259) Or fChr = ChrW(183) Or fChr = "+" Or fChr = "*" Then
pKGRight.Range.Characters(1).text = "-"
ElseIf fChr <> "-" Then
pKGRight.Range.InsertBefore "- "
End If
If pKGRight.Range.Characters(1).text = "-" And pKGRight.Range.Characters.Count > 2 Then
If pKGRight.Range.Characters(2).text <> " " Then
pKGRight.Range.Characters(1).InsertAfter " "
End If
End If
ElseIf lastLineKG = 1 Then
Dim fChr1 As String
fChr1 = pKGRight.Range.Characters(1).text
If fChr1 = "-" Or fChr1 = ChrW(8211) Or fChr1 = ChrW(8212) Or fChr1 = ChrW(8226) Or fChr1 = "+" Or fChr1 = "*" Then
pKGRight.Range.Characters(1).Delete
Call SafeTrimLeadingSpaces(doc, pKGRight)
End If
End If
With pKGRight.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pKGRight.Range.Font.Name = "Times New Roman"
pKGRight.Range.Font.Size = 14
pKGRight.Range.Font.Bold = False
pKGRight.Range.Font.Italic = False
pKGRight.Range.Font.ColorIndex = wdBlack
pKGRight.Range.Font.Underline = wdUnderlineNone
Call NormalizeLineEndingPunctuation(doc, pKGRight, (lineKGIdx = lastLineKG))
End If
lineKGIdx = lineKGIdx + 1
Next pKGRight
End If
End If
NextCellKG:
Next cKG
End If
Next tbl
Call TT_DebugLog("BUOC 3.5 bat dau")
For Each tbl In doc.Tables
Dim tAscFooter As String
tAscFooter = ToAscii(tbl.Range.text)
If IsFooterSignatureTable(tbl, tAscFooter) Then
On Error Resume Next
tbl.Borders.Enable = False
On Error Resume Next
Dim fCellIdx As Integer: fCellIdx = 1
Dim fTotalCells As Integer: fTotalCells = tbl.Range.Cells.Count
For Each c In tbl.Range.Cells
On Error Resume Next
c.Borders.Enable = False
c.Shading.BackgroundPatternColor = wdColorAutomatic
c.Shading.Texture = wdTextureNone
Call SafeReplaceSoftBreaksInRange(c.Range)
On Error Resume Next
Dim cAscFooter As String
cAscFooter = Trim(ToAscii(c.Range.text))
If Len(cAscFooter) = 0 Then GoTo NextFooterCell
Dim isSignatureCell As Boolean
isSignatureCell = False
If InStr(1, cAscFooter, "NOI NHAN") > 0 Or Left(CleanParaText(c.Range.text), 1) = "-" Then
isSignatureCell = False
ElseIf HasSignatureKeywordInCell(cAscFooter) Then
isSignatureCell = True
ElseIf fCellIdx = fTotalCells And fTotalCells >= 2 Then
isSignatureCell = True
Else
On Error Resume Next
If c.ColumnIndex > 1 And tbl.Columns.Count >= 2 And c.ColumnIndex = tbl.Columns.Count Then
isSignatureCell = True
End If
On Error Resume Next
End If
If Not isSignatureCell Then
Dim lastListItemIdx As Integer: lastListItemIdx = 0
Dim tempP As Paragraph
Dim tCounter As Integer: tCounter = 1
For Each tempP In c.Range.Paragraphs
Dim tVal As String
tVal = CleanParaText(tempP.Range.text)
If Len(tVal) > 0 Then
If Not (LTrim(ToAscii(tVal)) Like "NOI NHAN*") Then
lastListItemIdx = tCounter
End If
End If
tCounter = tCounter + 1
Next tempP
Dim pNoiNhan As Paragraph
Dim pIndexNN As Integer: pIndexNN = 1
For Each pNoiNhan In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pNoiNhan)
Dim pTextNN As String
pTextNN = CleanParaText(pNoiNhan.Range.text)
If Len(pTextNN) = 0 Then GoTo SkipNNLine
On Error Resume Next
pNoiNhan.Borders.Enable = False
pNoiNhan.Shading.BackgroundPatternColor = wdColorAutomatic
pNoiNhan.Shading.Texture = wdTextureNone
On Error Resume Next
With pNoiNhan.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pNoiNhan.Range.Font.Name = "Times New Roman"
pNoiNhan.Range.Font.ColorIndex = wdBlack
Dim pTextAsc As String
pTextAsc = ToAscii(pTextNN)
If LTrim(pTextAsc) Like "NOI NHAN*" Then
pNoiNhan.Range.Font.Size = TT_N("NN_LABEL_SIZE")
If TT_B("NN_LABEL_UNDERLINE") Then
pNoiNhan.Range.Font.Underline = wdUnderlineSingle
Else
pNoiNhan.Range.Font.Underline = wdUnderlineNone
End If
pNoiNhan.Range.Font.Bold = TT_B("NN_LABEL_BOLD")
pNoiNhan.Range.Font.Italic = TT_B("NN_LABEL_ITALIC")
Else
pNoiNhan.Range.Font.Size = TT_N("NN_ITEM_SIZE")
pNoiNhan.Range.Font.Underline = wdUnderlineNone
pNoiNhan.Range.Font.Bold = False
pNoiNhan.Range.Font.Italic = False
pNoiNhan.Format.LeftIndent = 0
pNoiNhan.Format.FirstLineIndent = 0
Dim firstChar As String
firstChar = Left(pTextNN, 1)
If firstChar <> "-" Then
If firstChar = ChrW(8211) Or firstChar = ChrW(8212) Or firstChar = "*" Or firstChar = "+" Or firstChar = ChrW(8226) Then
pNoiNhan.Range.Characters(1).text = "-"
Else
pNoiNhan.Range.InsertBefore "- "
End If
End If
If pNoiNhan.Range.Characters(1).text = "-" And pNoiNhan.Range.Characters.Count > 2 Then
If pNoiNhan.Range.Characters(2).text <> " " Then
pNoiNhan.Range.Characters(1).InsertAfter " "
End If
End If
Call NormalizeLineEndingPunctuation(doc, pNoiNhan, (pIndexNN = lastListItemIdx))
End If
SkipNNLine:
pIndexNN = pIndexNN + 1
Next pNoiNhan
Else
On Error Resume Next
Dim textIndices(1 To 30) As Integer
Dim textCount As Integer: textCount = 0
Dim pScan As Paragraph
Dim scanIdx As Integer: scanIdx = 1
For Each pScan In c.Range.Paragraphs
Call SafeTrimLeadingSpaces(doc, pScan)
If Len(CleanParaText(pScan.Range.text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
Dim lastLineHasSignerName As Boolean
lastLineHasSignerName = False
If textCount >= 2 Then
Dim lastSigLineAsc As String
lastSigLineAsc = Trim(ToAscii(CleanParaText(c.Range.Paragraphs(textIndices(textCount)).Range.text)))
If Not IsSignatureRoleOrAuth(lastSigLineAsc) Then
lastLineHasSignerName = True
End If
End If
If textCount >= 2 And lastLineHasSignerName Then
If textCount >= 3 Then
Dim delIdx As Integer
For delIdx = textIndices(textCount - 1) - 1 To textIndices(1) + 1 Step -1
If Len(CleanParaText(c.Range.Paragraphs(delIdx).Range.text)) = 0 Then
c.Range.Paragraphs(delIdx).Range.Delete
End If
Next delIdx
textCount = 0
scanIdx = 1
For Each pScan In c.Range.Paragraphs
If Len(CleanParaText(pScan.Range.text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
End If
Dim sigIdx As Integer: sigIdx = textIndices(textCount)
Dim prevTitleIdx As Integer: prevTitleIdx = textIndices(textCount - 1)
Dim curEmpty As Integer: curEmpty = sigIdx - prevTitleIdx - 1
If curEmpty < 4 Then
Dim nAdd As Integer
For nAdd = 1 To (4 - curEmpty)
Dim rIns As Range
Set rIns = doc.Range(c.Range.Paragraphs(sigIdx).Range.Start, c.Range.Paragraphs(sigIdx).Range.Start)
rIns.text = vbCr
Next nAdd
ElseIf curEmpty > 4 Then
Dim nSub As Integer
For nSub = 1 To (curEmpty - 4)
c.Range.Paragraphs(prevTitleIdx + 1).Range.Delete
Next nSub
End If
ElseIf textCount >= 1 And Not lastLineHasSignerName Then
If textCount >= 2 Then
Dim delMid As Integer
For delMid = textIndices(textCount) - 1 To textIndices(1) + 1 Step -1
If Len(CleanParaText(c.Range.Paragraphs(delMid).Range.text)) = 0 Then
c.Range.Paragraphs(delMid).Range.Delete
End If
Next delMid
textCount = 0
scanIdx = 1
For Each pScan In c.Range.Paragraphs
If Len(CleanParaText(pScan.Range.text)) > 0 Then
If textCount < 30 Then
textCount = textCount + 1
textIndices(textCount) = scanIdx
End If
End If
scanIdx = scanIdx + 1
Next pScan
End If
Dim trailingEmpty As Integer
trailingEmpty = c.Range.Paragraphs.Count - textIndices(textCount)
If trailingEmpty < 4 Then
Dim nTrail As Integer
For nTrail = 1 To (4 - trailingEmpty)
Call SafeAppendBlankParagraphInCell(doc, c)
Next nTrail
End If
End If
On Error Resume Next
Dim pSig As Paragraph
Dim nonEmptyOrder As Integer: nonEmptyOrder = 0
Dim sawAuthLine As Boolean: sawAuthLine = False
For Each pSig In c.Range.Paragraphs
On Error Resume Next
pSig.Borders.Enable = False
pSig.Shading.BackgroundPatternColor = wdColorAutomatic
pSig.Shading.Texture = wdTextureNone
On Error Resume Next
With pSig.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
pSig.Range.Font.Name = "Times New Roman"
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
pSig.Range.Font.ColorIndex = wdBlack
pSig.Range.Font.Underline = wdUnderlineNone
Dim curText As String
curText = CleanParaText(pSig.Range.text)
If Len(curText) > 0 Then
nonEmptyOrder = nonEmptyOrder + 1
Dim uText As String
uText = Trim(ToAscii(curText))
If uText Like "T/M *" Or uText Like "TM.*" Or uText Like "TM *" Or _
uText Like "K/T *" Or uText Like "KT.*" Or uText Like "KT *" Or _
uText Like "T/L *" Or uText Like "TL.*" Or uText Like "TL *" Or _
uText Like "TUQ.*" Or uText Like "T/UQ *" Or uText Like "Q. *" Or uText Like "Q.[A-Z]*" Then
sawAuthLine = True
Dim fixAuth As String
fixAuth = ChuanHoa_ThamQuyen(curText)
If fixAuth <> curText And Len(fixAuth) > 0 Then
Call SafeSetParagraphText(doc, pSig, fixAuth)
End If
pSig.Format.Alignment = wdAlignParagraphCenter
pSig.Range.Font.Name = "Times New Roman"
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
ElseIf Not IsSignatureRoleOrAuth(uText) And nonEmptyOrder = textCount Then
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
ElseIf sawAuthLine Then
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = TT_B("ROLE_UNDER_AUTH_BOLD")
pSig.Range.Font.Italic = False
Else
pSig.Range.Case = wdUpperCase
pSig.Range.Font.Bold = True
pSig.Range.Font.Italic = False
End If
Else
pSig.Range.Font.Bold = False
pSig.Range.Font.Size = TT_N("SIGN_SIZE")
End If
Next pSig
End If
NextFooterCell:
fCellIdx = fCellIdx + 1
Next c
End If
Next tbl
Call TT_DebugLog("BUOC 4 bat dau")
Dim p As Paragraph
Dim s As String
Dim titleState As Integer: titleState = 0
Dim nonTableTextParaCount As Integer: nonTableTextParaCount = 0
Dim inSalutationState As Boolean: inSalutationState = False
Dim outsideSawSo As Boolean: outsideSawSo = False
Dim bodySawDecisionVerb As Boolean: bodySawDecisionVerb = False
Dim afterTitleSep As Boolean: afterTitleSep = False
Dim sauChuong As Boolean: sauChuong = False
Dim inAppendix As Boolean: inAppendix = False
Dim giuCanLe As Long
Dim sigEnd As Long: sigEnd = TT_FindSignatureEnd(doc)
Dim sauChuKy As Boolean: sauChuKy = False
Dim dangCanGiua As Boolean
Dim tenLoaiTran As Boolean
Dim nnNgoaiBang As Boolean: nnNgoaiBang = False
Dim laHeading As Boolean
Dim pTenVB As Paragraph, pKeTieuDe As Paragraph
Dim keLien As Boolean, dsKe As Collection
keLien = (TT_S("SUMMARY_RULE") = "LINE")
Set dsKe = New Collection
Dim gianTruoc As Single, gianSau As Single
gianTruoc = TT_GianTruoc()
gianSau = TT_GianSau()
Dim gianDong As Single, khoiDang As Boolean
gianDong = TT_Opt("LINE_SPACING")
khoiDang = (TT_TemplateGroupOf(TT_Profile()) = "DANG")
For Each p In doc.Paragraphs
If p.Range.Information(wdWithInTable) Then
If titleState >= 1 And titleState <= 5 Then titleState = 6
GoTo NextParagraphMain
End If
Call SafeTrimLeadingSpaces(doc, p)
s = CleanParaText(p.Range.text)
On Error Resume Next
p.Borders.Enable = False
p.Borders(wdBorderTop).LineStyle = wdLineStyleNone
p.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
p.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
p.Borders(wdBorderRight).LineStyle = wdLineStyleNone
p.Shading.BackgroundPatternColor = wdColorAutomatic
p.Shading.Texture = wdTextureNone
On Error Resume Next
If Len(s) = 0 Then
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Format.SpaceBefore = 0
p.Format.SpaceAfter = 0
p.Format.LineSpacingRule = wdLineSpaceSingle
GoTo NextParagraphMain
End If
If InStr(1, p.Range.text, vbTab) > 0 Or InStr(1, p.Range.text, "  ") > 0 Then
Call SafeCollapseSpacesInParagraph(doc, p)
Call SafeTrimLeadingSpaces(doc, p)
s = CleanParaText(p.Range.text)
End If
nonTableTextParaCount = nonTableTextParaCount + 1
p.Range.Font.Name = "Times New Roman"
p.Range.Font.ColorIndex = wdBlack
Dim sAscPara As String
sAscPara = Trim(ToAscii(s))
laHeading = False
laHeading = (p.OutlineLevel >= 1 And p.OutlineLevel <= 9)
If laHeading And titleState >= 2 And titleState <= 5 Then titleState = 6
If Not sauChuKy And sigEnd > 0 Then
If p.Range.Start >= sigEnd Then
sauChuKy = True
inAppendix = True
End If
End If
dangCanGiua = False
If Left(Replace(sAscPara, " ", ""), 6) = "PHULUC" Or sauChuKy Then dangCanGiua = (p.Format.Alignment = wdAlignParagraphCenter)
If (titleState <> 0 Or nonTableTextParaCount > 4) And TT_IsAppendixLabel(s, sAscPara, sauChuKy, dangCanGiua) Then
inAppendix = True
titleState = 1
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
GoTo NextParagraphMain
End If
If inAppendix And titleState >= 1 And titleState <= 5 Then
If sAscPara Like "(KEM THEO*" Or sAscPara Like "(BAN HANH KEM THEO*" Or sAscPara Like "KEM THEO*" Or sAscPara Like "BAN HANH KEM THEO*" Then
titleState = titleState + 1
Call FormatCenteredSubNote(p)
GoTo NextParagraphMain
ElseIf titleState <= 3 And TT_IsAppendixTitleLine(p, s, sAscPara) Then
titleState = titleState + 1
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
Else
titleState = 6
End If
End If
If IsDocNumberLineAsc(sAscPara) And nonTableTextParaCount <= 4 And titleState = 0 Then
outsideSawSo = True
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("DOC_NUMBER_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = False
GoTo NextParagraphMain
ElseIf (outsideSawSo Or sAscPara Like "V/V*" Or sAscPara Like "V.V*") And nonTableTextParaCount <= 5 And Not IsKinhGuiStart(s, sAscPara) And titleState = 0 Then
outsideSawSo = False
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 2
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("CV_SUMMARY_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = TT_B("CV_SUMMARY_ITALIC")
p.Range.Font.ColorIndex = wdBlack
p.Range.Font.Underline = wdUnderlineNone
titleState = 6
GoTo NextParagraphMain
End If
If titleState = 0 Then
If nonTableTextParaCount <= 4 And (sAscPara = "DU THAO" Or sAscPara = "(DU THAO)") Then
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
End With
p.Range.Font.Size = 14
p.Range.Font.Bold = True
p.Range.Font.Italic = True
GoTo NextParagraphMain
ElseIf nonTableTextParaCount <= 4 And IsPartyDocTitle(sAscPara) Then
titleState = 1
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
Set pTenVB = p
p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("TITLE_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
GoTo NextParagraphMain
ElseIf nonTableTextParaCount <= 4 And TT_IsLooseDocTitle(p, s, sAscPara, tenLoaiTran) Then
titleState = 2
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
Set pTenVB = p
p.Range.Font.Size = TT_N("TITLE_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
GoTo NextParagraphMain
Else
titleState = 6
End If
ElseIf titleState >= 1 And titleState <= 5 Then
If IsSeparatorLine(s) Then
titleState = 6
afterTitleSep = True
Call FormatCenteredSeparator(p)
p.Format.SpaceAfter = 0
If Not pTenVB Is Nothing And pKeTieuDe Is Nothing Then Set pKeTieuDe = p
If keLien Then dsKe.Add p
GoTo NextParagraphMain
ElseIf IsKinhGuiStart(s, sAscPara) Then
titleState = 6
ElseIf sAscPara Like "(KEM THEO*" Or sAscPara Like "(BAN HANH KEM THEO*" Or sAscPara Like "KEM THEO*" Or sAscPara Like "BAN HANH KEM THEO*" Then
titleState = titleState + 1
Call FormatCenteredSubNote(p)
GoTo NextParagraphMain
ElseIf HasSeparatorWithinNextParas(p, 4) And Not IsStrictBodyOrHeadingStart(s, sAscPara) Then
titleState = titleState + 1
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
ElseIf titleState = 1 And Not IsStrictBodyOrHeadingStart(s, sAscPara) And Len(s) <= 350 Then
titleState = 2
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
ElseIf titleState = 2 And Not IsBodyStart(s, sAscPara) And Len(s) <= 130 And Right(s, 1) <> "." And Right(s, 1) <> ":" Then
titleState = 3
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
Else
titleState = 6
End If
End If
If afterTitleSep Then
afterTitleSep = False
If UCase(s) = s And Len(s) <= 150 And sAscPara Like "*[A-Z][A-Z]*" And _
Not IsStrictBodyOrHeadingStart(s, sAscPara) And Not IsKinhGuiStart(s, sAscPara) Then
Call FormatCenteredSummary(p)
GoTo NextParagraphMain
End If
End If
If sauChuong Then
sauChuong = False
If Len(s) <= 200 And Not IsStrictBodyOrHeadingStart(s, sAscPara) And Not (sAscPara Like "DIEU #*") And _
Not (sAscPara Like "MUC #*") And Not (sAscPara Like "CHUONG *") And Not (sAscPara Like "PHAN *") And _
Not (sAscPara Like "[A-Za-z])*") And Not (sAscPara Like "[A-Za-z]. *") And Not (sAscPara Like "#*") And _
InStr(1, ".;:", Right(s, 1)) = 0 Then
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Case = wdUpperCase
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
GoTo NextParagraphMain
End If
End If
If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Or _
sAscPara Like "PHAN I*" Or sAscPara Like "PHAN II*" Or sAscPara Like "PHAN III*" Or sAscPara Like "PHAN IV*" Or sAscPara Like "PHAN THU *" Or _
sAscPara Like "CHUONG I*" Or sAscPara Like "CHUONG II*" Or sAscPara Like "CHUONG III*" Or sAscPara Like "CHUONG IV*" Then
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
If sAscPara = "QUYET DINH:" Or sAscPara = "QUYET DINH :" Then
p.Range.Case = wdUpperCase
bodySawDecisionVerb = True
Else
sauChuong = TT_ChuThuongPhanChuong(p, s, sAscPara)
End If
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
GoTo NextParagraphMain
ElseIf IsSeparatorLine(s) Then
Call FormatCenteredSeparator(p)
GoTo NextParagraphMain
End If
If IsKinhGuiStart(s, sAscPara) Then
Call SafeTrimLeadingSpaces(doc, p)
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Underline = wdUnderlineNone
p.Range.Font.ColorIndex = wdBlack
Dim cPosColon As Long: cPosColon = InStr(1, p.Range.text, ":")
If Not TT_B("KINH_GUI_ITALIC") Then
p.Range.Font.Italic = False
ElseIf cPosColon > 0 Then
p.Range.Font.Italic = False
doc.Range(p.Range.Start, p.Range.Start + cPosColon).Font.Italic = True
Else
p.Range.Font.Italic = True
End If
Dim hasNextRecip As Boolean: hasNextRecip = False
Dim pNextCheck As Paragraph: Set pNextCheck = p.Next
Do While Not pNextCheck Is Nothing
If pNextCheck.Range.Information(wdWithInTable) Then Exit Do
Dim sNextC As String: sNextC = CleanParaText(pNextCheck.Range.text)
If Len(sNextC) > 0 Then
Dim sAscNextC As String: sAscNextC = Trim(ToAscii(sNextC))
If IsKinhGuiRecipientItem(sNextC, sAscNextC) Then
hasNextRecip = True
End If
Exit Do
End If
Set pNextCheck = pNextCheck.Next
Loop
If Not hasNextRecip Then
p.Format.SpaceAfter = 6
inSalutationState = False
Call NormalizeLineEndingPunctuation(doc, p, True)
Else
inSalutationState = True
If cPosColon > 0 And cPosColon < Len(CleanParaText(p.Range.text)) Then
Call NormalizeLineEndingPunctuation(doc, p, False)
End If
End If
GoTo NextParagraphMain
ElseIf inSalutationState And IsKinhGuiRecipientItem(s, sAscPara) Then
Call SafeTrimLeadingSpaces(doc, p)
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 6
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
p.Range.Font.ColorIndex = wdBlack
Dim nextIsRecipient As Boolean: nextIsRecipient = False
Dim pCheckNext As Paragraph: Set pCheckNext = p.Next
Do While Not pCheckNext Is Nothing
If pCheckNext.Range.Information(wdWithInTable) Then Exit Do
Dim sNextChk As String: sNextChk = CleanParaText(pCheckNext.Range.text)
If Len(sNextChk) > 0 Then
Dim sAscNextChk As String: sAscNextChk = Trim(ToAscii(sNextChk))
If IsKinhGuiRecipientItem(sNextChk, sAscNextChk) Then
nextIsRecipient = True
End If
Exit Do
End If
Set pCheckNext = pCheckNext.Next
Loop
If Not nextIsRecipient Then
p.Format.SpaceAfter = 6
inSalutationState = False
Call NormalizeLineEndingPunctuation(doc, p, True)
Else
Call NormalizeLineEndingPunctuation(doc, p, False)
End If
GoTo NextParagraphMain
Else
inSalutationState = False
End If
If nnNgoaiBang Then
If Left(s, 1) = "-" Or Left(s, 1) = ChrW(8211) Or Left(s, 1) = "+" Then
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Size = TT_N("NN_ITEM_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
GoTo NextParagraphMain
Else
nnNgoaiBang = False
End If
End If
If sAscPara Like "NOI NHAN:*" Or sAscPara = "NOI NHAN" Or sAscPara = "NOI NHAN :" Then
nnNgoaiBang = True
With p.Format
.Alignment = wdAlignParagraphLeft
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("NN_LABEL_SIZE")
p.Range.Font.Bold = TT_B("NN_LABEL_BOLD")
p.Range.Font.Italic = TT_B("NN_LABEL_ITALIC")
If TT_B("NN_LABEL_UNDERLINE") Then
p.Range.Font.Underline = wdUnderlineSingle
Else
p.Range.Font.Underline = wdUnderlineNone
End If
GoTo NextParagraphMain
End If
If Len(s) <= 60 And Not (sAscPara Like "*[.;:]") Then
If sAscPara Like "* NGAY*THANG*NAM ####" And InStr(1, s, ",") > 0 And InStr(1, s, ",") <= 32 And InStr(1, sAscPara, " NGAY") <= 34 Then
With p.Format
.Alignment = wdAlignParagraphRight
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 6
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.Bold = False
p.Range.Font.Italic = True
GoTo NextParagraphMain
End If
End If
giuCanLe = -1
If inAppendix Then
If p.Format.Alignment = wdAlignParagraphCenter Then giuCanLe = wdAlignParagraphCenter
If p.Format.Alignment = wdAlignParagraphRight Then giuCanLe = wdAlignParagraphRight
End If
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("BODY_SIZE")
p.Range.Font.ColorIndex = wdBlack
With p.Format
.Alignment = wdAlignParagraphJustify
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = CentimetersToPoints(TT_N("INDENT_CM"))
.SpaceBefore = gianTruoc
.SpaceAfter = gianSau
.SpaceBeforeAuto = False
.SpaceAfterAuto = False
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
If gianDong > 0 Then Call TT_ApGianDong(p, gianDong, khoiDang)
If inAppendix And giuCanLe >= 0 Then
p.Format.Alignment = giuCanLe
p.Format.FirstLineIndent = 0
End If
If TT_B("LEGAL_BASIS_ITALIC") And Not bodySawDecisionVerb And (sAscPara Like "CAN CU *" Or sAscPara Like "XET DE NGHI*" Or sAscPara Like "THEO DE NGHI*") Then
p.Range.Font.Italic = True
p.Range.Font.Bold = False
GoTo NextParagraphMain
End If
If sAscPara Like "I.*" Or sAscPara Like "II.*" Or sAscPara Like "III.*" Or sAscPara Like "IV.*" Or sAscPara Like "V.*" Or sAscPara Like "VI.*" Or sAscPara Like "VII.*" Or sAscPara Like "VIII.*" Or sAscPara Like "IX.*" Or sAscPara Like "X.*" Or _
sAscPara Like "I-*" Or sAscPara Like "II-*" Or sAscPara Like "III-*" Or sAscPara Like "IV-*" Or sAscPara Like "V-*" Then
p.Range.Case = wdUpperCase
p.Range.Font.Bold = True
ElseIf (s Like "#. *" Or s Like "##. *" Or s Like "#.#. *" Or s Like "#.##. *" Or s Like "##.#. *" Or s Like "#.#.#*" Or Left(s, 1) = "*") And ((Len(s) < 350 And Right(s, 1) <> "." And Right(s, 1) <> ";") Or (Len(s) < 120 And Right(s, 1) <> ";")) Then
p.Range.Font.Bold = True
ElseIf laHeading And Len(s) < 200 Then
p.Range.Font.Bold = True
ElseIf sAscPara Like "[A-H]. *" And Len(s) < 150 And Right(s, 1) <> "." And Right(s, 1) <> ";" Then
p.Range.Font.Bold = True
ElseIf sAscPara Like "DIEU #.*" Or sAscPara Like "DIEU ##.*" Or sAscPara Like "DIEU #:*" Or sAscPara Like "DIEU ##:*" Then
Dim colonOrDotPos As Long
colonOrDotPos = InStr(1, s, ".")
If colonOrDotPos = 0 Or colonOrDotPos > 10 Then colonOrDotPos = InStr(1, s, ":")
If colonOrDotPos > 0 And colonOrDotPos <= 10 Then
Dim rDieu As Range
Set rDieu = doc.Range(p.Range.Start, p.Range.Start + colonOrDotPos)
rDieu.Font.Bold = True
End If
ElseIf p.Range.Font.Bold = True And Not inAppendix Then
p.Range.Font.Bold = False
End If
NextParagraphMain:
Next p
Dim dsNeo As Collection, pKeCu As Paragraph, pNeoKe As Paragraph, iKe As Long, laKeTieuDe As Boolean
Set dsNeo = New Collection
If keLien Then
For iKe = 1 To dsKe.Count
Set pKeCu = dsKe(iKe)
laKeTieuDe = False
If Not pKeTieuDe Is Nothing Then laKeTieuDe = (pKeTieuDe.Range.Start = pKeCu.Range.Start)
Set pNeoKe = TT_NeoKeTrichYeu(pKeCu)
If Not pNeoKe Is Nothing Then
dsNeo.Add pNeoKe
If laKeTieuDe Then Set pKeTieuDe = pNeoKe
End If
Next iKe
End If
If Not pTenVB Is Nothing Then Call TT_DongTrongQuanhTieuDe(pTenVB, pKeTieuDe)
For iKe = 1 To dsNeo.Count
Call TT_VeKeDuoiTrichYeu(doc, dsNeo(iKe), iKe)
Next iKe
Call TT_DebugLog("BUOC 5 bat dau")
Dim tblAll As Table
For Each tblAll In doc.Tables
On Error Resume Next
If tblAll.Rows.LeftIndent >= -MillimetersToPoints(3) Then
tblAll.AutoFitBehavior 2
tblAll.AllowAutoFit = True
End If
tblAll.Range.Font.Name = "Times New Roman"
tblAll.Range.Font.ColorIndex = wdBlack
On Error Resume Next
Next tblAll
If TT_GomChuLe() Then
Call TT_DebugLog("BUOC 5b gom chu le bat dau")
Call TT_GomChuLeCuoiDoan(doc)
End If
Call TT_DebugLog("BUOC 6 bat dau")
Dim strWarnings As String
strWarnings = KiemTra_LogicTheThuc(doc)
On Error Resume Next
If cuChinhTa Then Application.Options.CheckSpellingAsYouType = True
If cuNguPhap Then Application.Options.CheckGrammarAsYouType = True
Application.ScreenUpdating = True
On Error GoTo 0
If Not (docPhanSau Is Nothing) Then
Application.ScreenUpdating = False
Call TT_GhepVanBanTiepTheo(doc, docPhanSau)
Application.ScreenUpdating = True
End If
Call TT_DebugLog("KET THUC")
If silentMode Then
On Error Resume Next
doc.Variables("TheThuc_BaoCao").Delete
doc.Variables.Add "TheThuc_BaoCao", TT_ReportHeader() & vbCrLf & "Canh bao:" & strWarnings
On Error GoTo 0
End If
If Not silentMode Then
If Len(strWarnings) > 0 Then
MsgBox "DA CHUAN HOA THE THUC VAN BAN!" & vbCrLf & TT_ReportHeader() & vbCrLf & vbCrLf & _
"Phat hien cac diem bat hop ly ve LOGIC can kiem tra truoc khi trinh ky:" & vbCrLf & _
strWarnings & vbCrLf & vbCrLf & _
"Created by VPDU xa Thanh Phu", vbExclamation, "Bao Cao Ra Soat Logic & The Thuc"
Else
MsgBox "Da chuan hoa the thuc van ban!" & vbCrLf & TT_ReportHeader() & vbCrLf & vbCrLf & _
"Van ban dat chuan 100%, khong phat hien loi logic." & vbCrLf & vbCrLf & _
"Created by VPDU xa Thanh Phu", vbInformation, "Chuan Hoa The Thuc Thanh Cong"
End If
End If
Exit Sub
CleanUpAndExit:
On Error Resume Next
If cuChinhTa Then Application.Options.CheckSpellingAsYouType = True
If cuNguPhap Then Application.Options.CheckGrammarAsYouType = True
Application.ScreenUpdating = True
On Error GoTo 0
End Sub
Private Sub SafeReplaceSoftBreaksInRange(ByVal rng As Range)
On Error Resume Next
If InStr(1, rng.text, Chr(11)) = 0 Then Exit Sub
Dim p As Paragraph, r As Range, guard As Long
Set p = Nothing
Set p = rng.Paragraphs(1)
Do While Not (p Is Nothing)
guard = guard + 1
If guard > 20000 Then Exit Do
If p.Range.Start >= rng.End Then Exit Do
Set r = p.Range
If InStr(1, r.text, Chr(11)) > 0 Then
Call SafeReplaceSoftBreaksInOnePara(r)
Set p = r.Paragraphs(r.Paragraphs.Count).Next
Else
Set p = p.Next
End If
Loop
On Error GoTo 0
End Sub
Private Sub SafeReplaceSoftBreaksInOnePara(ByVal r As Range)
On Error Resume Next
Dim t As String, i As Long, ch As Range, c As String, lech As Boolean
t = r.text
For i = Len(t) To 1 Step -1
If Mid(t, i, 1) = Chr(11) Then
Set ch = Nothing
Set ch = r.Document.Range(r.Start + i - 1, r.Start + i)
c = ""
If Not (ch Is Nothing) Then c = ch.text
If c = Chr(11) Then
ch.text = vbCr
Else
lech = True
End If
End If
Next i
If lech Then
For i = r.Characters.Count To 1 Step -1
Set ch = Nothing
Set ch = r.Characters(i)
c = ""
If Not (ch Is Nothing) Then c = ch.text
If c = Chr(11) Then ch.text = vbCr
Next i
End If
On Error GoTo 0
End Sub
Private Sub SafeCollapseSpacesInParagraph(ByVal doc As Document, ByVal p As Paragraph)
On Error Resume Next
Dim rawTxt As String
rawTxt = p.Range.text
If InStr(1, rawTxt, "  ") = 0 And InStr(1, rawTxt, vbTab) = 0 Then Exit Sub
Dim ascTxt As String
ascTxt = Trim(ToAscii(rawTxt))
If Left(ascTxt, 2) = "SO" And (InStr(1, ascTxt, "-") > 0 Or InStr(1, ascTxt, "/") > 0) Then Exit Sub
If InStr(1, ascTxt, "NGAY") > 0 And InStr(1, ascTxt, "THANG") > 0 And InStr(1, ascTxt, "NAM") > 0 Then Exit Sub
Dim k As Long
Dim nChars As Long
Dim rgDoan As Range, viTriDau As Long, kyTu As String, kyTuTruoc As String, rgKyTu As Range, lechViTri As Boolean
Set rgDoan = p.Range
viTriDau = rgDoan.Start
If rgDoan.End - rgDoan.Start = Len(rawTxt) Then
For k = Len(rawTxt) - 1 To 1 Step -1
kyTu = Mid$(rawTxt, k, 1)
If kyTu = vbTab Or kyTu = " " Then
kyTuTruoc = ""
If k > 1 Then kyTuTruoc = Mid$(rawTxt, k - 1, 1)
Set rgKyTu = Nothing
Set rgKyTu = doc.Range(viTriDau + k - 1, viTriDau + k)
If rgKyTu Is Nothing Then
lechViTri = True
Exit For
End If
If rgKyTu.text <> kyTu Then
lechViTri = True
Exit For
End If
If kyTuTruoc = " " Or kyTuTruoc = vbTab Then
rgKyTu.Delete
ElseIf kyTu = vbTab Then
rgKyTu.text = " "
End If
End If
Next k
If Not lechViTri Then
On Error GoTo 0
Exit Sub
End If
End If
nChars = p.Range.Characters.Count
For k = nChars - 1 To 1 Step -1
Dim chCur As String
chCur = p.Range.Characters(k).text
If chCur = vbTab Then
p.Range.Characters(k).text = " "
chCur = " "
End If
If chCur = " " And k > 1 Then
Dim chPrev As String
chPrev = p.Range.Characters(k - 1).text
If chPrev = " " Or chPrev = vbTab Then
p.Range.Characters(k).Delete
End If
End If
Next k
On Error GoTo 0
End Sub
Private Function CleanParaText(ByVal rawText As String) As String
Dim s As String
s = Replace(Replace(rawText, vbCr, ""), vbLf, "")
s = Replace(s, Chr(7), "")
s = Replace(s, Chr(11), "")
s = Replace(s, Chr(12), "")
s = Replace(s, ChrW(8203), "")
s = Replace(s, ChrW(65279), "")
s = Replace(s, Chr(160), " ")
s = Replace(s, vbTab, " ")
CleanParaText = Trim(s)
End Function
Private Sub RemoveEmptyParagraphsInCell(ByVal doc As Document, ByVal c As Cell)
On Error Resume Next
Dim idxP As Long
For idxP = c.Range.Paragraphs.Count To 1 Step -1
If c.Range.Paragraphs.Count <= 1 Then Exit For
Dim pCheck As Paragraph
Set pCheck = c.Range.Paragraphs(idxP)
If Len(CleanParaText(pCheck.Range.text)) = 0 Then
If idxP < c.Range.Paragraphs.Count Then
pCheck.Range.Delete
ElseIf idxP > 1 Then
Dim pPrev As Paragraph
Set pPrev = c.Range.Paragraphs(idxP - 1)
Dim rDelCr As Range
Set rDelCr = doc.Range(pPrev.Range.End - 1, pPrev.Range.End)
If rDelCr.text = vbCr Then rDelCr.Delete
End If
End If
Next idxP
On Error GoTo 0
End Sub
Private Sub SafeTrimLeadingSpaces(ByVal doc As Document, ByVal p As Paragraph)
On Error Resume Next
Dim rawTxt As String
rawTxt = p.Range.text
If Len(rawTxt) <= 1 Then Exit Sub
Dim firstCode As Long
firstCode = AscW(Left$(rawTxt, 1))
If firstCode < 0 Then firstCode = firstCode + 65536
If firstCode <> 32 And firstCode <> 9 And firstCode <> 160 Then Exit Sub
Dim guard As Integer: guard = 0
Dim prevCount As Long
Do While p.Range.Characters.Count > 1 And guard < 40
guard = guard + 1
Dim chTxt As String
chTxt = p.Range.Characters(1).text
If chTxt = " " Or chTxt = vbTab Or chTxt = Chr(160) Then
prevCount = p.Range.Characters.Count
p.Range.Characters(1).Delete
If p.Range.Characters.Count >= prevCount Then Exit Do
Else
Exit Do
End If
Loop
On Error GoTo 0
End Sub
Private Sub RemoveHeaderFloatingShapes(ByVal doc As Document, ByVal hdrTbl As Table)
On Error Resume Next
Dim sIdx As Long
For sIdx = doc.Shapes.Count To 1 Step -1
Dim sItem As Shape
Set sItem = Nothing
Set sItem = doc.Shapes(sIdx)
If Not sItem Is Nothing Then
Dim shouldDel As Boolean
shouldDel = False
If sItem.Name Like "LineTieuNguDang*" Then shouldDel = True
If sItem.Name Like "LineTieuNguQuocHieu*" Then shouldDel = True
If sItem.Name Like "LineCoQuanBanHanh*" Then shouldDel = True
If Not hdrTbl Is Nothing Then
If sItem.Anchor.InRange(hdrTbl.Range) Then shouldDel = True
End If
If sItem.Height < 8 And sItem.Top < 260 Then shouldDel = True
If (sItem.Type = 9 Or sItem.Type = 1) And sItem.Top < 260 And sItem.Height < 15 Then shouldDel = True
If shouldDel Then sItem.Delete
End If
Next sIdx
If Not hdrTbl Is Nothing Then
Dim iShpIdx As Long
For iShpIdx = hdrTbl.Range.InlineShapes.Count To 1 Step -1
hdrTbl.Range.InlineShapes(iShpIdx).Delete
Next iShpIdx
End If
On Error GoTo 0
End Sub
Private Sub SafeSetParagraphText(ByVal doc As Document, ByVal p As Paragraph, ByVal newText As String)
On Error Resume Next
Dim r As Range
Set r = p.Range
Do While r.End > r.Start
Dim code As Long
code = AscW(Right(r.text, 1))
If code < 0 Then code = code + 65536
If code = 13 Or code = 10 Or code = 7 Then
r.End = r.End - 1
Else
Exit Do
End If
Loop
r.text = newText
On Error GoTo 0
End Sub
Private Sub SafeAppendBlankParagraphInCell(ByVal doc As Document, ByVal c As Cell)
On Error Resume Next
Dim rEnd As Range
Set rEnd = c.Range
rEnd.End = rEnd.End - 1
rEnd.Collapse Direction:=wdCollapseEnd
rEnd.InsertAfter vbCr
On Error GoTo 0
End Sub
Private Sub AutoCreateHeaderTable(ByVal doc As Document)
On Error Resume Next
If doc.Tables.Count >= 1 Then
Dim firstTbl As Table
Set firstTbl = doc.Tables(1)
Dim t1Asc As String
t1Asc = ToAscii(firstTbl.Range.text)
Dim parasBeforeTbl1 As Integer: parasBeforeTbl1 = 0
Dim pCheckBefore As Paragraph
For Each pCheckBefore In doc.Paragraphs
If pCheckBefore.Range.Start >= firstTbl.Range.Start Then Exit For
If Len(CleanParaText(pCheckBefore.Range.text)) > 0 Then
parasBeforeTbl1 = parasBeforeTbl1 + 1
End If
Next pCheckBefore
Dim t1DetectStrong As Boolean
If parasBeforeTbl1 <= 1 And (InStr(1, t1Asc, "DANG CONG SAN") > 0 Or InStr(1, t1Asc, "DANG BO") > 0 Or InStr(1, t1Asc, "DANG UY") > 0 Or InStr(1, t1Asc, "CHI BO") > 0 Or InStr(1, t1Asc, "TINH UY") > 0 Or InStr(1, t1Asc, "THANH UY") > 0 Or _
InStr(1, t1Asc, "CONG HOA XA HOI") > 0 Or (firstTbl.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(t1Asc, t1DetectStrong)) > 0 And InStr(1, t1Asc, "NOI NHAN") = 0)) Then
Dim pAfterH As Paragraph
Set pAfterH = firstTbl.Range.Paragraphs(firstTbl.Range.Paragraphs.Count).Next
Dim scanAfterCount As Integer: scanAfterCount = 0
Do While Not pAfterH Is Nothing And scanAfterCount < 4
If pAfterH.Range.Information(wdWithInTable) Then Exit Do
Dim sAfterH As String
sAfterH = CleanParaText(pAfterH.Range.text)
Dim pNextH As Paragraph
Set pNextH = pAfterH.Next
If Len(sAfterH) > 0 Then
scanAfterCount = scanAfterCount + 1
Dim sAfterHAsc As String
sAfterHAsc = Trim(ToAscii(sAfterH))
If IsPartyDocTitle(sAfterHAsc) Or IsKinhGuiStart(sAfterH, sAfterHAsc) Or IsGluedTitleLine(sAfterH, sAfterHAsc) Then
Exit Do
End If
If InStr(1, sAfterHAsc, "NGAY") > 0 And InStr(1, sAfterHAsc, "THANG") > 0 And InStr(1, sAfterHAsc, "NAM") > 0 And Len(sAfterH) < 95 And InStr(1, t1Asc, "NGAY") = 0 Then
If firstTbl.Columns.Count >= 2 Then
Call AppendTextToCell(doc, firstTbl.Cell(1, firstTbl.Columns.Count), sAfterH)
pAfterH.Range.Delete
End If
ElseIf IsDocNumberLineAsc(sAfterHAsc) And Len(sAfterH) < 60 And InStr(1, t1Asc, "SO ") = 0 And InStr(1, t1Asc, "SO:") = 0 Then
Call AppendTextToCell(doc, firstTbl.Cell(1, 1), sAfterH)
pAfterH.Range.Delete
End If
End If
Set pAfterH = pNextH
Loop
Exit Sub
End If
End If
Dim p As Paragraph
Dim pHStart As Paragraph: Set pHStart = Nothing
Dim pHEnd As Paragraph: Set pHEnd = Nothing
Dim nonEmptyTopCount As Integer: nonEmptyTopCount = 0
Dim sawMottoOrOrgan As Boolean: sawMottoOrOrgan = False
Dim sawMotto As Boolean: sawMotto = False
Dim sawSoOrDateOrStar As Boolean: sawSoOrDateOrStar = False
For Each p In doc.Paragraphs
If p.Range.Information(wdWithInTable) Then Exit For
Dim sRaw As String
sRaw = CleanParaText(p.Range.text)
If Len(sRaw) > 0 Then
nonEmptyTopCount = nonEmptyTopCount + 1
If nonEmptyTopCount > 12 Then Exit For
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If IsPartyDocTitle(sAsc) Or IsGluedTitleLine(sRaw, sAsc) Or IsKinhGuiStart(sRaw, sAsc) Then
Exit For
End If
If IsStrictBodyOrHeadingStart(sRaw, sAsc) Then
Exit For
End If
Dim isHdrLine As Boolean: isHdrLine = False
If InStr(1, sAsc, "DANG CONG SAN") > 0 Or InStr(1, sAsc, "DOC LAP - TU DO") > 0 Or InStr(1, sAsc, "CONG HOA XA HOI") > 0 Then
isHdrLine = True
sawMottoOrOrgan = True
sawMotto = True
ElseIf (IsHeaderOrganLineAsc(sAsc) Or IsNationalOrganLine(sRaw, sAsc)) And Len(sRaw) < 110 Then
isHdrLine = True
sawMottoOrOrgan = True
ElseIf IsDocNumberLineAsc(sAsc) And Len(sRaw) < 95 Then
isHdrLine = True
sawSoOrDateOrStar = True
ElseIf InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 And Len(sRaw) < 110 Then
isHdrLine = True
sawSoOrDateOrStar = True
ElseIf (sAsc Like "V/V*" Or sAsc Like "V.V*" Or sAsc Like "VE VIEC*") And sawMottoOrOrgan Then
isHdrLine = True
ElseIf (sAsc = "*" Or sAsc = "***" Or IsSeparatorLine(sRaw)) And sawMottoOrOrgan Then
isHdrLine = True
sawSoOrDateOrStar = True
End If
If isHdrLine Then
If pHStart Is Nothing Then Set pHStart = p
Set pHEnd = p
ElseIf Not pHStart Is Nothing Then
Exit For
End If
End If
Next p
If pHStart Is Nothing Or pHEnd Is Nothing Then Exit Sub
If Not (sawMotto Or (sawMottoOrOrgan And sawSoOrDateOrStar)) Then Exit Sub
Dim leftLines As String: leftLines = ""
Dim rightLines As String: rightLines = ""
Dim pCurH As Paragraph
Set pCurH = pHStart
Do While Not pCurH Is Nothing
Dim rawLine As String
rawLine = Replace(Replace(Replace(pCurH.Range.text, vbCr, ""), vbLf, ""), Chr(7), "")
rawLine = Replace(rawLine, Chr(11), vbCr)
Dim subLines() As String
subLines = Split(rawLine, vbCr)
Dim slIdx As Long
For slIdx = LBound(subLines) To UBound(subLines)
Call ClassifyAndAppendHeaderLine(subLines(slIdx), leftLines, rightLines)
Next slIdx
If pCurH.Range.Start >= pHEnd.Range.Start Then Exit Do
Set pCurH = pCurH.Next
Loop
If Len(Trim(leftLines)) = 0 And Len(Trim(rightLines)) = 0 Then Exit Sub
Dim rightAsc As String
rightAsc = ToAscii(rightLines)
Dim hasAnyMotto As Boolean
hasAnyMotto = (InStr(1, rightAsc, "DANG CONG SAN") > 0 Or InStr(1, rightAsc, "CONG HOA XA HOI") > 0)
Dim mottoStr As String
mottoStr = ""
If TT_B("AUTO_INSERT_MOTTO") Then
If TT_S("MOTTO_KIND") = "DANG" And Not hasAnyMotto Then
mottoStr = TT_TextTieuNguDang()
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" Then
If Not hasAnyMotto Then
mottoStr = TT_TextQuocHieu() & vbCr & TT_TextTieuNgu()
ElseIf InStr(1, rightAsc, "CONG HOA XA HOI") > 0 And InStr(1, Replace(rightAsc, " ", ""), "DOCLAP-TUDO") = 0 Then
Dim rParts() As String
rParts = Split(rightLines, vbCr)
Dim rIdx As Long
Dim rebuilt As String
rebuilt = ""
For rIdx = LBound(rParts) To UBound(rParts)
Call AppendLineToBuffer(rebuilt, rParts(rIdx))
If InStr(1, ToAscii(rParts(rIdx)), "CONG HOA XA HOI") > 0 Then
Call AppendLineToBuffer(rebuilt, TT_TextTieuNgu())
End If
Next rIdx
rightLines = rebuilt
End If
End If
End If
If Len(mottoStr) > 0 Then
If Len(Trim(rightLines)) > 0 Then
rightLines = mottoStr & vbCr & rightLines
Else
rightLines = mottoStr
End If
End If
Dim rngReplaceH As Range
Set rngReplaceH = doc.Range(pHStart.Range.Start, pHEnd.Range.End)
Dim newHdrTbl As Table
Set newHdrTbl = doc.Tables.Add(Range:=rngReplaceH, NumRows:=1, NumColumns:=2)
With newHdrTbl
.Borders.Enable = False
.Columns(1).Width = MillimetersToPoints(TT_N("HDR_LEFT_MM"))
.Columns(2).Width = MillimetersToPoints(TT_N("HDR_RIGHT_MM"))
.Cell(1, 1).Range.text = leftLines
.Cell(1, 2).Range.text = rightLines
End With
On Error GoTo 0
End Sub
Private Sub ClassifyAndAppendHeaderLine(ByVal rawSubLine As String, ByRef leftLines As String, ByRef rightLines As String)
Dim sLine As String
sLine = Trim(Replace(rawSubLine, Chr(160), " "))
If Len(sLine) = 0 Then Exit Sub
If InStr(1, sLine, vbTab) > 0 Then
Dim tabParts() As String
tabParts = Split(sLine, vbTab)
Dim tIdx As Long
For tIdx = LBound(tabParts) To UBound(tabParts)
If Len(Trim(tabParts(tIdx))) > 0 Then
Call ClassifyAndAppendHeaderLine(Trim(tabParts(tIdx)), leftLines, rightLines)
End If
Next tIdx
Exit Sub
End If
Dim sAsc As String
sAsc = Trim(ToAscii(sLine))
If Len(sAsc) = 0 Then Exit Sub
Dim posDCS As Long
posDCS = InStr(1, sAsc, "DANG CONG SAN")
If posDCS = 0 Then posDCS = InStr(1, sAsc, "CONG HOA XA HOI")
If posDCS = 0 Then posDCS = InStr(1, sAsc, "DOC LAP - TU DO")
If posDCS > 3 Then
Dim leftPart1 As String, rightPart1 As String
leftPart1 = Trim(Left(sLine, posDCS - 1))
rightPart1 = Trim(Mid(sLine, posDCS))
If Len(leftPart1) > 0 Then Call AppendLineToBuffer(leftLines, leftPart1)
If Len(rightPart1) > 0 Then Call AppendLineToBuffer(rightLines, rightPart1)
Exit Sub
End If
If InStr(1, sAsc, "DANG CONG SAN") > 0 Or sAsc = "VIET NAM" Or InStr(1, sAsc, "DOC LAP - TU DO") > 0 Or InStr(1, sAsc, "CONG HOA XA HOI") > 0 Then
Call AppendLineToBuffer(rightLines, sLine)
ElseIf InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 And InStr(1, sAsc, "V/V") = 0 Then
Call AppendLineToBuffer(rightLines, sLine)
ElseIf sAsc = "*" Or sAsc = "***" Then
If InStr(1, leftLines, "*") = 0 And Not TT_B("ORGAN_UNDERLINE") Then
Call AppendLineToBuffer(leftLines, "*")
End If
ElseIf IsSeparatorLine(sLine) Then
If InStr(1, ToAscii(leftLines), "V/V") > 0 Or InStr(1, ToAscii(leftLines), "VE VIEC") > 0 Then
Call AppendLineToBuffer(leftLines, sLine)
End If
Else
Call AppendLineToBuffer(leftLines, sLine)
End If
End Sub
Private Sub AppendLineToBuffer(ByRef buf As String, ByVal newLine As String)
Dim cleanL As String
cleanL = Trim(newLine)
If Len(cleanL) = 0 Then Exit Sub
If Len(buf) > 0 Then
buf = buf & vbCr & cleanL
Else
buf = cleanL
End If
End Sub
Private Sub AppendTextToCell(ByVal doc As Document, ByVal targetCell As Cell, ByVal extraLine As String)
On Error Resume Next
Dim existingTxt As String
existingTxt = CleanParaText(targetCell.Range.text)
If Len(existingTxt) > 0 Then
Call SafeAppendBlankParagraphInCell(doc, targetCell)
Call SafeSetParagraphText(doc, targetCell.Range.Paragraphs(targetCell.Range.Paragraphs.Count), extraLine)
Else
targetCell.Range.text = extraLine
End If
On Error GoTo 0
End Sub
Private Sub SplitGluedHeaderAndTitleLines(ByVal doc As Document)
On Error Resume Next
If doc.Tables.Count >= 1 Then
Dim tblH As Table
Set tblH = doc.Tables(1)
Dim tHAsc As String
tHAsc = ToAscii(tblH.Range.text)
If InStr(1, tHAsc, "DANG CONG SAN") > 0 Or InStr(1, tHAsc, "DANG BO") > 0 Or InStr(1, tHAsc, "DANG UY") > 0 Then
Dim ch As Cell
For Each ch In tblH.Range.Cells
Dim pH As Paragraph
For Each pH In ch.Range.Paragraphs
Dim sH As String: sH = CleanParaText(pH.Range.text)
Dim sHAsc As String: sHAsc = Trim(ToAscii(sH))
If InStr(1, sHAsc, "DANG CONG SAN VIET NAM") > 0 And InStr(1, sHAsc, "NGAY") > 0 And InStr(1, sHAsc, "THANG") > 0 Then
Dim posVN As Long
posVN = InStr(1, sHAsc, "VIET NAM")
If posVN > 0 Then
Dim cutVN As Long: cutVN = posVN + 7
Dim leftMotto As String, rightDate As String
leftMotto = Trim(Left(sH, cutVN))
rightDate = Trim(Mid(sH, cutVN + 1))
Do While Len(rightDate) > 0 And (Left(rightDate, 1) = "*" Or Left(rightDate, 1) = "-" Or Left(rightDate, 1) = " ")
rightDate = Trim(Mid(rightDate, 2))
Loop
If Len(rightDate) > 0 Then
Call SafeSetParagraphText(doc, pH, leftMotto & vbCr & rightDate)
End If
End If
ElseIf sHAsc Like "DANG BO *" Or sHAsc Like "TINH UY *" Or sHAsc Like "THANH UY *" Or sHAsc Like "HUYEN UY *" Or sHAsc Like "DANG UY *" Then
Dim splitPos As Long: splitPos = 0
If InStr(12, sHAsc, " BAN THUONG VU") > 0 Then
splitPos = InStr(12, sHAsc, " BAN THUONG VU")
ElseIf InStr(12, sHAsc, " BAN CHAP HANH") > 0 Then
splitPos = InStr(12, sHAsc, " BAN CHAP HANH")
ElseIf InStr(12, sHAsc, " BAN XAY DUNG") > 0 Then
splitPos = InStr(12, sHAsc, " BAN XAY DUNG")
ElseIf InStr(12, sHAsc, " BAN TUYEN GIAO") > 0 Then
splitPos = InStr(12, sHAsc, " BAN TUYEN GIAO")
ElseIf InStr(12, sHAsc, " BAN TO CHUC") > 0 Then
splitPos = InStr(12, sHAsc, " BAN TO CHUC")
ElseIf InStr(12, sHAsc, " VAN PHONG") > 0 Then
splitPos = InStr(12, sHAsc, " VAN PHONG")
ElseIf InStr(12, sHAsc, " UY BAN KIEM TRA") > 0 Then
splitPos = InStr(12, sHAsc, " UY BAN KIEM TRA")
ElseIf InStr(12, sHAsc, " DANG UY ") > 0 And Not (sHAsc Like "BAN CHAP HANH*") Then
splitPos = InStr(12, sHAsc, " DANG UY ")
ElseIf InStr(12, sHAsc, " CHI BO ") > 0 Then
splitPos = InStr(12, sHAsc, " CHI BO ")
End If
If splitPos > 10 Then
Dim org1 As String, org2 As String
org1 = Trim(Left(sH, splitPos - 1))
org2 = Trim(Mid(sH, splitPos + 1))
If Len(org1) > 0 And Len(org2) > 0 Then
Call SafeSetParagraphText(doc, pH, org1 & vbCr & org2)
End If
End If
End If
Next pH
Next ch
End If
End If
Dim pBody As Paragraph
Dim nonTblChecked As Integer: nonTblChecked = 0
For Each pBody In doc.Paragraphs
If Not pBody.Range.Information(wdWithInTable) Then
Dim sRawB As String
sRawB = CleanParaText(pBody.Range.text)
If Len(sRawB) > 0 Then
nonTblChecked = nonTblChecked + 1
If nonTblChecked > 4 Then Exit For
Dim sAscB As String
sAscB = Trim(ToAscii(sRawB))
Dim titleLen As Integer
titleLen = GetGluedTitlePrefixLen(sRawB, sAscB)
If titleLen > 0 Then
Dim partTitle As String, partSummary As String
partTitle = Trim(Left(sRawB, titleLen))
partSummary = Trim(Mid(sRawB, titleLen + 1))
If Left(partSummary, 1) = ":" Or Left(partSummary, 1) = "-" Then
partSummary = Trim(Mid(partSummary, 2))
End If
If Len(partTitle) > 0 And Len(partSummary) > 0 Then
Call SafeSetParagraphText(doc, pBody, partTitle & vbCr & partSummary)
Exit For
End If
End If
End If
End If
Next pBody
On Error GoTo 0
End Sub
Private Function IsGluedTitleLine(ByVal sRaw As String, ByVal sAsc As String) As Boolean
IsGluedTitleLine = (GetGluedTitlePrefixLen(sRaw, sAsc) > 0)
End Function
Private Function GetGluedTitlePrefixLen(ByVal sRaw As String, ByVal sAsc As String) As Integer
GetGluedTitlePrefixLen = 0
If IsPartyDocTitle(sAsc) Then Exit Function
Dim titles(1 To 16) As String
titles(1) = "CHUONG TRINH HANH DONG"
titles(2) = "BAO CAO CHINH TRI"
titles(3) = "DE CUONG BAO CAO"
titles(4) = "CHUONG TRINH"
titles(5) = "NGHI QUYET"
titles(6) = "QUYET DINH"
titles(7) = "HUONG DAN"
titles(8) = "THONG BAO"
titles(9) = "THONG TRI"
titles(10) = "KE HOACH"
titles(11) = "KET LUAN"
titles(12) = "QUY DINH"
titles(13) = "TO TRINH"
titles(14) = "BIEN BAN"
titles(15) = "BAO CAO"
titles(16) = "QUY CHE"
Dim idx As Integer
For idx = 1 To 16
Dim kw As String: kw = titles(idx)
Dim kwLen As Integer: kwLen = Len(kw)
If Len(sAsc) > kwLen + 3 Then
If Left(sAsc, kwLen) = kw Then
Dim rawPrefix As String
rawPrefix = Left(sRaw, kwLen)
Dim rawRest As String
rawRest = Trim(Mid(sRaw, kwLen + 1))
If rawPrefix = UCase(rawPrefix) And Len(rawRest) > 2 And rawRest <> UCase(rawRest) Then
GetGluedTitlePrefixLen = kwLen
Exit Function
End If
End If
End If
Next idx
End Function
Private Function InsertGapIfPrecededByTable(ByVal doc As Document, ByVal pFirst As Paragraph) As Boolean
InsertGapIfPrecededByTable = False
On Error GoTo GapFail
Dim pPrev As Paragraph
Set pPrev = pFirst.Previous
If pPrev Is Nothing Then Exit Function
If Not pPrev.Range.Information(wdWithInTable) Then Exit Function
Dim posStart As Long
posStart = pFirst.Range.Start
doc.Range(posStart, posStart).InsertBefore vbCr
InsertGapIfPrecededByTable = True
Exit Function
GapFail:
InsertGapIfPrecededByTable = False
End Function
Private Sub AutoCreateKinhGuiTable(ByVal doc As Document)
On Error Resume Next
Dim prefixText As String
prefixText = "K" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.text)
If IsKinhGuiTable(tbl, tAsc) Then
Dim kgCellLeft As Cell: Set kgCellLeft = Nothing
Dim kgCellRight As Cell: Set kgCellRight = Nothing
Dim cScanKG As Cell
For Each cScanKG In tbl.Range.Cells
Dim cScanAsc As String
cScanAsc = Trim(ToAscii(cScanKG.Range.text))
If HasKinhGuiOrTrinh(cScanAsc) Then
Set kgCellLeft = cScanKG
ElseIf Not kgCellLeft Is Nothing And kgCellRight Is Nothing Then
Set kgCellRight = cScanKG
Exit For
End If
Next cScanKG
Dim pAfterTbl As Paragraph
Dim pAbsorbStart As Paragraph: Set pAbsorbStart = Nothing
Dim pAbsorbEnd As Paragraph: Set pAbsorbEnd = Nothing
Dim lastPInTbl As Paragraph
Set lastPInTbl = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count)
Set pAfterTbl = lastPInTbl.Next
Do While Not pAfterTbl Is Nothing
If pAfterTbl.Range.Information(wdWithInTable) Then Exit Do
Dim sAfterRaw As String
sAfterRaw = CleanParaText(pAfterTbl.Range.text)
If Len(sAfterRaw) > 0 Then
Dim sAfterAsc As String
sAfterAsc = Trim(ToAscii(sAfterRaw))
If IsKinhGuiRecipientItem(sAfterRaw, sAfterAsc) Then
If pAbsorbStart Is Nothing Then Set pAbsorbStart = pAfterTbl
Set pAbsorbEnd = pAfterTbl
Else
Exit Do
End If
End If
Set pAfterTbl = pAfterTbl.Next
Loop
If Not pAbsorbStart Is Nothing And Not pAbsorbEnd Is Nothing Then
Dim targetCell As Cell
Set targetCell = kgCellRight
If targetCell Is Nothing Then Set targetCell = kgCellLeft
If Not targetCell Is Nothing Then
Dim allRecipText As String: allRecipText = ""
Dim pInC As Paragraph
For Each pInC In targetCell.Range.Paragraphs
Dim sCLine As String
sCLine = CleanParaText(pInC.Range.text)
If Len(sCLine) > 0 Then
If Len(allRecipText) > 0 Then
allRecipText = allRecipText & vbCr & sCLine
Else
allRecipText = sCLine
End If
End If
Next pInC
Dim pCurAbsorb As Paragraph
Set pCurAbsorb = pAbsorbStart
Do While Not pCurAbsorb Is Nothing
Dim sAbsLine As String
sAbsLine = CleanParaText(pCurAbsorb.Range.text)
If Len(sAbsLine) > 0 Then
If Len(allRecipText) > 0 Then
allRecipText = allRecipText & vbCr & sAbsLine
Else
allRecipText = sAbsLine
End If
End If
If pCurAbsorb.Range.Start >= pAbsorbEnd.Range.Start Then Exit Do
Set pCurAbsorb = pCurAbsorb.Next
Loop
doc.Range(pAbsorbStart.Range.Start, pAbsorbEnd.Range.End).Delete
targetCell.Range.text = allRecipText
targetCell.Range.Font.Bold = False
targetCell.Range.Font.Italic = False
End If
End If
If Not kgCellLeft Is Nothing Then
Dim leftCellTxt As String
leftCellTxt = ToAscii(kgCellLeft.Range.text)
Call SafeSetParagraphText(doc, kgCellLeft.Range.Paragraphs(1), KinhGuiPrefixFor(leftCellTxt))
End If
Exit Sub
End If
Next tbl
Dim p As Paragraph
Dim pStart As Paragraph: Set pStart = Nothing
Dim pEnd As Paragraph: Set pEnd = Nothing
Dim foundStart As Boolean: foundStart = False
Dim scanCount As Integer: scanCount = 0
For Each p In doc.Paragraphs
If Not p.Range.Information(wdWithInTable) Then
Dim sRaw As String
sRaw = CleanParaText(p.Range.text)
If Len(sRaw) > 0 Then
scanCount = scanCount + 1
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If Not foundStart Then
If IsKinhGuiStart(sRaw, sAsc) Then
Set pStart = p
Set pEnd = p
foundStart = True
ElseIf scanCount > 25 Then
Exit Sub
End If
Else
If IsKinhGuiRecipientItem(sRaw, sAsc) Then
Set pEnd = p
Else
Exit For
End If
End If
End If
End If
Next p
If Not foundStart Or pStart Is Nothing Or pEnd Is Nothing Then
Exit Sub
End If
Dim recipientsText As String: recipientsText = ""
Dim currP As Paragraph
Set currP = pStart
Do While Not currP Is Nothing
Dim lineText As String
lineText = CleanParaText(currP.Range.text)
If Len(lineText) > 0 Then
Dim lineAsc As String
lineAsc = Trim(ToAscii(lineText))
If currP.Range.Start = pStart.Range.Start Then
prefixText = KinhGuiPrefixFor(lineAsc)
Dim cPos As Long
cPos = InStr(1, lineText, ":")
If cPos > 0 And cPos < Len(lineText) Then
Dim afterColon As String
afterColon = Trim(Mid(lineText, cPos + 1))
If Len(afterColon) > 0 Then
recipientsText = afterColon
End If
ElseIf cPos = 0 And (lineAsc Like "KINH GUI *" Or lineAsc Like "DONG KINH GUI *") Then
Dim firstSpacePos As Long
If lineAsc Like "DONG KINH GUI *" Then
firstSpacePos = InStr(15, lineText, " ")
If firstSpacePos = 0 Then firstSpacePos = 14
Else
firstSpacePos = InStr(9, lineText, " ")
If firstSpacePos = 0 Then firstSpacePos = 9
End If
Dim remText As String
remText = Trim(Mid(lineText, firstSpacePos + 1))
If Len(remText) > 0 Then
recipientsText = remText
End If
End If
Else
If Len(recipientsText) > 0 Then
recipientsText = recipientsText & vbCr & lineText
Else
recipientsText = lineText
End If
End If
End If
If currP.Range.Start >= pEnd.Range.Start Then Exit Do
Set currP = currP.Next
Loop
If Len(Trim(recipientsText)) = 0 Then
Exit Sub
End If
Dim kgStartPos As Long, kgEndPos As Long
kgStartPos = pStart.Range.Start
kgEndPos = pEnd.Range.End
If InsertGapIfPrecededByTable(doc, pStart) Then
kgStartPos = kgStartPos + 1
kgEndPos = kgEndPos + 1
End If
Dim rngReplace As Range
Set rngReplace = doc.Range(kgStartPos, kgEndPos)
Dim newTbl As Table
Set newTbl = doc.Tables.Add(Range:=rngReplace, NumRows:=1, NumColumns:=2)
With newTbl
.Borders.Enable = False
On Error Resume Next
.Rows(1).LeftIndent = 0
.Rows(1).Alignment = wdAlignRowLeft
.Columns(1).Width = MillimetersToPoints(52)
.Columns(2).Width = MillimetersToPoints(117)
On Error GoTo 0
.Cell(1, 1).Range.text = prefixText
.Cell(1, 1).Range.Font.Bold = False
.Cell(1, 1).Range.Font.Italic = TT_B("KINH_GUI_ITALIC")
.Cell(1, 2).Range.text = recipientsText
.Cell(1, 2).Range.Font.Bold = False
.Cell(1, 2).Range.Font.Italic = False
End With
On Error GoTo 0
End Sub
Private Sub AutoCreateFooterTable(ByVal doc As Document)
On Error Resume Next
Dim tbl As Table
For Each tbl In doc.Tables
Dim tAsc As String
tAsc = ToAscii(tbl.Range.text)
If IsFooterSignatureTable(tbl, tAsc) Then Exit Sub
Next tbl
Dim totalP As Long: totalP = doc.Paragraphs.Count
If totalP < 3 Then Exit Sub
Dim startScanIdx As Long
startScanIdx = totalP - 30
If startScanIdx < 2 Then startScanIdx = 2
Dim pFooterStart As Paragraph: Set pFooterStart = Nothing
Dim pFooterEnd As Paragraph: Set pFooterEnd = Nothing
Dim sawNoiNhan As Boolean: sawNoiNhan = False
Dim sawSigAuthOrRole As Boolean: sawSigAuthOrRole = False
Dim idx As Long
For idx = startScanIdx To totalP
Dim pCur As Paragraph
Set pCur = doc.Paragraphs(idx)
If pCur.Range.Information(wdWithInTable) Then
If Not pFooterStart Is Nothing Then Exit For
GoTo NextFooterScan
End If
Dim sRaw As String
sRaw = CleanParaText(pCur.Range.text)
If Len(sRaw) = 0 Then GoTo NextFooterScan
Dim sAsc As String
sAsc = Trim(ToAscii(sRaw))
If pFooterStart Is Nothing Then
If sAsc Like "NOI NHAN*" Or IsSignatureRoleOrAuth(sAsc) Or HasTabSeparatedFooterLine(pCur.Range.text) Then
Set pFooterStart = pCur
Set pFooterEnd = pCur
If InStr(1, sAsc, "NOI NHAN") > 0 Then sawNoiNhan = True
If IsSignatureRoleOrAuth(sAsc) Or HasTabSeparatedFooterLine(pCur.Range.text) Then sawSigAuthOrRole = True
End If
Else
If IsPartyDocTitle(sAsc) Or sAsc Like "PHU LUC*" Then Exit For
If Not sawNoiNhan And (sAsc Like "CAN CU *" Or sAsc Like "- CAN CU *" Or sAsc Like "THEO DE NGHI*" Or _
sAsc Like "XET DE NGHI*" Or sAsc Like "XET *" Or sAsc = "QUYET DINH:" Or sAsc = "QUYET DINH :" Or _
sAsc Like "DIEU #*" Or sAsc Like "I. *" Or sAsc Like "II. *" Or sAsc Like "CHUONG *" Or sAsc Like "PHAN *") Then
Set pFooterStart = Nothing
Set pFooterEnd = Nothing
sawSigAuthOrRole = False
GoTo NextFooterScan
End If
Set pFooterEnd = pCur
If InStr(1, sAsc, "NOI NHAN") > 0 Then sawNoiNhan = True
If IsSignatureRoleOrAuth(sAsc) Then sawSigAuthOrRole = True
End If
NextFooterScan:
Next idx
If pFooterStart Is Nothing Or pFooterEnd Is Nothing Then Exit Sub
If Not (sawNoiNhan Or sawSigAuthOrRole) Then Exit Sub
Dim leftFooter As String: leftFooter = ""
Dim rightFooter As String: rightFooter = ""
Dim inSigBlockSeq As Boolean: inSigBlockSeq = False
Dim pIter As Paragraph
Set pIter = pFooterStart
Do While Not pIter Is Nothing
Dim rawF As String
rawF = Replace(Replace(Replace(pIter.Range.text, vbCr, ""), vbLf, ""), Chr(7), "")
rawF = Replace(rawF, Chr(11), vbCr)
rawF = TT_SpacesToTab(rawF)
Dim fSubLines() As String
fSubLines = Split(rawF, vbCr)
Dim fIdx As Long
For fIdx = LBound(fSubLines) To UBound(fSubLines)
Call ClassifyAndAppendFooterLine(fSubLines(fIdx), leftFooter, rightFooter, inSigBlockSeq)
Next fIdx
If pIter.Range.Start >= pFooterEnd.Range.Start Then Exit Do
Set pIter = pIter.Next
Loop
If Len(Trim(leftFooter)) = 0 And Len(Trim(rightFooter)) = 0 Then Exit Sub
Dim ftStartPos As Long, ftEndPos As Long
ftStartPos = pFooterStart.Range.Start
ftEndPos = pFooterEnd.Range.End
If InsertGapIfPrecededByTable(doc, pFooterStart) Then
ftStartPos = ftStartPos + 1
ftEndPos = ftEndPos + 1
End If
Dim rngFooterReplace As Range
Set rngFooterReplace = doc.Range(ftStartPos, ftEndPos)
Dim newFooterTbl As Table
Set newFooterTbl = doc.Tables.Add(Range:=rngFooterReplace, NumRows:=1, NumColumns:=2)
With newFooterTbl
.Borders.Enable = False
.Columns(1).Width = MillimetersToPoints(80)
.Columns(2).Width = MillimetersToPoints(85)
.Cell(1, 1).Range.text = leftFooter
.Cell(1, 2).Range.text = rightFooter
End With
On Error GoTo 0
End Sub
Private Function HasTabSeparatedFooterLine(ByVal rawParaText As String) As Boolean
HasTabSeparatedFooterLine = False
If InStr(1, rawParaText, vbTab) = 0 Then Exit Function
Dim sAsc As String
sAsc = Trim(ToAscii(CleanParaText(rawParaText)))
If InStr(1, sAsc, "NOI NHAN") > 0 Or InStr(1, sAsc, "T/M ") > 0 Or InStr(1, sAsc, "TM.") > 0 Or _
InStr(1, sAsc, "K/T ") > 0 Or InStr(1, sAsc, "KT.") > 0 Or InStr(1, sAsc, "BI THU") > 0 Or _
InStr(1, sAsc, "TRUONG BAN") > 0 Or InStr(1, sAsc, "CHANH VAN PHONG") > 0 Or _
InStr(1, sAsc, "TUQ.") > 0 Or InStr(1, sAsc, "TL.") > 0 Or InStr(1, sAsc, "T/L ") > 0 Or _
InStr(1, sAsc, "CHU TICH") > 0 Or InStr(1, sAsc, "GIAM DOC") > 0 Then
HasTabSeparatedFooterLine = True
End If
End Function
Private Sub ClassifyAndAppendFooterLine(ByVal rawLine As String, ByRef leftFooter As String, ByRef rightFooter As String, ByRef inSigBlockSeq As Boolean)
Dim sCleanCheck As String
sCleanCheck = CleanParaText(rawLine)
If Len(sCleanCheck) = 0 Then Exit Sub
If InStr(1, rawLine, vbTab) > 0 Then
Dim firstTab As Long
firstTab = InStr(1, rawLine, vbTab)
Dim lPart As String, rPart As String
lPart = CleanParaText(Left(rawLine, firstTab - 1))
rPart = CleanParaText(Mid(rawLine, firstTab + 1))
If Len(lPart) > 0 Then
Dim lAsc As String: lAsc = Trim(ToAscii(lPart))
If IsSignatureRoleOrAuth(lAsc) Then
Call AppendLineToBuffer(rightFooter, lPart)
inSigBlockSeq = True
Else
Call AppendLineToBuffer(leftFooter, lPart)
End If
End If
If Len(rPart) > 0 Then
Call AppendLineToBuffer(rightFooter, rPart)
inSigBlockSeq = True
End If
Exit Sub
End If
Dim sAsc As String
sAsc = Trim(ToAscii(sCleanCheck))
If sAsc Like "NOI NHAN*" Or Left(sCleanCheck, 1) = "-" Or Left(sCleanCheck, 1) = ChrW(8211) Or _
sAsc Like "LUU:*" Or sAsc Like "LUU VT*" Or sAsc Like "NHU TREN*" Or sAsc Like "NHU DIEU*" Then
Call AppendLineToBuffer(leftFooter, sCleanCheck)
ElseIf IsSignatureRoleOrAuth(sAsc) Then
Call AppendLineToBuffer(rightFooter, sCleanCheck)
inSigBlockSeq = True
ElseIf inSigBlockSeq And Len(sCleanCheck) < 65 Then
Call AppendLineToBuffer(rightFooter, sCleanCheck)
Else
Call AppendLineToBuffer(leftFooter, sCleanCheck)
End If
End Sub
Private Function IsDocNumberLineAsc(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If Len(t) = 0 Or Len(t) > 35 Then
IsDocNumberLineAsc = False
Exit Function
End If
If (t Like "SO *" Or t Like "SO:*" Or t Like "SO-*") And _
(InStr(1, t, "/") > 0 Or t Like "SO #*" Or t Like "SO:*#*" Or t Like "SO -*") Then
IsDocNumberLineAsc = True
ElseIf Len(t) <= 20 And InStr(1, t, " ") = 0 And _
(InStr(1, t, "-CV/") > 0 Or InStr(1, t, "-BC/") > 0 Or InStr(1, t, "-KH/") > 0 Or _
InStr(1, t, "-QD/") > 0 Or InStr(1, t, "-NQ/") > 0 Or InStr(1, t, "-TTR/") > 0 Or _
InStr(1, t, "-TB/") > 0 Or InStr(1, t, "-KL/") > 0 Or InStr(1, t, "-CT/") > 0 Or _
InStr(1, t, "-HD/") > 0 Or InStr(1, t, "-QC/") > 0 Or InStr(1, t, "-QYD/") > 0 Or _
InStr(1, t, "-CTR/") > 0) Then
IsDocNumberLineAsc = True
Else
IsDocNumberLineAsc = False
End If
End Function
Private Function IsHeaderOrganLineAsc(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If t Like "DANG BO *" Or t Like "DANG UY *" Or t Like "TINH UY *" Or _
t Like "THANH UY *" Or t Like "HUYEN UY *" Or t Like "THI UY *" Or _
t Like "QUAN UY *" Or t Like "CHI BO *" Or t Like "BAN CHAP HANH*" Or _
t Like "BAN THUONG VU*" Or t Like "BAN XAY DUNG*" Or t Like "BAN TUYEN GIAO*" Or _
t Like "BAN TO CHUC*" Or t Like "BAN DAN VAN*" Or t Like "VAN PHONG*" Or _
t Like "UY BAN KIEM TRA*" Or t Like "TIEU BAN*" Or t Like "HOI DONG*" Or _
t Like "UY BAN NHAN DAN*" Or t Like "UBND *" Or t Like "HDND *" Or _
t Like "BO CONG AN*" Or t Like "CONG AN *" Or t Like "BO QUOC PHONG*" Or _
t Like "BO CHI HUY*" Or t Like "BAN CHI HUY*" Or t Like "BO TU LENH*" Or t Like "QUAN KHU*" Or _
t Like "UY BAN MAT TRAN*" Or t Like "MAT TRAN TO QUOC*" Or t Like "BAN THUONG TRUC*" Then
IsHeaderOrganLineAsc = True
Else
IsHeaderOrganLineAsc = False
End If
End Function
Private Function IsNationalOrganLine(ByVal sRaw As String, ByVal sAsc As String) As Boolean
IsNationalOrganLine = False
If Not TT_IsNational() Then Exit Function
Dim t As String
t = Trim(sRaw)
If Len(t) < 3 Or Len(t) > 90 Then Exit Function
If UCase(t) <> t Then Exit Function
If Not (sAsc Like "*[A-Z][A-Z]*") Then Exit Function
If IsDocNumberLineAsc(sAsc) Or IsSeparatorLine(t) Then Exit Function
If InStr(1, sAsc, "NGAY") > 0 And InStr(1, sAsc, "THANG") > 0 Then Exit Function
If IsPartyDocTitle(sAsc) Or IsKinhGuiStart(sRaw, sAsc) Then Exit Function
IsNationalOrganLine = True
End Function
Private Function IsKinhGuiTable(ByVal tbl As Table, ByVal tAscKG As String) As Boolean
IsKinhGuiTable = False
If Not HasKinhGuiOrTrinh(tAscKG) Or InStr(1, tAscKG, "NOI NHAN") > 0 Then Exit Function
If tbl.Columns.Count > 3 Or tbl.Rows.Count > 5 Then Exit Function
Dim cCheck As Cell
For Each cCheck In tbl.Range.Cells
Dim cAsc As String
cAsc = Trim(ToAscii(CleanParaText(cCheck.Range.text)))
If cAsc Like "KINH GUI*" Or cAsc Like "DONG KINH GUI*" Or IsKinhTrinhAsc(cAsc) Then
IsKinhGuiTable = True
Exit Function
End If
Next cCheck
End Function
Private Function IsFooterSignatureTable(ByVal tbl As Table, ByVal tAscFooter As String) As Boolean
IsFooterSignatureTable = False
If InStr(1, tAscFooter, "DANG CONG SAN") > 0 Then Exit Function
If InStr(1, tAscFooter, "CONG HOA XA HOI") > 0 Then Exit Function
If tbl.Rows.Count > 4 Or tbl.Columns.Count > 4 Then Exit Function
Dim cCheck As Cell
For Each cCheck In tbl.Range.Cells
Dim pCheck As Paragraph
For Each pCheck In cCheck.Range.Paragraphs
Dim pAsc As String
pAsc = Trim(ToAscii(CleanParaText(pCheck.Range.text)))
If pAsc Like "NOI NHAN*" Or IsSignatureRoleOrAuth(pAsc) Then
If InStr(1, tAscFooter, "NOI NHAN") = 0 And tbl.Rows.Count >= 3 Then
If tbl.Borders.Enable <> 0 Then Exit Function
End If
IsFooterSignatureTable = True
Exit Function
End If
Next pCheck
Next cCheck
End Function
Private Function HasSignatureKeywordInCell(ByVal cAsc As String) As Boolean
HasSignatureKeywordInCell = False
If InStr(1, cAsc, "T/M ") > 0 Or InStr(1, cAsc, "TM.") > 0 Or _
InStr(1, cAsc, "K/T ") > 0 Or InStr(1, cAsc, "KT.") > 0 Or _
InStr(1, cAsc, "T/L ") > 0 Or InStr(1, cAsc, "TL.") > 0 Or _
InStr(1, cAsc, "BI THU") > 0 Or InStr(1, cAsc, "CHANH VAN PHONG") > 0 Or _
InStr(1, cAsc, "TRUONG BAN") > 0 Or InStr(1, cAsc, "CHU NHIEM") > 0 Or _
InStr(1, cAsc, "CHU TICH") > 0 Or InStr(1, cAsc, "NGUOI LAP") > 0 Or _
InStr(1, cAsc, "TUQ.") > 0 Or InStr(1, cAsc, "GIAM DOC") > 0 Or _
InStr(1, cAsc, "TRUONG PHONG") > 0 Or InStr(1, cAsc, "THU TRUONG") > 0 Or _
InStr(1, cAsc, "TU LENH") > 0 Or InStr(1, cAsc, "CHINH UY") > 0 Or _
InStr(1, cAsc, "CHI HUY TRUONG") > 0 Or InStr(1, cAsc, "BO TRUONG") > 0 Or _
InStr(1, cAsc, "THU KY") > 0 Or InStr(1, cAsc, "CHU TOA") > 0 Or InStr(1, cAsc, "CHU TRI") > 0 Then
HasSignatureKeywordInCell = True
End If
End Function
Private Function IsSignatureRoleOrAuth(ByVal sAsc As String) As Boolean
Dim t As String: t = Trim(sAsc)
If Len(t) = 0 Or Len(t) > 65 Then
IsSignatureRoleOrAuth = False
Exit Function
End If
If t Like "T/M *" Or t Like "TM.*" Or t Like "TM *" Or _
t Like "K/T *" Or t Like "KT.*" Or t Like "KT *" Or _
t Like "T/L *" Or t Like "TL.*" Or t Like "TL *" Or _
t = "BI THU" Or t Like "PHO BI THU*" Or t Like "QUYEN BI THU*" Or _
t Like "CHANH VAN PHONG*" Or t Like "PHO CHANH VAN PHONG*" Or _
t Like "TRUONG BAN*" Or t Like "PHO TRUONG BAN*" Or _
t Like "CHU NHIEM*" Or t Like "PHO CHU NHIEM*" Or _
t Like "UY VIEN*" Or t = "TONG BI THU" Or _
t Like "CHU TICH*" Or t Like "PHO CHU TICH*" Or _
t Like "THU TRUONG*" Or t Like "NGUOI LAP*" Or t Like "NGUOI BAO CAO*" Or t Like "XAC NHAN*" Or _
t Like "TUQ.*" Or t Like "T/UQ *" Or t Like "Q. *" Or t Like "Q.[A-Z]*" Or _
t Like "GIAM DOC*" Or t Like "PHO GIAM DOC*" Or t Like "TRUONG PHONG*" Or t Like "PHO TRUONG PHONG*" Or _
t Like "BO TRUONG*" Or t Like "CUC TRUONG*" Or t Like "PHO CUC TRUONG*" Or _
t Like "TU LENH*" Or t Like "PHO TU LENH*" Or t Like "CHINH UY*" Or t Like "PHO CHINH UY*" Or _
t Like "CHI HUY TRUONG*" Or t Like "CHI HUY PHO*" Or t Like "TRUONG CONG AN*" Or t Like "PHO TRUONG CONG AN*" Or _
t Like "HIEU TRUONG*" Or t Like "PHO HIEU TRUONG*" Or t Like "VIEN TRUONG*" Or t Like "PHO VIEN TRUONG*" Or _
t = "THU KY" Or t = "CHU TOA" Or t = "CHU TRI" Then
IsSignatureRoleOrAuth = True
Else
IsSignatureRoleOrAuth = TT_IsGenericSignerRole(t)
End If
End Function
Private Function TT_IsGenericSignerRole(ByVal t As String) As Boolean
TT_IsGenericSignerRole = False
t = UCase(Trim(t))
If Len(t) < 5 Or Len(t) > 40 Then Exit Function
If t Like "*#*" Or t Like "*[.,;:]" Or t Like "[-+(]*" Then Exit Function
If UBound(Split(t, " ")) > 5 Then Exit Function
If t = "NGUOI LAM DON" Or t Like "NGUOI VIET*" Or t Like "NGUOI DE NGHI*" Or t Like "NGUOI KHAI*" Or _
t Like "NGUOI CAM KET*" Or t Like "NGUOI NHAN*" Or t Like "NGUOI GIAO*" Or t Like "NGUOI GHI*" Or _
t Like "NGUOI KIEM TRA*" Or t Like "NGUOI TUONG TRINH*" Or t Like "NGUOI LAM CHUNG*" Or t Like "NGUOI DAI DIEN*" Or _
t Like "NGUOI UY QUYEN*" Or t Like "NGUOI DUOC UY QUYEN*" Or t Like "NGUOI YEU CAU*" Or t Like "NGUOI THUC HIEN*" Or _
t Like "NGUOI XIN *" Or t Like "NGUOI DUNG DON*" Or t Like "NGUOI VIET DON*" Or _
t Like "DAI DIEN *" Or t = "BEN A" Or t = "BEN B" Or t Like "DAI DIEN BEN *" Or _
t Like "TRUONG TRAM*" Or t = "TRAM TRUONG" Or t Like "PHO TRUONG TRAM*" Or t Like "TO TRUONG*" Or t Like "TO PHO*" Or _
t Like "DOI TRUONG*" Or t Like "DOI PHO*" Or t Like "TRUONG AP*" Or t Like "TRUONG THON*" Or t Like "TRUONG KHOM*" Or _
t Like "TRUONG KHU*" Or t = "KE TOAN TRUONG" Or t = "KE TOAN" Or t = "THU QUY" Or t Like "CHANH AN*" Or _
t Like "TONG GIAM DOC*" Or t Like "PHO TONG GIAM DOC*" Or t Like "TONG BIEN TAP*" Or t Like "CHU TICH CONG DOAN*" Then
TT_IsGenericSignerRole = True
End If
End Function
Private Function IsKinhTrinhAsc(ByVal sAsc As String) As Boolean
Dim t As String
t = Trim(sAsc)
IsKinhTrinhAsc = (t Like "KINH TRINH:*" Or t Like "KINH TRINH :*")
End Function
Private Function HasKinhGuiOrTrinh(ByVal sAsc As String) As Boolean
HasKinhGuiOrTrinh = (InStr(1, sAsc, "KINH GUI") > 0 Or InStr(1, sAsc, "KINH TRINH:") > 0 Or InStr(1, sAsc, "KINH TRINH :") > 0)
End Function
Private Function KinhGuiPrefixFor(ByVal sAsc As String) As String
If InStr(1, sAsc, "DONG KINH GUI") > 0 Then
KinhGuiPrefixFor = ChrW(272) & ChrW(7891) & "ng k" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
ElseIf InStr(1, sAsc, "KINH TRINH") > 0 Then
KinhGuiPrefixFor = "K" & ChrW(237) & "nh tr" & ChrW(236) & "nh:"
Else
KinhGuiPrefixFor = "K" & ChrW(237) & "nh g" & ChrW(7917) & "i:"
End If
End Function
Private Function IsKinhGuiStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tAsc As String: tAsc = Trim(sAsc)
Dim tRawKG As String: tRawKG = Trim(rawText)
If InStr(1, Left(tRawKG, 16), ":") = 0 And InStr(1, tRawKG, ",") > 0 Then
If Right(tRawKG, 1) = "." Or Right(tRawKG, 1) = "!" Or Right(tRawKG, 1) = "?" Then
IsKinhGuiStart = False
Exit Function
End If
End If
If tAsc Like "KINH GUI:*" Or tAsc Like "KINH GUI :*" Or _
tAsc Like "DONG KINH GUI:*" Or tAsc Like "DONG KINH GUI :*" Or _
tAsc Like "KINH GUI *" Or tAsc Like "DONG KINH GUI *" Or _
tAsc = "KINH GUI" Or tAsc = "DONG KINH GUI" Or IsKinhTrinhAsc(tAsc) Then
IsKinhGuiStart = True
Else
IsKinhGuiStart = False
End If
End Function
Private Function IsKinhGuiRecipientItem(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tRaw As String: tRaw = CleanParaText(rawText)
Dim tAsc As String: tAsc = Trim(ToAscii(tRaw))
If Len(tRaw) = 0 Or Len(tRaw) > 300 Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If IsSeparatorLine(tRaw) Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If tAsc Like "CAN CU*" Or tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
tAsc Like "QUYET DINH:*" Or tAsc Like "QUYET DINH :*" Or _
tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Or _
tAsc Like "DE *" Or tAsc Like "NHAM *" Then
IsKinhGuiRecipientItem = False
Exit Function
End If
If tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
tAsc Like "1.*" Or tAsc Like "2.*" Or tAsc Like "3.*" Or _
tAsc Like "DIEU #*" Or tAsc Like "DIEU ##*" Or _
IsPartyDocTitle(tAsc) Then
IsKinhGuiRecipientItem = False
Exit Function
End If
Dim firstCh As String: firstCh = Left(tRaw, 1)
If firstCh = "-" Or firstCh = ChrW(8211) Or firstCh = ChrW(8212) Or _
firstCh = ChrW(8210) Or firstCh = ChrW(8722) Or firstCh = ChrW(8226) Or _
firstCh = ChrW(8259) Or firstCh = ChrW(183) Or firstCh = "+" Or firstCh = "*" Then
IsKinhGuiRecipientItem = True
Exit Function
End If
Dim lastCh As String: lastCh = Right(tRaw, 1)
If lastCh = ";" Or lastCh = "," Then
IsKinhGuiRecipientItem = True
Exit Function
End If
If lastCh = "." And Len(tRaw) < 180 Then
If tAsc Like "*DANG UY*" Or tAsc Like "*CHI BO*" Or tAsc Like "*DANG BO*" Or _
tAsc Like "*BAN *" Or tAsc Like "*UBND*" Or tAsc Like "*HDND*" Or _
tAsc Like "*MAT TRAN*" Or tAsc Like "*MTTQ*" Or tAsc Like "*DOAN *" Or _
tAsc Like "*HOI *" Or tAsc Like "*PHONG *" Or tAsc Like "*SO *" Or _
tAsc Like "*XA *" Or tAsc Like "*HUYEN *" Or tAsc Like "*TINH *" Or _
tAsc Like "*TRUONG *" Or tAsc Like "*DON VI*" Or tAsc Like "*CO QUAN*" Or _
tAsc Like "*TRUNG TAM*" Or tAsc Like "*TRUYEN THANH*" Or tAsc Like "*VAN HOA*" Then
IsKinhGuiRecipientItem = True
Exit Function
End If
End If
IsKinhGuiRecipientItem = False
End Function
Private Sub NormalizeLineEndingPunctuation(ByVal doc As Document, ByVal p As Paragraph, ByVal isLastLine As Boolean)
On Error Resume Next
Dim rngLine As Range
Set rngLine = p.Range
Do While rngLine.End > rngLine.Start
Dim cCode As Long
cCode = AscW(Right(rngLine.text, 1))
If cCode < 0 Then cCode = cCode + 65536
If cCode = 13 Or cCode = 10 Or cCode = 7 Then
rngLine.End = rngLine.End - 1
Else
Exit Do
End If
Loop
Dim rngEndPos As Long
rngEndPos = rngLine.End
Do While rngLine.End > rngLine.Start
Dim endChCode As Long
endChCode = AscW(Right(rngLine.text, 1))
If endChCode < 0 Then endChCode = endChCode + 65536
If endChCode = 32 Or endChCode = 160 Or endChCode = 9 Or endChCode = 46 Or endChCode = 59 Or endChCode = 44 Then
rngLine.End = rngLine.End - 1
Else
Exit Do
End If
Loop
Dim rngTail As Range
Set rngTail = doc.Range(rngLine.End, rngEndPos)
If isLastLine Then
rngTail.text = "."
Else
rngTail.text = ";"
End If
On Error GoTo 0
End Sub
Private Sub FormatCenteredSummary(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 0
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub FormatCenteredSubNote(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 6
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = 14
p.Range.Font.Bold = False
p.Range.Font.Italic = True
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Sub FormatCenteredSeparator(ByVal p As Paragraph)
With p.Format
.Alignment = wdAlignParagraphCenter
.LeftIndent = 0
.RightIndent = 0
.FirstLineIndent = 0
.SpaceBefore = 0
.SpaceAfter = 14
.LineSpacingRule = wdLineSpaceSingle
.DisableLineHeightGrid = True
End With
p.Range.Font.Name = "Times New Roman"
p.Range.Font.Size = TT_N("SUMMARY_SIZE")
p.Range.Font.Bold = True
p.Range.Font.Italic = False
p.Range.Font.Underline = wdUnderlineNone
End Sub
Private Function HasSeparatorWithinNextParas(ByVal pCur As Paragraph, ByVal maxLookahead As Integer) As Boolean
HasSeparatorWithinNextParas = False
On Error Resume Next
Dim pNext As Paragraph
Set pNext = pCur.Next
Dim stepCount As Integer: stepCount = 0
Do While Not pNext Is Nothing And stepCount < maxLookahead
If pNext.Range.Information(wdWithInTable) Then Exit Do
Dim sNext As String
sNext = CleanParaText(pNext.Range.text)
If Len(sNext) > 0 Then
stepCount = stepCount + 1
If IsSeparatorLine(sNext) Then
HasSeparatorWithinNextParas = True
Exit Function
End If
Dim sNextAsc As String
sNextAsc = Trim(ToAscii(sNext))
If IsStrictBodyOrHeadingStart(sNext, sNextAsc) Or IsKinhGuiStart(sNext, sNextAsc) Then
Exit Function
End If
End If
Set pNext = pNext.Next
Loop
On Error GoTo 0
End Function
Private Function TT_CanonDocCode(ByVal code As String, ByVal isDang As Boolean) As String
Dim u As String
u = UCase(code)
TT_CanonDocCode = u
If u = "TTR" Then
TT_CanonDocCode = "TTr"
ElseIf u = "CTR" Then
TT_CanonDocCode = "CTr"
ElseIf Not isDang Then
If u = "QY" & ChrW(272) Then
TT_CanonDocCode = "Qy" & ChrW(272)
ElseIf u = "TRS" Then
TT_CanonDocCode = "TrS"
End If
End If
End Function
Private Function ChuanHoa_SoKyHieu(ByVal sRaw As String) As String
If InStr(1, sRaw, "[") > 0 Then
ChuanHoa_SoKyHieu = Trim(sRaw)
Exit Function
End If
Select Case TT_S("DOC_NUMBER_STYLE")
Case "NHA_NUOC"
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_NhaNuoc(sRaw, False)
Case "LINH_HOAT"
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_LinhHoat(sRaw)
Case Else
ChuanHoa_SoKyHieu = ChuanHoa_SoKyHieu_Dang(sRaw)
End Select
End Function
Private Function DocNumberStyleOf(ByVal sRaw As String) As String
DocNumberStyleOf = ""
Dim rest As String
rest = Trim(sRaw)
If UCase(Left(ToAscii(rest), 2)) <> "SO" Then Exit Function
rest = Trim(Mid(rest, 3))
If Left(rest, 1) = ":" Then rest = Trim(Mid(rest, 2))
rest = Replace(Replace(Replace(rest, " ", ""), ChrW(8211), "-"), ChrW(8212), "-")
Dim k As Long
For k = 1 To Len(rest)
Dim ch As String
ch = Mid(rest, k, 1)
If ch = "-" Then
If InStr(k + 1, rest, "/") > 0 Then DocNumberStyleOf = "DANG"
Exit Function
ElseIf ch = "/" Then
DocNumberStyleOf = "NHA_NUOC"
Exit Function
ElseIf Not (ch Like "#" Or ch = "." Or ch = ChrW(8230)) Then
Exit Function
End If
Next k
End Function
Private Function ChuanHoa_SoKyHieu_LinhHoat(ByVal sRaw As String) As String
Select Case DocNumberStyleOf(sRaw)
Case "DANG": ChuanHoa_SoKyHieu_LinhHoat = ChuanHoa_SoKyHieu_Dang(sRaw)
Case "NHA_NUOC": ChuanHoa_SoKyHieu_LinhHoat = ChuanHoa_SoKyHieu_NhaNuoc(sRaw, True)
Case Else: ChuanHoa_SoKyHieu_LinhHoat = Trim(sRaw)
End Select
End Function
Private Function ChuanHoa_SoKyHieu_NhaNuoc(ByVal sRaw As String, ByVal keepPrefix As Boolean) As String
Dim s As String
s = Trim(sRaw)
ChuanHoa_SoKyHieu_NhaNuoc = s
If UCase(Left(ToAscii(s), 2)) <> "SO" Then Exit Function
Dim rest As String
rest = Trim(Mid(s, 3))
Dim hadColon As Boolean
hadColon = (Left(rest, 1) = ":")
If hadColon Then rest = Mid(rest, 2)
rest = Replace(Trim(rest), " ", "")
rest = Replace(rest, Chr(160), "")
rest = Replace(rest, ChrW(8211), "-")
rest = Replace(rest, ChrW(8212), "-")
If Len(rest) = 0 Then Exit Function
Dim k As Long
Dim numPart As String
k = 1
Do While k <= Len(rest)
Dim ch As String
ch = Mid(rest, k, 1)
If ch Like "#" Or ch = "." Or ch = ChrW(8230) Then
numPart = numPart & ch
k = k + 1
Else
Exit Do
End If
Loop
Dim tail As String
tail = Mid(rest, k)
If Len(tail) < 2 Then Exit Function
Dim sep As String
sep = Left(tail, 1)
If sep <> "/" And sep <> "-" Then Exit Function
tail = Mid(tail, 2)
If sep = "-" Then
Dim sl As Long
sl = InStr(1, tail, "/")
If sl <= 1 Then Exit Function
If InStr(1, Left(tail, sl - 1), "-") = 0 Then
tail = Left(tail, sl - 1) & "-" & Mid(tail, sl + 1)
End If
End If
tail = UCase(tail)
Dim dashCode As Long
dashCode = InStr(1, tail, "-")
If dashCode > 1 Then
tail = TT_CanonDocCode(Left(tail, dashCode - 1), False) & Mid(tail, dashCode)
Else
tail = TT_CanonDocCode(tail, False)
End If
Dim cleanNum As String
cleanNum = Trim(Replace(Replace(numPart, ".", ""), ChrW(8230), ""))
Dim prefix As String
If keepPrefix And Not hadColon Then
prefix = "S" & ChrW(7889) & " "
Else
prefix = "S" & ChrW(7889) & ": "
End If
If Len(cleanNum) = 0 Or Not IsNumeric(cleanNum) Then
ChuanHoa_SoKyHieu_NhaNuoc = RTrim(prefix) & "   /" & tail
Else
If Len(cleanNum) = 1 And Val(cleanNum) > 0 Then cleanNum = "0" & cleanNum
ChuanHoa_SoKyHieu_NhaNuoc = prefix & cleanNum & "/" & tail
End If
End Function
Private Function ChuanHoa_SoKyHieu_Dang(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
Dim sAsc As String
sAsc = ToAscii(s)
If Left(sAsc, 3) = "SO:" Then
s = "S" & ChrW(7889) & " " & Trim(Mid(s, 4))
End If
Dim p1 As Long
p1 = InStr(1, s, " ")
If p1 > 0 Then
Dim prefix As String, rest As String
prefix = Left(s, p1)
rest = Replace(Trim(Mid(s, p1 + 1)), " ", "")
rest = Replace(rest, ChrW(8211), "-")
rest = Replace(rest, ChrW(8212), "-")
Dim dashPos As Long, slashPos As Long
dashPos = InStr(1, rest, "-")
slashPos = InStr(1, rest, "/")
Dim numPart As String, codePart As String, orgPart As String
If slashPos > 0 And dashPos > slashPos Then
numPart = Trim(Left(rest, slashPos - 1))
codePart = Trim(Mid(rest, slashPos + 1, dashPos - slashPos - 1))
orgPart = Trim(Mid(rest, dashPos + 1))
Dim cleanNum As String
cleanNum = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum = Trim(cleanNum)
If Len(cleanNum) = 0 Or Not IsNumeric(cleanNum) Then
s = "S" & ChrW(7889) & "   -" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
Else
If Val(cleanNum) < 10 And Val(cleanNum) > 0 And Len(cleanNum) = 1 Then
cleanNum = "0" & cleanNum
End If
s = "S" & ChrW(7889) & " " & cleanNum & "-" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
End If
ElseIf dashPos > 0 And slashPos > dashPos Then
numPart = Trim(Left(rest, dashPos - 1))
codePart = Trim(Mid(rest, dashPos + 1, slashPos - dashPos - 1))
orgPart = Trim(Mid(rest, slashPos + 1))
Dim cleanNum2 As String
cleanNum2 = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum2 = Trim(cleanNum2)
If Len(cleanNum2) = 0 Or Not IsNumeric(cleanNum2) Then
s = "S" & ChrW(7889) & "   -" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
Else
If Val(cleanNum2) < 10 And Val(cleanNum2) > 0 And Len(cleanNum2) = 1 Then
cleanNum2 = "0" & cleanNum2
End If
s = "S" & ChrW(7889) & " " & cleanNum2 & "-" & TT_CanonDocCode(codePart, True) & "/" & UCase(orgPart)
End If
ElseIf dashPos > 0 Then
numPart = Trim(Left(rest, dashPos - 1))
Dim cleanNum3 As String
cleanNum3 = Replace(Replace(numPart, ".", ""), ChrW(8230), "")
cleanNum3 = Trim(cleanNum3)
If Len(cleanNum3) = 0 Or Not IsNumeric(cleanNum3) Then
s = "S" & ChrW(7889) & "   -" & Mid(rest, dashPos + 1)
Else
If Val(cleanNum3) < 10 And Val(cleanNum3) > 0 And Len(cleanNum3) = 1 Then
s = "S" & ChrW(7889) & " 0" & rest
Else
s = "S" & ChrW(7889) & " " & rest
End If
End If
End If
End If
ChuanHoa_SoKyHieu_Dang = s
End Function
Private Function ChuanHoa_NgayThang(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
Dim sAsc As String
sAsc = ToAscii(s)
Dim pNgay As Long, pThang As Long, pNam As Long
pNgay = InStr(1, sAsc, "NGAY")
pThang = InStr(1, sAsc, "THANG")
pNam = InStr(1, sAsc, "NAM")
If pNgay > 0 And pThang > pNgay And pNam > pThang Then
Dim locPart As String, dayPart As String, monthPart As String, yearPart As String
locPart = Left(s, pNgay - 1)
dayPart = Trim(Mid(s, pNgay + 4, pThang - (pNgay + 4)))
monthPart = Trim(Mid(s, pThang + 5, pNam - (pThang + 5)))
yearPart = Trim(Mid(s, pNam + 3))
Dim cleanDay As String
cleanDay = Replace(Replace(dayPart, ".", ""), ChrW(8230), "")
cleanDay = Trim(cleanDay)
Dim dayText As String
If Len(cleanDay) = 0 Or Not IsNumeric(cleanDay) Then
dayText = "   "
Else
Dim dVal As Long: dVal = Val(cleanDay)
If dVal < 10 And dVal > 0 And Len(cleanDay) = 1 Then
cleanDay = "0" & cleanDay
End If
dayText = " " & cleanDay & " "
End If
Dim cleanMonth As String
cleanMonth = Replace(Replace(monthPart, ".", ""), ChrW(8230), "")
cleanMonth = Trim(cleanMonth)
Dim monthText As String
If Len(cleanMonth) = 0 Or Not IsNumeric(cleanMonth) Then
monthText = "   "
Else
Dim mVal As Long: mVal = Val(cleanMonth)
If (mVal = 1 Or mVal = 2) And Len(cleanMonth) = 1 Then
cleanMonth = "0" & cleanMonth
ElseIf mVal >= 3 And mVal <= 9 And Len(cleanMonth) > 1 And Left(cleanMonth, 1) = "0" Then
cleanMonth = CStr(mVal)
End If
monthText = " " & cleanMonth & " "
End If
Dim strNgay As String, strThang As String, strNam As String
strNgay = "ng" & ChrW(224) & "y"
strThang = "th" & ChrW(225) & "ng"
strNam = "n" & ChrW(259) & "m"
s = locPart & strNgay & dayText & strThang & monthText & strNam & " " & yearPart
End If
ChuanHoa_NgayThang = s
End Function
Private Function ChuanHoa_ThamQuyen(ByVal sRaw As String) As String
Dim s As String
s = Trim(sRaw)
ChuanHoa_ThamQuyen = s
Dim style As String
style = TT_S("SIGN_PREFIX_STYLE")
If style = "GIU_NGUYEN" Then Exit Function
Dim sUp As String
sUp = ToAscii(s)
Dim kind As String, restText As String
kind = ""
If Left(sUp, 4) = "T/UQ" Then
kind = "TUQ": restText = Mid(s, 5)
ElseIf Left(sUp, 4) = "TUQ." Or Left(sUp, 4) = "TUQ " Then
kind = "TUQ": restText = Mid(s, 5)
ElseIf Left(sUp, 3) = "T/M" Then
kind = "TM": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "TM." Or Left(sUp, 3) = "TM " Then
kind = "TM": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "K/T" Then
kind = "KT": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "KT." Or Left(sUp, 3) = "KT " Then
kind = "KT": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "T/L" Then
kind = "TL": restText = Mid(s, 4)
ElseIf Left(sUp, 3) = "TL." Or Left(sUp, 3) = "TL " Then
kind = "TL": restText = Mid(s, 4)
ElseIf sUp Like "Q.*" Then
kind = "Q": restText = Mid(s, 3)
End If
If Len(kind) = 0 Then Exit Function
restText = Trim(restText)
If Len(restText) = 0 Then Exit Function
If style = "DOT" Then
ChuanHoa_ThamQuyen = kind & ". " & restText
Else
Select Case kind
Case "TM": ChuanHoa_ThamQuyen = "T/M " & restText
Case "KT": ChuanHoa_ThamQuyen = "K/T " & restText
Case "TL": ChuanHoa_ThamQuyen = "T/L " & restText
Case "Q": ChuanHoa_ThamQuyen = "Q. " & restText
Case Else: ChuanHoa_ThamQuyen = s
End Select
End If
End Function
Private Function KiemTra_LogicTheThuc(ByVal doc As Document) As String
Dim warnings As String: warnings = ""
Dim warnCount As Integer: warnCount = 0
Dim docTextFull As String
docTextFull = doc.Content.text
Dim docAscFull As String
docAscFull = ToAscii(docTextFull)
Dim mismatchMsg As String
mismatchMsg = TT_MismatchWarning()
If Len(mismatchMsg) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". " & mismatchMsg
End If
If Len(mTTHeaderNote) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". " & mTTHeaderNote
End If
If InStr(1, docTextFull, "...") > 0 Or InStr(1, docTextFull, ChrW(8230)) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Du lieu trong] Phat hien con dau '...' chua dien day du thong tin."
End If
Dim phCount As Long: phCount = 0
Dim phPos As Long: phPos = InStr(1, docTextFull, "[")
Do While phPos > 0 And phCount < 500
If InStr(phPos + 1, docTextFull, "]") = 0 Then Exit Do
phCount = phCount + 1
phPos = InStr(phPos + 1, docTextFull, "[")
Loop
If phCount > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Cho trong] Con " & phCount & " cho can dien trong dau [ ] (vi du: [Ho va ten])."
End If
Dim foundSoRaw As String: foundSoRaw = ""
Dim foundNumVal As String: foundNumVal = ""
Dim foundCode As String: foundCode = ""
Dim foundDateRaw As String: foundDateRaw = ""
Dim foundTitle As String: foundTitle = ""
Dim foundSuperiorOrgan As String: foundSuperiorOrgan = ""
Dim foundIssuingOrgan As String: foundIssuingOrgan = ""
Dim hasVvSummary As Boolean: hasVvSummary = False
Dim hasKinhGui As Boolean: hasKinhGui = False
If InStr(1, docAscFull, "V/V") > 0 Or InStr(1, docAscFull, "VE VIEC") > 0 Then hasVvSummary = True
If InStr(1, docAscFull, "KINH GUI") > 0 Then hasKinhGui = True
Dim hasQuyetDinhVerb As Boolean: hasQuyetDinhVerb = False
Dim hasNoiNhan As Boolean: hasNoiNhan = False
Dim hasLuuVT As Boolean: hasLuuVT = False
Dim hasLuuOther As Boolean: hasLuuOther = False
Dim hasNhuDieu As Boolean: hasNhuDieu = False
Dim foundSignAuth As String: foundSignAuth = ""
Dim foundSignRole As String: foundSignRole = ""
If doc.Tables.Count >= 1 Then
Dim tblH As Table
Set tblH = doc.Tables(1)
Dim tHAsc As String
tHAsc = ToAscii(tblH.Range.text)
Dim tHDetectStrong As Boolean
If InStr(1, tHAsc, "DANG CONG SAN") > 0 Or InStr(1, tHAsc, "DANG UY") > 0 Or InStr(1, tHAsc, "DANG BO") > 0 Or InStr(1, tHAsc, "CHI BO") > 0 Or _
InStr(1, tHAsc, "CONG HOA XA HOI") > 0 Or (tblH.Rows.Count <= 3 And Len(TT_DetectFromHeaderText(tHAsc, tHDetectStrong)) > 0 And InStr(1, tHAsc, "NOI NHAN") = 0) Then
Dim ch As Cell
Dim orgLinesCount As Integer: orgLinesCount = 0
For Each ch In tblH.Range.Cells
Dim pH As Paragraph
For Each pH In ch.Range.Paragraphs
Dim sH As String: sH = CleanParaText(pH.Range.text)
Dim sHAsc As String: sHAsc = Trim(ToAscii(sH))
If Len(sHAsc) > 0 Then
If IsDocNumberLineAsc(sHAsc) Then
foundSoRaw = sHAsc
ElseIf InStr(1, sHAsc, "NGAY") > 0 And InStr(1, sHAsc, "THANG") > 0 Then
foundDateRaw = sHAsc
ElseIf sHAsc Like "V/V*" Or sHAsc Like "VE VIEC*" Then
hasVvSummary = True
ElseIf InStr(1, sHAsc, "DANG CONG SAN") = 0 And InStr(1, sHAsc, "CONG HOA XA HOI") = 0 And InStr(1, Replace(sHAsc, " ", ""), "DOCLAP-TUDO") = 0 And _
InStr(1, sHAsc, "KINH GUI") = 0 And Left(sH, 1) <> "-" And Not IsSeparatorLine(sH) And sHAsc <> "*" And sHAsc <> "***" Then
orgLinesCount = orgLinesCount + 1
If orgLinesCount = 1 Then
foundIssuingOrgan = sHAsc
ElseIf orgLinesCount = 2 Then
foundSuperiorOrgan = foundIssuingOrgan
foundIssuingOrgan = sHAsc
End If
End If
End If
Next pH
Next ch
End If
End If
Dim pCount As Long: pCount = doc.Paragraphs.Count
Dim nonTblIdx As Long: nonTblIdx = 0
Dim romanOrder(1 To 10) As String
romanOrder(1) = "I": romanOrder(2) = "II": romanOrder(3) = "III": romanOrder(4) = "IV": romanOrder(5) = "V"
romanOrder(6) = "VI": romanOrder(7) = "VII": romanOrder(8) = "VIII": romanOrder(9) = "IX": romanOrder(10) = "X"
Dim nextExpectedRoman As Integer: nextExpectedRoman = 1
Dim nextExpectedArabic As Integer: nextExpectedArabic = 1
Dim nextExpectedDieu As Integer: nextExpectedDieu = 1
Dim totalDieuCount As Integer: totalDieuCount = 0
Dim idx As Long
Dim pItem As Paragraph
For Each pItem In doc.Paragraphs
idx = idx + 1
Dim pText As String
pText = CleanParaText(pItem.Range.text)
If Len(pText) = 0 Then GoTo NextParaCheck
Dim pAsc As String
pAsc = Trim(ToAscii(pText))
Dim inTbl As Boolean
inTbl = pItem.Range.Information(wdWithInTable)
If Not inTbl Then nonTblIdx = nonTblIdx + 1
If Not inTbl And nonTblIdx <= 4 And Len(foundTitle) = 0 Then
If IsPartyDocTitle(pAsc) Then
foundTitle = NormalizeDocTitleForCompare(pAsc)
GoTo NextParaCheck
End If
End If
If InStr(1, pAsc, "KINH GUI") > 0 Then hasKinhGui = True
If pAsc Like "V/V*" Or pAsc Like "VE VIEC*" Then hasVvSummary = True
If pAsc = "QUYET DINH:" Or pAsc = "QUYET DINH :" Then hasQuyetDinhVerb = True
If idx >= pCount - 35 Or inTbl Then
If InStr(1, pAsc, "NOI NHAN") > 0 Then hasNoiNhan = True
If InStr(1, pAsc, "LUU VT") > 0 Or InStr(1, pAsc, "LUU: VT") > 0 Or InStr(1, pAsc, "LUU VAN THU") > 0 Or InStr(1, pAsc, "LUU: VAN THU") > 0 Or _
(TT_S("MOTTO_KIND") = "DANG" And InStr(1, pAsc, "LUU VAN PHONG") > 0) Then
hasLuuVT = True
ElseIf InStr(1, pAsc, "LUU ") > 0 Or InStr(1, pAsc, "LUU:") > 0 Then
hasLuuOther = True
End If
If InStr(1, pAsc, "NHU DIEU") > 0 Then hasNhuDieu = True
If pAsc Like "T/M *" Or pAsc Like "TM.*" Or pAsc Like "K/T *" Or pAsc Like "KT.*" Or pAsc Like "T/L *" Or pAsc Like "TL.*" Or pAsc Like "TUQ.*" Then
foundSignAuth = pAsc
ElseIf pAsc = "BI THU" Or pAsc Like "PHO BI THU*" Or pAsc = "CHANH VAN PHONG" Or pAsc = "PHO CHANH VAN PHONG" Or pAsc Like "TRUONG BAN*" Or pAsc Like "CHU NHIEM*" Then
foundSignRole = pAsc
End If
End If
If Not inTbl Then
If pAsc Like "DIEU #.*" Or pAsc Like "DIEU ##.*" Or pAsc Like "DIEU #:*" Then
Dim dieuNumStr As String
dieuNumStr = Trim(Mid(pAsc, 5, 3))
dieuNumStr = Replace(Replace(dieuNumStr, ".", ""), ":", "")
If IsNumeric(dieuNumStr) Then
Dim actualDieu As Integer: actualDieu = Val(dieuNumStr)
totalDieuCount = totalDieuCount + 1
If actualDieu <> nextExpectedDieu Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Thu tu Dieu] Nhay coc hoac sai thu tu: Ky vong 'Dieu " & nextExpectedDieu & "' nhung gap 'Dieu " & actualDieu & "'."
End If
nextExpectedDieu = actualDieu + 1
End If
End If
Dim rIdx As Integer
Dim matchedRoman As Boolean: matchedRoman = False
For rIdx = 1 To 10
If pAsc Like romanOrder(rIdx) & ". *" Or pAsc Like romanOrder(rIdx) & "- *" Or pAsc Like romanOrder(rIdx) & " - *" Then
matchedRoman = True
If rIdx <> nextExpectedRoman Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [De muc La Ma] Nhay coc thu tu: Ky vong muc '" & romanOrder(nextExpectedRoman) & "' nhung gap muc '" & romanOrder(rIdx) & "'."
End If
nextExpectedRoman = rIdx + 1
nextExpectedArabic = 1
Exit For
End If
Next rIdx
If Not matchedRoman Then
If (pAsc Like "#. *" Or pAsc Like "##. *") And Not (pAsc Like "#.#*") Then
Dim dotPos As Long: dotPos = InStr(1, pAsc, ".")
If dotPos > 1 Then
Dim arabStr As String: arabStr = Left(pAsc, dotPos - 1)
If IsNumeric(arabStr) Then
Dim actualArab As Integer: actualArab = Val(arabStr)
If actualArab <> nextExpectedArabic Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [De muc So] Nhay coc thu tu khoan: Ky vong muc '" & nextExpectedArabic & ".' nhung gap '" & actualArab & ".' (" & Left(pAsc, 35) & "...)."
End If
nextExpectedArabic = actualArab + 1
End If
End If
End If
End If
End If
NextParaCheck:
Next pItem
Dim isStandaloneNonDispatch As Boolean
isStandaloneNonDispatch = False
If foundTitle = "PHU LUC" Or foundTitle = "DANH MUC" Or foundTitle = "BANG TONG HOP" Or _
foundTitle = "BIEU TONG HOP" Or foundTitle = "BAI THU HOACH" Or foundTitle = "BAI PHAT BIEU" Or _
foundTitle = "BAN KIEM DIEM" Or foundTitle = "BAN CAM KET" Or foundTitle = "BAN DANG KY" Then
isStandaloneNonDispatch = True
End If
Dim foundSoStyle As String: foundSoStyle = ""
If Len(foundSoRaw) > 0 Then
Dim soBody As String
soBody = Trim(Mid(foundSoRaw, 3))
If Left(soBody, 1) = ":" Then soBody = Trim(Mid(soBody, 2))
soBody = Replace(soBody, " ", "")
Dim dashIdx As Long, slashIdx As Long
dashIdx = InStr(1, soBody, "-")
slashIdx = InStr(1, soBody, "/")
If dashIdx > 0 And slashIdx > dashIdx Then
foundSoStyle = "DANG"
foundNumVal = Left(soBody, dashIdx - 1)
foundCode = Mid(soBody, dashIdx + 1, slashIdx - dashIdx - 1)
ElseIf slashIdx > 0 Then
foundSoStyle = "NHA_NUOC"
foundNumVal = Left(soBody, slashIdx - 1)
Dim afterSlash As String
afterSlash = Mid(soBody, slashIdx + 1)
If afterSlash Like "####/*" Then afterSlash = Mid(afterSlash, 6)
Dim dashAfter As Long
dashAfter = InStr(1, afterSlash, "-")
If dashAfter > 0 Then
foundCode = Left(afterSlash, dashAfter - 1)
Else
foundCode = afterSlash
End If
End If
foundNumVal = Trim(Replace(Replace(foundNumVal, ".", ""), ChrW(8230), ""))
If Len(foundNumVal) = 0 Or Not IsNumeric(foundNumVal) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So van ban] Dang de trong 3 khoang trang ('" & foundSoRaw & "') cho van thu cap so."
End If
If TT_S("DOC_NUMBER_STYLE") = "DANG" And foundSoStyle = "NHA_NUOC" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Nha nuoc ('" & foundSoRaw & "'); van ban Dang ghi dang 'So 12-QD/DU'."
ElseIf TT_S("DOC_NUMBER_STYLE") = "NHA_NUOC" And foundSoStyle = "DANG" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Dang ('" & foundSoRaw & "'); van ban hanh chinh ghi dang 'So: 12/QD-UBND' (ND 30/2020)."
End If
ElseIf Not isStandaloneNonDispatch Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [So van ban] Khong tim thay dong So, ky hieu van ban o phan Tieu de."
End If
If Len(foundDateRaw) > 0 Then
Dim pNg As Long, pTh As Long, pNm As Long
pNg = InStr(1, foundDateRaw, "NGAY")
pTh = InStr(1, foundDateRaw, "THANG")
pNm = InStr(1, foundDateRaw, "NAM")
If pNg > 0 And pTh > pNg And pNm > pTh Then
Dim sDay As String, sMonth As String, sYear As String
sDay = Trim(Mid(foundDateRaw, pNg + 4, pTh - (pNg + 4)))
sMonth = Trim(Mid(foundDateRaw, pTh + 5, pNm - (pTh + 5)))
sYear = Trim(Left(Trim(Mid(foundDateRaw, pNm + 3)), 4))
If Not IsNumeric(sDay) Or Not IsNumeric(sMonth) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Dang de trong ngay/thang ban hanh ('" & foundDateRaw & "') cho ngay ky duyet."
Else
Dim dInt As Integer, mInt As Integer, yInt As Integer
dInt = Val(sDay): mInt = Val(sMonth): yInt = Val(sYear)
If mInt < 1 Or mInt > 12 Or dInt < 1 Or dInt > 31 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Ngay thang khong hop le: ngay " & dInt & " thang " & mInt & "."
ElseIf (mInt = 4 Or mInt = 6 Or mInt = 9 Or mInt = 11) And dInt > 30 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang " & mInt & " chi co toi da 30 ngay (dang ghi ngay " & dInt & ")."
ElseIf mInt = 2 And dInt > 29 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang 02 khong the co ngay " & dInt & "."
End If
End If
End If
End If
If foundCode = "SY" Or foundCode = "TRS" Or foundCode = "SL" Then foundCode = ""
If Len(foundCode) > 0 Then
Dim expectTitle As String
expectTitle = TT_DocCodeTitle(foundCode)
Select Case foundCode
Case "TTR": foundCode = "TTr"
Case "CTR": foundCode = "CTr"
Case "QYD": foundCode = "QyD"
End Select
If Len(expectTitle) = 0 And foundSoStyle = "NHA_NUOC" And Len(foundTitle) = 0 And hasVvSummary Then expectTitle = "CONG VAN"
If expectTitle = "CONG VAN" Then
If Len(foundTitle) > 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Loai van ban] Ky hieu la Cong van ('CV') nhung lai co tieu de '" & foundTitle & "' o giua bai."
End If
If Not hasVvSummary Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Trich yeu Cong van] Cong van dang thieu dong trich yeu 'V/v...' duoi so van ban."
End If
If Not hasKinhGui Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Kinh gui] Cong van dang thieu muc 'Kinh gui:'."
End If
ElseIf Len(expectTitle) > 0 Then
If Len(foundTitle) = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Thieu tieu de] Ky hieu van ban la '" & foundCode & "' nhung thieu dong ten loai van ban ('" & Replace(expectTitle, "|", " / ") & "') o giua trang."
ElseIf InStr(1, "|" & expectTitle & "|", "|" & foundTitle & "|") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Mau thuan loai VB] Ky hieu ghi la '" & foundCode & "' nhung tieu de van ban lai ghi la '" & foundTitle & "'."
End If
End If
End If
If foundTitle = "QUYET DINH" Or (foundCode = "QD" And Len(foundTitle) = 0) Then
If Not hasQuyetDinhVerb Or totalDieuCount = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Cau truc Quyet dinh] Van ban Quyet dinh nhung thieu dong 'QUYET DINH:' hoac thieu cac 'Dieu 1.', 'Dieu 2.'."
End If
If totalDieuCount > 0 And hasNoiNhan And Not hasNhuDieu Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Logic Noi nhan] Van ban Quyet dinh co Dieu khoan thi hanh nhung phan Noi nhan thieu dong '- Nhu Dieu ...;'."
End If
End If
If TT_B("CHECK_COLLECTIVE") And Len(foundIssuingOrgan) > 0 And (Len(foundSignAuth) > 0 Or Len(foundSignRole) > 0) Then
Dim isExecutive As Boolean: isExecutive = False
If foundIssuingOrgan Like "VAN PHONG*" Or foundIssuingOrgan Like "BAN TO CHUC*" Or _
foundIssuingOrgan Like "BAN TUYEN GIAO*" Or foundIssuingOrgan Like "BAN XAY DUNG*" Or _
foundIssuingOrgan Like "BAN DAN VAN*" Or foundIssuingOrgan Like "TRUNG TAM*" Then
isExecutive = True
End If
Dim isCollective As Boolean: isCollective = False
If Not isExecutive And (InStr(1, foundIssuingOrgan, "DANG UY") > 0 Or InStr(1, foundIssuingOrgan, "BAN CHAP HANH") > 0 Or InStr(1, foundIssuingOrgan, "BAN THUONG VU") > 0 Or InStr(1, foundIssuingOrgan, "UY BAN KIEM TRA") > 0 Or InStr(1, foundIssuingOrgan, "CHI BO") > 0) Then
isCollective = True
End If
If isCollective And Not (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la tap the ('" & foundIssuingOrgan & "') nhung chu ky thieu tham quyen 'T/M ...'."
ElseIf isExecutive And (foundSignAuth Like "T/M *" Or foundSignAuth Like "TM.*") Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la '" & foundIssuingOrgan & "' (che do thu truong) nen khong duoc ky 'T/M' (" & foundSignAuth & ")."
End If
End If
If hasNoiNhan Then
If Not hasLuuVT And Not hasLuuOther Then
warnCount = warnCount + 1
If TT_S("MOTTO_KIND") = "DANG" Then
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong luu o cuoi (vi du: '- Luu Van phong Tinh uy.')."
Else
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong '- Luu: VT...' o cuoi."
End If
ElseIf Not hasLuuVT And hasLuuOther Then
warnCount = warnCount + 1
If TT_S("MOTTO_KIND") = "DANG" Then
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen ghi noi luu (vi du theo HD 05-HD/VPTW: '- Luu Van phong Tinh uy.')."
Else
warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen bo sung ky hieu 'VT' (Van thu) (vi du: '- Luu: VT, ...')."
End If
End If
End If
If TT_S("MOTTO_KIND") = "DANG" And Len(foundIssuingOrgan) > 0 Then
Dim issIsLocalCommittee As Boolean
issIsLocalCommittee = (foundIssuingOrgan Like "DANG UY XA *" Or foundIssuingOrgan Like "DANG UY PHUONG *" Or _
foundIssuingOrgan Like "DANG UY DAC KHU *" Or foundIssuingOrgan Like "CHI BO *")
Dim issIsStaffOrgan As Boolean
issIsStaffOrgan = (foundIssuingOrgan Like "VAN PHONG*" Or foundIssuingOrgan Like "BAN XAY DUNG DANG*" Or _
foundIssuingOrgan Like "BAN TO CHUC*" Or foundIssuingOrgan Like "BAN TUYEN GIAO*" Or _
foundIssuingOrgan Like "BAN DAN VAN*" Or foundIssuingOrgan Like "CO QUAN UY BAN KIEM TRA*")
If (issIsLocalCommittee Or issIsStaffOrgan) And Len(foundSuperiorOrgan) = 0 Then
warnCount = warnCount + 1
If issIsStaffOrgan Then
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] '" & foundIssuingOrgan & "' la co quan tham muu, giup viec: dong tren phai ghi ten cap uy ma co quan do truc thuoc (HD 05-HD/VPTW muc 2.1.3), vi du 'DANG UY XA ...'."
Else
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Van ban cua '" & foundIssuingOrgan & "' phai ghi ten dang bo cap tren truc tiep o dong tren (HD 05-HD/VPTW muc 2.1.2), vi du 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."
End If
ElseIf issIsLocalCommittee And Not (foundIssuingOrgan Like "CHI BO *") Then
If foundSuperiorOrgan Like "HUYEN UY*" Or foundSuperiorOrgan Like "QUAN UY*" Or foundSuperiorOrgan Like "THI UY*" Or _
foundSuperiorOrgan Like "DANG BO HUYEN*" Or foundSuperiorOrgan Like "DANG BO QUAN*" Or foundSuperiorOrgan Like "DANG BO THI XA*" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren dang ghi '" & foundSuperiorOrgan & "' (mo hinh cap huyen cu). Dang bo xa, phuong, dac khu truc thuoc truc tiep Tinh uy, Thanh uy: ghi 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."
ElseIf foundSuperiorOrgan Like "TINH UY*" Or foundSuperiorOrgan Like "THANH UY*" Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren ghi ten DANG BO cap tren truc tiep (vi du 'DANG BO TINH ...'), khong ghi '" & foundSuperiorOrgan & "' (HD 05-HD/VPTW muc 2.1.2)."
End If
End If
End If
If Not isStandaloneNonDispatch Then
If TT_S("MOTTO_KIND") = "DANG" And InStr(1, docAscFull, "DANG CONG SAN VIET NAM") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Tieu ngu] Van ban Dang chua co dong 'DANG CONG SAN VIET NAM' o dau trang."
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" And InStr(1, docAscFull, "CONG HOA XA HOI CHU NGHIA") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Van ban hanh chinh chua co Quoc hieu 'CONG HOA XA HOI CHU NGHIA VIET NAM' va Tieu ngu."
ElseIf TT_S("MOTTO_KIND") = "QUOC_HIEU" And InStr(1, Replace(docAscFull, " ", ""), "DOCLAP-TUDO-HANHPHUC") = 0 Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Thieu dong Tieu ngu 'Doc lap - Tu do - Hanh phuc' duoi Quoc hieu."
End If
End If
If TT_B("RANK_BEFORE_NAME") Then
Dim signerName As String
signerName = TT_FindSignerName(doc)
If Len(signerName) > 0 Then
If Not TT_HasRankPrefix(ToAscii(signerName)) Then
warnCount = warnCount + 1
warnings = warnings & vbCrLf & warnCount & ". [Nguoi ky] Ho ten nguoi ky ('" & ToAscii(signerName) & "') chua ghi cap bac ham phia truoc (vi du: 'Dai ta Nguyen Van A'). Kiem tra lai neu don vi co quy dinh ghi cap bac."
End If
End If
End If
KiemTra_LogicTheThuc = warnings
End Function
Private Function ToAscii(ByVal s As String) As String
Dim n As Long
n = Len(s)
If n = 0 Then
ToAscii = ""
Exit Function
End If
Dim res As String
res = Space$(n)
Dim k As Long
Dim w As Long
For k = 1 To n
w = AscW(Mid$(s, k, 1))
If w < 0 Then w = w + 65536
Select Case w
Case 272, 273
Mid$(res, k, 1) = "D"
Case 192 To 195, 224 To 227, 258, 259, 7840 To 7863
Mid$(res, k, 1) = "A"
Case 200 To 202, 232 To 234, 278, 279, 7864 To 7879
Mid$(res, k, 1) = "E"
Case 204, 205, 236, 237, 296, 297, 7880 To 7883
Mid$(res, k, 1) = "I"
Case 210 To 214, 242 To 246, 416, 417, 7884 To 7907
Mid$(res, k, 1) = "O"
Case 217, 218, 249, 250, 360, 361, 431, 432, 7908 To 7921
Mid$(res, k, 1) = "U"
Case 221, 253, 7922 To 7929
Mid$(res, k, 1) = "Y"
Case 8211, 8212
Mid$(res, k, 1) = "-"
Case Else
If w >= 97 And w <= 122 Then
Mid$(res, k, 1) = Chr$(w - 32)
ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or _
w = 32 Or w = 40 Or w = 41 Or w = 42 Or w = 45 Or w = 46 Or w = 47 Or w = 58 Or w = 59 Then
Mid$(res, k, 1) = Chr$(w)
Else
Mid$(res, k, 1) = " "
End If
End Select
Next k
ToAscii = res
End Function
Private Function IsPartyDocTitle(ByVal sAsc As String) As Boolean
Dim t As String
t = Trim(sAsc)
If Left(t, 10) = "(DU THAO) " Then t = Trim(Mid(t, 11))
If Left(t, 8) = "DU THAO " Then t = Trim(Mid(t, 9))
If Len(t) = 0 Or Len(t) > 45 Then
IsPartyDocTitle = False
Exit Function
End If
Select Case t
Case "NGHI QUYET", "QUYET DINH", "CHI THI", "KET LUAN", "QUY CHE", _
"QUY DINH", "THONG TRI", "HUONG DAN", "THONG BAO", "KE HOACH", _
"CHUONG TRINH", "BAO CAO", "BIEN BAN", "TO TRINH", "CONG DIEN", _
"DE AN", "PHUONG AN", "GIAY MOI", "GIAY GIOI THIEU", _
"BAN CAM KET", "BAN DANG KY", "DIEU LE", "CHUONG TRINH HANH DONG", _
"BAO CAO CHINH TRI", "BAO CAO TOM TAT", "BAO CAO KIEM DIEM", _
"BAO CAO TONG KET", "BAO CAO SO KET", "DE CUONG", "DE CUONG BAO CAO", _
"DE CUONG TUYEN TRUYEN", "QUY CHE LAM VIEC", "QUY CHE HOAT DONG", _
"QUY CHE PHOI HOP", "CHUONG TRINH LAM VIEC", "CHUONG TRINH CONG TAC", _
"CHUONG TRINH KIEM TRA", "PHU LUC", "DANH MUC", "BANG TONG HOP", _
"BIEU TONG HOP", "BAI THU HOACH", "BAI PHAT BIEU", "BAN KIEM DIEM", _
"THU MOI", "THU TRIEU TAP", "LICH LAM VIEC", "LICH CONG TAC", _
"THONG CAO", "TUYEN BO", "LOI KEU GOI", "DU AN", "HOP DONG", _
"BAN GHI NHO", "BAN THOA THUAN", "GIAY UY QUYEN", "GIAY NGHI PHEP", _
"PHIEU GUI", "PHIEU CHUYEN", "PHIEU BAO", "QUY HOACH", "GIAY CHUNG NHAN"
IsPartyDocTitle = True
Case Else
If t Like "PHU LUC *" And Len(t) <= 20 Then
IsPartyDocTitle = True
Else
IsPartyDocTitle = False
End If
End Select
End Function
Private Sub TT_GoBangBoc(ByVal doc As Document)
On Error Resume Next
Dim i As Long, t As Table, soDong As Long, soCot As Long, soCon As Long, soDoan As Long, vong As Long
Dim canGo As Boolean, raw As String, a As String, pp As Paragraph, soDeMuc As Long
For vong = 1 To 3
For i = doc.Tables.Count To 1 Step -1
Set t = Nothing
Set t = doc.Tables(i)
If Not (t Is Nothing) Then
soDong = 0: soCot = 0: soCon = 0: soDoan = 0
soDong = t.Rows.Count
soCot = t.Columns.Count
soCon = t.Tables.Count
soDoan = t.Range.Paragraphs.Count
canGo = False
If soDong = 1 And soCot = 1 And (soCon > 0 Or soDoan > 12) Then canGo = True
If Not canGo And soCot = 1 And soCon = 0 Then
raw = ""
raw = t.Range.text
If i = 1 And soDong <= 4 And Len(raw) <= 600 Then
a = UCase(ToAscii(raw))
If InStr(1, a, "CONG HOA XA HOI") > 0 Or InStr(1, a, "DANG CONG SAN VIET NAM") > 0 Then canGo = True
ElseIf soDong >= 3 And Len(raw) > 300 Then
soDeMuc = 0
For Each pp In t.Range.Paragraphs
raw = CleanParaText(pp.Range.text)
If Len(raw) > 120 Then canGo = True
If Len(raw) > 0 And Len(raw) <= 120 Then
a = Trim(ToAscii(raw))
If IsPartyDocTitle(a) Then canGo = True
If IsStrictBodyOrHeadingStart(raw, a) Then soDeMuc = soDeMuc + 1
End If
If soDeMuc >= 2 Then canGo = True
If canGo Then Exit For
Next pp
End If
End If
If canGo Then
t.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
End If
End If
Next i
Next vong
On Error GoTo 0
End Sub
Private Function TT_LaDongMoDauVanBan(ByVal rawText As String) As Boolean
TT_LaDongMoDauVanBan = False
Dim t As String, a As String, vt As Long
t = Replace(Replace(Replace(rawText, vbCr, ""), Chr(7), ""), Chr(11), vbTab)
If Len(t) < 15 Or Len(t) > 110 Then Exit Function
vt = InStrRev(t, vbTab)
If vt > 0 Then t = Mid(t, vt + 1)
a = UCase(Trim(ToAscii(t)))
If a Like "CONG HOA XA HOI CHU NGHIA VIET NAM*" And Len(a) <= 45 Then TT_LaDongMoDauVanBan = True
If a = "DANG CONG SAN VIET NAM" Then TT_LaDongMoDauVanBan = True
End Function
Private Function TT_TachVanBanTiepTheo(ByVal doc As Document) As Document
Set TT_TachVanBanTiepTheo = Nothing
On Error Resume Next
Dim p As Paragraph, dem As Long, lanDau As Long, viTri As Long, t As String, tb As Table
viTri = -1
For Each p In doc.Paragraphs
dem = dem + 1
t = p.Range.text
If Len(t) >= 15 And Len(t) <= 112 Then
If TT_LaDongMoDauVanBan(t) Then
If lanDau = 0 Then
lanDau = dem
ElseIf dem - lanDau >= 8 Then
viTri = p.Range.Start
If p.Range.Information(wdWithInTable) Then
Set tb = Nothing
Set tb = p.Range.Tables(1)
If Not (tb Is Nothing) Then viTri = tb.Range.Start
End If
Exit For
End If
End If
End If
Next p
If viTri <= 0 Then Exit Function
Dim rng As Range, d2 As Document, doDai As Long
Set rng = Nothing
Set rng = doc.Range(viTri, doc.Content.End - 1)
If rng Is Nothing Then Exit Function
doDai = Len(rng.text)
If doDai < 40 Then Exit Function
Set d2 = Nothing
Set d2 = Documents.Add
If d2 Is Nothing Then Exit Function
d2.Content.FormattedText = rng.FormattedText
If Len(d2.Content.text) < doDai * 0.9 Then
d2.Close 0
Exit Function
End If
rng.Delete
Set TT_TachVanBanTiepTheo = d2
On Error GoTo 0
End Function
Private Sub TT_GhepVanBanTiepTheo(ByVal doc As Document, ByVal d2 As Document)
On Error Resume Next
Dim truoc As Long, them As Long, r As Range
doc.Activate
truoc = Len(doc.Content.text)
them = Len(d2.Content.text)
Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
If InStr(1, Right(doc.Content.text, 4), Chr(12)) > 0 Then
r.InsertBreak wdSectionBreakContinuous
Else
r.InsertBreak wdSectionBreakNextPage
End If
Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
r.FormattedText = d2.Content.FormattedText
If Len(doc.Content.text) >= truoc + them * 0.9 Then
d2.Close 0
Else
MsgBox "Van ban thu hai trong file da duoc chuan hoa nhung khong ghep lai duoc." & vbCrLf & _
"Noi dung dang nam o cua so van ban moi, hay chep thu cong vao cuoi van ban.", vbExclamation, "Chuan hoa the thuc"
End If
doc.Activate
On Error GoTo 0
End Sub
Private Sub TT_GoHopVanBanDauTrang(ByVal doc As Document)
On Error Resume Next
If doc.Shapes.Count = 0 Then Exit Sub
Dim gioiHan As Long, n As Long
n = doc.Paragraphs.Count
If n > 25 Then n = 25
gioiHan = doc.Paragraphs(n).Range.End
Dim i As Long, shp As Shape, t As String, a As String, r As Range, viTri As Long
For i = doc.Shapes.Count To 1 Step -1
Set shp = Nothing
Set shp = doc.Shapes(i)
If Not (shp Is Nothing) Then
t = ""
t = shp.TextFrame.TextRange.text
Do While Len(t) > 0 And (Right(t, 1) = vbCr Or Right(t, 1) = vbLf Or Right(t, 1) = " ")
t = Left(t, Len(t) - 1)
Loop
If Len(Trim(t)) >= 3 And Len(t) <= 400 Then
viTri = -1
viTri = shp.Anchor.Start
If viTri >= 0 And viTri < gioiHan Then
a = UCase(ToAscii(t))
If InStr(1, a, "CONG HOA XA HOI") > 0 Or InStr(1, a, "DOC LAP") > 0 Or InStr(1, a, "DANG CONG SAN") > 0 Or _
IsHeaderOrganLineAsc(Trim(Split(a, vbCr)(0))) Or IsDocNumberLineAsc(Trim(Split(a, vbCr)(0))) Or _
a Like "*NGAY*THANG*NAM*" Then
Set r = Nothing
Set r = shp.Anchor.Paragraphs(1).Range
If Not (r Is Nothing) Then
shp.Delete
r.InsertBefore t & vbCr
End If
End If
End If
End If
End If
Next i
On Error GoTo 0
End Sub
Private Function TT_SpacesToTab(ByVal s As String) As String
Dim t As String
t = Replace(s, ChrW(160), " ")
If InStr(1, t, "     ") = 0 Then
TT_SpacesToTab = s
Exit Function
End If
Do While InStr(1, t, "      ") > 0
t = Replace(t, "      ", "     ")
Loop
t = Replace(t, "     ", vbTab)
Do While InStr(1, t, vbTab & " ") > 0
t = Replace(t, vbTab & " ", vbTab)
Loop
Do While InStr(1, t, " " & vbTab) > 0
t = Replace(t, " " & vbTab, vbTab)
Loop
TT_SpacesToTab = t
End Function
Private Sub TT_ChuanKinhGui(ByVal doc As Document)
On Error Resume Next
Dim p As Paragraph, dem As Long, t As String, a As String, vt As Long, r As Range
Set p = Nothing
Set p = doc.Paragraphs(1)
Do While Not (p Is Nothing)
dem = dem + 1
If dem > 120 Then Exit Do
t = Left(p.Range.text, 20)
If Len(t) >= 8 Then
a = UCase(ToAscii(t))
If LTrim(a) Like "KINH GOI*" Or LTrim(a) Like "DONG KINH GOI*" Then
vt = InStr(1, t, ChrW(&H1EDF))
If vt > 0 Then
Set r = doc.Range(p.Range.Start + vt - 1, p.Range.Start + vt)
If r.text = ChrW(&H1EDF) Then r.text = ChrW(&H1EED)
Else
vt = InStr(1, t, ChrW(&H1EDE))
If vt > 0 Then
Set r = doc.Range(p.Range.Start + vt - 1, p.Range.Start + vt)
If r.text = ChrW(&H1EDE) Then r.text = ChrW(&H1EEC)
End If
End If
End If
End If
Set p = p.Next
Loop
On Error GoTo 0
End Sub
Private Function TT_StartsWithDocType(ByVal sAsc As String) As Boolean
TT_StartsWithDocType = False
Dim w() As String, k As Long, dau As String
w = Split(Trim(sAsc), " ")
For k = 0 To UBound(w)
If k > 3 Then Exit For
If k = 0 Then dau = w(0) Else dau = dau & " " & w(k)
If k >= 1 Then
If IsPartyDocTitle(dau) Then
TT_StartsWithDocType = True
Exit Function
End If
End If
Next k
End Function
Private Function TT_IsLooseDocTitle(ByVal p As Paragraph, ByVal rawText As String, ByVal sAsc As String, ByRef tenLoaiTran As Boolean) As Boolean
TT_IsLooseDocTitle = False
tenLoaiTran = False
Dim r As String, a As String
r = Trim(rawText)
a = UCase(Trim(sAsc))
If Len(r) < 4 Or Len(r) > 150 Then Exit Function
If Right(r, 1) = "." Or Right(r, 1) = ";" Or Right(r, 1) = ":" Or Right(r, 1) = "," Then Exit Function
If IsStrictBodyOrHeadingStart(rawText, sAsc) Or IsKinhGuiStart(rawText, sAsc) Or IsSeparatorLine(rawText) Then Exit Function
If IsDocNumberLineAsc(a) Or a Like "V/V*" Or a Like "V.V*" Or a Like "#*" Or a Like "[A-H]. *" Then Exit Function
If a Like "CONG HOA XA HOI*" Or a Like "DOC LAP*" Or a Like "DANG CONG SAN*" Or a Like "*NGAY*THANG*NAM*" Then Exit Function
If IsHeaderOrganLineAsc(a) Or a Like "NOI NHAN*" Or IsSignatureRoleOrAuth(a) Then Exit Function
Dim giua As Boolean, dam As Boolean
On Error Resume Next
giua = (p.Format.Alignment = wdAlignParagraphCenter)
dam = (p.Range.Font.Bold = True)
On Error GoTo 0
Dim kq As Boolean
If UCase(r) = r And a Like "*[A-Z][A-Z]*" Then
kq = True
ElseIf Len(r) <= 120 And (giua Or dam) Then
If TT_StartsWithDocType(a) Then kq = True
End If
If kq Then tenLoaiTran = (Len(r) <= 25)
TT_IsLooseDocTitle = kq
End Function
Private Function TT_FindSignatureEnd(ByVal doc As Document) As Long
TT_FindSignatureEnd = 0
On Error Resume Next
Dim tbl As Table, raw As String, a As String, idx As Long, soDong As Long, duPhong As Long
For Each tbl In doc.Tables
idx = idx + 1
soDong = 99
soDong = tbl.Rows.Count
raw = ""
If soDong <= 4 Then raw = tbl.Range.text
If Len(raw) > 0 And Len(raw) <= 3000 Then
a = UCase(ToAscii(raw))
If InStr(1, a, "NOI NHAN") > 0 Then
TT_FindSignatureEnd = tbl.Range.End
Exit Function
End If
If duPhong = 0 And idx > 1 And soDong <= 2 Then
If HasSignatureKeywordInCell(a) Then duPhong = tbl.Range.End
End If
End If
Next tbl
TT_FindSignatureEnd = duPhong
On Error GoTo 0
End Function
Private Function TT_IsAppendixLabel(ByVal rawText As String, ByVal sAsc As String, Optional ByVal sauChuKy As Boolean = False, Optional ByVal dangCanGiua As Boolean = False) As Boolean
TT_IsAppendixLabel = False
Dim t As String, gon As String, r As String
t = UCase(Trim(sAsc))
r = Trim(rawText)
gon = Replace(t, " ", "")
If Left(gon, 6) <> "PHULUC" Then Exit Function
If Len(t) > 200 Then Exit Function
If sauChuKy Or dangCanGiua Then
If Len(r) <= 150 And Right(r, 1) <> ";" And Not (Right(r, 1) = "." And Len(r) > 40) Then TT_IsAppendixLabel = True
Exit Function
End If
Do While Len(t) > 0 And (Right(t, 1) = ":" Or Right(t, 1) = ".")
t = Trim(Left(t, Len(t) - 1))
Loop
If t = "PHU LUC" Or gon = "PHULUC" Then
TT_IsAppendixLabel = True
Exit Function
End If
Dim rest As String
rest = Trim(Mid(gon, 7))
If Left(t, 8) = "PHU LUC " Then rest = Trim(Mid(t, 9))
If Left(rest, 3) = "SO " Then rest = Trim(Mid(rest, 4))
If Len(rest) = 0 Then Exit Function
Dim tok As String, i As Long, ch As String, soKyTu As Long
For i = 1 To Len(rest)
ch = Mid(rest, i, 1)
If InStr(1, "0123456789IVXLC", ch) > 0 Then
tok = tok & ch
Else
Exit For
End If
Next i
soKyTu = Len(tok)
If soKyTu = 0 Or soKyTu > 6 Then Exit Function
Dim sau As String
sau = Mid(rest, soKyTu + 1)
If Len(sau) > 0 Then
If Left(sau, 1) Like "[A-Z]" And (Len(sau) = 1 Or Not (Mid(sau, 2, 1) Like "[A-Z0-9]")) Then sau = Mid(sau, 2)
End If
sau = Trim(sau)
If Len(sau) = 0 Then
TT_IsAppendixLabel = True
Exit Function
End If
If Not (Left(sau, 1) Like "[A-Z0-9]") Then
If UCase(r) = r And Len(r) <= 200 Then TT_IsAppendixLabel = True
End If
End Function
Private Function TT_IsAppendixTitleLine(ByVal p As Paragraph, ByVal rawText As String, ByVal sAsc As String) As Boolean
TT_IsAppendixTitleLine = False
Dim r As String, a As String
r = Trim(rawText)
a = UCase(Trim(sAsc))
If Len(r) = 0 Or Len(r) > 250 Then Exit Function
If Right(r, 1) = "." Or Right(r, 1) = ";" Or Right(r, 1) = ":" Then Exit Function
If IsStrictBodyOrHeadingStart(rawText, sAsc) Then Exit Function
If a Like "BANG #*" Or a Like "BIEU #*" Or a Like "BIEU SO*" Or a Like "MAU #*" Or a Like "MAU SO*" Or a Like "#*" Then Exit Function
Dim giua As Boolean, dam As Boolean
On Error Resume Next
giua = (p.Format.Alignment = wdAlignParagraphCenter)
dam = (p.Range.Font.Bold = True)
On Error GoTo 0
If UCase(r) = r Or giua Or dam Then TT_IsAppendixTitleLine = True
End Function
Private Function NormalizeDocTitleForCompare(ByVal sAsc As String) As String
Dim t As String
t = Trim(sAsc)
If Left(t, 10) = "(DU THAO) " Then t = Trim(Mid(t, 11))
If Left(t, 8) = "DU THAO " Then t = Trim(Mid(t, 9))
If t Like "BAO CAO *" Then t = "BAO CAO"
If t Like "CHUONG TRINH *" Then t = "CHUONG TRINH"
If t Like "QUY CHE *" Then t = "QUY CHE"
If t Like "DE CUONG *" Then t = "DE CUONG"
If t Like "PHU LUC *" Then t = "PHU LUC"
NormalizeDocTitleForCompare = t
End Function
Private Function IsSeparatorLine(ByVal rawText As String) As Boolean
Dim cleanStr As String
cleanStr = Trim(Replace(Replace(Replace(rawText, vbCr, ""), vbLf, ""), " ", ""))
cleanStr = Replace(cleanStr, Chr(160), "")
cleanStr = Replace(cleanStr, vbTab, "")
If Len(cleanStr) < 2 Or Len(cleanStr) > 45 Then
IsSeparatorLine = False
Exit Function
End If
Dim isAllDash As Boolean: isAllDash = True
Dim i As Long
For i = 1 To Len(cleanStr)
Dim wCode As Long
wCode = AscW(Mid(cleanStr, i, 1))
If wCode < 0 Then wCode = wCode + 65536
If wCode <> 45 And wCode <> 95 And wCode <> 42 And _
wCode <> 8211 And wCode <> 8212 And wCode <> 8213 Then
isAllDash = False
Exit For
End If
Next i
IsSeparatorLine = isAllDash
End Function
Private Function TT_ChuThuongPhanChuong(ByVal p As Paragraph, ByVal s As String, ByVal sAsc As String) As Boolean
TT_ChuThuongPhanChuong = False
On Error Resume Next
Dim tk() As String
tk = Split(Trim(sAsc), " ")
If UBound(tk) < 1 Then Exit Function
Dim nTu As Integer
If tk(0) = "PHAN" And tk(1) = "THU" Then
If UBound(tk) < 2 Then Exit Function
nTu = 3
Else
nTu = 1
End If
Dim i As Integer, dai As Long
dai = 0
For i = 0 To nTu - 1
dai = dai + Len(tk(i)) + IIf(i > 0, 1, 0)
Next i
Dim r As Range
Set r = p.Range.Duplicate
If dai > 1 Then
r.SetRange p.Range.Start + 1, p.Range.Start + dai
r.Case = wdLowerCase
End If
r.SetRange p.Range.Start, p.Range.Start + 1
r.Case = wdUpperCase
Dim conLai As Integer
conLai = UBound(tk) - nTu
If nTu = 1 Then conLai = conLai - 1
TT_ChuThuongPhanChuong = (conLai <= 0)
End Function
Private Function IsStrictBodyOrHeadingStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
Dim tAsc As String: tAsc = Trim(sAsc)
Dim tRaw As String: tRaw = Trim(rawText)
If Len(tAsc) = 0 Then
IsStrictBodyOrHeadingStart = False
Exit Function
End If
If tAsc Like "CAN CU *" Or tAsc Like "- CAN CU *" Or _
tAsc Like "XET DE NGHI*" Or tAsc Like "THEO DE NGHI*" Or _
tAsc = "QUYET DINH:" Or tAsc = "QUYET DINH :" Or _
tAsc Like "I.*" Or tAsc Like "II.*" Or tAsc Like "III.*" Or _
tAsc Like "I-*" Or tAsc Like "II-*" Or tAsc Like "III-*" Or _
tAsc Like "PHAN I*" Or tAsc Like "PHAN 1*" Or tAsc Like "PHAN THU*" Or _
tAsc Like "MUC I*" Or tAsc Like "MUC 1*" Or _
tAsc Like "CHUONG I*" Or tAsc Like "CHUONG 1*" Or _
tAsc Like "DIEU 1*" Or tAsc Like "DIEU 2*" Or _
tAsc Like "1. *" Or tAsc Like "2. *" Then
IsStrictBodyOrHeadingStart = True
ElseIf Left(tRaw, 1) = "-" Or Left(tRaw, 1) = "+" Then
IsStrictBodyOrHeadingStart = True
Else
IsStrictBodyOrHeadingStart = False
End If
End Function
Private Function IsBodyStart(ByVal rawText As String, ByVal sAsc As String) As Boolean
If IsStrictBodyOrHeadingStart(rawText, sAsc) Then
IsBodyStart = True
Exit Function
End If
Dim tAsc As String: tAsc = Trim(sAsc)
Dim tRaw As String: tRaw = Trim(rawText)
If tAsc Like "THUC HIEN*" Or tAsc Like "DE THUC HIEN*" Or _
tAsc Like "BAN THUONG VU*" Or tAsc Like "BAN CHAP HANH*" Or _
tAsc Like "DANG UY*" Or tAsc Like "CHI BO*" Or _
tAsc Like "HOI NGHI*" Or tAsc Like "DAI HOI*" Or _
tAsc Like "YEU CAU:*" Or tAsc Like "YEU CAU :*" Then
IsBodyStart = True
Exit Function
End If
Dim lastCh As String: lastCh = Right(tRaw, 1)
If lastCh = ";" Or lastCh = ":" Or Len(tRaw) > 130 Then
IsBodyStart = True
Else
IsBodyStart = False
End If
End Function
Private Function LicApiUrl() As String
LicApiUrl = "https://script.google.com/macros/s/AKfycbxSZyf-7FeVZ6xbhBwuT5aVVGJDdqP8AJCYek_Wp3Ila5yWLGnmq_rIua3kL1dFsvzOug/exec"
End Function
Private Function CheckLicenseActivation() As Boolean
Dim appName As String, sectionName As String
Dim savedKey As String, currentMachine As String
Dim lastCheckDate As String
Dim apiUrl As String
appName = "OfficePartySuite"
sectionName = "License"
apiUrl = LicApiUrl()
On Error Resume Next
savedKey = GetSetting(appName, sectionName, "ActiveKey", "")
lastCheckDate = GetSetting(appName, sectionName, "LastCheckDate", "")
On Error GoTo 0
If Len(savedKey) > 0 Then
If Len(lastCheckDate) > 0 Then
Dim isRecent As Boolean: isRecent = False
On Error Resume Next
If IsDate(lastCheckDate) Then
If Abs(DateDiff("d", CDate(lastCheckDate), Now)) < 7 Then
isRecent = True
End If
End If
On Error GoTo 0
If isRecent Then
CheckLicenseActivation = True
Exit Function
End If
End If
currentMachine = GetMachineName()
Dim checkUrl As String
checkUrl = apiUrl & "?key=" & EncodeUrl(savedKey) & "&machine=" & EncodeUrl(currentMachine)
Dim checkResp As String
checkResp = CallHttpApi(checkUrl)
If Len(checkResp) > 0 Then
Dim checkStatus As String
checkStatus = ParseJsonField(checkResp, "status")
If checkStatus = "ERROR" Then
SaveSetting appName, sectionName, "ActiveKey", ""
SaveSetting appName, sectionName, "ActiveMachine", ""
SaveSetting appName, sectionName, "LastCheckDate", ""
Dim errReason As String
errReason = ParseJsonField(checkResp, "message")
If Len(errReason) = 0 Then errReason = "Ban quyen cua ban da bi khoa boi VPDU xa Thanh Phu!"
MsgBox "CANH BAO BAN QUYEN:" & vbCrLf & errReason & vbCrLf & vbCrLf & _
"Tien ich da bi dung hoat dong tren may nay.", vbCritical, "Ban Quyen Bi Khoa"
CheckLicenseActivation = False
Exit Function
Else
On Error Resume Next
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
On Error GoTo 0
End If
End If
CheckLicenseActivation = True
Exit Function
End If
currentMachine = GetMachineName()
If Not LIC_CAN_NHAP_MA Then
If GetSetting(appName, sectionName, "AutoTryDate", "") = Format(Now, "yyyy-mm-dd") Then
CheckLicenseActivation = True
Exit Function
End If
Dim autoResp As String
autoResp = CallHttpApi(apiUrl & "?machine=" & EncodeUrl(currentMachine))
If ParseJsonField(autoResp, "khoa") = "1" Then
MsgBox "CANH BAO BAN QUYEN:" & vbCrLf & ParseJsonField(autoResp, "message"), vbCritical, "Ban Quyen Add-in"
CheckLicenseActivation = False
Exit Function
End If
If ParseJsonField(autoResp, "status") = "SUCCESS" And Len(ParseJsonField(autoResp, "key")) > 0 Then
SaveSetting appName, sectionName, "ActiveKey", ParseJsonField(autoResp, "key")
SaveSetting appName, sectionName, "ActiveMachine", currentMachine
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
Else
SaveSetting appName, sectionName, "AutoTryDate", Format(Now, "yyyy-mm-dd")
End If
CheckLicenseActivation = True
Exit Function
End If
Dim userInput As String
userInput = InputBox("Vui long nhap Ma kich hoat ban quyen de su dung tien ich:" & vbCrLf & _
"(Moi may tinh chi can nhap kich hoat 1 lan duy nhat)" & vbCrLf & vbCrLf & _
"Ten may tinh: " & currentMachine & vbCrLf & _
"Don vi phat trien: Van phong Dang uy xa Thanh Phu", _
"Kich Hoat Ban Quyen Add-in")
userInput = Trim(UCase(userInput))
If Len(userInput) = 0 Then
MsgBox "Ban chua nhap ma kich hoat! Khong the thuc hien chuan hoa.", vbExclamation, "Ban Quyen Add-in"
CheckLicenseActivation = False
Exit Function
End If
Dim actUrl As String
actUrl = apiUrl & "?key=" & EncodeUrl(userInput) & "&machine=" & EncodeUrl(currentMachine)
Dim actResp As String
actResp = CallHttpApi(actUrl)
If Len(actResp) = 0 Then
MsgBox "Khong the ket noi den may chu xac thuc ban quyen!" & vbCrLf & _
"Vui long kiem tra ket noi Internet va thu lai.", vbExclamation, "Loi Ket Noi"
CheckLicenseActivation = False
Exit Function
End If
Dim actStatus As String, actMsg As String
actStatus = ParseJsonField(actResp, "status")
actMsg = ParseJsonField(actResp, "message")
If actStatus = "SUCCESS" Then
SaveSetting appName, sectionName, "ActiveKey", userInput
SaveSetting appName, sectionName, "ActiveMachine", currentMachine
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
MsgBox "Kich hoat ban quyen thanh cong!" & vbCrLf & _
"May tinh: " & currentMachine & vbCrLf & _
"Ban co the su dung tien ich vinh vien tren may nay.", vbInformation, "Kich Hoat Thanh Cong"
CheckLicenseActivation = True
Else
MsgBox "Kich hoat that bai!" & vbCrLf & actMsg & vbCrLf & vbCrLf & _
"Vui long lien he Van phong Dang uy xa Thanh Phu de duoc cap ma hop le.", vbCritical, "Kich Hoat That Bai"
CheckLicenseActivation = False
End If
End Function
Private Function CallHttpApi(ByVal url As String) As String
Dim res As String
res = ""
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
End If
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 2000, 3000, 3000, 3000
http.Send
If Err.Number = 0 Then
res = http.responseText
End If
On Error GoTo 0
CallHttpApi = Trim(res)
End Function
Private Function GetMachineName() As String
Dim mName As String
mName = Environ("COMPUTERNAME")
If Len(mName) = 0 Then mName = Environ("USERNAME")
If Len(mName) = 0 Then mName = "UNKNOWN_PC"
GetMachineName = mName
End Function
Private Function ParseJsonField(ByVal json As String, ByVal fieldName As String) As String
Dim pos As Long, startPos As Long, endPos As Long
pos = InStr(json, """" & fieldName & """")
If pos > 0 Then
startPos = InStr(pos + Len(fieldName) + 2, json, """")
If startPos > 0 Then
endPos = InStr(startPos + 1, json, """")
If endPos > startPos Then
ParseJsonField = Mid(json, startPos + 1, endPos - startPos - 1)
End If
End If
End If
End Function
Private Function EncodeUrl(ByVal s As String) As String
Dim i As Long, ch As String, code As Long
Dim res As String
res = ""
For i = 1 To Len(s)
ch = Mid(s, i, 1)
code = AscW(ch) And &HFFFF&
If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
res = res & ch
ElseIf code < 128 Then
res = res & "%" & Right("0" & Hex(code), 2)
Else
res = res & "_"
End If
Next i
EncodeUrl = res
End Function
Public Sub AutoExec()
Call CN_HenKiemTraKhiMo
End Sub
Public Function CN_PhienBan() As String
CN_PhienBan = "2026.11.12.2"
End Function
Private Function CN_ManifestUrl() As String
CN_ManifestUrl = "https://github.com/CuongBC195/add-in-office/releases/latest/download/version.json"
End Function
Private Function CN_SoSanh(ByVal a As String, ByVal b As String) As Integer
Dim pa() As String, pb() As String
pa = Split(a, ".")
pb = Split(b, ".")
Dim i As Integer, va As Long, vb As Long
For i = 0 To 5
va = 0: vb = 0
If i <= UBound(pa) Then va = Val(pa(i))
If i <= UBound(pb) Then vb = Val(pb(i))
If va > vb Then
CN_SoSanh = 1
Exit Function
ElseIf va < vb Then
CN_SoSanh = -1
Exit Function
End If
Next i
CN_SoSanh = 0
End Function
Private Function CN_FileTuChay() As String
On Error Resume Next
CN_FileTuChay = ThisDocument.FullName
On Error GoTo 0
End Function
Private Function CN_DungFileAddIn() As Boolean
Dim n As String
n = LCase(Mid(CN_FileTuChay(), InStrRev(Replace(CN_FileTuChay(), "\", "/"), "/") + 1))
CN_DungFileAddIn = (n = "chuanhoa_vanbandang.dotm")
End Function
Private Function CN_ShQ(ByVal s As String) As String
CN_ShQ = "'" & Replace(s, "'", "'\''") & "'"
End Function
Private Function CN_TaiManifest(Optional ByVal nhanh As Boolean = False) As String
Dim i As Integer, res As String, soLan As Integer, cho As Integer
soLan = IIf(nhanh, 1, 2)
cho = IIf(nhanh, 3, 15)
For i = 1 To soLan
res = ""
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", CN_ManifestUrl(), False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 5000, cho * 1000, cho * 1000
http.Send
If Err.Number = 0 Then
If http.Status = 200 Then res = http.responseText
End If
On Error GoTo 0
res = Trim(res)
If Len(ParseJsonField(res, "version")) > 0 Then Exit For
Next i
CN_TaiManifest = res
End Function
Public Sub RibbonCapNhat(control As Object)
Call CN_KiemTraCapNhat(True)
End Sub
Public Sub CN_HenKiemTraKhiMo()
Static daHen As Boolean
If daHen Then Exit Sub
daHen = True
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 20), "CN_KiemTraNhanhKhiMo"
End Sub
Public Sub CN_KiemTraNhanhKhiMo()
Call CN_TuKiemTraHangTuan(True)
End Sub
Public Sub CN_TuKiemTraHangTuan(Optional ByVal nhanh As Boolean = False)
On Error Resume Next
Dim lastChk As String, daDocDuoc As Boolean
lastChk = GetSetting("OfficePartySuite", "CapNhat", "LanKiemTra_WORD", "")
If Len(lastChk) > 0 Then
If IsDate(lastChk) Then
If DateDiff("d", CDate(lastChk), Date) < 1 Then Exit Sub
End If
End If
Dim lanThu As String
lanThu = GetSetting("OfficePartySuite", "CapNhat", "LanThu_WORD", "")
If Len(lanThu) > 0 Then
If IsDate(lanThu) Then
If DateDiff("n", CDate(lanThu), Now) < 60 Then Exit Sub
End If
End If
SaveSetting "OfficePartySuite", "CapNhat", "LanThu_WORD", Format(Now, "yyyy-mm-dd hh:nn:ss")
On Error GoTo 0
Call CN_KiemTraCapNhat(False, nhanh, daDocDuoc)
If daDocDuoc Then SaveSetting "OfficePartySuite", "CapNhat", "LanKiemTra_WORD", Format(Date, "yyyy-mm-dd")
End Sub
Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
Dim js As String
js = CN_TaiManifest(nhanh)
Dim ver As String, url As String, sha As String, notes As String, kt As Long
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "word_url")
sha = LCase(ParseJsonField(js, "word_sha256"))
notes = ParseJsonField(js, "notes_ascii")
kt = Val(ParseJsonField(js, "word_size"))
If Len(ParseJsonField(js, "word_url_win")) > 0 Then
url = ParseJsonField(js, "word_url_win")
sha = LCase(ParseJsonField(js, "word_sha256_win"))
kt = Val(ParseJsonField(js, "word_size_win"))
End If
If Len(ver) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Then
If hoiKhiKhongCo Then MsgBox "Khong ket noi duoc may chu cap nhat (GitHub). Kiem tra mang roi thu lai." & vbCrLf & "Phien ban dang dung: " & CN_PhienBan(), vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
daDocDuoc = True
If CN_SoSanh(ver, CN_PhienBan()) <= 0 Then
If hoiKhiKhongCo Then MsgBox "Ban dang dung phien ban moi nhat: " & CN_PhienBan(), vbInformation, "Cap nhat Add-in"
Exit Sub
End If
If Not CN_DungFileAddIn() Then
MsgBox "Co phien ban moi " & ver & " nhung macro dang chay tu '" & CN_FileTuChay() & "', khong phai file add-in da cai." & vbCrLf & _
"Hay cai lai bang bo cai dat.", vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
Dim ans As VbMsgBoxResult
ans = MsgBox("Co phien ban moi: " & ver & " (dang dung " & CN_PhienBan() & ")." & vbCrLf & vbCrLf & notes & vbCrLf & vbCrLf & "Cap nhat ngay?", vbYesNo + vbQuestion, "Cap nhat Add-in")
If ans <> vbYes Then Exit Sub
Dim msg As String
msg = CN_TaiVaCai(url, sha, kt)
If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cap nhat Add-in"
End Sub
Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String, Optional ByVal kichThuoc As Long = 0) As String
Dim target As String
target = CN_FileTuChay()
On Error GoTo WinFail
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 10000, 15000, 60000, 180000
http.Send
If http.Status <> 200 Then GoTo WinFail
Dim b() As Byte, n As Long
b = http.responseBody
n = UBound(b) - LBound(b) + 1
If n < 10240 Then GoTo WinBad
If b(LBound(b)) <> 80 Or b(LBound(b) + 1) <> 75 Then GoTo WinBad
If kichThuoc > 0 And n <> kichThuoc Then GoTo WinBad
Dim thuMuc As String
thuMuc = Environ("APPDATA") & "\OfficePartySuite"
If Len(Dir(thuMuc, vbDirectory)) = 0 Then MkDir thuMuc
Dim moiW As String
moiW = thuMuc & "\capnhat_WORD.moi"
If Len(Dir(moiW)) > 0 Then Kill moiW
Dim f As Integer
f = FreeFile
Open moiW For Binary Access Write As #f
Put #f, 1, b
Close #f
If FileLen(moiW) <> n Then GoTo WinBad
Dim tgtW As String
tgtW = target
If Len(Environ("APPDATA")) > 0 And LCase(Left(tgtW, Len(Environ("APPDATA")))) = LCase(Environ("APPDATA")) Then
tgtW = "%APPDATA%" & Mid(tgtW, Len(Environ("APPDATA")) + 1)
End If
Dim cmdPath As String
cmdPath = thuMuc & "\capnhat_WORD.cmd"
f = FreeFile
Open cmdPath For Output As #f
Print #f, "@echo off"
Print #f, ":cho"
Print #f, "tasklist /FI ""IMAGENAME eq WINWORD.EXE"" 2>nul | find /I ""WINWORD.EXE"" >nul"
Print #f, "if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto cho)"
Print #f, "copy /Y ""%APPDATA%\OfficePartySuite\capnhat_WORD.moi"" """ & tgtW & """ >nul"
Print #f, "del ""%APPDATA%\OfficePartySuite\capnhat_WORD.moi"" >nul 2>&1"
Close #f
Shell "cmd.exe /c """ & cmdPath & """", vbMinimizedNoFocus
CN_TaiVaCai = "Da tai ban moi. Hay DONG HET cua so Word: ban moi se tu cai dat, mo lai Word la dung duoc."
Exit Function
WinBad:
On Error Resume Next
Close #f
If Len(moiW) > 0 Then Kill moiW
CN_TaiVaCai = "File tai ve khong dung (thieu hoac hong). Chua thay doi gi, hay thu lai sau."
Exit Function
WinFail:
CN_TaiVaCai = "Tai ban cap nhat that bai (kiem tra mang). Chua thay doi gi."
End Function
Public Sub CN_KiemThuCapNhat()
Dim h As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim js As String
js = CN_TaiManifest()
Dim ver As String, url As String, sha As String, res As String
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "word_url")
sha = LCase(ParseJsonField(js, "word_sha256"))
res = "file=" & CN_FileTuChay() & " | dungFile=" & CN_DungFileAddIn() & " | dangDung=" & CN_PhienBan() & " | banMoi=" & ver
If Len(ver) > 0 And CN_SoSanh(ver, CN_PhienBan()) > 0 And CN_DungFileAddIn() Then res = res & " | " & CN_TaiVaCai(url, sha)
Dim f As Integer
f = FreeFile
Open h & "/cn_test_WORD.log" For Output As #f
Print #f, res
Close #f
End Sub
