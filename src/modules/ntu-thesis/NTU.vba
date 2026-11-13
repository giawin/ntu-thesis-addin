Attribute VB_Name = "NTU"
Option Explicit

' ===================================================================
'  Chuan hoa van ban theo quy cach Truong Dai hoc Nha Trang
'  Quy dinh trinh bay luan van, do an, de an tot nghiep thac si
'  (Quyet dinh so 900/QD-DHNT ngay 31/7/2023)
'
'  Tom tat quy cach dang ap dung:
'    - Kho A4 210 x 297 mm, chieu doc, in mot mat
'    - Le: tren 2 cm, duoi 2 cm, trai 3 cm, phai 2 cm
'    - Times New Roman co 13 (co dinh), can deu 2 ben (justified)
'    - Gian dong 1,5 lines; dong dau moi doan thut 1 cm
'    - Cach doan: 3-6 pt phia tren, 0 pt phia duoi
'    - KHONG dung header; so trang o giua phia DUOI, footer cach 1 cm
'    - Tieu de chuong: IN HOA, dam, can giua, co 14
'    - Tieu de muc/tieu muc: chu thuong, dam, can deu, khong thut, co 13
'    - Bang/hinh danh so theo chuong; tieu de bang TREN, hinh DUOI
'
'  LUU Y KHI SUA FILE NAY:
'  - File duoc luu UTF-8 de doc cho de, nhung luc build se doi moi
'    chuoi trong dau nhay kep thanh dang ASCII + ChrW() vi VBA luu
'    code theo codepage ANSI (tieng Viet se bi mat dau neu ghi truc tiep).
'  - Comment nen viet khong dau; comment co dau cung khong anh huong.
' ===================================================================

Private Const APP As String = "NTUChuanHoa"
Private Const MAX_DEPTH As Long = 4
Private Const FONT_NAME As String = "Times New Roman"

' Buoc dang chay — de bao loi biet chet o dau (xem NTU_Apply).
Private lastStep As String

' So caption da chen trong lan ApplyCaptions gan nhat (doc lai o NTU_ChuanHoaFull).
Private LastCaptionCount As Long

' Cua so nhap (Document tam) dung cho RibbonPhanTich — giu tham chieu de dong lai.
Private LastScratch As Document

' Ket qua luot kiem tra gan nhat: so loi, tai lieu goc, va vung cua tung loi
' (de nhap so la nhay toi).
Private LastDemLoi As Long
Private LastGoc As Document
Private LastChoLoi() As Range

' ---------------------------------------------------------------- Vet loi
' Ghi tung buoc ra %TEMP%\NTUChuanHoa.log de biet chet o dau.
' Khong bao gio nem loi: add-in hong phai im lang.

Public Sub Trace(ByVal s As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\NTUChuanHoa.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
End Sub

' ---------------------------------------------------------------- Ribbon

Public Sub RibbonLoad(ribbon As IRibbonUI)
    ' Word chi can ham ton tai cho onLoad; ghi trace de biet ribbon da nap.
    Trace "RibbonLoad: ribbon da nap"
    ' Hen kiem tra ban moi mot lan moi khi mo Word (chay ngam sau 20 giay).
    Call NTU_HenKiemTraKhiMo
End Sub

Public Sub RibbonHuongDan(control As IRibbonControl)
    Dim s As String
    s = "QUY CÁCH TRÌNH BÀY — TRƯỜNG ĐẠI HỌC NHA TRANG" & vbCrLf & _
        "(Quyết định số 900/QĐ-ĐHNT ngày 31/7/2023)" & vbCrLf & String(46, "-") & vbCrLf & _
        "Khổ giấy: A4 (210 × 297 mm), chiều dọc, in một mặt" & vbCrLf & _
        "Lề: trên 2 · dưới 2 · trái 3 · phải 2 cm" & vbCrLf & _
        "Phông: Times New Roman, cỡ 13; căn đều 2 bên" & vbCrLf & _
        "Giãn dòng 1,5 · dòng đầu mỗi đoạn thụt 1 cm" & vbCrLf & _
        "Cách đoạn: 3–6 pt phía trên, 0 pt phía dưới" & vbCrLf & _
        "Không dùng header · số trang ở giữa phía DƯỚI" & vbCrLf & vbCrLf & _
        "Tiêu đề chương: IN HOA, in đậm, căn giữa, cỡ 14." & vbCrLf & _
        "Tiêu đề mục/tiểu mục: chữ thường, in đậm, căn đều, không thụt, cỡ 13." & vbCrLf & _
        "Tiểu mục tối đa 4 cấp; mỗi nhóm phải có ít nhất 2 mục." & vbCrLf & _
        "Bảng/hình: số gắn với chương (Hình 3.4). Đầu đề bảng ở TRÊN, hình ở DƯỚI." & vbCrLf & _
        "Trích dẫn: ngoặc vuông theo số thứ tự, ví dụ [15, tr.32-35]." & vbCrLf & vbCrLf & _
        "Công cụ tự sửa: khổ giấy, lề, phông, cỡ chữ, căn lề, thụt dòng, giãn dòng, " & _
        "cách đoạn, header, số trang, gạch dài." & vbCrLf & _
        "Phải tự sửa: số bảng/hình theo chương, nhóm chỉ 1 tiểu mục, ""bảng dưới đây"", " & _
        "danh mục tài liệu tham khảo, bảng chữ viết tắt."
    MsgBox s, vbInformation, "Chuẩn NTU"
End Sub

' Nut [Chuan hoa luan an]: 1-CLICK du 5 buoc (QD 900 muc 2.b).
Public Sub RibbonChuanHoa(control As IRibbonControl)
    If Application.Documents.Count = 0 Then
        MsgBox "Chưa có tài liệu nào đang mở.", vbExclamation, "Chuẩn NTU"
        Exit Sub
    End If

    Dim d As Document
    Set d = ActiveDocument

    Dim msg As String
    msg = "Chuẩn hóa tài liệu theo quy cách Trường Đại học Nha Trang" & vbCrLf & _
          "(QĐ 900/QĐ-ĐHNT ngày 31/7/2023):" & vbCrLf & vbCrLf & _
          "1. Khổ A4 dọc · lề trên 2 · dưới 2 · trái 3 · phải 2 cm" & vbCrLf & _
          "2. Thân bài Times New Roman 13 pt, căn đều, giãn 1,5, thụt dòng đầu, cách đoạn 6/0 pt" & vbCrLf & _
          "3. Tiêu đề chương in hoa đậm 14 pt căn giữa; mục/tiểu mục đậm 13 pt" & vbCrLf & _
          "4. Caption ""Bảng X.Y"" trên bảng, ""Hình X.Y"" dưới hình, kèm dòng Nguồn" & vbCrLf & _
          "5. Số trang La Mã ở phần đầu, Ả Rập từ chương 1 · bỏ header" & vbCrLf & vbCrLf & _
          "Thao tác này sửa trực tiếp tài liệu đang mở (Ctrl+Z hoàn tác được). Tiếp tục?"
    If MsgBox(msg, vbOKCancel + vbQuestion, "Chuẩn NTU") <> vbOK Then Exit Sub

    Dim rep As String
    rep = NTU_ChuanHoaFull(d, OptSize(), OptKeepHeading(), True)

    ' Buoc 5: tu soat lai roi tom tat chung mot hop thoai.
    Dim kq As String
    kq = NTU_KiemTraSilent()
    MsgBox rep & vbCrLf & "----- Tự kiểm tra -----" & vbCrLf & kq, vbInformation, "Chuẩn NTU"
End Sub

' Ten goi thay the cho nut [Phan tich] (giu de tuong thich voi manifest cu).
Public Sub RibbonPhuLuc(control As Object)
    RibbonPhanTich control
End Sub

Public Sub RibbonKiemTra(control As IRibbonControl)
    NTU_KiemTra
End Sub

Public Sub RibbonTuyChon(control As IRibbonControl)
    NTU_TuyChon
End Sub

' ---------------------------------------------------------------- Tuy chon

Private Function OptSize() As Single
    ' QD 900/2023: co chu 13, khong con "13 hoac 14".
    OptSize = 13
End Function

Private Function OptKeepHeading() As Boolean
    OptKeepHeading = (GetSetting(APP, "Opt", "KeepHeading", "1") <> "0")
End Function

Public Sub NTU_TuyChon()
    Dim s As String

    s = InputBox("Giữ nguyên cỡ chữ của tiêu đề đã đặt tay (1 = có, 0 = đưa tất cả về 13)." & vbCrLf & _
                 "Quy cách quy định: tiêu đề chương 14, mục/tiểu mục 13.", "Chuẩn NTU — tùy chọn", IIf(OptKeepHeading(), "1", "0"))
    If StrPtr(s) = 0 Then Exit Sub
    SaveSetting APP, "Opt", "KeepHeading", IIf(Trim$(s) = "0", "0", "1")

    MsgBox "Đã lưu: thân bài 13 pt · tiêu đề " & _
           IIf(OptKeepHeading(), "giữ cỡ riêng", "đưa về 13 pt") & ".", vbInformation, "Chuẩn NTU"
End Sub

' ---------------------------------------------------------------- Vao viec

Public Sub NTU_ChuanHoa()
    Dim d As Document
    If Application.Documents.Count = 0 Then
        MsgBox "Chưa có tài liệu nào đang mở.", vbExclamation, "Chuẩn NTU"
        Exit Sub
    End If
    Set d = ActiveDocument

    Dim msg As String
    msg = "Chuẩn hóa tài liệu theo quy cách Trường Đại học Nha Trang" & vbCrLf & _
          "(QĐ 900/QĐ-ĐHNT ngày 31/7/2023):" & vbCrLf & vbCrLf & _
          "• Times New Roman 13 pt, căn đều 2 bên, giãn dòng 1,5" & vbCrLf & _
          "• Dòng đầu mỗi đoạn thụt 1 cm; cách đoạn 6 pt trên, 0 pt dưới" & vbCrLf & _
          "• A4 · lề trên 2 · dưới 2 · trái 3 · phải 2 cm" & vbCrLf & _
          "• Bỏ header · số trang ở giữa phía DƯỚI mỗi trang" & vbCrLf & _
          "• Gạch dài (— – ‒ ―) đổi thành gạch nối (-)" & vbCrLf & vbCrLf & _
          "Thao tác này sửa trực tiếp tài liệu đang mở (Ctrl+Z hoàn tác được). Tiếp tục?"
    If MsgBox(msg, vbOKCancel + vbQuestion, "Chuẩn NTU") <> vbOK Then Exit Sub

    Dim rep As String
    rep = NTU_Apply(d, OptSize(), OptKeepHeading(), True)
    MsgBox rep, vbInformation, "Chuẩn NTU"
End Sub

' Entry point khong hoi giac (dung cho kiem thu tu dong)
Public Function NTU_ChuanHoaSilent() As String
    On Error GoTo Loi
    If Application.Documents.Count = 0 Then
        NTU_ChuanHoaSilent = "KHONG CO TAI LIEU"
        Exit Function
    End If
    NTU_ChuanHoaSilent = NTU_Apply(ActiveDocument, OptSize(), OptKeepHeading(), False)
    Exit Function
Loi:
    NTU_ChuanHoaSilent = "LOI: " & Err.Description
End Function

' --------------------------------------------------------- Loi chinh

Public Function NTU_Apply(ByVal d As Document, ByVal sz As Single, ByVal keepH As Boolean, ByVal allowUndo As Boolean) As String
    Dim rep As String
    Dim ur As UndoRecord
    Dim n As Long

    lastStep = "bat dau"
    Trace "== NTU_Apply sz=" & sz & " keepH=" & keepH
    On Error GoTo Loi

    If allowUndo Then
        On Error Resume Next
        Set ur = Application.UndoRecord
        If Not ur Is Nothing Then ur.StartCustomRecord "Chuan NTU"
        On Error GoTo Loi
    End If

    Application.ScreenUpdating = False

    ' Moi buoc chay doc lap: mot buoc hong khong duoc huy cac buoc sau.
    ' (VBA khong co try/catch nen dung On Error Resume Next + kiem tra Err.)

    ' 1. Kho giay + le
    lastStep = "le"
    On Error Resume Next
    Err.Clear
    OptimizePageSetup d
    If Err.Number <> 0 Then
        rep = rep & "! Lề / khổ giấy: " & Err.Description & vbCrLf
        Trace "step le ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "✓ Khổ A4 · lề trên 2 · dưới 2 · trái 3 · phải 2 cm" & vbCrLf
    End If
    On Error GoTo Loi

    ' 2. Phong chu + co chu + gian dong
    lastStep = "chu"
    On Error Resume Next
    Err.Clear
    n = NormalizeBody(d, sz, keepH)
    If Err.Number <> 0 Then
        rep = rep & "! Phông / cỡ chữ: " & Err.Description & vbCrLf
        Trace "step chu ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "✓ Đã định dạng " & Format$(n, "#,##0") & " đoạn" & _
          IIf(keepH, " (giữ cỡ chữ của tiêu đề)", "") & vbCrLf
    On Error GoTo Loi

    ' 3. Kieu Normal de go tiep cung dung chuan
    lastStep = "style"
    On Error Resume Next
    Err.Clear
    SetNormalStyle d, sz
    If Err.Number <> 0 Then
        rep = rep & "! Kiểu Normal: " & Err.Description & vbCrLf
        Trace "step style ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo Loi

    ' 4. So trang
    lastStep = "sotrang"
    On Error Resume Next
    Err.Clear
    Trace "  -> sotrang"
    Dim s4 As String
    s4 = SetPageNumberFooter(d)
    Trace "  <- sotrang err=" & Err.Number & " " & Err.Description
    rep = rep & s4 & vbCrLf
    If Err.Number <> 0 Then
        rep = rep & "! Số trang: " & Err.Description & vbCrLf
        Trace "step sotrang ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo Loi

    ' 5. Gach dai
    lastStep = "gachdai"
    On Error Resume Next
    Err.Clear
    Trace "  -> gachdai"
    n = FixDashes(d, rep)
    Trace "  <- gachdai err=" & Err.Number & " " & Err.Description
    If Err.Number <> 0 Then
        rep = rep & "! Gạch dài: " & Err.Description & vbCrLf
        Trace "step gachdai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    ElseIf n > 0 Then
        rep = rep & "✓ Đã đổi " & Format$(n, "#,##0") & " dấu gạch dài thành gạch nối (-)" & vbCrLf
    Else
        rep = rep & "✓ Không có dấu gạch dài nào cần đổi" & vbCrLf
    End If
    On Error GoTo Loi

    Application.ScreenUpdating = True
    Trace "  -> ScreenUpdating ok"

    If allowUndo Then
        On Error Resume Next
        ur.EndCustomRecord
        On Error GoTo Loi
    End If

    Trace "== xong"
    NTU_Apply = rep & vbCrLf & "Xong. Bấm [Kiểm Tra] để soát lại."
    Exit Function
Loi:
    Trace "FATAL at " & lastStep & ": " & Err.Description
    Application.ScreenUpdating = True
    NTU_Apply = rep & "! Lỗi ở bước [" & lastStep & "]: " & Err.Description
End Function

' --------------------------------------------------------- Tung buoc

Private Function Cm(ByVal v As Double) As Single
    Cm = CentimetersToPoints(v)
End Function

Private Sub OptimizePageSetup(ByVal d As Document)
    Dim i As Long
    For i = 1 To d.Sections.Count
        With d.Sections(i).PageSetup
            .PaperSize = wdPaperA4
            .Orientation = wdOrientPortrait
            .PageWidth = Cm(21#)
            .PageHeight = Cm(29.7)
            .TopMargin = Cm(2#)
            .BottomMargin = Cm(2#)
            .LeftMargin = Cm(3#)
            .RightMargin = Cm(2#)
            .Gutter = 0
            .MirrorMargins = False
            .HeaderDistance = Cm(1#)
            .FooterDistance = Cm(1#)
        End With
    Next
End Sub

Private Sub SetNormalStyle(ByVal d As Document, ByVal sz As Single)
    On Error Resume Next
    With d.Styles(wdStyleNormal)
        .Font.Name = FONT_NAME
        .Font.NameFarEast = FONT_NAME
        .Font.Size = sz
        .Font.Spacing = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .ParagraphFormat.FirstLineIndent = Cm(1#)
        .ParagraphFormat.SpaceBefore = 6
        .ParagraphFormat.SpaceAfter = 0
    End With
    On Error GoTo 0
End Sub

Private Function NormalizeBody(ByVal d As Document, ByVal sz As Single, ByVal keepH As Boolean) As Long
    Dim n As Long, sec As Section, hf As HeaderFooter

    n = NormalizeRange(d.Content, sz, keepH)

    For Each sec In d.Sections
        For Each hf In sec.Headers
            If Not hf.LinkToPrevious Then n = n + NormalizeRange(hf.Range, sz, keepH)
        Next
        For Each hf In sec.Footers
            If Not hf.LinkToPrevious Then n = n + NormalizeRange(hf.Range, sz, keepH)
        Next
    Next

    NormalizeBody = n
End Function

Private Function NormalizeRange(ByVal r As Range, ByVal sz As Single, ByVal keepH As Boolean) As Long
    Dim p As Paragraph, n As Long, bad As Long, firstErr As String
    Dim inTable As Boolean, lvl As Long

    ' 1. Phong chu cho ca range mot lan — nhanh va khong phu thuoc tung doan.
    On Error Resume Next
    Err.Clear
    SetRunFont r
    If Err.Number <> 0 Then
        Trace "SetRunFont range ERR " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' 2. Dinh dang tung doan. Mot doan loi (o bang gop, field, textbox)
    '    khong duoc phep huy ca tai lieu — dem vao bad roi di tiep.
    '
    '    QD 900/2023:
    '      - than bai: co 13, gian dong 1,5, can deu 2 ben, thut dong dau 1 cm,
    '        cach doan 6 pt tren / 0 pt duoi
    '      - trong bang: gian dong 1 line, khong thut dong dau (bang co quy dinh rieng)
    '      - tieu de chuong: IN HOA, dam, can giua, co 14, khong thut dong
    '      - tieu de muc/tieu muc: dam, can deu, co 13, khong thut dong
    For Each p In r.Paragraphs
        On Error Resume Next
        Err.Clear
        inTable = False
        inTable = p.Range.Information(wdWithInTable)
        Err.Clear
        lvl = 0
        If keepH Then lvl = HeadingLevel(p)
        Err.Clear

        With p.Range.ParagraphFormat
            If inTable Then
                .LineSpacingRule = wdLineSpaceSingle
                .SpaceBefore = 0
            Else
                .LineSpacingRule = wdLineSpace1pt5
                .SpaceBefore = 6
            End If
            .SpaceAfter = 0

            If lvl = 1 Then
                .Alignment = wdAlignParagraphCenter
                .FirstLineIndent = 0
            ElseIf lvl > 1 Then
                .Alignment = wdAlignParagraphJustify
                .FirstLineIndent = 0
            ElseIf inTable Then
                .Alignment = wdAlignParagraphJustify
                .FirstLineIndent = 0
            Else
                .Alignment = wdAlignParagraphJustify
                .FirstLineIndent = Cm(1#)
            End If
        End With

        If lvl = 1 Then
            p.Range.Font.Size = 14
            p.Range.Font.Bold = True
            p.Range.Font.AllCaps = True
        ElseIf lvl > 1 Then
            p.Range.Font.Size = sz
            p.Range.Font.Bold = True
        Else
            p.Range.Font.Size = sz
        End If

        If Err.Number <> 0 Then
            bad = bad + 1
            If firstErr = "" Then firstErr = Err.Number & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        n = n + 1
    Next

    If bad > 0 Then Trace "NormalizeRange: " & bad & "/" & n & " doan loi, dau tien " & firstErr
    NormalizeRange = n
End Function

Private Sub SetRunFont(ByVal r As Range)
    ' NameFarEast bi tu choi (loi 5844) tren mot so tai lieu — bo qua, khong lam hong ca buoc.
    On Error Resume Next
    Err.Clear
    With r.Font
        .Name = FONT_NAME
        .NameAscii = FONT_NAME
        .NameOther = FONT_NAME
        .Spacing = 0
        .Kerning = 0
        .NameFarEast = FONT_NAME
    End With
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Function IsHeadingLike(ByVal p As Paragraph) As Boolean
    IsHeadingLike = (HeadingLevel(p) > 0)
End Function

' Cap cua tieu de: 1 = tieu de chuong, 2..9 = muc/tieu muc, 0 = than bai.
' Nhan dien theo OutlineLevel (Word tu dat khi dung style Heading), va theo
' ten style / so hieu dau dong khi tac gia go tay.
Private Function HeadingLevel(ByVal p As Paragraph) As Long
    Dim lvl As Long, nm As String, t As String

    On Error Resume Next
    Err.Clear
    lvl = p.Range.ParagraphFormat.OutlineLevel
    If Err.Number <> 0 Then
        Err.Clear
        lvl = wdOutlineLevelBodyText
    End If
    If lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9 Then
        HeadingLevel = lvl - wdOutlineLevel1 + 1
        Err.Clear
        Exit Function
    End If

    nm = p.Style.NameLocal
    Err.Clear
    If Err.Number = 0 Then
        nm = LCase$(nm)
        If InStr(nm, "heading 1") > 0 Or InStr(nm, "tiêu đề 1") > 0 Then HeadingLevel = 1 : Err.Clear : Exit Function
        If InStr(nm, "heading") > 0 Or InStr(nm, "tiêu đề") > 0 Then HeadingLevel = 2 : Err.Clear : Exit Function
    End If
    Err.Clear

    ' "Chương 1" / "CHƯƠNG 1" go tay
    t = CleanText(p.Range.Text)
    Err.Clear
    If Len(t) <= 60 Then
        If Left$(LCase$(t), 6) = "chương" Or Left$(LCase$(t), 5) = "chuong" Then HeadingLevel = 1 : Err.Clear : Exit Function
    End If
    Err.Clear

    HeadingLevel = 0
    On Error GoTo 0
End Function

' Dau ket thuc o (Chr(13) & Chr(7)) va dau ket thuc bang KHONG phai doan van that:
' Word khong cho dat ParagraphFormat tren chung, nen kiem tra chung se bao sai.
Private Function IsTableMark(ByVal p As Paragraph) As Boolean
    Dim t As String, i As Long, c As Integer, digits As Long

    On Error Resume Next
    t = p.Range.Text
    If Err.Number <> 0 Then
        Err.Clear
        IsTableMark = False
        Exit Function
    End If
    On Error GoTo 0

    IsTableMark = False
    If InStr(t, Chr(7)) = 0 Then Exit Function

    ' Chi co ky tu dieu khien/dau cach/so thu tu dong — tuc la khong co chu that.
    For i = 1 To Len(t)
        c = Asc(Mid(t, i, 1))
        Select Case c
            Case 13, 7, 32, 9, 160
            Case 48 To 57
                digits = digits + 1
            Case Else
                Exit Function
        End Select
    Next
    IsTableMark = True
End Function

' QD 900/2023:
'   - KHONG dung ghi dau trang (header) -> xoa sach header
'   - so trang o GIUA phia DUOI moi trang, footer dat cach mep duoi 1 cm
Private Function SetPageNumberFooter(ByVal d As Document) As String
    Dim sec As Section, hf As HeaderFooter
    Dim r As Range, txt As String
    Dim done As Long

    For Each sec In d.Sections
        ' 1. Xoa header: quy cach khong cho dung.
        For Each hf In sec.Headers
            If Not hf.LinkToPrevious Then
                On Error Resume Next
                Err.Clear
                hf.Range.Text = ""
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        Next

        ' 2. Dat so trang o footer, can giua.
        For Each hf In sec.Footers
            If Not hf.LinkToPrevious Then
                On Error Resume Next
                Err.Clear
                DelFields hf.Range, wdFieldPage
                DelFields hf.Range, wdFieldNumPages
                txt = CleanText(hf.Range.Text)
                ' Footer con chu khac (ten truong, ten de tai...) thi giu, chi chen them dong.
                If txt <> "" And Not IsPageNumberish(txt) Then
                    hf.Range.InsertParagraphAfter
                End If
                Set r = hf.Range
                r.Text = ""
                Set r = hf.Range          ' lay lai: range cu da co lai sau khi xoa
                r.Fields.Add r, wdFieldPage
                With r.Paragraphs(1)
                    .Alignment = wdAlignParagraphCenter
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceSingle
                End With
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                done = done + 1
            End If
        Next
    Next

    ' Cap nhat field de footer hien so dung ngay (khong thi hien 0)
    On Error Resume Next
    d.Fields.Update
    d.Repaginate
    Err.Clear
    On Error GoTo 0

    SetPageNumberFooter = "✓ Số trang: đặt ở giữa phía dưới (" & done & " chân trang) · đã bỏ header"
End Function

Private Sub DelFields(ByVal r As Range, ByVal typ As Long)
    Dim i As Long
    For i = r.Fields.Count To 1 Step -1
        If r.Fields(i).Type = typ Then r.Fields(i).Delete
    Next
End Sub

Private Function CleanText(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, Chr$(7), " ")
    t = Replace(t, Chr$(12), " ")
    t = Replace(t, vbTab, " ")
    CleanText = Trim$(t)
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long, a As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        a = AscW(Mid$(s, i, 1))
        If a < 48 Or a > 57 Then Exit Function
    Next
    IsDigits = True
End Function

' Footer chi con so trang (vd "Trang  / 1", "1", "3 / 10") thi coi la so trang
Private Function IsPageNumberish(ByVal s As String) As Boolean
    Dim i As Long, a As Long, seen As Boolean, t As String
    If Len(s) = 0 Then IsPageNumberish = True : Exit Function
    t = LCase$(s)
    t = Replace(t, "trang", " ")
    t = Replace(t, "page", " ")
    t = Replace(t, "tr.", " ")
    For i = 1 To Len(t)
        a = AscW(Mid$(t, i, 1))
        If (a >= 48 And a <= 57) Then
            seen = True
        ElseIf a = 32 Or a = 47 Or a = 45 Or a = 46 Or a = 58 Then  ' space / - . :
            ' bo qua
        Else
            Exit Function
        End If
    Next
    ' Di het vong lap = chi con so/dau phan cach (hoac chu "trang"/"page" da thay bang space)
    ' => chinh la so trang. Khong co so (vd "Trang  / " sau khi xoa field) van phai xoa.
    IsPageNumberish = True
End Function

' --------------------------------------------------------- Gach dai

Private Function FixDashes(ByVal d As Document, ByRef rep As String) As Long
    Dim sec As Section, hf As HeaderFooter, i As Long, r As Range

    FixDashes = 0
    ReplaceDashesIn d.Content, FixDashes

    For Each sec In d.Sections
        For Each hf In sec.Headers
            If Not hf.LinkToPrevious Then ReplaceDashesIn hf.Range, FixDashes
        Next
        For Each hf In sec.Footers
            If Not hf.LinkToPrevious Then ReplaceDashesIn hf.Range, FixDashes
        Next
    Next

    ' Chu thich / chu thich cuoi / textbox. Duyet NGUOC va lay lai range moi vong:
    ' vua doc vua Find-Replace tren StoryRanges se lam lech collection (va co the treo).
    ' So 2,3,4,5 = footnotes, endnotes, comments, text frames.
    For i = d.StoryRanges.Count To 1 Step -1
        Set r = Nothing
        On Error Resume Next
        Set r = d.StoryRanges(i)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not r Is Nothing Then
            Select Case r.StoryType
                Case 2, 3, 4, 5
                    On Error Resume Next
                    ReplaceDashesIn r, FixDashes
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
            End Select
        End If
    Next
End Function

Private Sub ReplaceDashesIn(ByVal r As Range, ByRef count As Long)
    Dim ch As Variant, t As String, k As Long

    For Each ch In Array(ChrW(&H2014), ChrW(&H2013), ChrW(&H2012), ChrW(&H2015))
        t = r.Text
        If Len(t) > 0 Then
            k = UBound(Split(t, CStr(ch)))
            If k > 0 Then
                count = count + k
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = CStr(ch)
                    .Replacement.Text = "-"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
    Next
End Sub

' --------------------------------------------------------- Kiem tra

' Nut [Kiem tra]: bao cao chi tiet trong cua so rieng, danh so de nhay toi loi,
' va xuat duoc ra tep %TEMP%\NTUChuanHoa_KiemTra.txt (UTF-16).
Public Sub NTU_KiemTra()
    Dim d As Document
    If Application.Documents.Count = 0 Then
        MsgBox "Chưa có tài liệu nào đang mở.", vbExclamation, "Chuẩn NTU"
        Exit Sub
    End If
    Set d = ActiveDocument

    Dim rep As String
    rep = NTU_CheckDayDu(d, OptSize())

    ' Hien bao cao trong cua so rieng (khong phai MsgBox: bao cao dai, can cuon).
    Dim w As Document
    Set w = MoCuaSoNhap()
    If Not w Is Nothing Then
        w.Content.Text = rep
        w.Content.Font.Name = FONT_NAME
        w.Content.Font.Size = 10
    Else
        MsgBox rep, vbInformation, "Chuẩn NTU — kết quả kiểm tra"
    End If

    XuatTxt rep
    NhayToiLoi
End Sub

' Hoi so thu tu roi nhay toi dung cho loi do. Lap lai duoc nhieu lan, Cancel de thoat.
Private Sub NhayToiLoi()
    Dim s As String, k As Long
    If LastDemLoi = 0 Then Exit Sub
    If LastGoc Is Nothing Then Exit Sub

    Do
        s = InputBox("Nhập số thứ tự lỗi cần xem (1 - " & LastDemLoi & ")." & vbCrLf & _
                     "Bỏ trống hoặc Cancel để thôi.", "Chuẩn NTU — nhảy tới lỗi")
        If Len(Trim$(s)) = 0 Then Exit Do
        If IsNumeric(s) Then
            k = CLng(Val(s))
            If k >= 1 And k <= LastDemLoi Then
                On Error Resume Next
                LastGoc.Activate
                LastChoLoi(k).Select
                If Err.Number <> 0 Then
                    MsgBox "Không nhảy tới được mục này (tài liệu đã thay đổi?).", vbExclamation, "Chuẩn NTU"
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Loop
End Sub

Public Function NTU_KiemTraSilent() As String
    On Error GoTo Loi
    If Application.Documents.Count = 0 Then
        NTU_KiemTraSilent = "KHONG CO TAI LIEU"
        Exit Function
    End If
    NTU_KiemTraSilent = NTU_Check(ActiveDocument, OptSize())
    Exit Function
Loi:
    NTU_KiemTraSilent = "LOI: " & Err.Description
End Function

Private Sub AddL(ByRef rep As String, ByVal s As String)
    rep = rep & vbCrLf & s
End Sub

Public Function NTU_Check(ByVal d As Document, ByVal sz As Single) As String
    Dim rep As String
    Dim p As Paragraph
    Dim badFont As Long, badSize As Long, badLine As Long
    Dim badAlign As Long, badIndent As Long, badHead As Long, badHeadAlign As Long
    Dim n As Long, lvl As Long, t As String, line As String
    Dim firstFont As String, firstSize As String
    Dim headCount As Long, dashCount As Long
    Dim deepCount As Long, refCount As Long
    Dim footerOk As Boolean

    On Error GoTo Loi
    Application.ScreenUpdating = False

    If PageSetupOk(d) Then
        rep = "Lề / khổ giấy: đúng"
    Else
        rep = "Lề / khổ giấy: SAI — cần trên 2 · dưới 2 · trái 3 · phải 2 cm, khổ A4 dọc"
    End If

    For Each p In d.Content.Paragraphs
        If Not IsTableMark(p) Then
            n = n + 1
            t = p.Range.Text
            If p.Range.Font.Name <> FONT_NAME Then
                If p.Range.Font.Name <> "" Then
                    badFont = badFont + 1
                    If firstFont = "" Then firstFont = Snippet(t)
                End If
            End If
            lvl = HeadingLevel(p)
            If lvl = 0 Then
                If p.Range.Information(wdWithInTable) Then
                    If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Then badLine = badLine + 1
                Else
                    If Abs(p.Range.Font.Size - sz) > 0.5 Then
                        badSize = badSize + 1
                        If firstSize = "" Then firstSize = Snippet(t)
                    End If
                    If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 Then badLine = badLine + 1
                    If p.Range.ParagraphFormat.Alignment <> wdAlignParagraphJustify Then badAlign = badAlign + 1
                    If Abs(p.Range.ParagraphFormat.FirstLineIndent - CentimetersToPoints(1)) > 3 Then badIndent = badIndent + 1
                End If
            ElseIf lvl = 1 Then
                headCount = headCount + 1
                If Abs(p.Range.Font.Size - 14) > 0.5 Then badHead = badHead + 1
                If p.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then badHeadAlign = badHeadAlign + 1
            Else
                headCount = headCount + 1
                If Abs(p.Range.Font.Size - sz) > 0.5 Then badHead = badHead + 1
            End If
        End If
    Next p

    If badFont > 0 Then
        line = "Phông chữ khác Times New Roman: " & badFont
        If firstFont <> "" Then line = line & " (vd: " & firstFont & ")"
        AddL rep, line
    End If
    If badSize > 0 Then
        line = "Cỡ chữ khác " & sz & " pt: " & badSize
        If firstSize <> "" Then line = line & " (vd: " & firstSize & ")"
        AddL rep, line
    End If
    If badLine > 0 Then AddL rep, "Đoạn chưa dãn dòng đúng (thân bài 1,5 · trong bảng 1): " & badLine
    If badAlign > 0 Then AddL rep, "Đoạn chưa căn đều 2 bên: " & badAlign
    If badIndent > 0 Then AddL rep, "Đoạn chưa thụt dòng đầu 1 cm: " & badIndent
    If badHead > 0 Then AddL rep, "Tiêu đề sai cỡ chữ (chương 14 · mục 13): " & badHead
    If badHeadAlign > 0 Then AddL rep, "Tiêu đề chương chưa căn giữa: " & badHeadAlign

    footerOk = HasPageNumberBottom(d)
    If Not footerOk Then AddL rep, "Số trang: chưa đặt ở giữa phía dưới"

    If CountHeadersWithText(d) > 0 Then AddL rep, "Header còn nội dung (quy cách không dùng header)"

    dashCount = CountDashes(d)
    If dashCount > 0 Then AddL rep, "Còn dấu gạch dài: " & dashCount & " (dùng gạch nối -)"

    deepCount = CountDeepHeadings(d)
    If deepCount > 0 Then AddL rep, "Mục quá 4 cấp: " & deepCount

    refCount = CountBadRefs(d)
    If refCount > 0 Then AddL rep, "Chỗ viết ""bảng/hình dưới đây"" (phải nêu số): " & refCount

    rep = rep & vbCrLf & String(58, "-")
    If badFont = 0 And badSize = 0 And badLine = 0 And badAlign = 0 And badIndent = 0 _
       And badHead = 0 And badHeadAlign = 0 And footerOk And dashCount = 0 _
       And CountHeadersWithText(d) = 0 Then
        rep = rep & vbCrLf & "ĐẠT thể thức. Còn phải tự soát: đánh số bảng/hình theo chương, danh mục tài liệu tham khảo, số trang La Mã ở phần đầu."
    Else
        rep = rep & vbCrLf & "Bấm [Chuẩn Hóa] để sửa các mục thể thức ở trên; mục nội dung phải sửa tay."
    End If
    If deepCount > 0 Or refCount > 0 Then
        rep = rep & vbCrLf & "Còn " & (deepCount + refCount) & " lỗi nội dung — không tự sửa được."
    End If

    AddL rep, ""
    rep = "KẾT QUẢ KIỂM TRA - chuẩn NTU (QĐ 900/QĐ-ĐHNT 31/7/2023)" & rep
    rep = rep & vbCrLf & "(đã soát " & n & " đoạn, " & headCount & " tiêu đề)"
    Application.ScreenUpdating = True
    NTU_Check = rep
    Exit Function
Loi:
    Application.ScreenUpdating = True
    NTU_Check = "Lỗi khi kiểm tra: " & Err.Description
End Function

Private Function PageSetupOk(ByVal d As Document) As Boolean
    Dim ok As Boolean, i As Long
    ok = True
    For i = 1 To d.Sections.Count
        With d.Sections(i).PageSetup
            If Abs(.PageWidth - Cm(21#)) > 2 Then ok = False
            If Abs(.PageHeight - Cm(29.7)) > 2 Then ok = False
            If Abs(.TopMargin - Cm(2#)) > 2 Then ok = False
            If Abs(.BottomMargin - Cm(2#)) > 2 Then ok = False
            If Abs(.LeftMargin - Cm(3#)) > 2 Then ok = False
            If Abs(.RightMargin - Cm(2#)) > 2 Then ok = False
        End With
    Next
    PageSetupOk = ok
End Function

' So trang phai o GIUA phia DUOI (QD 900/2023 muc 7.a).
Private Function HasPageNumberBottom(ByVal d As Document) As Boolean
    Dim sec As Section, hf As HeaderFooter, i As Long
    For Each sec In d.Sections
        For Each hf In sec.Footers
            For i = 1 To hf.Range.Fields.Count
                If hf.Range.Fields(i).Type = wdFieldPage Then
                    HasPageNumberBottom = True
                    Exit Function
                End If
            Next
        Next
    Next
End Function

' Quy cach KHONG dung header (muc 6.a). Tra ve so header con chu.
Private Function CountHeadersWithText(ByVal d As Document) As Long
    Dim sec As Section, hf As HeaderFooter, t As String
    For Each sec In d.Sections
        For Each hf In sec.Headers
            If Not hf.LinkToPrevious Then
                t = CleanText(hf.Range.Text)
                If t <> "" Then CountHeadersWithText = CountHeadersWithText + 1
            End If
        Next
    Next
End Function

Private Function CountDashes(ByVal r As Range) As Long
    Dim t As String, ch As Variant, k As Long
    t = r.Text
    For Each ch In Array(ChrW(&H2014), ChrW(&H2013), ChrW(&H2012), ChrW(&H2015))
        k = UBound(Split(t, CStr(ch)))
        CountDashes = CountDashes + k
    Next
End Function

Private Function CountDeepHeadings(ByVal d As Document) As Long
    Dim p As Paragraph, num As String, re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "^\s*(\d+(?:\.\d+)*)"
    For Each p In d.Content.Paragraphs
        If IsHeadingLike(p) Then
            num = ""
            On Error Resume Next
            num = p.Range.ListFormat.ListString
            On Error GoTo 0
            If Trim$(num) = "" Then num = p.Range.Text
            If re.Test(num) Then
                Set m = re.Execute(num)(0)
                If (Len(m.SubMatches(0)) - Len(Replace(m.SubMatches(0), ".", ""))) + 1 > MAX_DEPTH Then
                    CountDeepHeadings = CountDeepHeadings + 1
                End If
            End If
        End If
    Next
End Function

Private Function CountBadRefs(ByVal d As Document) As Long
    Dim p As Paragraph, txt As String
    Dim a As String, b As String
    a = "bảng"
    b = "dưới đây"
    For Each p In d.Content.Paragraphs
        txt = LCase$(p.Range.Text)
        If InStr(txt, a) > 0 And InStr(txt, b) > 0 Then CountBadRefs = CountBadRefs + 1
    Next
End Function

Private Function Snippet(ByVal s As String) As String
    Dim t As String
    t = CleanText(s)
    If Len(t) > 50 Then t = Left$(t, 47) & "..."
    Snippet = t
End Function

' ===================================================================
'  PHAN MO RONG v2
'  - 1-click 5 buoc: kho/le + than bai + tieu de + caption + so trang
'  - Phan tich (chi doc), kiem tra mo rong, xuat bao cao .txt
'  - Cap nhat qua version.json (doc lap, khong dung Word core)
' ===================================================================

' -------------------------------------------------------- Tuy chon mo rong

' Co tu in hoa tieu de chuong hay khong (mac dinh bat).
Private Function OptUpperChapter() As Boolean
    OptUpperChapter = (GetSetting(APP, "Opt", "UpperChapter", "1") <> "0")
End Function

' Thut dong dau: 1,0 cm theo QD 900; cho phep 1,25 cm neu nguoi dung chon.
Private Function OptIndent() As Double
    Dim s As String, v As Double
    s = Trim$(GetSetting(APP, "Opt", "Indent", "1.00"))
    If Len(s) = 0 Then
        OptIndent = 1#
        Exit Function
    End If
    v = Val(Replace(s, ",", "."))
    If v < 0.5 Or v > 2 Then v = 1#
    OptIndent = v
End Function

' -------------------------------------------------------- 1-click 5 buoc

' Duong chay day du: moi buoc chay doc lap (mot buoc hong khong huy cac buoc sau).
' ponytail: one-click all 5 steps; tach thanh task pane tung buoc khi nguoi dung can.
Public Function NTU_ChuanHoaFull(ByVal d As Document, ByVal sz As Single, ByVal keepH As Boolean, ByVal allowUndo As Boolean) As String
    Dim rep As String
    Dim ur As UndoRecord
    Dim n As Long, soSec As Long
    Dim olds As Boolean
    Dim oldal As WdAlertLevel
    Dim em As String

    lastStep = "bat dau"
    LastCaptionCount = 0
    Trace "== NTU_ChuanHoaFull sz=" & sz & " keepH=" & keepH
    On Error GoTo Cleanup

    olds = Application.ScreenUpdating
    oldal = Application.DisplayAlerts
    Application.ScreenUpdating = False

    If allowUndo Then
        On Error Resume Next
        Set ur = Application.UndoRecord
        If Not ur Is Nothing Then ur.StartCustomRecord "Chuan NTU"
        Err.Clear
        On Error GoTo Cleanup
    End If

    ' 1. Kho giay + le + kieu Normal
    lastStep = "kho/le"
    On Error Resume Next
    Err.Clear
    OptimizePageSetup d
    SetNormalStyle d, sz
    If Err.Number <> 0 Then
        rep = rep & "! Khổ giấy / lề / kiểu Normal: " & Err.Description & vbCrLf
        Trace "step kho/le ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "✓ A4 dọc · lề trên 2 · dưới 2 · trái 3 · phải 2 cm · kiểu Normal đã đặt" & vbCrLf
    End If
    On Error GoTo Cleanup

    ' 2. Than bai
    lastStep = "than bai"
    On Error Resume Next
    Err.Clear
    n = NormalizeBody(d, sz, keepH)
    If Err.Number <> 0 Then
        rep = rep & "! Thân bài: " & Err.Description & vbCrLf
        Trace "step than bai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "✓ Đã định dạng " & Format$(n, "#,##0") & " đoạn" & _
          IIf(keepH, " (giữ cỡ chữ của tiêu đề)", "") & vbCrLf
    On Error GoTo Cleanup

    ' 3. Tieu de muc
    lastStep = "tieu de"
    On Error Resume Next
    Err.Clear
    ApplyHeadingStyles d
    If Err.Number <> 0 Then
        rep = rep & "! Tiêu đề mục: " & Err.Description & vbCrLf
        Trace "step tieu de ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "✓ Tiêu đề: chương IN HOA đậm 14 căn giữa · mục đậm 13 không thụt dòng" & vbCrLf
    End If
    On Error GoTo Cleanup

    ' 4. Caption bang/hinh
    lastStep = "caption"
    On Error Resume Next
    Err.Clear
    ApplyCaptions d
    If Err.Number <> 0 Then
        rep = rep & "! Caption bảng/hình: " & Err.Description & vbCrLf
        Trace "step caption ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "✓ Đã chèn " & LastCaptionCount & " caption bảng/hình còn thiếu" & vbCrLf
    On Error GoTo Cleanup

    ' 5. So trang theo section (La Ma phan dau, A Rap phan than)
    lastStep = "so trang"
    On Error Resume Next
    Err.Clear
    ApplySectionPageNumbers d
    If Err.Number <> 0 Then
        rep = rep & "! Số trang theo phần: " & Err.Description & vbCrLf
        Trace "step so trang ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    soSec = 0
    soSec = d.Sections.Count
    rep = rep & "✓ Số trang: " & soSec & " phần (section), đánh số ở giữa phía dưới" & vbCrLf
    On Error GoTo Cleanup

    ' 6. Gach dai
    lastStep = "gach dai"
    On Error Resume Next
    Err.Clear
    n = FixDashes(d, rep)
    If Err.Number <> 0 Then
        rep = rep & "! Gạch dài: " & Err.Description & vbCrLf
        Trace "step gach dai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    ElseIf n > 0 Then
        rep = rep & "✓ Đã đổi " & Format$(n, "#,##0") & " gạch dài thành gạch nối (-)" & vbCrLf
    End If
    On Error GoTo Cleanup

    Trace "== xong"
    On Error Resume Next
    If allowUndo Then
        If Not ur Is Nothing Then ur.EndCustomRecord
    End If
    Application.ScreenUpdating = olds
    Application.DisplayAlerts = oldal
    Err.Clear
    NTU_ChuanHoaFull = rep & vbCrLf & "Xong. Bấm [Kiểm tra] để soát lại và xem lỗi có đánh số."
    Exit Function

Cleanup:
    em = Err.Description
    Trace "FATAL at " & lastStep & ": " & em
    On Error Resume Next
    If allowUndo Then
        If Not ur Is Nothing Then ur.EndCustomRecord
    End If
    Application.ScreenUpdating = olds
    Application.DisplayAlerts = oldal
    Err.Clear
    NTU_ChuanHoaFull = rep & "! Lỗi ở bước [" & lastStep & "]: " & em
End Function

' -------------------------------------------------------- Tieu de muc

' Heading1: IN HOA, dam, 14, can giua, khong thut dong.
' Heading2..4: dam, 13, can deu, khong thut dong.
Private Sub ApplyHeadingStyles(ByVal d As Document)
    Dim p As Paragraph
    Dim lvl As Long
    Dim upC As Boolean

    upC = OptUpperChapter()
    For Each p In d.Content.Paragraphs
        On Error Resume Next
        Err.Clear
        lvl = HeadingLevel(p)
        If Err.Number <> 0 Then Err.Clear
        If lvl >= 1 And lvl <= MAX_DEPTH Then
            If Not IsTableMark(p) Then
                With p.Range.ParagraphFormat
                    .FirstLineIndent = 0
                    .SpaceBefore = 6
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpace1pt5
                    If lvl = 1 Then
                        .Alignment = wdAlignParagraphCenter
                    Else
                        .Alignment = wdAlignParagraphJustify
                    End If
                End With
                SetRunFont p.Range
                p.Range.Font.Bold = True
                If lvl = 1 Then
                    p.Range.Font.Size = 14
                    p.Range.Font.AllCaps = True
                    If upC Then UpperChapterText p
                Else
                    p.Range.Font.Size = 13
                    p.Range.Font.AllCaps = False
                End If
            End If
        End If
        If Err.Number <> 0 Then
            Trace "ApplyHeadingStyles ERR " & Err.Number & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next
End Sub

' In hoa chu cua tieu de chuong. Chi lam khi: doan ngan (<120 ky tu), chua in hoa,
' va khong chua field/hinh (tranh pha dinh dang).
Private Sub UpperChapterText(ByVal p As Paragraph)
    Dim r As Range, t As String
    On Error Resume Next
    t = CleanText(p.Range.Text)
    If Len(t) = 0 Or Len(t) >= 120 Then Exit Sub
    If UCase$(t) = t Then Exit Sub
    Set r = p.Range
    If r.End > r.Start Then r.MoveEnd wdCharacter, -1
    If r.End <= r.Start Then Exit Sub
    If r.Fields.Count > 0 Then Exit Sub
    If r.InlineShapes.Count > 0 Then Exit Sub
    r.Text = UCase$(t)
    If Err.Number <> 0 Then
        Trace "UpperChapterText ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
End Sub

' -------------------------------------------------------- Caption bang/hinh

' So chuong tu text tieu de ("CHUONG 2 ..." -> 2); khong doc duoc thi dung so thu tu.
Private Function ChapterNoFromText(ByVal t As String, ByVal fallback As Long) As Long
    Dim i As Long, a As Long, s As String, seen As Boolean
    For i = 1 To Len(t)
        a = AscW(Mid$(t, i, 1))
        If a >= 48 And a <= 57 Then
            s = s & Chr$(a)
            seen = True
        ElseIf seen Then
            Exit For
        End If
    Next
    If Len(s) > 0 Then
        ChapterNoFromText = CLng(Val(s))
    Else
        ChapterNoFromText = fallback
    End If
    If ChapterNoFromText < 1 Then ChapterNoFromText = fallback
End Function

' So chuong dang hieu tai vi tri p (mm). 0 = truoc chuong 1 (phan dau: bo qua).
Private Function ChuongAt(ByRef pos() As Long, ByRef no() As Long, ByVal n As Long, ByVal p As Long) As Long
    Dim i As Long
    ChuongAt = 0
    For i = 1 To n
        If pos(i) <= p Then
            ChuongAt = no(i)
        Else
            Exit For
        End If
    Next
End Function

' "Bang 2.3 ..." -> 3 ; "Hinh 2.3 ..." -> 3 ; sai dang -> 0
Private Function CaptionNo(ByVal t As String, ByVal kind As String) As Long
    Dim s As String, dot As Long, i As Long, prefix As String
    CaptionNo = 0
    s = Trim$(t)
    prefix = kind & " "
    If Len(s) <= Len(prefix) Then Exit Function
    If Left$(s, Len(prefix)) <> prefix Then Exit Function
    s = Mid$(s, Len(prefix) + 1)
    dot = InStr(s, ".")
    If dot <= 1 Then Exit Function
    If Not IsDigits(Left$(s, dot - 1)) Then Exit Function
    s = Mid$(s, dot + 1)
    i = 1
    Do While i <= Len(s)
        If Not IsDigits(Mid$(s, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If i = 1 Then Exit Function
    CaptionNo = CLng(Val(Left$(s, i - 1)))
End Function

' Doan ngay TREN bang (thuong la caption cua bang).
Private Function ParaBeforeTable(ByVal t As Table) As Paragraph
    Dim r As Range
    On Error Resume Next
    Set r = t.Range
    If r.Start <= 0 Then Exit Function
    r.Collapse wdCollapseStart
    r.MoveStart wdCharacter, -1
    Set ParaBeforeTable = r.Paragraphs(1)
    Err.Clear
End Function

' Doan ngay DUOI hinh.
Private Function ParaAfterShape(ByVal sh As InlineShape) As Paragraph
    Dim p As Paragraph
    On Error Resume Next
    Set p = sh.Range.Paragraphs(1)
    If p Is Nothing Then Exit Function
    Set ParaAfterShape = p.Next
    Err.Clear
End Function

' Chen caption ngay TREN bang: chen "txt" + xuong dong vao dung vi tri bat dau bang.
Private Function ChenCaptionTren(ByVal d As Document, ByVal t As Table, ByVal txt As String) As Range
    Dim pos As Long, r As Range
    On Error Resume Next
    pos = 0
    pos = t.Range.Start
    Set r = d.Range(pos, pos)
    r.InsertBefore txt & vbCr
    Set ChenCaptionTren = d.Range(pos, pos + Len(txt))
    Err.Clear
End Function

' Chen caption ngay DUOI hinh: chen xuong dong + "txt" ngay truoc dau doan cua hinh.
Private Function ChenCaptionDuoi(ByVal d As Document, ByVal sh As InlineShape, ByVal txt As String) As Range
    Dim p As Paragraph, pos As Long, r As Range
    On Error Resume Next
    Set p = sh.Range.Paragraphs(1)
    If p Is Nothing Then Exit Function
    pos = 0
    pos = p.Range.End - 1
    If pos < 0 Then Exit Function
    Set r = d.Range(pos, pos)
    r.InsertAfter vbCr & txt
    Set ChenCaptionDuoi = d.Range(pos + 1, pos + 1 + Len(txt))
    Err.Clear
End Function

Private Sub StyleCaptionRange(ByVal r As Range)
    On Error Resume Next
    r.Style = r.Document.Styles(wdStyleCaption)
    If Err.Number <> 0 Then Err.Clear
    SetRunFont r
    With r.Font
        .Bold = True
        .Italic = False
        .Size = 13
        .AllCaps = False
    End With
    With r.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .FirstLineIndent = 0
        .LeftIndent = 0
        .RightIndent = 0
        .SpaceBefore = 6
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .OutlineLevel = wdOutlineLevelBodyText
    End With
    If Err.Number <> 0 Then
        Trace "StyleCaptionRange ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
End Sub

' Dong ngay duoi caption bat dau bang "Nguon:" -> nghieng 11, can trai.
Private Sub StyleSourceLine(ByVal capRange As Range)
    Dim p As Paragraph
    On Error Resume Next
    Set p = capRange.Paragraphs(1).Next
    If p Is Nothing Then Exit Sub
    If p.Range.Information(wdWithInTable) Then Exit Sub
    If Not IsNguonLine(p.Range.Text) Then Exit Sub
    With p.Range.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
    p.Range.Font.Size = 11
    p.Range.Font.Italic = True
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function IsNguonLine(ByVal t As String) As Boolean
    Dim s As String
    s = LCase$(CleanText(t))
    IsNguonLine = (Left$(s, 6) = "nguồn:" Or Left$(s, 6) = "nguon:")
End Function

Private Function IsCaptionPara(ByVal p As Paragraph) As Boolean
    Dim t As String
    On Error Resume Next
    t = CleanText(p.Range.Text)
    If Len(t) > 200 Then Exit Function
    If IsNguonLine(t) Then
        IsCaptionPara = True
        Exit Function
    End If
    If CaptionNo(t, "Bảng") > 0 Then
        IsCaptionPara = True
        Exit Function
    End If
    If CaptionNo(t, "Hình") > 0 Then IsCaptionPara = True
    Err.Clear
End Function

' Chen caption cho moi bang/hinh con thieu; danh so theo chuong; idempotent.
Private Sub ApplyCaptions(ByVal d As Document)
    Dim p As Paragraph, r As Range
    Dim i As Long, k As Long, nCh As Long, ch As Long, datNo As Long
    Dim pos() As Long, no() As Long
    Dim demB() As Long, demH() As Long
    Dim caps As Collection, a As Variant
    Dim t As String, txt As String, kind As String

    LastCaptionCount = 0
    On Error Resume Next
    Err.Clear

    ' 1. Vi tri + so hieu cac chuong (quet Heading1 mot lan; dung vi tri tuyet doi
    '    de khong phu thuoc viec chen/xoa o cac buoc sau).
    nCh = 0
    For Each p In d.Content.Paragraphs
        If HeadingLevel(p) = 1 Then
            nCh = nCh + 1
            ReDim Preserve pos(1 To nCh)
            ReDim Preserve no(1 To nCh)
            pos(nCh) = p.Range.Start
            no(nCh) = ChapterNoFromText(CleanText(p.Range.Text), nCh)
        End If
        If Err.Number <> 0 Then Err.Clear
    Next
    If nCh = 0 Then
        Trace "ApplyCaptions: khong co Heading1 -> bo qua"
        On Error GoTo 0
        Exit Sub
    End If

    ReDim demB(1 To nCh)
    ReDim demH(1 To nCh)
    Set caps = New Collection

    ' 2. Bang: duyet TIEN de tinh so, chua chen gi.
    For i = 1 To d.Tables.Count
        Err.Clear
        If Not d.Tables(i).Range.Information(wdWithInTable) Then
            ch = ChuongAt(pos, no, nCh, d.Tables(i).Range.Start)
            If ch > 0 Then
                Set p = ParaBeforeTable(d.Tables(i))
                t = ""
                If Not p Is Nothing Then t = CleanText(p.Range.Text)
                datNo = CaptionNo(t, "Bảng")
                If datNo > 0 Then
                    If datNo > demB(ch) Then demB(ch) = datNo
                    caps.Add Array("B", ch, i, 0)
                Else
                    demB(ch) = demB(ch) + 1
                    caps.Add Array("B", ch, i, demB(ch))
                End If
            End If
        End If
        If Err.Number <> 0 Then
            Trace "ApplyCaptions table " & i & " ERR " & Err.Description
            Err.Clear
        End If
    Next

    ' 3. Hinh
    For i = 1 To d.InlineShapes.Count
        Err.Clear
        If Not d.InlineShapes(i).Range.Information(wdWithInTable) Then
            ch = ChuongAt(pos, no, nCh, d.InlineShapes(i).Range.Start)
            If ch > 0 Then
                Set p = ParaAfterShape(d.InlineShapes(i))
                t = ""
                If Not p Is Nothing Then t = CleanText(p.Range.Text)
                datNo = CaptionNo(t, "Hình")
                If datNo > 0 Then
                    If datNo > demH(ch) Then demH(ch) = datNo
                    caps.Add Array("H", ch, i, 0)
                Else
                    demH(ch) = demH(ch) + 1
                    caps.Add Array("H", ch, i, demH(ch))
                End If
            End If
        End If
        If Err.Number <> 0 Then
            Trace "ApplyCaptions shape " & i & " ERR " & Err.Description
            Err.Clear
        End If
    Next

    ' 4. Chen NGUOC tu cuoi len: vi tri cac doi tuong phia truoc khong doi.
    For k = caps.Count To 1 Step -1
        a = caps(k)
        kind = CStr(a(0))
        ch = CLng(a(1))
        i = CLng(a(2))
        datNo = CLng(a(3))
        txt = IIf(kind = "B", "Bảng ", "Hình ") & ch & "." & datNo
        Set r = Nothing
        Set p = Nothing
        If datNo > 0 Then
            If kind = "B" Then
                Set r = ChenCaptionTren(d, d.Tables(i), txt)
            Else
                Set r = ChenCaptionDuoi(d, d.InlineShapes(i), txt)
            End If
        Else
            If kind = "B" Then
                Set p = ParaBeforeTable(d.Tables(i))
            Else
                Set p = ParaAfterShape(d.InlineShapes(i))
            End If
            If Not p Is Nothing Then Set r = p.Range
        End If
        If Not r Is Nothing Then
            StyleCaptionRange r
            StyleSourceLine r
            If datNo > 0 Then LastCaptionCount = LastCaptionCount + 1
        End If
        If Err.Number <> 0 Then
            Trace "ApplyCaptions chen " & kind & i & " ERR " & Err.Description
            Err.Clear
        End If
    Next

    Trace "ApplyCaptions: chen " & LastCaptionCount & " caption / " & caps.Count & " doi tuong"
    On Error GoTo 0
End Sub

' -------------------------------------------------------- So trang theo phan

' Phan dau (bia, muc luc, ...) danh so La Ma; tu chuong 1 danh so A Rap tu 1.
Private Sub ApplySectionPageNumbers(ByVal d As Document)
    Dim p As Paragraph, first As Paragraph
    Dim r As Range, t As String
    Dim i As Long, nSec As Long

    On Error Resume Next
    Err.Clear

    ' 1. Tieu de chuong dau tien
    For Each p In d.Content.Paragraphs
        If HeadingLevel(p) = 1 Then
            Set first = p
            Exit For
        End If
    Next

    If first Is Nothing Then
        Trace "ApplySectionPageNumbers: khong co Heading1"
        MsgBox "Không tìm thấy tiêu đề chương (Heading 1 hoặc đoạn mở đầu ""CHƯƠNG ..."") nên chưa đặt lại số trang theo phần." & vbCrLf & _
               "Quy cách: phần đầu đánh số La Mã (i, ii...), từ chương 1 đánh số 1, 2, 3...", vbExclamation, "Chuẩn NTU"
        On Error GoTo 0
        Exit Sub
    End If

    ' 2. Chen section break ngay TRUOC chuong 1 neu chua co
    If first.Range.Start > 0 Then
        t = ""
        t = d.Range(first.Range.Start - 1, first.Range.Start).Text
        If t <> Chr$(12) Then
            Set r = d.Range(first.Range.Start, first.Range.Start)
            r.InsertBreak wdSectionBreakNextPage
            Err.Clear
        End If
    End If
    If Err.Number <> 0 Then
        Trace "ApplySectionPageNumbers: chen break ERR " & Err.Description
        Err.Clear
    End If

    ' 3. So trang tung section
    nSec = 0
    nSec = d.Sections.Count
    If nSec <= 1 Then
        ' Khong co phan dau -> danh so A Rap binh thuong, khong bat dau bang La Ma.
        With d.Sections(1).PageNumbers
            .NumberStyle = wdPageNumberStyleArabic
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
    Else
        With d.Sections(1).PageNumbers
            .NumberStyle = wdPageNumberStyleLowercaseRoman
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
        For i = 2 To nSec
            With d.Sections(i).PageNumbers
                .NumberStyle = wdPageNumberStyleArabic
                If i = 2 Then
                    .RestartNumberingAtSection = True
                    .StartingNumber = 1
                Else
                    .RestartNumberingAtSection = False
                End If
            End With
        Next
    End If
    If Err.Number <> 0 Then
        Trace "ApplySectionPageNumbers: danh so ERR " & Err.Description
        Err.Clear
    End If

    ' 4. Xoa header + dat lai field PAGE o footer cho tung section
    SetPageNumberFooter d
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

' -------------------------------------------------------- Cua so nhap

' Mo cua so nhap (dong cu so truoc). Khong co UserForm nen dung Document tam.
Private Function MoCuaSoNhap() As Document
    Dim d As Document
    On Error Resume Next
    DongCuaSoNhap
    Set d = Documents.Add
    If Err.Number <> 0 Then
        Trace "MoCuaSoNhap ERR " & Err.Description
        Err.Clear
        Exit Function
    End If
    Set LastScratch = d
    Set MoCuaSoNhap = d
End Function

' Dong cua so nhap: khong hoi luu, khong dung toi tai lieu goc.
Private Sub DongCuaSoNhap()
    Dim i As Long, found As Boolean
    If LastScratch Is Nothing Then Exit Sub
    On Error Resume Next
    found = False
    For i = 1 To Application.Documents.Count
        If Application.Documents(i) Is LastScratch Then found = True
    Next
    If found Then
        Application.DisplayAlerts = wdAlertsNone
        LastScratch.Close 0
        Application.DisplayAlerts = wdAlertsAll
    End If
    Set LastScratch = Nothing
    Err.Clear
End Sub

' -------------------------------------------------------- Phan tich (chi doc)

' ponytail: dung cua so nhap thay task pane; nang len task pane that khi chuyen VSTO/Office.js
Public Sub RibbonPhanTich(control As Object)
    Dim src As Document, kq As Document
    Dim p As Paragraph, sh As InlineShape
    Dim rows As Collection, a As Variant
    Dim i As Long, n As Long, lvl As Long, prev As Long
    Dim nCh As Long, soBang As Long, soHinh As Long
    Dim thieuCap As Long, nhayCap As Long
    Dim pos() As Long, no() As Long
    Dim t As String, hienTai As String
    Dim soFont As Long, soCo As Long, soDong As Long, soCan As Long, soThut As Long
    Dim r As Range

    On Error GoTo Loi
    If Application.Documents.Count = 0 Then
        MsgBox "Chưa có tài liệu nào đang mở.", vbExclamation, "Chuẩn NTU"
        Exit Sub
    End If
    Set src = ActiveDocument
    Application.ScreenUpdating = False
    Set rows = New Collection

    ' 1. Chuong (Heading1)
    For Each p In src.Content.Paragraphs
        If HeadingLevel(p) = 1 Then
            nCh = nCh + 1
            ReDim Preserve pos(1 To nCh)
            ReDim Preserve no(1 To nCh)
            pos(nCh) = p.Range.Start
            no(nCh) = ChapterNoFromText(CleanText(p.Range.Text), nCh)
            hienTai = "Cỡ " & FmtSo(p.Range.Font.Size) & " · đậm: " & YesNo(p.Range.Font.Bold) & _
                      " · in hoa: " & YesNo(p.Range.Font.AllCaps) & _
                      " · căn: " & TenCan(p.Range.ParagraphFormat.Alignment)
            rows.Add Array("Chương " & no(nCh), Snippet80(p.Range.Text), hienTai, "In hoa, đậm, 14 pt, căn giữa, không thụt dòng")
        End If
        If rows.Count >= 300 Then Exit For
    Next

    ' 2. Than bai + tieu de muc: chi ghi cho nao lech chuan
    prev = 0
    For Each p In src.Content.Paragraphs
        If Not IsTableMark(p) Then
            n = n + 1
            lvl = HeadingLevel(p)
            If lvl > 0 Then
                If prev > 0 Then
                    If lvl > prev + 1 Then
                        nhayCap = nhayCap + 1
                        If rows.Count < 300 Then
                            rows.Add Array("Tiêu đề (đoạn " & n & ")", Snippet80(p.Range.Text), _
                                           "H" & prev & " → H" & lvl, "Không nhảy cấp: thêm mục cấp trung gian")
                        End If
                    End If
                End If
                prev = lvl
            End If
            If lvl = 0 And Not p.Range.Information(wdWithInTable) And Not IsCaptionPara(p) Then
                t = p.Range.Text
                If p.Range.Font.Name <> FONT_NAME And p.Range.Font.Name <> "" Then
                    soFont = soFont + 1
                    If rows.Count < 300 Then rows.Add Array("Đoạn " & n, Snippet80(t), _
                        "Phông " & p.Range.Font.Name, "Times New Roman")
                End If
                If Abs(p.Range.Font.Size - 13) > 0.5 Then
                    soCo = soCo + 1
                    If rows.Count < 300 Then rows.Add Array("Đoạn " & n, Snippet80(t), _
                        "Cỡ " & FmtSo(p.Range.Font.Size), "13 pt")
                End If
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 Then
                    soDong = soDong + 1
                    If rows.Count < 300 Then rows.Add Array("Đoạn " & n, Snippet80(t), _
                        "Giãn dòng " & TenGian(p.Range.ParagraphFormat.LineSpacingRule), "Giãn dòng 1,5")
                End If
                If p.Range.ParagraphFormat.Alignment <> wdAlignParagraphJustify Then
                    soCan = soCan + 1
                    If rows.Count < 300 Then rows.Add Array("Đoạn " & n, Snippet80(t), _
                        "Căn " & TenCan(p.Range.ParagraphFormat.Alignment), "Căn đều 2 bên")
                End If
                If Abs(p.Range.ParagraphFormat.FirstLineIndent - Cm(OptIndent())) > 3 Then
                    soThut = soThut + 1
                    If rows.Count < 300 Then rows.Add Array("Đoạn " & n, Snippet80(t), _
                        "Thụt dòng đầu " & FmtSo(p.Range.ParagraphFormat.FirstLineIndent / Cm(1#)) & " cm", _
                        "Thụt dòng đầu " & Format$(OptIndent(), "0.00") & " cm")
                End If
            End If
        End If
        If rows.Count >= 300 Then Exit For
    Next

    ' 3. Bang / hinh thieu caption
    For i = 1 To src.Tables.Count
        If Not src.Tables(i).Range.Information(wdWithInTable) Then
            soBang = soBang + 1
            Set p = ParaBeforeTable(src.Tables(i))
            t = ""
            If Not p Is Nothing Then t = p.Range.Text
            If CaptionNo(CleanText(t), "Bảng") = 0 Then
                thieuCap = thieuCap + 1
                If rows.Count < 300 Then
                    rows.Add Array("Bảng #" & soBang & " (" & src.Tables(i).Rows.Count & "×" & src.Tables(i).Columns.Count & ")", _
                                   Snippet80(t), "Không có caption", "Chèn [Bảng X.Y] ngay TRÊN bảng, đậm 13, căn giữa")
                End If
            End If
        End If
    Next

    For Each sh In src.InlineShapes
        If Not sh.Range.Information(wdWithInTable) Then
            soHinh = soHinh + 1
            Set p = ParaAfterShape(sh)
            t = ""
            If Not p Is Nothing Then t = p.Range.Text
            If CaptionNo(CleanText(t), "Hình") = 0 Then
                thieuCap = thieuCap + 1
                If rows.Count < 300 Then
                    rows.Add Array("Hình #" & soHinh, "", "Không có caption", _
                                   "Chèn [Hình X.Y] ngay DƯỚI hình, đậm 13, căn giữa")
                End If
            End If
        End If
    Next

    ' 4. Cua so nhap
    Set kq = MoCuaSoNhap()
    If kq Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Không mở được cửa sổ nháp để ghi báo cáo.", vbExclamation, "Chuẩn NTU"
        Exit Sub
    End If
    kq.Content.InsertAfter "BÁO CÁO PHÂN TÍCH — QUY CÁCH NTU" & vbCr & _
        "(QĐ 900/QĐ-ĐHNT ngày 31/7/2023) · " & src.Name & " · " & Format$(Now, "dd/mm/yyyy hh:nn") & vbCr & _
        "Chế độ CHỈ ĐỌC: không sửa gì trong tài liệu gốc. Bảng chi tiết ở dưới." & vbCr & vbCr & _
        "TỔNG KẾT" & vbCr & _
        "- Chương (Heading 1): " & nCh & vbCr & _
        "- Bảng: " & soBang & " · Hình: " & soHinh & vbCr & _
        "- Bảng/hình thiếu caption: " & thieuCap & vbCr & _
        "- Tiêu đề nhảy cấp: " & nhayCap & vbCr & _
        "- Đoạn lệch phông: " & soFont & " · lệch cỡ chữ: " & soCo & _
        " · lệch giãn dòng: " & soDong & " · lệch căn lề: " & soCan & _
        " · lệch thụt dòng: " & soThut & vbCr & _
        "- Ảnh: " & src.InlineShapes.Count & " — DPI KHÔNG xác định được (Word không phơi DPI cho InlineShape)." & vbCr & vbCr

    Set r = kq.Content
    r.Collapse wdCollapseEnd
    Dim tbl As Table
    Set tbl = kq.Tables.Add(r, rows.Count + 1, 4)
    tbl.Borders.Enable = True
    tbl.Cell(1, 1).Range.Text = "Vị trí"
    tbl.Cell(1, 2).Range.Text = "Đoạn (80 ký tự)"
    tbl.Cell(1, 3).Range.Text = "Hiện tại"
    tbl.Cell(1, 4).Range.Text = "Gợi ý"
    For i = 1 To rows.Count
        a = rows(i)
        tbl.Cell(i + 1, 1).Range.Text = CStr(a(0))
        tbl.Cell(i + 1, 2).Range.Text = CStr(a(1))
        tbl.Cell(i + 1, 3).Range.Text = CStr(a(2))
        tbl.Cell(i + 1, 4).Range.Text = CStr(a(3))
    Next
    On Error Resume Next
    tbl.Range.Font.Name = FONT_NAME
    tbl.Range.Font.Size = 11
    tbl.Rows(1).HeadingFormat = True
    Err.Clear

    kq.Saved = True
    Application.ScreenUpdating = True
    Trace "RibbonPhanTich: " & rows.Count & " dong / " & soBang & " bang / " & soHinh & " hinh"
    Exit Sub

Loi:
    Application.ScreenUpdating = True
    MsgBox "Lỗi khi phân tích: " & Err.Description, vbExclamation, "Chuẩn NTU"
End Sub

Private Function Snippet80(ByVal s As String) As String
    Dim t As String
    t = CleanText(s)
    If Len(t) > 80 Then t = Left$(t, 77) & "..."
    Snippet80 = t
End Function

Private Function FmtSo(ByVal v As Single) As String
    FmtSo = Trim$(Format$(v, "0.##"))
End Function

Private Function YesNo(ByVal b As Boolean) As String
    If b Then
        YesNo = "có"
    Else
        YesNo = "không"
    End If
End Function

Private Function TenCan(ByVal al As Long) As String
    Select Case al
        Case wdAlignParagraphLeft
            TenCan = "trái"
        Case wdAlignParagraphRight
            TenCan = "phải"
        Case wdAlignParagraphCenter
            TenCan = "giữa"
        Case wdAlignParagraphJustify
            TenCan = "đều 2 bên"
        Case Else
            TenCan = "khác (" & al & ")"
    End Select
End Function

Private Function TenGian(ByVal v As Long) As String
    Select Case v
        Case wdLineSpaceSingle
            TenGian = "1,0"
        Case wdLineSpace1pt5
            TenGian = "1,5"
        Case wdLineSpaceDouble
            TenGian = "2,0"
        Case wdLineSpaceAtLeast
            TenGian = "tối thiểu"
        Case wdLineSpaceExactly
            TenGian = "chính xác"
        Case wdLineSpaceMultiple
            TenGian = "bội số"
        Case Else
            TenGian = "khác (" & v & ")"
    End Select
End Function

' ============================================================================
' TU CAP NHAT (doc lap voi add-in Word core; dung manifest version.json cua repo)
' ============================================================================

Public Function NTU_PhienBan() As String
    NTU_PhienBan = "1.0.2"
End Function

Private Function NTU_ManifestUrl() As String
    NTU_ManifestUrl = "https://github.com/giawin/ntu-thesis-addin/releases/latest/download/version.json"
End Function

Private Function NTU_ParseJson(ByVal json As String, ByVal fieldName As String) As String
    Dim pos As Long, startPos As Long, endPos As Long
    pos = InStr(json, """" & fieldName & """")
    If pos > 0 Then
        startPos = InStr(pos + Len(fieldName) + 2, json, """")
        If startPos > 0 Then
            endPos = InStr(startPos + 1, json, """")
            If endPos > startPos Then
                NTU_ParseJson = Mid$(json, startPos + 1, endPos - startPos - 1)
            End If
        End If
    End If
End Function

Private Function NTU_SoSanh(ByVal a As String, ByVal b As String) As Integer
    Dim pa() As String, pb() As String
    pa = Split(a, ".")
    pb = Split(b, ".")
    Dim i As Integer, va As Long, vb As Long
    For i = 0 To 5
        va = 0: vb = 0
        If i <= UBound(pa) Then va = Val(pa(i))
        If i <= UBound(pb) Then vb = Val(pb(i))
        If va > vb Then
            NTU_SoSanh = 1
            Exit Function
        ElseIf va < vb Then
            NTU_SoSanh = -1
            Exit Function
        End If
    Next i
    NTU_SoSanh = 0
End Function

Private Function NTU_FileTuChay() As String
    On Error Resume Next
    NTU_FileTuChay = ThisDocument.FullName
    On Error GoTo 0
End Function

Private Function NTU_DungFileAddIn() As Boolean
    Dim n As String, p As String
    p = NTU_FileTuChay()
    n = LCase$(Mid$(p, InStrRev(Replace(p, "\", "/"), "/") + 1))
    NTU_DungFileAddIn = (n = "ntuthesis.dotm")
End Function

Private Function NTU_TaiManifest(Optional ByVal nhanh As Boolean = False) As String
    Dim i As Integer, res As String, soLan As Integer, cho As Integer
    soLan = IIf(nhanh, 1, 2)
    cho = IIf(nhanh, 3, 15)
    For i = 1 To soLan
        res = ""
        On Error Resume Next
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", NTU_ManifestUrl(), False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.setTimeouts 5000, 5000, cho * 1000, cho * 1000
        http.Send
        If Err.Number = 0 Then
            If http.Status = 200 Then res = http.responseText
        End If
        On Error GoTo 0
        res = Trim$(res)
        If Len(NTU_ParseJson(res, "version")) > 0 Then Exit For
    Next i
    NTU_TaiManifest = res
End Function

Public Sub RibbonCapNhat(control As Object)
    ' Nut [Cap Nhat] bam tay: luon hoi ket qua, ke ca khi da moi nhat.
    Call NTU_KiemTraCapNhat(True)
End Sub

Public Sub NTU_HenKiemTraKhiMo()
    Static daHen As Boolean
    If daHen Then Exit Sub
    daHen = True
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 20), "NTU_KiemTraNhanhKhiMo"
End Sub

Public Sub NTU_KiemTraNhanhKhiMo()
    Call NTU_TuKiemTraHangTuan(True)
End Sub

Public Sub NTU_TuKiemTraHangTuan(Optional ByVal nhanh As Boolean = False)
    ' Moi ngay kiem tra mot lan; giua hai lan thu cach nhau it nhat 60 phut.
    On Error Resume Next
    Dim lastChk As String, daDocDuoc As Boolean
    lastChk = GetSetting(APP, "CapNhat", "LanKiemTra_NTU", "")
    If Len(lastChk) > 0 Then
        If IsDate(lastChk) Then
            If DateDiff("d", CDate(lastChk), Date) < 1 Then Exit Sub
        End If
    End If
    Dim lanThu As String
    lanThu = GetSetting(APP, "CapNhat", "LanThu_NTU", "")
    If Len(lanThu) > 0 Then
        If IsDate(lanThu) Then
            If DateDiff("n", CDate(lanThu), Now) < 60 Then Exit Sub
        End If
    End If
    SaveSetting APP, "CapNhat", "LanThu_NTU", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error GoTo 0
    Call NTU_KiemTraCapNhat(False, nhanh, daDocDuoc)
    If daDocDuoc Then SaveSetting APP, "CapNhat", "LanKiemTra_NTU", Format$(Date, "yyyy-mm-dd")
End Sub

Public Sub NTU_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, _
                              Optional ByVal nhanh As Boolean = False, _
                              Optional ByRef daDocDuoc As Boolean = False)
    Dim js As String
    js = NTU_TaiManifest(nhanh)
    Dim ver As String, url As String, sha As String, notes As String, kt As Long
    ver = NTU_ParseJson(js, "version")
    ' Manifest co the co ca ban chung va ban _win; uu tien _win khi co.
    url = NTU_ParseJson(js, "ntu_url_win")
    If Len(url) = 0 Then url = NTU_ParseJson(js, "ntu_url")
    sha = LCase$(NTU_ParseJson(js, "ntu_sha256_win"))
    If Len(sha) = 0 Then sha = LCase$(NTU_ParseJson(js, "ntu_sha256"))
    notes = NTU_ParseJson(js, "notes_ascii")
    kt = Val(NTU_ParseJson(js, "ntu_size_win"))
    If kt = 0 Then kt = Val(NTU_ParseJson(js, "ntu_size"))
    If Len(ver) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Then
        If hoiKhiKhongCo Then
            MsgBox "Không kết nối được máy chủ cập nhật (GitHub). Kiểm tra mạng rồi thử lại." & vbCrLf & _
                   "Phiên bản đang dùng: " & NTU_PhienBan(), vbExclamation, "Cập nhật Chuẩn NTU"
        End If
        Exit Sub
    End If
    daDocDuoc = True
    If NTU_SoSanh(ver, NTU_PhienBan()) <= 0 Then
        If hoiKhiKhongCo Then
            MsgBox "Bạn đang dùng phiên bản mới nhất: " & NTU_PhienBan(), vbInformation, "Cập nhật Chuẩn NTU"
        End If
        Exit Sub
    End If
    If Not NTU_DungFileAddIn() Then
        Dim q As String
        q = Chr$(39)
        MsgBox "Có phiên bản mới " & ver & " nhưng macro đang chạy từ " & q & NTU_FileTuChay() & q & " - không phải tệp add-in đã cài." & vbCrLf & "Hãy cài lại bằng bộ cài đặt.", vbExclamation, "Cập nhật Chuẩn NTU"
        Exit Sub
    End If
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Có phiên bản mới: " & ver & " (đang dùng " & NTU_PhienBan() & ")." & vbCrLf & vbCrLf & _
                 notes & vbCrLf & vbCrLf & "Cập nhật ngay?", vbYesNo + vbQuestion, "Cập nhật Chuẩn NTU")
    If ans <> vbYes Then Exit Sub
    Dim msg As String
    msg = NTU_TaiVaCai(url, sha, kt)
    If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cập nhật Chuẩn NTU"
End Sub

Private Function NTU_TaiVaCai(ByVal url As String, ByVal sha As String, _
                              Optional ByVal kichThuoc As Long = 0) As String
    ' Tai ban moi, kiem tinh toan ven (ZIP magic + kich thuoc), roi de mot .cmd
    ' cho Word tat han moi copy de — khong the ghi de tep .dotm dang mo.
    Dim target As String
    target = NTU_FileTuChay()
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
    thuMuc = Environ("APPDATA") & "\NTUThesisAddin"
    If Len(Dir(thuMuc, vbDirectory)) = 0 Then MkDir thuMuc
    Dim moiW As String
    moiW = thuMuc & "\capnhat_NTU.moi"
    If Len(Dir(moiW)) > 0 Then Kill moiW
    Dim f As Integer
    f = FreeFile
    Open moiW For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    If FileLen(moiW) <> n Then GoTo WinBad
    Dim tgtW As String
    tgtW = target
    If Len(Environ("APPDATA")) > 0 And LCase$(Left$(tgtW, Len(Environ("APPDATA")))) = LCase$(Environ("APPDATA")) Then
        tgtW = "%APPDATA%" & Mid$(tgtW, Len(Environ("APPDATA")) + 1)
    End If
    Dim cmdPath As String
    cmdPath = thuMuc & "\capnhat_NTU.cmd"
    f = FreeFile
    Open cmdPath For Output As #f
    Print #f, "@echo off"
    Print #f, ":cho"
    Print #f, "tasklist /FI ""IMAGENAME eq WINWORD.EXE"" 2>nul | find /I ""WINWORD.EXE"" >nul"
    Print #f, "if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto cho)"
    Print #f, "copy /Y ""%APPDATA%\NTUThesisAddin\capnhat_NTU.moi"" """ & tgtW & """ >nul"
    Print #f, "del ""%APPDATA%\NTUThesisAddin\capnhat_NTU.moi"" >nul 2>&1"
    Close #f
    Shell "cmd.exe /c """ & cmdPath & """", vbMinimizedNoFocus
    NTU_TaiVaCai = "Đã tải bản mới. Hãy ĐÓNG HẾT cửa sổ Word: bản mới sẽ tự cài đặt, mở lại Word là dùng được."
    Exit Function
WinBad:
    On Error Resume Next
    Close #f
    If Len(moiW) > 0 Then Kill moiW
    NTU_TaiVaCai = "Tệp tải về không đúng (thiếu hoặc hỏng). Chưa thay đổi gì, hãy thử lại sau."
    Exit Function
WinFail:
    NTU_TaiVaCai = "Tải bản cập nhật thất bại (kiểm tra mạng). Chưa thay đổi gì."
End Function

' ================================================== KIEM TRA NANG CAO (QD 900)

' Bao cao day du: tom tat cua NTU_Check + phan CHI TIET co danh so, moi muc
' giu lai vung cua no de NhayToiLoi nhay toi duoc.
Public Function NTU_CheckDayDu(ByVal d As Document, ByVal sz As Single) As String
    Dim rep As String, chiTiet As String
    Dim p As Paragraph, t As String
    Dim lvl As Long, lvlTruoc As Long
    Dim i As Long, n As Long
    Dim coAnh As Long

    rep = NTU_Check(d, sz)

    Set LastGoc = d
    LastDemLoi = 0
    ReDim LastChoLoi(1 To 1)

    On Error GoTo Loi
    Application.ScreenUpdating = False

    ' (i)(ii)(iii) phong / co chu / gian dong — quet lai de lay vung tung loi.
    For Each p In d.Content.Paragraphs
        If Not IsTableMark(p) Then
            n = n + 1
            t = CleanText(p.Range.Text)
            lvl = HeadingLevel(p)

            If p.Range.Font.Name <> "" And p.Range.Font.Name <> FONT_NAME Then
                GhiLoi p.Range, chiTiet, "Phông " & p.Range.Font.Name & " (cần " & FONT_NAME & ")"
            ElseIf p.Range.Font.Name = "" And Len(t) > 0 Then
                GhiLoi p.Range, chiTiet, "Phông không đồng nhất trong đoạn (cần " & FONT_NAME & ")"
            End If

            ' Co chu hop le: than bai = sz; tieu de chuong 14; caption 11 hoac 13;
            ' trong bang 13. Ngoai bo do la loi.
            If Not CoChuHopLe(p, lvl, sz) Then
                If p.Range.Font.Size >= 9999999 Then
                    GhiLoi p.Range, chiTiet, "Cỡ chữ không đồng nhất trong đoạn"
                Else
                    GhiLoi p.Range, chiTiet, "Cỡ chữ " & FmtSo(p.Range.Font.Size) & " pt (cần " & _
                           IIf(lvl = 1, "14", FmtSo(sz)) & " pt" & _
                           IIf(IsCaptionPara(p), "; caption 11 hoặc 13", "") & ")"
                End If
            End If

            If lvl = 0 And Not p.Range.Information(wdWithInTable) Then
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 Then
                    GhiLoi p.Range, chiTiet, "Giãn dòng " & TenGian(p.Range.ParagraphFormat.LineSpacingRule) & " (cần 1,5)"
                End If
            ElseIf p.Range.Information(wdWithInTable) Then
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Then
                    GhiLoi p.Range, chiTiet, "Giãn dòng trong bảng chưa bằng 1,0"
                End If
            End If

            ' (v) nhay cap tieu de: H1 -> H3 tro len la loi.
            If lvl >= 3 And lvlTruoc = 1 Then
                GhiLoi p.Range, chiTiet, "Nhảy cấp tiêu đề: cấp 1 -> cấp " & lvl & " (thiếu cấp 2)"
            End If
            If lvl > 0 Then lvlTruoc = lvl
        End If
    Next p

    ' (iv) bang/hinh thieu caption.
    For i = 1 To d.Tables.Count
        Set p = ParaBeforeTable(d.Tables(i))
        If Not p Is Nothing Then
            t = CleanText(p.Range.Text)
            If Len(Trim$(t)) = 0 Or InStr(1, t, "Bảng", vbTextCompare) = 0 Then
                GhiLoi d.Tables(i).Range, chiTiet, "Bảng chưa có caption ""Bảng X.Y"" ngay TRÊN bảng"
            End If
        End If
    Next i

    For i = 1 To d.InlineShapes.Count
        Set p = ParaAfterShape(d.InlineShapes(i))
        If Not p Is Nothing Then
            t = CleanText(p.Range.Text)
            If Len(Trim$(t)) = 0 Or InStr(1, t, "Hình", vbTextCompare) = 0 Then
                GhiLoi d.InlineShapes(i).Range, chiTiet, "Hình chưa có caption ""Hình X.Y"" ngay DƯỚI hình"
            End If
        End If
    Next i

    ' (vi) do phan giai anh: Word KHONG cho doc DPI qua COM -> noi thang la
    ' khong xac dinh duoc, khong doan bua.
    coAnh = d.InlineShapes.Count
    If coAnh > 0 Then
        chiTiet = chiTiet & "  - Ảnh: " & coAnh & " hình — độ phân giải không xác định được qua COM;" & _
                  " mở Word, chuột phải > Size để tự soát (cần >= 300 dpi khi in)." & vbCrLf
    End If

    If LastDemLoi = 0 Then
        chiTiet = "  (không có mục nào cần chỉ tay)"
    Else
        chiTiet = "  Nhập số trong [1.." & LastDemLoi & "] để nhảy tới lỗi." & vbCrLf & chiTiet
    End If

    rep = rep & vbCrLf & String(58, "-") & vbCrLf & _
          "CHI TIẾT LỖI (đã soát " & n & " đoạn)" & vbCrLf & chiTiet

    Application.ScreenUpdating = True
    NTU_CheckDayDu = rep
    Exit Function
Loi:
    Application.ScreenUpdating = True
    NTU_CheckDayDu = rep & vbCrLf & "! Lỗi khi kiểm tra chi tiết: " & Err.Description
End Function

' Co chu co nam trong bo hop le khong.
Private Function CoChuHopLe(ByVal p As Paragraph, ByVal lvl As Long, ByVal sz As Single) As Boolean
    Dim v As Single
    v = p.Range.Font.Size
    If v <= 0 Then CoChuHopLe = True: Exit Function

    If lvl = 1 Then
        CoChuHopLe = (Abs(v - 14) <= 0.5)
        Exit Function
    End If
    If IsCaptionPara(p) Then
        CoChuHopLe = (Abs(v - 11) <= 0.5 Or Abs(v - 13) <= 0.5)
        Exit Function
    End If
    If p.Range.Information(wdWithInTable) Then
        CoChuHopLe = (Abs(v - 13) <= 0.5)
        Exit Function
    End If
    CoChuHopLe = (Abs(v - sz) <= 0.5)
End Function

' Ghi 1 loi vao bao cao chi tiet va giu vung de nhay toi.
Private Sub GhiLoi(ByVal r As Range, ByRef out As String, ByVal lyDo As String)
    Dim t As String
    LastDemLoi = LastDemLoi + 1
    ReDim Preserve LastChoLoi(1 To LastDemLoi)
    Set LastChoLoi(LastDemLoi) = r.Duplicate

    t = CleanText(r.Text)
    If Len(t) > 60 Then t = Left$(t, 57) & "..."
    out = out & "  [" & LastDemLoi & "] " & lyDo & vbCrLf & _
          "      """ & t & """" & vbCrLf
End Sub

' Xuat bao cao ra tep de gui kem / luu lai. UTF-16 cho Word mo lai khong loi font.
Private Sub XuatTxt(ByVal rep As String)
    Dim p As String, f As Object
    On Error Resume Next
    p = Environ$("TEMP") & "\NTUChuanHoa_KiemTra.txt"
    Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile(p, 2, True, -1)
    If Err.Number <> 0 Then
        Trace "XuatTxt ERR " & Err.Description
        Err.Clear
        Exit Sub
    End If
    f.Write rep
    f.Close
    Set f = Nothing
    Trace "XuatTxt -> " & p
    Err.Clear
End Sub
