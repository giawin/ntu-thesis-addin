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

' Buoc dang chay â€” de bao loi biet chet o dau (xem NTU_Apply).
Private lastStep As String

' So caption da chen trong lan ApplyCaptions gan nhat (doc lai o NTU_ChuanHoaFull).
Private LastCaptionCount As Long

' Cua so nhap (Document tam) dung cho RibbonPhanTich â€” giu tham chieu de dong lai.
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
    s = "QUY CÃCH TRÃŒNH BÃ€Y â€” TRÆ¯á»œNG Äáº I Há»ŒC NHA TRANG" & vbCrLf & _
        "(Quyáº¿t Ä‘á»‹nh sá»‘ 900/QÄ-ÄHNT ngÃ y 31/7/2023)" & vbCrLf & String(46, "-") & vbCrLf & _
        "Khá»• giáº¥y: A4 (210 Ã— 297 mm), chiá»u dá»c, in má»™t máº·t" & vbCrLf & _
        "Lá»: trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm" & vbCrLf & _
        "PhÃ´ng: Times New Roman, cá»¡ 13; cÄƒn Ä‘á»u 2 bÃªn" & vbCrLf & _
        "GiÃ£n dÃ²ng 1,5 Â· dÃ²ng Ä‘áº§u má»—i Ä‘oáº¡n thá»¥t 1 cm" & vbCrLf & _
        "CÃ¡ch Ä‘oáº¡n: 3â€“6 pt phÃ­a trÃªn, 0 pt phÃ­a dÆ°á»›i" & vbCrLf & _
        "KhÃ´ng dÃ¹ng header Â· sá»‘ trang á»Ÿ giá»¯a phÃ­a DÆ¯á»šI" & vbCrLf & vbCrLf & _
        "TiÃªu Ä‘á» chÆ°Æ¡ng: IN HOA, in Ä‘áº­m, cÄƒn giá»¯a, cá»¡ 14." & vbCrLf & _
        "TiÃªu Ä‘á» má»¥c/tiá»ƒu má»¥c: chá»¯ thÆ°á»ng, in Ä‘áº­m, cÄƒn Ä‘á»u, khÃ´ng thá»¥t, cá»¡ 13." & vbCrLf & _
        "Tiá»ƒu má»¥c tá»‘i Ä‘a 4 cáº¥p; má»—i nhÃ³m pháº£i cÃ³ Ã­t nháº¥t 2 má»¥c." & vbCrLf & _
        "Báº£ng/hÃ¬nh: sá»‘ gáº¯n vá»›i chÆ°Æ¡ng (HÃ¬nh 3.4). Äáº§u Ä‘á» báº£ng á»Ÿ TRÃŠN, hÃ¬nh á»Ÿ DÆ¯á»šI." & vbCrLf & _
        "TrÃ­ch dáº«n: ngoáº·c vuÃ´ng theo sá»‘ thá»© tá»±, vÃ­ dá»¥ [15, tr.32-35]." & vbCrLf & vbCrLf & _
        "CÃ´ng cá»¥ tá»± sá»­a: khá»• giáº¥y, lá», phÃ´ng, cá»¡ chá»¯, cÄƒn lá», thá»¥t dÃ²ng, giÃ£n dÃ²ng, " & _
        "cÃ¡ch Ä‘oáº¡n, header, sá»‘ trang, gáº¡ch dÃ i." & vbCrLf & _
        "Pháº£i tá»± sá»­a: sá»‘ báº£ng/hÃ¬nh theo chÆ°Æ¡ng, nhÃ³m chá»‰ 1 tiá»ƒu má»¥c, ""báº£ng dÆ°á»›i Ä‘Ã¢y"", " & _
        "danh má»¥c tÃ i liá»‡u tham kháº£o, báº£ng chá»¯ viáº¿t táº¯t."
    MsgBox s, vbInformation, "Chuáº©n NTU"
End Sub

' Nut [Chuan hoa luan an]: 1-CLICK du 5 buoc (QD 900 muc 2.b).
Public Sub RibbonChuanHoa(control As IRibbonControl)
    If Application.Documents.Count = 0 Then
        MsgBox "ChÆ°a cÃ³ tÃ i liá»‡u nÃ o Ä‘ang má»Ÿ.", vbExclamation, "Chuáº©n NTU"
        Exit Sub
    End If

    Dim d As Document
    Set d = ActiveDocument

    Dim msg As String
    msg = "Chuáº©n hÃ³a tÃ i liá»‡u theo quy cÃ¡ch TrÆ°á»ng Äáº¡i há»c Nha Trang" & vbCrLf & _
          "(QÄ 900/QÄ-ÄHNT ngÃ y 31/7/2023):" & vbCrLf & vbCrLf & _
          "1. Khá»• A4 dá»c Â· lá» trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm" & vbCrLf & _
          "2. ThÃ¢n bÃ i Times New Roman 13 pt, cÄƒn Ä‘á»u, giÃ£n 1,5, thá»¥t dÃ²ng Ä‘áº§u, cÃ¡ch Ä‘oáº¡n 6/0 pt" & vbCrLf & _
          "3. TiÃªu Ä‘á» chÆ°Æ¡ng in hoa Ä‘áº­m 14 pt cÄƒn giá»¯a; má»¥c/tiá»ƒu má»¥c Ä‘áº­m 13 pt" & vbCrLf & _
          "4. Caption ""Báº£ng X.Y"" trÃªn báº£ng, ""HÃ¬nh X.Y"" dÆ°á»›i hÃ¬nh, kÃ¨m dÃ²ng Nguá»“n" & vbCrLf & _
          "5. Sá»‘ trang La MÃ£ á»Ÿ pháº§n Ä‘áº§u, áº¢ Ráº­p tá»« chÆ°Æ¡ng 1 Â· bá» header" & vbCrLf & vbCrLf & _
          "Thao tÃ¡c nÃ y sá»­a trá»±c tiáº¿p tÃ i liá»‡u Ä‘ang má»Ÿ (Ctrl+Z hoÃ n tÃ¡c Ä‘Æ°á»£c). Tiáº¿p tá»¥c?"
    If MsgBox(msg, vbOKCancel + vbQuestion, "Chuáº©n NTU") <> vbOK Then Exit Sub

    Dim rep As String
    rep = NTU_ChuanHoaFull(d, OptSize(), OptKeepHeading(), True)

    ' Buoc 5: tu soat lai roi tom tat chung mot hop thoai.
    Dim kq As String
    kq = NTU_KiemTraSilent()
    MsgBox rep & vbCrLf & "----- Tá»± kiá»ƒm tra -----" & vbCrLf & kq, vbInformation, "Chuáº©n NTU"
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

    s = InputBox("Giá»¯ nguyÃªn cá»¡ chá»¯ cá»§a tiÃªu Ä‘á» Ä‘Ã£ Ä‘áº·t tay (1 = cÃ³, 0 = Ä‘Æ°a táº¥t cáº£ vá» 13)." & vbCrLf & _
                 "Quy cÃ¡ch quy Ä‘á»‹nh: tiÃªu Ä‘á» chÆ°Æ¡ng 14, má»¥c/tiá»ƒu má»¥c 13.", "Chuáº©n NTU â€” tÃ¹y chá»n", IIf(OptKeepHeading(), "1", "0"))
    If StrPtr(s) = 0 Then Exit Sub
    SaveSetting APP, "Opt", "KeepHeading", IIf(Trim$(s) = "0", "0", "1")

    MsgBox "ÄÃ£ lÆ°u: thÃ¢n bÃ i 13 pt Â· tiÃªu Ä‘á» " & _
           IIf(OptKeepHeading(), "giá»¯ cá»¡ riÃªng", "Ä‘Æ°a vá» 13 pt") & ".", vbInformation, "Chuáº©n NTU"
End Sub

' ---------------------------------------------------------------- Vao viec

Public Sub NTU_ChuanHoa()
    Dim d As Document
    If Application.Documents.Count = 0 Then
        MsgBox "ChÆ°a cÃ³ tÃ i liá»‡u nÃ o Ä‘ang má»Ÿ.", vbExclamation, "Chuáº©n NTU"
        Exit Sub
    End If
    Set d = ActiveDocument

    Dim msg As String
    msg = "Chuáº©n hÃ³a tÃ i liá»‡u theo quy cÃ¡ch TrÆ°á»ng Äáº¡i há»c Nha Trang" & vbCrLf & _
          "(QÄ 900/QÄ-ÄHNT ngÃ y 31/7/2023):" & vbCrLf & vbCrLf & _
          "â€¢ Times New Roman 13 pt, cÄƒn Ä‘á»u 2 bÃªn, giÃ£n dÃ²ng 1,5" & vbCrLf & _
          "â€¢ DÃ²ng Ä‘áº§u má»—i Ä‘oáº¡n thá»¥t 1 cm; cÃ¡ch Ä‘oáº¡n 6 pt trÃªn, 0 pt dÆ°á»›i" & vbCrLf & _
          "â€¢ A4 Â· lá» trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm" & vbCrLf & _
          "â€¢ Bá» header Â· sá»‘ trang á»Ÿ giá»¯a phÃ­a DÆ¯á»šI má»—i trang" & vbCrLf & _
          "â€¢ Gáº¡ch dÃ i (â€” â€“ â€’ â€•) Ä‘á»•i thÃ nh gáº¡ch ná»‘i (-)" & vbCrLf & vbCrLf & _
          "Thao tÃ¡c nÃ y sá»­a trá»±c tiáº¿p tÃ i liá»‡u Ä‘ang má»Ÿ (Ctrl+Z hoÃ n tÃ¡c Ä‘Æ°á»£c). Tiáº¿p tá»¥c?"
    If MsgBox(msg, vbOKCancel + vbQuestion, "Chuáº©n NTU") <> vbOK Then Exit Sub

    Dim rep As String
    rep = NTU_Apply(d, OptSize(), OptKeepHeading(), True)
    MsgBox rep, vbInformation, "Chuáº©n NTU"
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
        rep = rep & "! Lá» / khá»• giáº¥y: " & Err.Description & vbCrLf
        Trace "step le ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "âœ“ Khá»• A4 Â· lá» trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm" & vbCrLf
    End If
    On Error GoTo Loi

    ' 2. Phong chu + co chu + gian dong
    lastStep = "chu"
    On Error Resume Next
    Err.Clear
    n = NormalizeBody(d, sz, keepH)
    If Err.Number <> 0 Then
        rep = rep & "! PhÃ´ng / cá»¡ chá»¯: " & Err.Description & vbCrLf
        Trace "step chu ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "âœ“ ÄÃ£ Ä‘á»‹nh dáº¡ng " & Format$(n, "#,##0") & " Ä‘oáº¡n" & _
          IIf(keepH, " (giá»¯ cá»¡ chá»¯ cá»§a tiÃªu Ä‘á»)", "") & vbCrLf
    On Error GoTo Loi

    ' 3. Kieu Normal de go tiep cung dung chuan
    lastStep = "style"
    On Error Resume Next
    Err.Clear
    SetNormalStyle d, sz
    If Err.Number <> 0 Then
        rep = rep & "! Kiá»ƒu Normal: " & Err.Description & vbCrLf
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
        rep = rep & "! Sá»‘ trang: " & Err.Description & vbCrLf
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
        rep = rep & "! Gáº¡ch dÃ i: " & Err.Description & vbCrLf
        Trace "step gachdai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    ElseIf n > 0 Then
        rep = rep & "âœ“ ÄÃ£ Ä‘á»•i " & Format$(n, "#,##0") & " dáº¥u gáº¡ch dÃ i thÃ nh gáº¡ch ná»‘i (-)" & vbCrLf
    Else
        rep = rep & "âœ“ KhÃ´ng cÃ³ dáº¥u gáº¡ch dÃ i nÃ o cáº§n Ä‘á»•i" & vbCrLf
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
    NTU_Apply = rep & vbCrLf & "Xong. Báº¥m [Kiá»ƒm Tra] Ä‘á»ƒ soÃ¡t láº¡i."
    Exit Function
Loi:
    Trace "FATAL at " & lastStep & ": " & Err.Description
    Application.ScreenUpdating = True
    NTU_Apply = rep & "! Lá»—i á»Ÿ bÆ°á»›c [" & lastStep & "]: " & Err.Description
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

    ' 1. Phong chu cho ca range mot lan â€” nhanh va khong phu thuoc tung doan.
    On Error Resume Next
    Err.Clear
    SetRunFont r
    If Err.Number <> 0 Then
        Trace "SetRunFont range ERR " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' 2. Dinh dang tung doan. Mot doan loi (o bang gop, field, textbox)
    '    khong duoc phep huy ca tai lieu â€” dem vao bad roi di tiep.
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
    ' NameFarEast bi tu choi (loi 5844) tren mot so tai lieu â€” bo qua, khong lam hong ca buoc.
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
        If InStr(nm, "heading 1") > 0 Or InStr(nm, "tiÃªu Ä‘á» 1") > 0 Then HeadingLevel = 1 : Err.Clear : Exit Function
        If InStr(nm, "heading") > 0 Or InStr(nm, "tiÃªu Ä‘á»") > 0 Then HeadingLevel = 2 : Err.Clear : Exit Function
    End If
    Err.Clear

    ' "ChÆ°Æ¡ng 1" / "CHÆ¯Æ NG 1" go tay
    t = CleanText(p.Range.Text)
    Err.Clear
    If Len(t) <= 60 Then
        If Left$(LCase$(t), 6) = "chÆ°Æ¡ng" Or Left$(LCase$(t), 5) = "chuong" Then HeadingLevel = 1 : Err.Clear : Exit Function
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

    ' Chi co ky tu dieu khien/dau cach/so thu tu dong â€” tuc la khong co chu that.
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

    SetPageNumberFooter = "âœ“ Sá»‘ trang: Ä‘áº·t á»Ÿ giá»¯a phÃ­a dÆ°á»›i (" & done & " chÃ¢n trang) Â· Ä‘Ã£ bá» header"
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
        MsgBox "ChÆ°a cÃ³ tÃ i liá»‡u nÃ o Ä‘ang má»Ÿ.", vbExclamation, "Chuáº©n NTU"
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
        MsgBox rep, vbInformation, "Chuáº©n NTU â€” káº¿t quáº£ kiá»ƒm tra"
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
        s = InputBox("Nháº­p sá»‘ thá»© tá»± lá»—i cáº§n xem (1 - " & LastDemLoi & ")." & vbCrLf & _
                     "Bá» trá»‘ng hoáº·c Cancel Ä‘á»ƒ thÃ´i.", "Chuáº©n NTU â€” nháº£y tá»›i lá»—i")
        If Len(Trim$(s)) = 0 Then Exit Do
        If IsNumeric(s) Then
            k = CLng(Val(s))
            If k >= 1 And k <= LastDemLoi Then
                On Error Resume Next
                LastGoc.Activate
                LastChoLoi(k).Select
                If Err.Number <> 0 Then
                    MsgBox "KhÃ´ng nháº£y tá»›i Ä‘Æ°á»£c má»¥c nÃ y (tÃ i liá»‡u Ä‘Ã£ thay Ä‘á»•i?).", vbExclamation, "Chuáº©n NTU"
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
        rep = "Lá» / khá»• giáº¥y: Ä‘Ãºng"
    Else
        rep = "Lá» / khá»• giáº¥y: SAI â€” cáº§n trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm, khá»• A4 dá»c"
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
        line = "PhÃ´ng chá»¯ khÃ¡c Times New Roman: " & badFont
        If firstFont <> "" Then line = line & " (vd: " & firstFont & ")"
        AddL rep, line
    End If
    If badSize > 0 Then
        line = "Cá»¡ chá»¯ khÃ¡c " & sz & " pt: " & badSize
        If firstSize <> "" Then line = line & " (vd: " & firstSize & ")"
        AddL rep, line
    End If
    If badLine > 0 Then AddL rep, "Äoáº¡n chÆ°a dÃ£n dÃ²ng Ä‘Ãºng (thÃ¢n bÃ i 1,5 Â· trong báº£ng 1): " & badLine
    If badAlign > 0 Then AddL rep, "Äoáº¡n chÆ°a cÄƒn Ä‘á»u 2 bÃªn: " & badAlign
    If badIndent > 0 Then AddL rep, "Äoáº¡n chÆ°a thá»¥t dÃ²ng Ä‘áº§u 1 cm: " & badIndent
    If badHead > 0 Then AddL rep, "TiÃªu Ä‘á» sai cá»¡ chá»¯ (chÆ°Æ¡ng 14 Â· má»¥c 13): " & badHead
    If badHeadAlign > 0 Then AddL rep, "TiÃªu Ä‘á» chÆ°Æ¡ng chÆ°a cÄƒn giá»¯a: " & badHeadAlign

    footerOk = HasPageNumberBottom(d)
    If Not footerOk Then AddL rep, "Sá»‘ trang: chÆ°a Ä‘áº·t á»Ÿ giá»¯a phÃ­a dÆ°á»›i"

    If CountHeadersWithText(d) > 0 Then AddL rep, "Header cÃ²n ná»™i dung (quy cÃ¡ch khÃ´ng dÃ¹ng header)"

    dashCount = CountDashes(d)
    If dashCount > 0 Then AddL rep, "CÃ²n dáº¥u gáº¡ch dÃ i: " & dashCount & " (dÃ¹ng gáº¡ch ná»‘i -)"

    deepCount = CountDeepHeadings(d)
    If deepCount > 0 Then AddL rep, "Má»¥c quÃ¡ 4 cáº¥p: " & deepCount

    refCount = CountBadRefs(d)
    If refCount > 0 Then AddL rep, "Chá»— viáº¿t ""báº£ng/hÃ¬nh dÆ°á»›i Ä‘Ã¢y"" (pháº£i nÃªu sá»‘): " & refCount

    rep = rep & vbCrLf & String(58, "-")
    If badFont = 0 And badSize = 0 And badLine = 0 And badAlign = 0 And badIndent = 0 _
       And badHead = 0 And badHeadAlign = 0 And footerOk And dashCount = 0 _
       And CountHeadersWithText(d) = 0 Then
        rep = rep & vbCrLf & "Äáº T thá»ƒ thá»©c. CÃ²n pháº£i tá»± soÃ¡t: Ä‘Ã¡nh sá»‘ báº£ng/hÃ¬nh theo chÆ°Æ¡ng, danh má»¥c tÃ i liá»‡u tham kháº£o, sá»‘ trang La MÃ£ á»Ÿ pháº§n Ä‘áº§u."
    Else
        rep = rep & vbCrLf & "Báº¥m [Chuáº©n HÃ³a] Ä‘á»ƒ sá»­a cÃ¡c má»¥c thá»ƒ thá»©c á»Ÿ trÃªn; má»¥c ná»™i dung pháº£i sá»­a tay."
    End If
    If deepCount > 0 Or refCount > 0 Then
        rep = rep & vbCrLf & "CÃ²n " & (deepCount + refCount) & " lá»—i ná»™i dung â€” khÃ´ng tá»± sá»­a Ä‘Æ°á»£c."
    End If

    AddL rep, ""
    rep = "Káº¾T QUáº¢ KIá»‚M TRA - chuáº©n NTU (QÄ 900/QÄ-ÄHNT 31/7/2023)" & rep
    rep = rep & vbCrLf & "(Ä‘Ã£ soÃ¡t " & n & " Ä‘oáº¡n, " & headCount & " tiÃªu Ä‘á»)"
    Application.ScreenUpdating = True
    NTU_Check = rep
    Exit Function
Loi:
    Application.ScreenUpdating = True
    NTU_Check = "Lá»—i khi kiá»ƒm tra: " & Err.Description
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
    a = "báº£ng"
    b = "dÆ°á»›i Ä‘Ã¢y"
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
        rep = rep & "! Khá»• giáº¥y / lá» / kiá»ƒu Normal: " & Err.Description & vbCrLf
        Trace "step kho/le ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "âœ“ A4 dá»c Â· lá» trÃªn 2 Â· dÆ°á»›i 2 Â· trÃ¡i 3 Â· pháº£i 2 cm Â· kiá»ƒu Normal Ä‘Ã£ Ä‘áº·t" & vbCrLf
    End If
    On Error GoTo Cleanup

    ' 2. Than bai
    lastStep = "than bai"
    On Error Resume Next
    Err.Clear
    n = NormalizeBody(d, sz, keepH)
    If Err.Number <> 0 Then
        rep = rep & "! ThÃ¢n bÃ i: " & Err.Description & vbCrLf
        Trace "step than bai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "âœ“ ÄÃ£ Ä‘á»‹nh dáº¡ng " & Format$(n, "#,##0") & " Ä‘oáº¡n" & _
          IIf(keepH, " (giá»¯ cá»¡ chá»¯ cá»§a tiÃªu Ä‘á»)", "") & vbCrLf
    On Error GoTo Cleanup

    ' 3. Tieu de muc
    lastStep = "tieu de"
    On Error Resume Next
    Err.Clear
    ApplyHeadingStyles d
    If Err.Number <> 0 Then
        rep = rep & "! TiÃªu Ä‘á» má»¥c: " & Err.Description & vbCrLf
        Trace "step tieu de ERR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        rep = rep & "âœ“ TiÃªu Ä‘á»: chÆ°Æ¡ng IN HOA Ä‘áº­m 14 cÄƒn giá»¯a Â· má»¥c Ä‘áº­m 13 khÃ´ng thá»¥t dÃ²ng" & vbCrLf
    End If
    On Error GoTo Cleanup

    ' 4. Caption bang/hinh
    lastStep = "caption"
    On Error Resume Next
    Err.Clear
    ApplyCaptions d
    If Err.Number <> 0 Then
        rep = rep & "! Caption báº£ng/hÃ¬nh: " & Err.Description & vbCrLf
        Trace "step caption ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    rep = rep & "âœ“ ÄÃ£ chÃ¨n " & LastCaptionCount & " caption báº£ng/hÃ¬nh cÃ²n thiáº¿u" & vbCrLf
    On Error GoTo Cleanup

    ' 5. So trang theo section (La Ma phan dau, A Rap phan than)
    lastStep = "so trang"
    On Error Resume Next
    Err.Clear
    ApplySectionPageNumbers d
    If Err.Number <> 0 Then
        rep = rep & "! Sá»‘ trang theo pháº§n: " & Err.Description & vbCrLf
        Trace "step so trang ERR " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    soSec = 0
    soSec = d.Sections.Count
    rep = rep & "âœ“ Sá»‘ trang: " & soSec & " pháº§n (section), Ä‘Ã¡nh sá»‘ á»Ÿ giá»¯a phÃ­a dÆ°á»›i" & vbCrLf
    On Error GoTo Cleanup

    ' 6. Gach dai
    lastStep = "gach dai"
    On Error Resume Next
    Err.Clear
    n = FixDashes(d, rep)
    If Err.Number <> 0 Then
        rep = rep & "! Gáº¡ch dÃ i: " & Err.Description & vbCrLf
        Trace "step gach dai ERR " & Err.Number & " " & Err.Description
        Err.Clear
    ElseIf n > 0 Then
        rep = rep & "âœ“ ÄÃ£ Ä‘á»•i " & Format$(n, "#,##0") & " gáº¡ch dÃ i thÃ nh gáº¡ch ná»‘i (-)" & vbCrLf
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
    NTU_ChuanHoaFull = rep & vbCrLf & "Xong. Báº¥m [Kiá»ƒm tra] Ä‘á»ƒ soÃ¡t láº¡i vÃ  xem lá»—i cÃ³ Ä‘Ã¡nh sá»‘."
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
    NTU_ChuanHoaFull = rep & "! Lá»—i á»Ÿ bÆ°á»›c [" & lastStep & "]: " & em
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
    IsNguonLine = (Left$(s, 6) = "nguá»“n:" Or Left$(s, 6) = "nguon:")
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
    If CaptionNo(t, "Báº£ng") > 0 Then
        IsCaptionPara = True
        Exit Function
    End If
    If CaptionNo(t, "HÃ¬nh") > 0 Then IsCaptionPara = True
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
                datNo = CaptionNo(t, "Báº£ng")
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
                datNo = CaptionNo(t, "HÃ¬nh")
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
        txt = IIf(kind = "B", "Báº£ng ", "HÃ¬nh ") & ch & "." & datNo
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
        MsgBox "KhÃ´ng tÃ¬m tháº¥y tiÃªu Ä‘á» chÆ°Æ¡ng (Heading 1 hoáº·c Ä‘oáº¡n má»Ÿ Ä‘áº§u ""CHÆ¯Æ NG ..."") nÃªn chÆ°a Ä‘áº·t láº¡i sá»‘ trang theo pháº§n." & vbCrLf & _
               "Quy cÃ¡ch: pháº§n Ä‘áº§u Ä‘Ã¡nh sá»‘ La MÃ£ (i, ii...), tá»« chÆ°Æ¡ng 1 Ä‘Ã¡nh sá»‘ 1, 2, 3...", vbExclamation, "Chuáº©n NTU"
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
        MsgBox "ChÆ°a cÃ³ tÃ i liá»‡u nÃ o Ä‘ang má»Ÿ.", vbExclamation, "Chuáº©n NTU"
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
            hienTai = "Cá»¡ " & FmtSo(p.Range.Font.Size) & " Â· Ä‘áº­m: " & YesNo(p.Range.Font.Bold) & _
                      " Â· in hoa: " & YesNo(p.Range.Font.AllCaps) & _
                      " Â· cÄƒn: " & TenCan(p.Range.ParagraphFormat.Alignment)
            rows.Add Array("ChÆ°Æ¡ng " & no(nCh), Snippet80(p.Range.Text), hienTai, "In hoa, Ä‘áº­m, 14 pt, cÄƒn giá»¯a, khÃ´ng thá»¥t dÃ²ng")
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
                            rows.Add Array("TiÃªu Ä‘á» (Ä‘oáº¡n " & n & ")", Snippet80(p.Range.Text), _
                                           "H" & prev & " â†’ H" & lvl, "KhÃ´ng nháº£y cáº¥p: thÃªm má»¥c cáº¥p trung gian")
                        End If
                    End If
                End If
                prev = lvl
            End If
            If lvl = 0 And Not p.Range.Information(wdWithInTable) And Not IsCaptionPara(p) Then
                t = p.Range.Text
                If p.Range.Font.Name <> FONT_NAME And p.Range.Font.Name <> "" Then
                    soFont = soFont + 1
                    If rows.Count < 300 Then rows.Add Array("Äoáº¡n " & n, Snippet80(t), _
                        "PhÃ´ng " & p.Range.Font.Name, "Times New Roman")
                End If
                If Abs(p.Range.Font.Size - 13) > 0.5 Then
                    soCo = soCo + 1
                    If rows.Count < 300 Then rows.Add Array("Äoáº¡n " & n, Snippet80(t), _
                        "Cá»¡ " & FmtSo(p.Range.Font.Size), "13 pt")
                End If
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 Then
                    soDong = soDong + 1
                    If rows.Count < 300 Then rows.Add Array("Äoáº¡n " & n, Snippet80(t), _
                        "GiÃ£n dÃ²ng " & TenGian(p.Range.ParagraphFormat.LineSpacingRule), "GiÃ£n dÃ²ng 1,5")
                End If
                If p.Range.ParagraphFormat.Alignment <> wdAlignParagraphJustify Then
                    soCan = soCan + 1
                    If rows.Count < 300 Then rows.Add Array("Äoáº¡n " & n, Snippet80(t), _
                        "CÄƒn " & TenCan(p.Range.ParagraphFormat.Alignment), "CÄƒn Ä‘á»u 2 bÃªn")
                End If
                If Abs(p.Range.ParagraphFormat.FirstLineIndent - Cm(OptIndent())) > 3 Then
                    soThut = soThut + 1
                    If rows.Count < 300 Then rows.Add Array("Äoáº¡n " & n, Snippet80(t), _
                        "Thá»¥t dÃ²ng Ä‘áº§u " & FmtSo(p.Range.ParagraphFormat.FirstLineIndent / Cm(1#)) & " cm", _
                        "Thá»¥t dÃ²ng Ä‘áº§u " & Format$(OptIndent(), "0.00") & " cm")
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
            If CaptionNo(CleanText(t), "Báº£ng") = 0 Then
                thieuCap = thieuCap + 1
                If rows.Count < 300 Then
                    rows.Add Array("Báº£ng #" & soBang & " (" & src.Tables(i).Rows.Count & "Ã—" & src.Tables(i).Columns.Count & ")", _
                                   Snippet80(t), "KhÃ´ng cÃ³ caption", "ChÃ¨n [Báº£ng X.Y] ngay TRÃŠN báº£ng, Ä‘áº­m 13, cÄƒn giá»¯a")
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
            If CaptionNo(CleanText(t), "HÃ¬nh") = 0 Then
                thieuCap = thieuCap + 1
                If rows.Count < 300 Then
                    rows.Add Array("HÃ¬nh #" & soHinh, "", "KhÃ´ng cÃ³ caption", _
                                   "ChÃ¨n [HÃ¬nh X.Y] ngay DÆ¯á»šI hÃ¬nh, Ä‘áº­m 13, cÄƒn giá»¯a")
                End If
            End If
        End If
    Next

    ' 4. Cua so nhap
    Set kq = MoCuaSoNhap()
    If kq Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "KhÃ´ng má»Ÿ Ä‘Æ°á»£c cá»­a sá»• nhÃ¡p Ä‘á»ƒ ghi bÃ¡o cÃ¡o.", vbExclamation, "Chuáº©n NTU"
        Exit Sub
    End If
    kq.Content.InsertAfter "BÃO CÃO PHÃ‚N TÃCH â€” QUY CÃCH NTU" & vbCr & _
        "(QÄ 900/QÄ-ÄHNT ngÃ y 31/7/2023) Â· " & src.Name & " Â· " & Format$(Now, "dd/mm/yyyy hh:nn") & vbCr & _
        "Cháº¿ Ä‘á»™ CHá»ˆ Äá»ŒC: khÃ´ng sá»­a gÃ¬ trong tÃ i liá»‡u gá»‘c. Báº£ng chi tiáº¿t á»Ÿ dÆ°á»›i." & vbCr & vbCr & _
        "Tá»”NG Káº¾T" & vbCr & _
        "- ChÆ°Æ¡ng (Heading 1): " & nCh & vbCr & _
        "- Báº£ng: " & soBang & " Â· HÃ¬nh: " & soHinh & vbCr & _
        "- Báº£ng/hÃ¬nh thiáº¿u caption: " & thieuCap & vbCr & _
        "- TiÃªu Ä‘á» nháº£y cáº¥p: " & nhayCap & vbCr & _
        "- Äoáº¡n lá»‡ch phÃ´ng: " & soFont & " Â· lá»‡ch cá»¡ chá»¯: " & soCo & _
        " Â· lá»‡ch giÃ£n dÃ²ng: " & soDong & " Â· lá»‡ch cÄƒn lá»: " & soCan & _
        " Â· lá»‡ch thá»¥t dÃ²ng: " & soThut & vbCr & _
        "- áº¢nh: " & src.InlineShapes.Count & " â€” DPI KHÃ”NG xÃ¡c Ä‘á»‹nh Ä‘Æ°á»£c (Word khÃ´ng phÆ¡i DPI cho InlineShape)." & vbCr & vbCr

    Set r = kq.Content
    r.Collapse wdCollapseEnd
    Dim tbl As Table
    Set tbl = kq.Tables.Add(r, rows.Count + 1, 4)
    tbl.Borders.Enable = True
    tbl.Cell(1, 1).Range.Text = "Vá»‹ trÃ­"
    tbl.Cell(1, 2).Range.Text = "Äoáº¡n (80 kÃ½ tá»±)"
    tbl.Cell(1, 3).Range.Text = "Hiá»‡n táº¡i"
    tbl.Cell(1, 4).Range.Text = "Gá»£i Ã½"
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
    MsgBox "Lá»—i khi phÃ¢n tÃ­ch: " & Err.Description, vbExclamation, "Chuáº©n NTU"
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
        YesNo = "cÃ³"
    Else
        YesNo = "khÃ´ng"
    End If
End Function

Private Function TenCan(ByVal al As Long) As String
    Select Case al
        Case wdAlignParagraphLeft
            TenCan = "trÃ¡i"
        Case wdAlignParagraphRight
            TenCan = "pháº£i"
        Case wdAlignParagraphCenter
            TenCan = "giá»¯a"
        Case wdAlignParagraphJustify
            TenCan = "Ä‘á»u 2 bÃªn"
        Case Else
            TenCan = "khÃ¡c (" & al & ")"
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
            TenGian = "tá»‘i thiá»ƒu"
        Case wdLineSpaceExactly
            TenGian = "chÃ­nh xÃ¡c"
        Case wdLineSpaceMultiple
            TenGian = "bá»™i sá»‘"
        Case Else
            TenGian = "khÃ¡c (" & v & ")"
    End Select
End Function

' ============================================================================
' TU CAP NHAT (doc lap voi add-in Word core; dung manifest version.json cua repo)
' ============================================================================

Public Function NTU_PhienBan() As String
    NTU_PhienBan = "1.0.0"
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
            MsgBox "KhÃ´ng káº¿t ná»‘i Ä‘Æ°á»£c mÃ¡y chá»§ cáº­p nháº­t (GitHub). Kiá»ƒm tra máº¡ng rá»“i thá»­ láº¡i." & vbCrLf & _
                   "PhiÃªn báº£n Ä‘ang dÃ¹ng: " & NTU_PhienBan(), vbExclamation, "Cáº­p nháº­t Chuáº©n NTU"
        End If
        Exit Sub
    End If
    daDocDuoc = True
    If NTU_SoSanh(ver, NTU_PhienBan()) <= 0 Then
        If hoiKhiKhongCo Then
            MsgBox "Báº¡n Ä‘ang dÃ¹ng phiÃªn báº£n má»›i nháº¥t: " & NTU_PhienBan(), vbInformation, "Cáº­p nháº­t Chuáº©n NTU"
        End If
        Exit Sub
    End If
    If Not NTU_DungFileAddIn() Then
        Dim q As String
        q = Chr$(39)
        MsgBox "CÃ³ phiÃªn báº£n má»›i " & ver & " nhÆ°ng macro Ä‘ang cháº¡y tá»« " & q & NTU_FileTuChay() & q & " - khÃ´ng pháº£i tá»‡p add-in Ä‘Ã£ cÃ i." & vbCrLf & "HÃ£y cÃ i láº¡i báº±ng bá»™ cÃ i Ä‘áº·t.", vbExclamation, "Cáº­p nháº­t Chuáº©n NTU"
        Exit Sub
    End If
    Dim ans As VbMsgBoxResult
    ans = MsgBox("CÃ³ phiÃªn báº£n má»›i: " & ver & " (Ä‘ang dÃ¹ng " & NTU_PhienBan() & ")." & vbCrLf & vbCrLf & _
                 notes & vbCrLf & vbCrLf & "Cáº­p nháº­t ngay?", vbYesNo + vbQuestion, "Cáº­p nháº­t Chuáº©n NTU")
    If ans <> vbYes Then Exit Sub
    Dim msg As String
    msg = NTU_TaiVaCai(url, sha, kt)
    If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cáº­p nháº­t Chuáº©n NTU"
End Sub

Private Function NTU_TaiVaCai(ByVal url As String, ByVal sha As String, _
                              Optional ByVal kichThuoc As Long = 0) As String
    ' Tai ban moi, kiem tinh toan ven (ZIP magic + kich thuoc), roi de mot .cmd
    ' cho Word tat han moi copy de â€” khong the ghi de tep .dotm dang mo.
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
    NTU_TaiVaCai = "ÄÃ£ táº£i báº£n má»›i. HÃ£y ÄÃ“NG Háº¾T cá»­a sá»• Word: báº£n má»›i sáº½ tá»± cÃ i Ä‘áº·t, má»Ÿ láº¡i Word lÃ  dÃ¹ng Ä‘Æ°á»£c."
    Exit Function
WinBad:
    On Error Resume Next
    Close #f
    If Len(moiW) > 0 Then Kill moiW
    NTU_TaiVaCai = "Tá»‡p táº£i vá» khÃ´ng Ä‘Ãºng (thiáº¿u hoáº·c há»ng). ChÆ°a thay Ä‘á»•i gÃ¬, hÃ£y thá»­ láº¡i sau."
    Exit Function
WinFail:
    NTU_TaiVaCai = "Táº£i báº£n cáº­p nháº­t tháº¥t báº¡i (kiá»ƒm tra máº¡ng). ChÆ°a thay Ä‘á»•i gÃ¬."
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

    ' (i)(ii)(iii) phong / co chu / gian dong â€” quet lai de lay vung tung loi.
    For Each p In d.Content.Paragraphs
        If Not IsTableMark(p) Then
            n = n + 1
            t = CleanText(p.Range.Text)
            lvl = HeadingLevel(p)

            If p.Range.Font.Name <> "" And p.Range.Font.Name <> FONT_NAME Then
                GhiLoi p.Range, chiTiet, "PhÃ´ng " & p.Range.Font.Name & " (cáº§n " & FONT_NAME & ")"
            ElseIf p.Range.Font.Name = "" And Len(t) > 0 Then
                GhiLoi p.Range, chiTiet, "PhÃ´ng khÃ´ng Ä‘á»“ng nháº¥t trong Ä‘oáº¡n (cáº§n " & FONT_NAME & ")"
            End If

            ' Co chu hop le: than bai = sz; tieu de chuong 14; caption 11 hoac 13;
            ' trong bang 13. Ngoai bo do la loi.
            If Not CoChuHopLe(p, lvl, sz) Then
                If p.Range.Font.Size >= 9999999 Then
                    GhiLoi p.Range, chiTiet, "Cá»¡ chá»¯ khÃ´ng Ä‘á»“ng nháº¥t trong Ä‘oáº¡n"
                Else
                    GhiLoi p.Range, chiTiet, "Cá»¡ chá»¯ " & FmtSo(p.Range.Font.Size) & " pt (cáº§n " & _
                           IIf(lvl = 1, "14", FmtSo(sz)) & " pt" & _
                           IIf(IsCaptionPara(p), "; caption 11 hoáº·c 13", "") & ")"
                End If
            End If

            If lvl = 0 And Not p.Range.Information(wdWithInTable) Then
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 Then
                    GhiLoi p.Range, chiTiet, "GiÃ£n dÃ²ng " & TenGian(p.Range.ParagraphFormat.LineSpacingRule) & " (cáº§n 1,5)"
                End If
            ElseIf p.Range.Information(wdWithInTable) Then
                If p.Range.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Then
                    GhiLoi p.Range, chiTiet, "GiÃ£n dÃ²ng trong báº£ng chÆ°a báº±ng 1,0"
                End If
            End If

            ' (v) nhay cap tieu de: H1 -> H3 tro len la loi.
            If lvl >= 3 And lvlTruoc = 1 Then
                GhiLoi p.Range, chiTiet, "Nháº£y cáº¥p tiÃªu Ä‘á»: cáº¥p 1 -> cáº¥p " & lvl & " (thiáº¿u cáº¥p 2)"
            End If
            If lvl > 0 Then lvlTruoc = lvl
        End If
    Next p

    ' (iv) bang/hinh thieu caption.
    For i = 1 To d.Tables.Count
        Set p = ParaBeforeTable(d.Tables(i))
        If Not p Is Nothing Then
            t = CleanText(p.Range.Text)
            If Len(Trim$(t)) = 0 Or InStr(1, t, "Báº£ng", vbTextCompare) = 0 Then
                GhiLoi d.Tables(i).Range, chiTiet, "Báº£ng chÆ°a cÃ³ caption ""Báº£ng X.Y"" ngay TRÃŠN báº£ng"
            End If
        End If
    Next i

    For i = 1 To d.InlineShapes.Count
        Set p = ParaAfterShape(d.InlineShapes(i))
        If Not p Is Nothing Then
            t = CleanText(p.Range.Text)
            If Len(Trim$(t)) = 0 Or InStr(1, t, "HÃ¬nh", vbTextCompare) = 0 Then
                GhiLoi d.InlineShapes(i).Range, chiTiet, "HÃ¬nh chÆ°a cÃ³ caption ""HÃ¬nh X.Y"" ngay DÆ¯á»šI hÃ¬nh"
            End If
        End If
    Next i

    ' (vi) do phan giai anh: Word KHONG cho doc DPI qua COM -> noi thang la
    ' khong xac dinh duoc, khong doan bua.
    coAnh = d.InlineShapes.Count
    If coAnh > 0 Then
        chiTiet = chiTiet & "  - áº¢nh: " & coAnh & " hÃ¬nh â€” Ä‘á»™ phÃ¢n giáº£i khÃ´ng xÃ¡c Ä‘á»‹nh Ä‘Æ°á»£c qua COM;" & _
                  " má»Ÿ Word, chuá»™t pháº£i > Size Ä‘á»ƒ tá»± soÃ¡t (cáº§n >= 300 dpi khi in)." & vbCrLf
    End If

    If LastDemLoi = 0 Then
        chiTiet = "  (khÃ´ng cÃ³ má»¥c nÃ o cáº§n chá»‰ tay)"
    Else
        chiTiet = "  Nháº­p sá»‘ trong [1.." & LastDemLoi & "] Ä‘á»ƒ nháº£y tá»›i lá»—i." & vbCrLf & chiTiet
    End If

    rep = rep & vbCrLf & String(58, "-") & vbCrLf & _
          "CHI TIáº¾T Lá»–I (Ä‘Ã£ soÃ¡t " & n & " Ä‘oáº¡n)" & vbCrLf & chiTiet

    Application.ScreenUpdating = True
    NTU_CheckDayDu = rep
    Exit Function
Loi:
    Application.ScreenUpdating = True
    NTU_CheckDayDu = rep & vbCrLf & "! Lá»—i khi kiá»ƒm tra chi tiáº¿t: " & Err.Description
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
