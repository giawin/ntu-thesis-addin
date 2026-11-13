# Quy cách QĐ 900/QĐ-ĐHNT (31/7/2023) → ánh xạ sang thủ tục VBA

Mô-đun: `NTU.vba` (`Attribute VB_Name = "NTU"`), ribbon khai trong `customUI.xml`.
Hằng số dùng chung: `APP = "NTUChuanHoa"`, `FONT_NAME = "Times New Roman"`, `MAX_DEPTH = 4`.

## 1. Thể thức trang giấy

| Quy định | Thủ tục | Cách kiểm chứng |
|---|---|---|
| A4 dọc 210×297 mm | `OptimizePageSetup` | `d.PageWidth = 21cm`, `d.PageHeight = 29.7cm`, `d.PageSetup.Orientation = wdOrientPortrait` |
| Lề: trên 2,0 / dưới 2,0 / trái 3,0 / phải 2,0 cm | `OptimizePageSetup` | `d.PageSetup.TopMargin = 2`, `BottomMargin = 2`, `LeftMargin = 3`, `RightMargin = 2` (cm) |
| Không dùng header | `OptimizePageSetup` + `SetPageNumberFooter` | `d.Sections(i).Headers(wdHeaderFooterPrimary).Range.Text = ""` |
| Số trang ở giữa, phía dưới, field `PAGE` | `SetPageNumberFooter` | `Footers(wdHeaderFooterPrimary)` chứa field `wdFieldPage`, `ParagraphFormat.Alignment = wdAlignParagraphCenter` |

## 2. Phông và cỡ chữ

| Quy định | Thủ tục | Cách kiểm chứng |
|---|---|---|
| Times New Roman toàn văn | `SetRunFont`, `SetNormalStyle` | `p.Range.Font.Name = "Times New Roman"` |
| 13 pt CỐ ĐỊNH cho thân bài | `OptSize` trả về hằng `13`, `NormalizeRange` | `Abs(p.Range.Font.Size - 13) <= 0.5` |
| Tiêu đề chương 14 pt đậm in hoa căn giữa | `ApplyHeadingStyles`, `UpperChapterText` | `Font.Size = 14`, `Font.Bold = True`, `ParagraphFormat.Alignment = wdAlignParagraphCenter`, `text = UCase(text)` |
| Tiêu đề mục cấp 2–4: 13 pt đậm, không thụt đầu dòng, căn đều | `ApplyHeadingStyles` (`HeadingLevel >= 2`) | `Font.Size = 13`, `Font.Bold = True`, `FirstLineIndent = 0` |
| Caption bảng/hình: 13 pt đậm căn giữa | `StyleCaptionRange` | `Font.Size = 13`, `Font.Bold = True`, `Alignment = center` |
| Dòng "Nguồn:" dưới caption: 11 pt nghiêng căn trái | `StyleSourceLine`, `IsNguonLine` | `Font.Size = 11`, `Font.Italic = True`, `Alignment = wdAlignParagraphLeft`, `FirstLineIndent = 0` |

## 3. Đoạn văn

| Quy định | Thủ tục | Cách kiểm chứng |
|---|---|---|
| Giãn dòng 1,5 cho văn bản thường | `SetNormalStyle`, `NormalizeRange` | `ParagraphFormat.LineSpacingRule = wdLineSpace1pt5` |
| Giãn dòng 1,0 cho bảng / bìa / tài liệu tham khảo / phụ lục | `NormalizeRange`, `NTU_CheckDayDu` | `wdLineSpaceSingle` khi `p.Range.Information(wdWithInTable)` |
| Căn đều hai bên | `NormalizeRange` | `ParagraphFormat.Alignment = wdAlignParagraphJustify` |
| Thụt dòng đầu 1,0 cm (tùy chọn 1,25) | `OptIndent` đọc `GetSetting(APP, "Opt", "Indent", "1")` | `Abs(FirstLineIndent - CentimetersToPoints(v)) <= 3` |
| Cách đoạn: trước 6 pt, sau 0 pt | `SetNormalStyle` | `SpaceBefore = 6`, `SpaceAfter = 0` |

## 4. Bảng và hình

| Quy định | Thủ tục | Cách kiểm chứng |
|---|---|---|
| Caption bảng "Bảng X.Y" đặt TRÊN bảng | `ChenCaptionTren`, `ParaBeforeTable`, `CaptionNo` | Đoạn ngay trước `Tables(i)` bắt đầu bằng "Bảng"; X = số chương |
| Caption hình "Hình X.Y" đặt DƯỚI hình | `ChenCaptionDuoi`, `ParaAfterShape`, `CaptionNo` | Đoạn ngay sau `InlineShapes(i)` bắt đầu bằng "Hình" |
| Chèn caption không sinh trùng (idempotent) | `ApplyCaptions` + biến module `LastCaptionCount` | Chạy 2 lần: lần 2 không thêm đoạn nào, `LastCaptionCount` không tăng |
| Đánh số X theo chương | `ChapterNoFromText`, `ChuongAt` | Quét tiêu đề chương phía trước vị trí bảng/hình |

## 5. Chi tiết khác

| Quy định | Thủ tục | Cách kiểm chứng |
|---|---|---|
| Gạch dài (U+2012/2013/2014/2015) → `-` | `FixDashes`, `ReplaceDashesIn` | `CountDashes` sau khi chuẩn hóa = 0 |
| Đánh số trang: phần 1 La Mã thường, phần 2 Ả Rập từ 1 | `ApplySectionPageNumbers` | Section 1 `PageNumberStyle = wdPageNumberStyleLowercaseRoman`, `RestartNumberingAtSection = True`, `StartingNumber = 1`; section 2 Ả Rập, restart 1; các section sau không restart |

## 6. Kiểm tra

| Nội dung | Thủ tục |
|---|---|
| Tổng hợp nhanh (đạt/không đạt + đếm) | `NTU_Check`, `NTU_KiemTraSilent` |
| Chi tiết có đánh số, nhảy tới lỗi, xuất .txt | `NTU_CheckDayDu`, `GhiLoi`, `NhayToiLoi`, `XuatTxt` |
| Sáu nhóm lỗi mới | phông ≠ TNR; cỡ chữ sai (trừ caption 11/13, tiêu đề 14, bảng 13); giãn dòng ≠ 1,5 ở thân bài; bảng/hình thiếu caption; nhảy cấp tiêu đề (1 → 3); ảnh không xác định được DPI (ghi thẳng "không xác định được qua COM", không đoán) |

## 7. Cơ chế cập nhật (nút "Cập nhật")

| Bước | Thủ tục |
|---|---|
| Phiên bản hiện tại | `NTU_PhienBan()` → `"1.0.2"` |
| URL manifest | `NTU_ManifestUrl()` → `…/giawin/ntu-thesis-addin/releases/latest/download/version.json` |
| Tải manifest (WinHttp 5.1, có `setTimeouts`) | `NTU_TaiManifest` |
| Đọc trường JSON (`version`, `ntu_url_win`, `ntu_sha256_win`, `ntu_size_win`, `notes_ascii`) | `NTU_ParseJson` |
| So sánh phiên bản | `NTU_SoSanh` |
| Chặn khi đang chạy tệp tạm | `NTU_FileTuChay`, `NTU_DungFileAddIn` |
| Tải + kiểm `PK` + sinh `.cmd` chờ WINWORD.EXE rồi `copy /Y` | `NTU_TaiVaCai` (`Environ("APPDATA") & "\NTUThesisAddin\"`) |
| Nút ribbon | `RibbonCapNhat` (worker: `NTU_KiemTraCapNhat`) |

## 8. Ràng buộc kỹ thuật

- Tệp UTF-8 tiếng Việt thật; `vba2bas.mjs` chuyển ký tự >127 thành `ChrW` trước khi nạp vào VBE.
- Dấu nháy kép bên trong chuỗi VBA phải nhân đôi.
- Không dùng UserForm (không có designer): cửa sổ nháp tạo bằng `Documents.Add` + `Tables.Add`.
- Vòng lặp xóa/chèn chạy ngược (For i = Count To 1 Step -1) để không lệch chỉ số.
- Bỏ qua đoạn đánh dấu ô bảng (`IsTableMark`).
- Mọi thủ tục dài dùng `On Error GoTo Cleanup` + nhãn `Cleanup:` phục hồi `ScreenUpdating` / `DisplayAlerts`.