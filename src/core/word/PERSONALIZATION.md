# PERSONALIZATION — Module1.upstream.bas → Module1.vba

Nguồn: `Module1.upstream.bas` (8.688 dòng, 383.797 B, ASCII thuần, CRLF)
Kết quả: `Module1.vba` (8.575 dòng, 379.243 B, ASCII thuần, CRLF)

Chênh lệch dòng: **-113** = khối license bị thay (115 dòng → 2 dòng). Mọi thay đổi khác là thay tại chỗ (1 dòng → 1 dòng), không đổi số dòng.

Quy ước cột "Dòng mới": với các dòng **trước** khối license (dòng nguồn ≤ 8309) số dòng không đổi; với các dòng **sau** khối license, số dòng mới = số dòng cũ - 113.

## Bảng thay đổi

| # | Dòng gốc | Chuỗi gốc | Chuỗi mới | Dòng mới | Lý do |
|---|---|---|---|---|---|
| 1 | 16 | `Private Const TT_SETTING_APP As String = "OfficePartySuite"` | `Private Const TT_SETTING_APP As String = "NTUThesisAddin"` | 16 | Đổi tên app registry/settings; bắt buộc theo bảng thay thế |
| 2 | 5785 | `"Created by VPDU xa Thanh Phu"` (MsgBox báo cáo rà soát) | `"Created by Add-in NTU Thesis"` | 5785 | Xóa tên đơn vị gốc khỏi chữ ký hiển thị |
| 3 | 5789 | `"Created by VPDU xa Thanh Phu"` (MsgBox chuẩn hóa thành công) | `"Created by Add-in NTU Thesis"` | 5789 | Cùng lý do #2 |
| 4 | 8307 | `LicApiUrl = "https://script.google.com/macros/s/AKfycbx.../exec"` | `LicApiUrl = ""` | 8307 | Vô hiệu endpoint license của tác giả gốc; giữ nguyên `Private Function`/`End Function` |
| 5 | 8310–8424 | Toàn bộ thân `CheckLicenseActivation` (115 dòng, gồm `appName = "OfficePartySuite"` tại dòng 8314, `sectionName = "License"` 8315, các MsgBox `"Kich Hoat That Bai"` / `"Kich Hoat Thanh Cong"`, chuỗi `"lien he Van phong Dang uy xa Thanh Phu de duoc cap ma hop le."`, `lastCheckDate`, `savedKey`, `CallHttpApi(checkUrl)`, `ParseJsonField`…) | 2 dòng: `' Ban NTU: bo co che ban quyen theo may cua tac gia goc.` + `CheckLicenseActivation = True` | 8310–8311 | Bỏ cổng license theo máy. Hàm vẫn tồn tại và luôn trả `True` để 4 điểm gọi (`3827`, `3992`, `4307`, và chính nó) không đổi hành vi. Dòng khai báo 8309 và `End Function` (gốc 8425 → mới 8312) giữ nguyên |
| 6 | 8486 | `CN_PhienBan = "2026.11.12.2"` | `CN_PhienBan = "1.0.1"` | 8373 | Version riêng của bản NTU |
| 7 | 8489 | `"https://github.com/CuongBC195/add-in-office/releases/latest/download/version.json"` | `"https://github.com/giawin/ntu-thesis-addin/releases/latest/download/version.json"` | 8376 | Trỏ manifest cập nhật sang repo NTU |
| 8 | 8518 | `CN_DungFileAddIn = (n = "chuanhoa_vanbandang.dotm")` | `CN_DungFileAddIn = (n = "giatwin_h_ntu_thesis_word.dotm")` | 8405 | Nhận diện đúng tên file .dotm của bản NTU |
| 9 | 8561 | `GetSetting("OfficePartySuite", "CapNhat", "LanKiemTra_WORD", "")` | `GetSetting("NTUThesisAddin", "CapNhat", "LanKiemTra_WORD", "")` | 8448 | Đổi section settings, khớp #1 |
| 10 | 8568 | `GetSetting("OfficePartySuite", "CapNhat", "LanThu_WORD", "")` | `GetSetting("NTUThesisAddin", "CapNhat", "LanThu_WORD", "")` | 8455 | Cùng lý do #9 |
| 11 | 8574 | `SaveSetting "OfficePartySuite", "CapNhat", "LanThu_WORD", …` | `SaveSetting "NTUThesisAddin", "CapNhat", "LanThu_WORD", …` | 8461 | Cùng lý do #9 |
| 12 | 8577 | `SaveSetting "OfficePartySuite", "CapNhat", "LanKiemTra_WORD", …` | `SaveSetting "NTUThesisAddin", "CapNhat", "LanKiemTra_WORD", …` | 8464 | Cùng lý do #9 |
| 13 | 8594 | `"Cap nhat Add-in"` (tiêu đề MsgBox) | `"Cap nhat NTU Thesis Add-in"` | 8481 | Thương hiệu hóa tiêu đề hộp thoại cập nhật |
| 14 | 8599 | `"Cap nhat Add-in"` | `"Cap nhat NTU Thesis Add-in"` | 8486 | Cùng lý do #13 |
| 15 | 8604 | `"Cap nhat Add-in"` | `"Cap nhat NTU Thesis Add-in"` | 8491 | Cùng lý do #13 |
| 16 | 8608 | `"Cap nhat Add-in"` | `"Cap nhat NTU Thesis Add-in"` | 8495 | Cùng lý do #13 |
| 17 | 8612 | `"Cap nhat Add-in"` | `"Cap nhat NTU Thesis Add-in"` | 8499 | Cùng lý do #13 |
| 18 | 8632 | `thuMuc = Environ("APPDATA") & "\OfficePartySuite"` | `thuMuc = Environ("APPDATA") & "\NTUThesisAddin"` | 8519 | Thư mục tải bản cập nhật riêng của NTU |
| 19 | 8656 | `Print #f, "copy /Y ""%APPDATA%\OfficePartySuite\capnhat_WORD.moi"" …` | `… %APPDATA%\NTUThesisAddin\capnhat_WORD.moi …` | 8543 | Script .cmd phải trỏ đúng thư mục mới (khớp #18) |
| 20 | 8657 | `Print #f, "del ""%APPDATA%\OfficePartySuite\capnhat_WORD.moi"" >nul 2>&1"` | `… %APPDATA%\NTUThesisAddin\capnhat_WORD.moi …` | 8544 | Cùng lý do #19 |

Tổng: **20 mục** sửa đổi (18 phép thay chuỗi tại chỗ + 1 khối thân hàm thay bằng 2 dòng + 1 dòng `appName` biến mất cùng khối #5).

## customUI.xml

`customUI.upstream.xml` (25.687 B, UTF-8 **có BOM**, 218 dòng, 1 `<customUI>`, 1 `<ribbon>`, 2 `<tab>` (1 `idMso="TabHome"` + `tabVanBanDang`), 5 `<group>`, 129 `<button>`, 132 `onAction=`) **không chứa bất kỳ chuỗi thương hiệu nào** của tác giả gốc — đã kiểm bằng grep: `VPDU` 0, `Thanh Phu` 0, `OfficePartySuite` 0, `Office Party` 0, `CuongBC195` 0, `script.google.com` 0.

4 lần `Đảng ủy` và 2 lần `Ủy ban` đều là **ví dụ minh họa trong supertip** của dropdown đơn vị (chuỗi: `VD: Đảng ủy xã, Văn phòng Đảng ủy, Ban Xây dựng Đảng, Ủy ban Kiểm tra`), không phải tên đơn vị tác giả → giữ nguyên.

Do đó `customUI.xml` **giữ 100% nội dung**, chỉ chuẩn hóa encoding:

| Khía cạnh | Nguồn | Kết quả | Lý do |
|---|---|---|---|
| BOM | Có (`EF BB BF`) | Không | Word/Office đọc ribbon XML ổn định hơn khi UTF-8 không BOM |
| Kích thước | 25.687 B | 25.684 B | Đúng 3 B chênh = BOM bị bỏ; nội dung ký tự không đổi |
| `onAction` | 132 | 132 | Không mất liên kết macro nào |
| Node XML | 162 | 162 | Parse `[xml]` OK |

## Những gì CỐ Ý KHÔNG đổi

- `Attribute VB_Name = "Module1"` (dòng 1) — giữ nguyên để `onAction` trong ribbon vẫn gọi được.
- `LIC_CAN_NHAP_MA As Boolean = False` (dòng 15) — brief quy định giữ `False`.
- `TT_SETTING_SECTION = "TheThuc"` (dòng 17) — không nằm trong bảng thay thế.
- `VPDU` tại dòng **3174** (`Call TT_UnitSet(g, "VIETTAT", "VPDU")`) và **3291** (`d.Tables(1).Cell(3, 2).Range.text = "VPDU"`) — đây là **dữ liệu test** trong thủ tục kiểm thử của `TT_Profile`, không phải chuỗi thương hiệu hiển thị cho người dùng.
- Toàn bộ thân `ChuanHoa_VanBanDang`, `ChuanHoa_VanBanDang_Silent`, các `Ribbon*`, `Mau*`, `DonVi*`, `DVBH*`, `TT_*` — không sửa một byte.
- Toàn bộ hàm tiện ích dùng chung với `CN_KiemTraCapNhat` (`ParseJsonField`, `CallHttpApi`, `EncodeUrl`, `GetMachineId`…) — giữ nguyên vì tính năng cập nhật vẫn cần.
- Chuỗi tiếng Việt dạng `ChrW(...)` (3.917 lần) — giữ nguyên, không đụng tới.