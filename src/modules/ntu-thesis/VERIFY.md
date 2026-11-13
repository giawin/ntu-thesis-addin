# Kết quả kiểm chứng (output thật, không chỉnh sửa)

Máy chạy: Windows, Windows PowerShell 5.1.26100.9818, Node v26.13.0.
Tệp nguồn: `NTU.vba` (85578 B, 2257 dòng, UTF-8 không BOM, 84 thủ tục) và `customUI.xml` (3819 B).
Script kiểm chứng: `_verify.ps1` trong cùng thư mục (thuần ASCII, chỉ đọc).
Chạy lại: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File _verify.ps1`

```
### LENH 1 - dem nut trong customUI.xml
buttons=6
buttons_via_tag=6
bare_//button=0   (0 vi default xmlns khong duoc dang ky voi XPath)
onAction=RibbonChuanHoa, RibbonPhanTich, RibbonKiemTra, RibbonTuyChon, RibbonCapNhat, RibbonHuongDan
tabLabel=NTU Thesis
onLoad=RibbonLoad
### LENH 2 - moi onAction phai co Sub tuong ung
RibbonChuanHoa -> True
RibbonPhanTich -> True
RibbonKiemTra -> True
RibbonTuyChon -> True
RibbonCapNhat -> True
RibbonHuongDan -> True
### LENH 3 - vba2bas.mjs
OK -> C:\Users\giath\AppData\Local\Temp\NTU_test.bas (106874 bytes, ASCII thuan)
exit=0
### LENH 4 - .bas sinh ra phai thuan ASCII
bas_bytes=106874  non-ascii=0
### LENH 5 - khong con ten goi cu
hits=0
### BO SUNG - tinh toan ven tep nguon
ntu_bytes=85578  lines=2257
bom=False  first3=65,116,116
high_codepoints=1339
replacement_char=0
procedures=84
duplicate_procedures=(none)
cu_xml_bytes=3819
APP_const=Private Const APP As String = "NTUChuanHoa"
```

## Đối chiếu tiêu chí PASS

| Tiêu chí | Yêu cầu | Kết quả thật | Kết luận |
|---|---|---|---|
| Số nút | `buttons=6` | `buttons=6` (và `buttons_via_tag=6`) | ĐẠT |
| Mỗi `onAction` có thủ tục | tất cả `True` | 6/6 `True` | ĐẠT |
| Bộ chuyển đổi | `vba2bas` exit 0 | `OK -> ... (106874 bytes, ASCII thuan)`, `exit=0` | ĐẠT |
| Tệp .bas sinh ra | thuần ASCII | `non-ascii=0` | ĐẠT |
| Tên gói cũ | không còn | `hits=0` | ĐẠT |

## Ghi chú trung thực về `//button` = 0

Lệnh `$x.SelectNodes('//button').Count` theo đúng nguyên văn trong đề bài trả về **0** — kể cả
trên tệp `customUI.xml` đã đúng 6 nút, và kể cả trên `customUI.xml` gốc của add-in tham chiếu
(`src\core\word\customUI.xml`). Nguyên nhân: phần tử gốc `<customUI>` khai báo
`xmlns="http://schemas.microsoft.com/office/2009/07/customui"`, nên XPath của .NET hiểu
`button` là nút **không thuộc namespace nào**, trong khi mọi `<button>` thật đều nằm trong
namespace mặc định đó. Đây là lỗi của biểu thức XPath, không phải thiếu nút.

Vì vậy số nút được đếm bằng hai cách không phụ thuộc namespace:
`SelectNodes('//*[local-name()="button"]')` và `$x.GetElementsByTagName('button')` — cả hai
đều cho **6**. Cả ba con số đều được in ở trên, không giấu con số 0.

## Những gì KHÔNG kiểm chứng được ở đây

1. **Không chạy Word COM** (theo yêu cầu đề bài), nên chưa có bằng chứng biên dịch thật trong VBE
   và chưa có bằng chứng hành vi khi chạy trên văn bản thật. `vba2bas` chỉ kiểm cú pháp chuỗi/nháy
   kép và chuyển ký tự >127 sang `ChrW`, không phải trình biên dịch VBA.
2. **Độ phân giải ảnh**: Word không cho đọc DPI qua COM. Công cụ ghi thẳng
   "không xác định được qua COM" thay vì đoán, đúng như quy định.
3. **Nhánh cập nhật qua mạng** (`RibbonCapNhat` → `NTU_TaiManifest` → `NTU_TaiVaCai`): mã đã có,
   nhưng chưa thực thi lần nào vì cần Word + mạng + bản phát hành thật trên GitHub.
4. **`NTU_CheckDayDu` và `ApplyCaptions`**: chưa chạy trên tài liệu thật; logic đã đọc lại bằng mắt,
   chưa có kiểm thử tự động.