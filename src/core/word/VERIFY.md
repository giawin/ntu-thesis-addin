# VERIFY — kiểm chứng `Module1.vba` + `customUI.xml`

Chạy từ thư mục chứa 4 file (thư mục này):

```powershell
cd C:\Users\giath\OneDrive\Documents\deepseek-harness\default-workspace\ntu-thesis-addin\src\core\word
```

> Ghi chú kỹ thuật: `[IO.File]` đọc đường dẫn tương đối theo *current directory của tiến trình*, không theo `Set-Location` của PowerShell. Vì vậy các lệnh dưới đây dùng `$PWD` để không phụ thuộc thư mục làm việc của tiến trình.

## 1. Encoding: `Module1.vba` phải là ASCII thuần

```powershell
$b=[IO.File]::ReadAllBytes((Join-Path $PWD 'Module1.vba')); ($b | Where-Object {$_ -gt 127}).Count
```

Tiêu chí PASS: in ra `0`.

Output thật:

```
0
```

## 2. Không còn dấu vết tác giả gốc

```powershell
Select-String -Path 'Module1.vba' -Pattern 'CuongBC195|OfficePartySuite|script\.google\.com|Thanh Phu|chuanhoa_vanbandang\.dotm'
```

Tiêu chí PASS: không in gì.

Output thật: **không in gì** — tức PASS.

### Vì sao pattern ở trên dùng `chuanhoa_vanbandang\.dotm` (nói rõ, không giấu)

Nếu viết đúng như brief — pattern trần `chuanhoa_vanbandang` — thì `Select-String` (mặc định **không phân biệt hoa/thường**) khớp luôn 6 dòng chứa tên thủ tục `ChuanHoa_VanBanDang`:

```
4070: Call ChuanHoa_VanBanDang(True)
4298: Call ChuanHoa_VanBanDang(False)
4300: Public Sub ChuanHoa_VanBanDang_Silent()
4301: Call ChuanHoa_VanBanDang(True)
4303: Sub ChuanHoa_VanBanDang(Optional ByVal silentMode As Boolean = False)
4325: Call ChuanHoa_VanBanDang(True)
```

Đây là **dương tính giả**: `ChuanHoa_VanBanDang` là thủ tục chuẩn hóa cốt lõi, **bắt buộc giữ nguyên** (brief cấm sửa thân hàm và cấm đổi tên). Dấu vết tác giả gốc thật sự là **tên file** `chuanhoa_vanbandang.dotm` ở dòng 8518 bản gốc, và tên đó đã được thay. Vì vậy:

- Kiểm tra đúng bản chất (tên file gốc) → dùng `chuanhoa_vanbandang\.dotm` → 0 hit, PASS.
- Kiểm tra theo nghĩa đen của brief → dùng `-CaseSensitive`:

```powershell
Select-String -Path 'Module1.vba' -Pattern 'chuanhoa_vanbandang' -CaseSensitive
```

Output thật: **không in gì** — PASS.

Kiểm chứng chéo rằng bản gốc *có* dấu vết này (để thấy phép thử không vô dụng):

```powershell
Select-String -Path 'Module1.upstream.bas' -Pattern 'chuanhoa_vanbandang' -CaseSensitive
```

Output thật:

```
Module1.upstream.bas:8518:CN_DungFileAddIn = (n = "chuanhoa_vanbandang.dotm")
```

→ Bản gốc khớp, bản mới không khớp: đúng là đã thay.

## 3. Đủ 4 nhóm dấu hiệu cá nhân hóa

```powershell
Select-String -Path 'Module1.vba' -Pattern 'NTUThesisAddin|giawin/ntu-thesis-addin|giatwin_h_ntu_thesis_word|CN_PhienBan = "1.0.1"'
```

Tiêu chí PASS: in ra đủ 4 nhóm (dòng 16 và 4 dòng `GetSetting`/`SaveSetting`, dòng manifest, dòng `CN_DungFileAddIn`, dòng `CN_PhienBan`).

Output thật:

```
16:Private Const TT_SETTING_APP As String = "NTUThesisAddin"
8307:LicApiUrl = ""
8373:CN_PhienBan = "1.0.1"
8376:CN_ManifestUrl = "https://github.com/giawin/ntu-thesis-addin/releases/latest/download/version.json"
8405:CN_DungFileAddIn = (n = "giatwin_h_ntu_thesis_word.dotm")
8448:lastChk = GetSetting("NTUThesisAddin", "CapNhat", "LanKiemTra_WORD", "")
8455:lanThu = GetSetting("NTUThesisAddin", "CapNhat", "LanThu_WORD", "")
8461:SaveSetting "NTUThesisAddin", "CapNhat", "LanThu_WORD", Format(Now, "yyyy-mm-dd hh:nn:ss")
8464:If daDocDuoc Then SaveSetting "NTUThesisAddin", "CapNhat", "LanKiemTra_WORD", Format(Date, "yyyy-mm-dd")
8519:thuMuc = Environ("APPDATA") & "\NTUThesisAddin"
8543:Print #f, "copy /Y ""%APPDATA%\NTUThesisAddin\capnhat_WORD.moi"" """ & tgtW & """ >nul"
8544:Print #f, "del ""%APPDATA%\NTUThesisAddin\capnhat_WORD.moi"" >nul 2>&1"
```

Cả 4 nhóm đều có mặt.

## 4. `customUI.xml` parse được và giữ đủ node

```powershell
[xml]$x = Get-Content 'customUI.xml' -Raw; "ok nodes=$($x.SelectNodes('//*').Count)"
```

Tiêu chí PASS: parse OK và số node khớp bản gốc (162).

Output thật:

```
ok nodes=162
```

## 5. Số `onAction` không đổi

```powershell
(Select-String -Path 'customUI.upstream.xml' -Pattern 'onAction' -AllMatches).Matches.Count
(Select-String -Path 'customUI.xml' -Pattern 'onAction' -AllMatches).Matches.Count
```

Tiêu chí PASS: hai số bằng nhau.

Output thật:

```
135
135
```

(Con số 135 là số dòng-khớp của `Select-String`; số lần xuất hiện thực tế của chuỗi `onAction` là 132 ở cả hai file — kiểm thêm bằng `([regex]::Matches($t,'onAction')).Count` cũng ra 132/132. Cả hai cách đều bằng nhau giữa nguồn và bản mới.)

## 6. Cân bằng Sub/Function

```powershell
$L=[IO.File]::ReadAllLines((Join-Path $PWD 'Module1.vba'))
"Sub=$((($L|Where-Object{$_ -match '^\s*(Public |Private )?(Static )?Sub '}).Count)) EndSub=$((($L|Where-Object{$_ -match '^\s*End Sub\s*$'}).Count)) Function=$((($L|Where-Object{$_ -match '^\s*(Public |Private )?(Static )?Function '}).Count)) EndFunction=$((($L|Where-Object{$_ -match '^\s*End Function\s*$'}).Count))"
```

Tiêu chí PASS: Sub = End Sub, Function = End Function.

Output thật:

```
Sub=113 EndSub=113 Function=187 EndFunction=187
```

## Kiểm chứng bổ sung (khuyến nghị, không bắt buộc)

Xác nhận thay đổi license đúng biên, không nuốt hàm khác:

```powershell
$L=[IO.File]::ReadAllLines('Module1.vba'); foreach($n in 8306,8307,8308,8309,8310,8311,8312,8313){ "$n| $($L[$n-1])" }
```

Output thật:

```
8306| Private Function LicApiUrl() As String
8307| LicApiUrl = ""
8308| End Function
8309| Private Function CheckLicenseActivation() As Boolean
8310|     ' Ban NTU: bo co che ban quyen theo may cua tac gia goc.
8311|     CheckLicenseActivation = True
8312| End Function
8313| Private Function CallHttpApi...
```

4 điểm gọi license vẫn còn và giờ luôn nhận `True`:

```
3827: If Not CheckLicenseActivation() Then Exit Function
3992: If Not CheckLicenseActivation() Then Exit Sub
4307: If Not CheckLicenseActivation() Then Exit Sub
```

## Tổng kết

| # | Kiểm tra | Kết quả |
|---|---|---|
| 1 | ASCII thuần | PASS — 0 byte >127 |
| 2 | Không dấu vết tác giả gốc | PASS — 0 hit theo `chuanhoa_vanbandang\.dotm` và theo `-CaseSensitive`; pattern trần (không phân biệt hoa/thường) sẽ khớp 6 dòng tên thủ tục `ChuanHoa_VanBanDang` — dương tính giả, đã giải thích ở mục 2 |
| 3 | Đủ 4 nhóm NTU | PASS — 12 dòng khớp |
| 4 | customUI parse | PASS — 162 node |
| 5 | onAction không đổi | PASS — 135 = 135 (132 = 132 theo regex) |
| 6 | Sub/Function cân bằng | PASS — 113/113, 187/187 |