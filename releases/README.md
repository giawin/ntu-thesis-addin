# Quy trình phát hành

Thư mục này chứa tài liệu hướng dẫn phát hành. Bản thân thư mục không chứa tệp
đóng gói — các tệp `.dotm` được sinh ra bởi quy trình đóng gói cục bộ rồi được
workflow đọc trực tiếp từ hai thư mục `dist/`.

## Vì sao phải đóng gói trên máy

Tệp `.dotm` và `.xlam` là gói Open XML có chứa `vbaProject.bin`. Muốn tạo hoặc
sửa phần VBA bên trong, bắt buộc phải có Microsoft Word/Excel thật đang chạy,
vì thư viện VBA của Office chỉ hoạt động trong tiến trình Office. GitHub Actions
không có Word, cũng không có cách nào hợp pháp để tự động hóa việc đó trên
runner. Do vậy:

- **Đóng gói**: chạy trên máy Windows có Microsoft Office.
- **Kiểm tra, tính mã băm, sinh `version.json`, tạo bản phát hành**: chạy trên
  GitHub Actions.

Ranh giới này là có chủ đích: nó giữ cho workflow phát hành chạy được ở chế độ
chỉ đọc đối với mã nguồn, và không bao giờ tự sinh ra một tệp `.dotm` giả.

## Thứ tự công việc trên máy

Chạy trong thư mục gốc của kho, bằng Windows PowerShell.

### 1. Đóng gói add-in Word (bản cá nhân hóa)

```powershell
cscript //nologo src\build\word\build-word.vbs
powershell -NoProfile -ExecutionPolicy Bypass -File src\build\word\pack-word.ps1
```

`build-word.vbs` mở Microsoft Word ẩn, nạp mã VBA đã cá nhân hóa vào tệp gốc
rồi lưu ra `src\core\word\dist\GiaTwinH_NTU_Thesis_Word.dotm`. `pack-word.ps1`
chèn `src\core\word\customUI.xml` vào tệp đó, vì Word không tự lưu Ribbon khi
ghi tệp.

Kết quả: `src\core\word\dist\GiaTwinH_NTU_Thesis_Word.dotm`

### 2. Đóng gói add-in NTU Thesis

```powershell
cscript //nologo src\build\ntu\build-ntu.vbs
powershell -NoProfile -ExecutionPolicy Bypass -File src\build\ntu\pack-ntu.ps1
```

Kết quả: `src\modules\ntu-thesis\dist\NTUThesis.dotm`

### 3. Kiểm tra bằng mắt

Mở từng tệp `.dotm` vừa tạo, xác nhận tab lệnh hiện đúng, bấm thử từng nút trên
một tài liệu nháp. Không bỏ qua bước này: mã băm chỉ chứng minh tệp không bị
thay đổi sau khi tạo, không chứng minh tệp chạy đúng.

### 4. Ghi nhận thay đổi

Cập nhật `CHANGELOG.md` theo mục phiên bản mới. Phần mô tả ngắn dùng cho hộp
thoại cập nhật phải là **ASCII thuần, không dấu** — workflow sẽ đưa nguyên khối
đó vào trường `notes_ascii` của `version.json`.

### 5. Commit và tạo thẻ

```powershell
git add src\core\word\dist\GiaTwinH_NTU_Thesis_Word.dotm
git add src\modules\ntu-thesis\dist\NTUThesis.dotm
git add CHANGELOG.md
git commit -m "chore(release): 1.0.1"
git tag v1.0.1
git push origin main --tags
```

Hai thư mục `dist/` nằm trong danh sách ngoại lệ của `.gitignore`, nên tệp đóng
gói được đưa lên kho như mọi tệp khác. Đây là lựa chọn có chủ đích: workflow
phát hành cần tệp nhị phân đã kiểm tra, và việc commit chúng cho phép truy vết
đúng phiên bản đã phát hành.

### 6. Workflow tiếp nhận

Đẩy thẻ `v*` sẽ kích hoạt `release.yml`. Workflow lần lượt:

1. Xác định số phiên bản từ tên thẻ, bắt buộc dạng `X.Y.Z`.
2. Chạy kiểm thử `src/build/vba2bas.test.mjs`.
3. Tập hợp sáu tệp vào `dist/`:

   | Tệp | Nguồn |
   |---|---|
   | `GiaTwinH_NTU_Thesis_Word.dotm` | `src/core/word/dist/` |
   | `NTUThesis.dotm` | `src/modules/ntu-thesis/dist/` |
   | `ChuanHoa_Excel.xlam` | `src/core/vendor/upstream/` (bản gốc, không sửa) |
   | `CaiDat.bat` | `src/build/installer/` |
   | `GoCaiDat.bat` | `src/build/installer/` |
   | `HuongDan_CaiDat.txt` | thư mục gốc kho |

4. Đối chiếu mã băm SHA-256 của `ChuanHoa_Excel.xlam` với bản trong kho, nhằm
   phát hiện mọi chỉnh sửa ngoài ý muốn đối với tệp gốc của tác giả.
5. Tính mã băm và kích thước của cả sáu tệp.
6. Sinh `dist/version.json`.
7. Đóng gói `out/giatwin_h-ntu-thesis-addin-vX.Y.Z.zip`.
8. Tạo bản phát hành trên GitHub, đính kèm sáu tệp riêng lẻ, tệp `.zip` và
   `version.json`. Nội dung mô tả lấy từ `CHANGELOG.md`.

## Nội dung `version.json`

Đây là tệp mà cả hai add-in đọc khi người dùng bấm "Cập nhật". Cấu trúc cố định,
giữ đúng thứ tự khoá:

```json
{
  "version": "1.0.1",
  "notes_ascii": "Sua loi chuan hoa tieu de chuong.",
  "word_url":       "https://github.com/giawin/ntu-thesis-addin/releases/download/v1.0.1/GiaTwinH_NTU_Thesis_Word.dotm",
  "word_sha256":    "<64 ký tự hex>",
  "word_size":      123456,
  "ntu_url":        "https://github.com/giawin/ntu-thesis-addin/releases/download/v1.0.1/NTUThesis.dotm",
  "ntu_sha256":     "<64 ký tự hex>",
  "ntu_size":       234567,
  "xlam_url":       "https://github.com/giawin/ntu-thesis-addin/releases/download/v1.0.1/ChuanHoa_Excel.xlam",
  "xlam_sha256":    "<64 ký tự hex>",
  "xlam_size":      44684
}
```

Hai add-in đọc hai nhóm khoá khác nhau nên có thể cập nhật độc lập:

- `GiaTwinH_NTU_Thesis_Word.dotm` đọc `word_*`.
- `NTUThesis.dotm` đọc `ntu_*`.

Nhóm `xlam_*` được ghi để phục vụ tra cứu và đối chiếu; add-in Excel không tham
gia cơ chế tự cập nhật và được thay thế bằng cách cài lại.

Địa chỉ add-in dùng để kiểm tra phiên bản luôn là bản phát hành mới nhất:

```text
https://github.com/giawin/ntu-thesis-addin/releases/latest/download/version.json
```

## Các điểm cần lưu ý

**Không sửa `ChuanHoa_Excel.xlam`.** Đây là tác phẩm của tác giả gốc, được dùng
nguyên trạng. Workflow sẽ dừng nếu phát hiện mã băm khác với bản trong kho. Mọi
yêu cầu chỉnh sửa hành vi của add-in này nên được chuyển thành góp ý cho tác
giả gốc, không tự sửa.

**Tên tệp là một phần của hợp đồng.** `CaiDat.bat`, `GoCaiDat.bat` và cơ chế tự
cập nhật đều dựa vào đúng sáu tên tệp nêu trên. Đổi tên tệp đòi hỏi sửa đồng
thời cả ba nơi: workflow, bộ cài, và hằng số trong mã VBA.

**Bản phát hành phải có đủ sáu tệp.** Workflow đặt `fail_on_unmatched_files`
thành `true`; thiếu một tệp là bản phát hành thất bại, không phải bản phát hành
thiếu tệp.

**Thẻ phiên bản không được xóa.** Địa chỉ `releases/latest/download/` phụ thuộc
vào bản phát hành mới nhất còn tồn tại. Xóa bản phát hành đang là mới nhất sẽ
làm hỏng chức năng cập nhật của mọi máy đã cài.

## Phát hành lại khi workflow lỗi

Nếu workflow thất bại sau khi thẻ đã được đẩy lên, có hai cách xử lý:

- Sửa nguyên nhân, xóa thẻ cả ở máy và ở kho từ xa, rồi tạo lại thẻ. Cách này
  giữ đúng số phiên bản.
- Chạy lại workflow bằng tay từ trang Actions, truyền số phiên bản vào ô nhập.

Bản phát hành đã tạo nhưng sai nội dung phải được xóa hẳn trước khi chạy lại,
nếu không GitHub sẽ từ chối ghi đè tệp cùng tên.

## Đồng bộ kho gốc

Workflow `sync-upstream.yml` chạy hằng ngày lúc 06:00 UTC. Workflow đọc
`version.json` của `CuongBC195/add-in-office`, so sánh với tệp
`src/core/vendor/upstream/VERSION`, và nếu khác thì tải hai tệp gốc, đối chiếu
mã băm SHA-256 rồi mở yêu cầu trộn chỉ chứa bốn tệp trong thư mục đó.

Yêu cầu trộn do workflow tạo **không được trộn tự động**. Phần VBA cá nhân hóa
nằm trong tệp `.bas` chứ không nằm trong `.dotm`, nên sau khi trộn phải chạy
`src/build/extract-vba.ps1` để trích lại mã và áp lại bảy điểm cá nhân hóa. Danh
sách kiểm tra nằm ngay trong thân yêu cầu trộn.

## Liên hệ

Câu hỏi về quy trình phát hành xin gửi qua mục Issues của kho
`giawin/ntu-thesis-addin`.