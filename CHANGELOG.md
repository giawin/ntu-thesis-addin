# Lịch sử thay đổi

Mọi thay đổi đáng chú ý của bộ add-in được ghi tại tệp này.

Định dạng tuân theo [Keep a Changelog](https://keepachangelog.com/vi/1.1.0/),
phiên bản tuân theo [Semantic Versioning](https://semver.org/lang/vi/).

Quy ước: phần mô tả trong tệp này phải là **văn bản ASCII thuần**, không dấu.
Workflow phát hành đọc nguyên khối phiên bản để đưa vào trường `notes_ascii`
của `version.json`, và add-in hiển thị nội dung đó trong hộp thoại cập nhật.
Dòng có dấu tiếng Việt vẫn được giữ ở phần thân tài liệu, nhưng sẽ không được
đưa vào `notes_ascii` — hãy đặt phần mô tả ngắn bằng ASCII ở đầu mỗi mục.

## [Chưa phát hành]

### Đã sửa
- `version.json` sinh khoá `xlam_url` / `xlam_sha256` / `xlam_size` trong khi
  add-in Excel nguyên bản đọc `excel_url` / `excel_sha256` / `excel_size`.
  Hệ quả: nút tự cập nhật trong Excel nhận chuỗi rỗng và dừng lặng lẽ.
  Nay manifest dùng đúng tên khoá `excel_*`.
- Thêm `src/build/verify-manifest-contract.ps1` chạy trong CI: nó trích khoá
  manifest từ chính mã VBA của cả ba add-in rồi đối chiếu với khoá workflow
  sinh ra, nên lỗi lệch tên khoá không thể tái diễn mà không làm CI đỏ.

### Sẽ bổ sung
- Bổ sung kiểm thử hồi quy cho `src/build/vba2bas.mjs` khi gặp tệp có ký tự
  ngoài bảng mã ASCII.

## [1.0.1] - 2026-11-13

Ban sua loi. Chi thay doi cach dong goi, khong doi hanh vi chuan hoa.

### Đã sửa
- **Add-in lõi Word không chạy được trên bản 1.0.0.** `src/build/word/build-word.vbs`
  nạp mã vào mô-đun bằng `CodeModule.DeleteLines` + `AddFromString`, cách này để
  dự án ở trạng thái không biên dịch được, khiến Word ẩn *toàn bộ* macro: dải
  lệnh vẫn hiện nhưng bấm nút không có phản ứng, và `Application.Run` báo
  "Không thể chạy macro đã nêu". Nay dùng `VBComponents.Import` — cùng cách
  `src/build/ntu/build-ntu.vbs` vẫn làm, và add-in NTU chưa bao giờ gặp lỗi này.
- `src/build/word/verify-word.vbs` chỉ kiểm sự có mặt của chuỗi trong mã, nên
  lỗi trên lọt qua toàn bộ pipeline, CI và bản phát hành 1.0.0. Nay có thêm pha
  kiểm tra thật: nạp add-in rồi gọi `CN_PhienBan`, `TT_ProfileCount`,
  `CN_ManifestUrl` và yêu cầu đúng giá trị. Bản 1.0.0 sẽ trượt bước này.
- **Danh sách chặn của Office.** Khi một add-in từng nạp thất bại, Word ghi
  đường dẫn vào `Resiliency\DisabledItems` và im lặng từ chối nạp lại vĩnh viễn.
  Cài lại không xoá được mục này. Thêm `DonDepBlacklist.vbs`: chỉ xoá mục trỏ
  tới tệp của bộ cài này, giữ nguyên mục của nhà cung cấp khác. `CaiDat.bat` và
  `GoCaiDat.bat` đều gọi nó.
- Banner `CaiDat.bat` ghi sai tên trường ("Nong Lam" thay vì "Nha Trang").

## [1.0.0] - 2026-11-13

Phien ban dau tien. Bo cai gom ba add-in, cai dat cung mot lan.

### Bổ sung
- **Add-in bộ cài**: `CaiDat.bat` và `GoCaiDat.bat` cài và gỡ cùng lúc ba
  thành phần vào `%APPDATA%\Microsoft\Word\STARTUP` và
  `%APPDATA%\Microsoft\Excel\XLSTART`, tự gọi `Unblock-File` để tránh bị
  chặn macro, không ghi vào Registry.
- **GiaTwinH_NTU_Thesis_Word.dotm**: cá nhân hóa từ dự án
  `CuongBC195/add-in-office`; tab "THỂ THỨC VĂN BẢN" trên dải lệnh Word,
  chuẩn hóa thể thức văn bản hành chính, có nút "Cập nhật" đọc các trường
  `word_*` từ `version.json`.
- **NTUThesis.dotm**: module mới, tab "NTU Thesis" với sáu lệnh — "Chuẩn hóa
  luận án", "Phân tích", "Kiểm tra", "Tùy chọn", "Cập nhật", "Hướng dẫn" —
  chuẩn hóa luận án theo Quyết định số 900/QĐ-ĐHNT ngày 31/7/2023, kèm cơ chế
  tự cập nhật riêng đọc các khoá `ntu_*` từ `version.json`.
- **ChuanHoa_Excel.xlam**: giữ nguyên bản gốc của tác giả, không sửa một byte;
  tab "CHUẨN HÓA IN" căn chỉnh trang in khổ A4.
- **Bộ chuyển đổi mã**: `src/build/vba2bas.mjs` chuyển tệp `.vba` UTF-8 thành
  `.bas` thuần ASCII, kèm bộ kiểm thử tự chứa.
- **Tích hợp liên tục**: workflow `ci.yml` chạy kiểm thử bộ chuyển đổi, kiểm
  tra bảng mã của mã nguồn, hợp lệ hóa Ribbon XML và sự tồn tại của tài liệu.
- **Phát hành**: workflow `release.yml` tính mã băm SHA-256, sinh
  `version.json`, đóng gói `.zip` và tạo bản phát hành trên GitHub.
- **Đồng bộ kho gốc**: workflow `sync-upstream.yml` kiểm tra bản phát hành của
  `CuongBC195/add-in-office` mỗi ngày, đối chiếu mã băm rồi mở yêu cầu trộn
  chỉ chứa bốn tệp trong `src/core/vendor/upstream/`.
- **Tài liệu**: `README.md`, `HuongDan_CaiDat.txt` (12 mục), `releases/README.md`.

### Ghi chú
- Ba add-in dùng chung một tệp `version.json`. Hai tệp `.dotm` đọc hai nhóm khoá khác
  nhau (`word_*` và `ntu_*`) nên có thể cập nhật độc lập.
- Trong phiên bản này, cả hai tệp `.dotm` đều có cơ chế đọc `version.json` và tự cập nhật;
  mỗi tệp chỉ đọc nhóm khoá của mình nên không phụ thuộc vào tệp còn lại.
- `ChuanHoa_Excel.xlam` không tham gia cơ chế tự cập nhật; muốn có bản mới thì
  cài lại từ bộ phát hành chính thức.