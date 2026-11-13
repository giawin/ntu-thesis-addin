# NTU Thesis Add-in

Bộ công cụ hỗ trợ soạn thảo dành cho Microsoft Word và Microsoft Excel, phục vụ việc chuẩn hóa
thể thức văn bản và trình bày luận án theo Quyết định số 900/QĐ-ĐHNT ngày 31/7/2023 của
Trường Đại học Nha Trang.

Phiên bản hiện hành: **1.0.1** — Mã định danh add-in: `giatwin_h-ntu-thesis-addin`.

Kho mã này là kho công khai (public). Toàn bộ tệp cài đặt được phát hành kèm theo trong mục
[Releases](../../releases).

## Giới thiệu

Bộ cài gồm ba add-in độc lập, được cài đặt trong cùng một lần và xuất hiện trên dải lệnh
(Ribbon) của Word và Excel ngay sau khi khởi động lại ứng dụng. Mỗi add-in đảm nhận một
nhóm công việc riêng, không chồng chéo lên nhau.

Mục tiêu thiết kế là giữ cho thao tác của người dùng ở mức tối thiểu: mở tài liệu, chọn
đúng tab, bấm một nút. Mọi thao tác định dạng phức tạp được xử lý bên trong add-in.

## Thành phần bộ cài

| Tệp | Nguồn | Chức năng | Cơ chế cập nhật |
|---|---|---|---|
| `GiaTwinH_NTU_Thesis_Word.dotm` | Cá nhân hóa từ `CuongBC195/add-in-office` | Chuẩn hóa thể thức văn bản Word — tab **THỂ THỨC VĂN BẢN** | Nút **Cập nhật** trong Word, đọc các khóa `word_*` từ `version.json` |
| `ChuanHoa_Excel.xlam` | **Bản gốc, không sửa một byte** | Căn chỉnh trang in A4 trong Excel — tab **CHUẨN HÓA IN** | Giữ nguyên cơ chế gốc của tác giả |
| `NTUThesis.dotm` | Module mới của nhóm thực hiện | Chuẩn hóa luận án NTU theo QĐ 900/QĐ-ĐHNT — tab **NTU Thesis**, 5 nút | Nút **Cập Nhật** trong Word, đọc các khóa `ntu_*` từ `version.json` |

Kèm theo bộ cài là `CaiDat.bat`, `GoCaiDat.bat` và `HuongDan_CaiDat.txt`.

## Yêu cầu hệ thống

- Hệ điều hành Windows 10 hoặc Windows 11.
- Microsoft Word và Microsoft Excel thuộc bộ Microsoft 365 hoặc Office 2016 trở lên.
- Quyền ghi vào thư mục hồ sơ người dùng (`%APPDATA%`). Không cần quyền quản trị.
- Kết nối Internet cho lần cài đặt đầu tiên và cho chức năng cập nhật.

## Cài đặt nhanh

1. Tải tệp `.zip` của bản phát hành mới nhất và giải nén ra một thư mục bất kỳ.
2. Bấm chuột phải vào tệp `CaiDat.bat`, chọn **Run as administrator** nếu máy yêu cầu.
3. Làm theo hướng dẫn hiển thị trong cửa sổ dòng lệnh.
4. Khởi động lại Microsoft Word và Microsoft Excel.

Hướng dẫn đầy đủ, kể cả phần xử lý sự cố, được trình bày trong tệp
[`HuongDan_CaiDat.txt`](HuongDan_CaiDat.txt).

## Ba tab lệnh

### Tab THỂ THỨC VĂN BẢN (Word)

Phục vụ việc trình bày văn bản hành chính theo Nghị định số 30/2020/NĐ-CP. Add-in hỗ trợ
nhiều khối thể thức khác nhau, bao gồm khối chính quyền, khối Công an nhân dân, khối Quân đội
nhân dân và khối Mặt trận Tổ quốc. Người dùng có thể để add-in tự nhận diện khối thể thức
theo dòng tiêu ngữ hoặc cơ quan ban hành, hoặc chọn thủ công.

### Tab CHUẨN HÓA IN (Excel)

Căn chỉnh vùng in theo khổ giấy A4, xử lý các tình huống bảng biểu bị tràn trang, dòng tiêu đề
bị ngắt và lề in không đồng nhất. Đây là bản gốc của tác giả, được giữ nguyên trạng nhằm bảo
toàn hành vi đã được kiểm chứng trong thực tế.

### Tab NTU Thesis (Word)

Sáu nút lệnh phục vụ việc chuẩn hóa luận án theo quy định của Trường Đại học Nha Trang:

| Nút | Chức năng |
|---|---|
| **Chuẩn hóa luận án** | Áp dụng toàn bộ quy cách trình bày luận án cho tài liệu đang mở: khổ giấy và lề, thân bài, tiêu đề chương, caption, số trang, gạch nối. Sửa trực tiếp tài liệu và hoàn tác được bằng `Ctrl+Z`. |
| **Phân tích** | Mở bảng liệt kê từng đoạn lệch chuẩn kèm vị trí, nội dung rút gọn, định dạng hiện tại và gợi ý sửa. Chỉ đọc, không sửa tài liệu. |
| **Kiểm tra** | Soát lỗi thể thức có đánh số để nhảy thẳng tới chỗ lỗi: phông, cỡ chữ, giãn dòng, caption thiếu, nhảy cấp tiêu đề, lề, khổ giấy, số trang, header, gạch dài. Không sửa nội dung. |
| **Tùy chọn** | Chọn thụt dòng đầu 1,0 hoặc 1,25 cm; có in hoa tiêu đề chương hay không; giữ nguyên hay đưa về 13 pt cỡ chữ tiêu đề đã đặt tay. |
| **Cập nhật** | Đọc `version.json` trên kho phát hành (khoá `ntu_*`), so với phiên bản đang dùng, tải và tự cài đặt khi đóng Word. Độc lập với add-in thể thức văn bản. |
| **Hướng dẫn** | Hiển thị bản tóm tắt quy định trình bày kèm theo Quyết định số 900/QĐ-ĐHNT ngày 31/7/2023. |
tu_*) và cài đặt nếu có. Hoạt động độc lập với add-in thể thức văn bản. |

Các tham số trình bày được áp dụng: phông Times New Roman 13 pt, căn đều hai bên, giãn dòng
1,5, dòng đầu thụt 1 cm; khổ A4 với lề trên 2 cm, dưới 2 cm, trái 3 cm, phải 2 cm; bỏ phần
đầu trang, số trang đặt giữa phía dưới; tiêu đề chương in hoa, đậm, cỡ 14 pt; các loại gạch
dài được chuyển về gạch nối.

## Cơ chế cập nhật

Một tệp `version.json` duy nhất tại địa chỉ phát hành mới nhất của kho mã phục vụ cả bộ
add-in. Tệp này do quy trình phát hành sinh ra tự động và chứa các trường sau:

| Khóa | Ý nghĩa |
|---|---|
| `version` | Số phiên bản hiện hành, dạng `X.Y.Z` |
| `notes_ascii` | Ghi chú phát hành, chỉ dùng ký tự ASCII |
| `word_url`, `word_sha256`, `word_size` | Địa chỉ, mã băm và dung lượng của `GiaTwinH_NTU_Thesis_Word.dotm` |
| `ntu_url`, `ntu_sha256`, `ntu_size` | Địa chỉ, mã băm và dung lượng của `NTUThesis.dotm` |
| `xlam_url`, `xlam_sha256`, `xlam_size` | Địa chỉ, mã băm và dung lượng của `ChuanHoa_Excel.xlam` |

Việc tách nhóm khóa theo tiền tố cho phép hai add-in Word cập nhật độc lập: mỗi add-in chỉ
đọc nhóm khóa của mình, nên có thể phát hành bản mới cho một thành phần mà không buộc thành
phần kia phải đổi theo. Nhóm `xlam_*` được ghi để tra cứu và đối chiếu; add-in Excel không
tham gia cơ chế tự cập nhật và được thay thế bằng cách cài lại.

Trình tự cập nhật của add-in Word: tải tệp mới về thư mục tạm, đối chiếu mã băm SHA-256,
sau đó tạo một tập lệnh chờ Microsoft Word đóng hẳn rồi mới thay thế tệp cũ. Nhờ vậy tệp
đang được ứng dụng giữ không bị ghi đè giữa chừng.

Tính đến phiên bản 1.0.1, cơ chế này được cài đặt trong cả hai tệp: `GiaTwinH_NTU_Thesis_Word.dotm`
đọc các khóa `word_*`, còn `NTUThesis.dotm` đọc các khóa `ntu_*`. Hai tệp dùng chung một
`version.json` nhưng hoàn toàn độc lập với nhau — không tệp nào gọi mã của tệp kia, nên có thể
cập nhật riêng từng tệp.

## Cấu trúc kho mã

```
ntu-thesis-addin/
├── .github/workflows/       Quy trình tích hợp, phát hành và đồng bộ kho gốc
├── releases/                Tài liệu quy trình phát hành
├── src/
│   ├── build/
│   │   ├── installer/       Bộ cài CaiDat.bat và GoCaiDat.bat
│   │   ├── ntu/             Đóng gói module luận án NTU
│   │   ├── word/            Đóng gói module Word
│   │   └── vba2bas.mjs      Bộ chuyển đổi mã VBA sang ASCII
│   ├── core/
│   │   ├── word/            Mã nguồn module Word và Ribbon
│   │   ├── excel/           Mã nguồn module Excel
│   │   └── vendor/upstream/ Bản gốc của tác giả, giữ nguyên trạng
│   └── modules/ntu-thesis/  Mã nguồn module luận án NTU
├── CHANGELOG.md             Lịch sử thay đổi
├── HuongDan_CaiDat.txt      Hướng dẫn cài đặt cho người dùng cuối
└── README.md
```

## Quy trình phát triển

Mã VBA được lưu dưới dạng tệp văn bản thuần để dễ so sánh khác biệt và dễ rà soát. Trước khi
đưa vào Microsoft Word, mã được chuyển qua công cụ `src/build/vba2bas.mjs`. Công cụ này thực
hiện hai việc: chuyển các ký tự tiếng Việt sang lời gọi `ChrW`, và xác nhận đầu ra là ASCII
thuần. Bước xác nhận này là bắt buộc, vì trình biên dịch VBA của Word từ chối các tệp `.bas`
chứa byte ngoài ASCII.

Quy trình đóng gói gồm bốn bước, thực hiện trên máy có cài Microsoft Word:

1. `cscript //nologo src/build/word/build-word.vbs` — tạo tệp `.dotm` từ mã nguồn.
2. `pwsh src/build/word/pack-word.ps1` — đưa Ribbon vào trong tệp vừa tạo.
3. `cscript //nologo src/build/ntu/build-ntu.vbs` — tạo tệp `.dotm` của module NTU.
4. `pwsh src/build/ntu/pack-ntu.ps1` — đưa Ribbon vào tệp vừa tạo.

Bước đưa Ribbon vào là bắt buộc vì Microsoft Word không lưu phần `customUI` khi tự ghi tệp.
Kết quả nằm ở `src/core/word/dist/GiaTwinH_NTU_Thesis_Word.dotm` và
`src/modules/ntu-thesis/dist/NTUThesis.dotm`; hai thư mục này được commit vì quy trình phát
hành đọc trực tiếp từ đó. Chi tiết được trình bày trong [`releases/README.md`](releases/README.md).

## Kiểm thử tự động

Quy trình tích hợp liên tục chạy trên GitHub Actions và không cần Microsoft Office. Quy trình
này kiểm tra bốn nhóm nội dung:

- Bộ chuyển đổi `vba2bas` vượt qua toàn bộ kiểm thử đi kèm.
- Mã nguồn `NTU.vba` chuyển đổi thành công và đầu ra là ASCII thuần.
- Tệp `customUI.xml` hợp lệ về cấu trúc và mọi nút lệnh đều khai báo thủ tục xử lý.
- Kho mã không lẫn tệp tạm của Office và không thiếu tệp tài liệu bắt buộc.

## Phát hành phiên bản

Quy trình phát hành được kích hoạt bằng cách đẩy một thẻ phiên bản:

```
git tag v1.0.1
git push origin v1.0.1
```

Trước khi tạo thẻ, hai tệp `.dotm` phải được đóng gói sẵn theo quy trình nêu trên và được
commit vào thư mục `dist/` tương ứng. Quy trình phát hành sẽ kiểm tra sự đầy đủ của sáu tệp,
tính mã băm SHA-256, sinh `version.json`, đóng gói tệp `.zip` và tạo bản phát hành trên
GitHub. Nếu thiếu bất kỳ tệp nào, quy trình dừng lại và báo rõ tệp còn thiếu thay vì phát
hành một bộ cài khiếm khuyết.

## Xử lý sự cố

| Hiện tượng | Nguyên nhân thường gặp | Cách xử lý |
|---|---|---|
| Không thấy tab sau khi cài | Word hoặc Excel chưa được khởi động lại | Đóng hẳn ứng dụng rồi mở lại |
| Báo lỗi macro bị chặn | Thiếu tin cậy thư mục STARTUP | Thêm thư mục vào Trust Center, hoặc bấm **Unblock** trong thuộc tính tệp |
| Tab hiện nhưng nút không phản hồi | Tệp bị chặn do tải từ Internet | Bấm chuột phải tệp, chọn **Properties**, tích **Unblock** |
| Chức năng cập nhật báo lỗi | Không có kết nối Internet hoặc bản phát hành chưa có `version.json` | Kiểm tra kết nối và sự tồn tại của bản phát hành |
| Excel báo lỗi khi mở tệp `.xlam` | Tệp bị chỉnh sửa ngoài ý muốn | Cài lại từ bộ phát hành chính thức |

## Gỡ cài đặt

Chạy `GoCaiDat.bat` trong thư mục cài đặt và khởi động lại Word, Excel. Tập lệnh này xóa ba
tệp add-in khỏi `%APPDATA%\Microsoft\Word\STARTUP` và `%APPDATA%\Microsoft\Excel\XLSTART`,
đồng thời xóa các tệp cấu hình đi kèm. Thao tác gỡ cài đặt không ảnh hưởng đến tài liệu của
người dùng.

## Quyền riêng tư

Add-in xử lý tài liệu ngay trên máy của người dùng. Không có nội dung tài liệu nào được gửi
ra ngoài. Kết nối mạng chỉ được dùng cho chức năng kiểm tra và tải bản cập nhật.

## Ghi công và giấy phép

Module Word được cá nhân hóa từ dự án `CuongBC195/add-in-office` của tác giả gốc. Phần lõi
chuẩn hóa thể thức văn bản và toàn bộ add-in Excel `ChuanHoa_Excel.xlam` thuộc quyền của tác
giả gốc và được sử dụng nguyên trạng, không chỉnh sửa. Sự ghi công này được giữ nguyên trong
mọi bản phát hành.

Module `NTUThesis.dotm` do nhóm thực hiện kho mã này phát triển.

## Liên hệ

Mọi góp ý, báo lỗi và đề nghị bổ sung tính năng xin gửi qua mục Issues của kho mã này.
Khi báo lỗi, vui lòng nêu rõ phiên bản add-in, phiên bản Office và các bước tái hiện.
