# 工程表ツール（m5-gantt-addin）— Add-in Excel cho mục 5 (工程表/Gantt)

Add-in Excel tách riêng từ 施工計画書 Z1.html — **mục 5 (工程表)**. Nhập công việc ngay
trong task pane, sinh Gantt chart sang sheet **ガントチャート** dạng biểu đồ mũi tên
（矢印）vẽ bằng shape, tự tính **lịch nghỉ Nhật** (đúng thuật toán đã kiểm chứng trong
Z1.html: 18/18 ngày lễ 2026, gồm 振替休日・国民の休日).

## Cấu trúc

```
m5-gantt-addin/
├── manifest.xml       ← manifest của add-in (trỏ tới GitHub Pages, sideload qua registry)
├── taskpane.html/.js/.css ← panel điều khiển + logic (host trên GitHub Pages)
├── install.bat         ← cài đặt 1-click (không cần Node.js) — chỉ cần file này + manifest.xml
├── uninstall.bat        ← gỡ add-in (không cần Node.js)
├── install.cjs / uninstall.cjs ← bản Node.js tương đương (tùy chọn)
├── server.js, start-server.bat ← chỉ dùng khi PHÁT TRIỂN cục bộ (xem mục cuối)
└── assets/            ← icon
```

**Không cần build, không cần webpack, không cần node_modules.**

## Nội dung được host ở đâu

Toàn bộ taskpane (html/js/css/icon) được host tĩnh trên **GitHub Pages**, tại
`https://manh25xx08xx99-sys.github.io/Gantt-chart/` — build từ repo
[manh25xx08xx99-sys/Gantt-chart](https://github.com/manh25xx08xx99-sys/Gantt-chart),
nhánh `main`, thư mục gốc. `manifest.xml` đã trỏ sẵn tới URL này.

**Cập nhật tự động:** mỗi khi push code mới lên nhánh `main`, GitHub Pages tự
build lại (thường trong ~1 phút). Máy nào đã sideload `manifest.xml` sẽ tự lấy
bản mới nhất ở lần mở task pane tiếp theo — **không cần cài lại, không cần
đồng nghiệp làm gì thêm.**

## Cài đặt (máy mình hoặc máy đồng nghiệp)

Không cần Node.js, không cần chạy server — taskpane đã host sẵn trên GitHub
Pages, máy chỉ cần đăng ký 1 dòng registry trỏ vào `manifest.xml`.

**Gửi cho đồng nghiệp: chỉ cần 2 file, để chung 1 thư mục:**
- `manifest.xml`
- `install.bat`

Cách cài:

1. Copy 2 file trên (email, USB, chia sẻ file...) vào 1 thư mục bất kỳ trên máy
   cần cài — vị trí không quan trọng, miễn 2 file nằm **cùng chỗ**.
2. Double-click **`install.bat`** — tự đăng ký add-in vào Excel (chỉ ảnh hưởng
   user hiện tại trên máy đó, không cần quyền admin, không cần cài gì thêm).
3. Mở Excel → tab **Home** → nút **工程表ツール** (nhóm 施工計画書) → task pane mở ra.

Muốn gỡ: double-click **`uninstall.bat`** (cùng thư mục, cũng không cần Node.js).

> Có Node.js thì cũng có thể dùng `node install.cjs` / `node uninstall.cjs`
> (làm y hệt install.bat/uninstall.bat) — hoặc đăng ký thủ công qua PowerShell:
> ```powershell
> New-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer' `
>   -Name 'dfc3fd23-ff19-4d33-b312-a15d117dd27d' `
>   -Value '<đường dẫn đầy đủ tới manifest.xml>' -PropertyType String -Force
> ```
> (Đừng dùng `reg add` trong Git Bash — nó bị biến các switch `/v /t /d /f`
> thành đường dẫn.)

⚠ Đây vẫn là add-in **sideload kiểu dev** (không qua Store/M365 admin), nên mỗi
máy cần tự đăng ký 1 lần như trên; khác với trước là **không cần chạy server**
và **cập nhật code thì tự động** (nhờ GitHub Pages) — không cần cài lại khi có
bản mới.

## Cách dùng

1. Nhập công việc ngay trong task pane — mỗi dòng: **作業内容 / 開始日 / 終了日 /
   所要日数 / 備考 / 色** (chấm màu 7 màu; bấm chấm để đổi màu thanh).
   - **開始日・終了日**: chọn ngày qua **ô lịch hệ thống**（`dd/mm/yyyy` + icon 🗓，
     hiển thị theo ngôn ngữ máy）hoặc gõ tay vào ô.
   - **所要日数**: nhập số → tự tính 終了日 = ngày làm việc thứ N tính từ 開始日
     (bỏ CN・lễ・thứ 7 nếu tắt). Ngược lại, điền đủ 開始/終了 thì 所要日数 tự hiện.
   - Dữ liệu tự lưu vào localStorage — đóng/mở lại task pane không mất.
2. Chọn **土曜日も稼働する** (mặc định bật = làm thứ 7).
3. Thêm **特別休業日**: chọn ngày trong ô date rồi bấm **休日追加** (hoặc Enter) — ngày đó được coi
   là ngày nghỉ giống Chủ nhật・lễ (thanh Gantt bỏ qua, không tính vào 所要日数, cột bị
   shade xám, chữ đỏ kể cả khi rơi vào thứ 7). Xóa bằng nút ✕ trên chip ngày.
3. Bấm **🔄 ガントチャートを生成** → sheet 「ガントチャート」được tạo mới hoàn toàn:
   - Tiêu đề 「工　程　表」 merge toàn bộ hàng đầu
   - Header 4 dòng: 年（merge theo năm）/ 月（merge theo tháng）/ 日 / 曜日, ô góc trái
     vẽ **đường chéo + nhãn 名称等**
   - Mỗi công việc 2 dòng (dòng 備考 + dòng 矢印), tên công việc merge 2 dòng ở cột A
   - Thanh Gantt = **hình mũi tên (shape)** đè lên ngày làm việc — Chủ nhật・ngày lễ・
     thứ 7 (khi tắt) bị bỏ qua, shade xám toàn cột; chữ CN đỏ, thứ 7 xanh
   - 備考 hiển thị ở ngày làm việc đầu tiên của công việc
   - Khối **行事・備考等** 3 dòng cuối: ô mẫu (trắng・xám・trắng) + dòng chú thích
   - Ẩn gridline; layout in mặc định **ngang・A3・fit 1 trang**
4. Sinh lại nhiều lần không sao: sheet ガントチャート xóa tạo lại, dữ liệu nhập
   trong task pane giữ nguyên.

## Gỡ add-in

Chạy `node uninstall.cjs` trong thư mục này (xóa đăng ký registry), rồi khởi
động lại Excel. Muốn dùng lại: chạy `install.bat` (hoặc lệnh đăng ký PowerShell
ở phần trên).

⚠ `manifest.xml` phải là **file cục bộ** trên máy đó (registry trỏ tới đường
dẫn file, không phải URL), dù nội dung nó tham chiếu tới GitHub Pages.

## Phát triển / sửa code cục bộ

Khi cần sửa và xem trước TRƯỚC KHI push lên GitHub:

1. Double-click **`start-server.bat`** để chạy server tĩnh cục bộ (port 8080).
2. Tạo 1 bản `manifest.xml` khác (hoặc sửa tạm) trỏ về `http://localhost:8080/...`
   thay vì URL GitHub Pages, đăng ký sideload bản đó riêng để test.
3. Test xong, sửa code thật trong repo, `git add . && git commit && git push` —
   GitHub Pages tự build lại và mọi máy (kể cả máy dùng bản manifest GitHub Pages)
   sẽ thấy bản mới ở lần mở task pane tiếp theo.

## Khắc phục sự cố

| Hiện tượng | Nguyên nhân / cách xử lý |
|---|---|
| `エラー：... Worksheet.delete ... GeneralException (0xA7120001)` khi sinh chart | Sheet cũ không xóa được (sách đang 構成保護 / đang chỉnh sửa ô / sheet duy nhất). Add-in sẽ tự **tái sử dụng sheet cũ** (gỡ merge, xóa shape, xóa dữ liệu) nên thường tự vượt qua; nếu vẫn lỗi: Enter/Esc để thoát chế độ chỉnh sửa ô, hoặc kiểm tra レビュー→シートの保護/ブックの保護 |
| Task pane trắng / "ADD-IN ERROR" | Kiểm tra máy có mạng để tải GitHub Pages không. Excel cache add-in cũ: xóa thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` rồi khởi động lại Excel |
| Không tìm thấy add-in / nút 工程表ツール | Chưa khởi động lại Excel hoàn toàn sau khi sideload; hoặc registry value bị xóa → chạy lại `install.bat` hoặc lệnh đăng ký PowerShell ở trên |
| Sửa code xong nhưng Excel vẫn hiện bản cũ | Chờ GitHub Pages build xong (~1-2 phút, xem tab **Actions** trên GitHub), rồi đóng hẳn task pane và mở lại (không chỉ ẩn/hiện) |
| `reg add` báo Invalid syntax | Đang chạy trong Git Bash (switch `/v` bị hiểu là đường dẫn) → dùng PowerShell như ở trên |

## Mức độ đã kiểm chứng

| Phần | Trạng thái |
|---|---|
| Logic lịch nghỉ Nhật 2026 (18 ngày, 振替/国民の休日) | ✅ test bằng Node, khớp 18/18 |
| Parse ngày, đếm/ cộng ngày làm việc, tự tính 所要日数, 特別休業日 né ngày làm việc | ✅ test bằng Node |
| Model Gantt (bar bỏ CN/lễ・特別休業日, bật-tắt thứ 7, nhiều tháng, ghi chú ngày đầu) | ✅ test bằng Node |
| Taskpane UI hiển thị, không lỗi console | ✅ test bằng Edge headless thật |
| Server tĩnh (taskpane/js/css/icon, chặn path traversal) | ✅ test bằng curl |
| **Chạy trong Excel thật** (sideload, mở taskpane, nhập liệu, sinh sheet) | ✅ **đã kiểm chứng 2026-09-14 trên Excel desktop thật** (64-bit Windows, Office x86 C2R): sinh sheet đúng 21 hàng × 46 cột cho 2 công việc 4/1–5/15, 曜日 đúng (4/1=水), mũi tên shape có run start/width khớp từng pixel với model Node, 備考 nằm đúng ô ngày làm việc đầu, shade cột CN/lễ (4/29 昭和の日) |
