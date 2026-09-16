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
├── install.bat         ← cài đặt 1-click (không cần Node.js) + bật tự động cập nhật
├── uninstall.bat        ← gỡ add-in + xoá tác vụ tự động cập nhật
├── auto-update.ps1      ← tải manifest.xml mới nhất từ GitHub (install.bat tự gọi/tự tải)
├── server.js, start-server.bat ← chỉ dùng khi PHÁT TRIỂN cục bộ (xem mục cuối)
└── assets/            ← icon
```

**Không cần build, không cần webpack, không cần node_modules.**

## Nội dung được host ở đâu

Toàn bộ taskpane (html/js/css/icon) được host tĩnh trên **GitHub Pages**, tại
`https://manh25xx08xx99-sys.github.io/Gantt-chart/` — build từ repo
[manh25xx08xx99-sys/Gantt-chart](https://github.com/manh25xx08xx99-sys/Gantt-chart),
nhánh `main`, thư mục gốc. `manifest.xml` đã trỏ sẵn tới URL này.

**Cập nhật tự động — 2 tầng:**

| Phần | Cách cập nhật |
|---|---|
| `taskpane.html/.js/.css`, icon | Excel tải qua mạng mỗi lần mở task pane → push lên `main`, đợi GitHub Pages build (~1 phút), đóng/mở lại task pane là có bản mới |
| `manifest.xml` (file **cục bộ** trên từng máy) | Tác vụ Windows `GanttAddin_AutoUpdateManifest` do `install.bat` tạo: chạy **mỗi ngày 9:00 + mỗi lần đăng nhập Windows**, tải `manifest.xml` mới nhất từ GitHub và đăng ký lại registry |

Nhờ đó máy đồng nghiệp **không cần làm gì thêm** khi có bản mới — kể cả khi
`manifest.xml` thay đổi (đổi icon, đổi tên add-in, thêm nút ribbon...).
Tác vụ này chỉ ghi đè `manifest.xml` khi nội dung thật sự khác, và chỉ nhận file
có đúng `<Id>` của add-in (tránh ghi đè bằng file lỗi/file lạ).

## Cài đặt (máy mình hoặc máy đồng nghiệp)

Không cần Node.js, không cần chạy server — taskpane đã host sẵn trên GitHub
Pages, máy chỉ cần đăng ký 1 dòng registry trỏ vào `manifest.xml`.

**Gửi cho đồng nghiệp: chỉ cần `install.bat`** (để ở thư mục nào cũng được, miễn
là thư mục cố định — không xoá/di chuyển sau khi cài).

Cách cài:

1. Copy `install.bat` vào 1 thư mục cố định trên máy cần cài.
2. Double-click **`install.bat`**. Nó sẽ tự:
   - tải `auto-update.ps1` + `manifest.xml` mới nhất từ GitHub vào cùng thư mục,
   - đăng ký add-in vào Excel (chỉ user hiện tại, không cần quyền admin),
   - tạo tác vụ tự động cập nhật `manifest.xml` mỗi ngày.
3. Mở Excel → tab **Home** → nút **工程表ツール** (nhóm 施工計画書) → task pane mở ra.

> Nếu máy đó **không có mạng** lúc cài, hãy gửi kèm cả `manifest.xml` để cùng thư
> mục — install.bat sẽ đăng ký bằng file có sẵn (nhưng khi đó không bật được tự
> động cập nhật).

Muốn gỡ: double-click **`uninstall.bat`** — xoá cả đăng ký add-in và tác vụ tự
động cập nhật.

> Đăng ký thủ công qua PowerShell (nếu cần):
> ```powershell
> New-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer' `
>   -Name 'dfc3fd23-ff19-4d33-b312-a15d117dd27d' `
>   -Value '<đường dẫn đầy đủ tới manifest.xml>' -PropertyType String -Force
> ```
> (Đừng dùng `reg add` trong Git Bash — nó bị biến các switch `/v /t /d /f`
> thành đường dẫn.)
>
> Kiểm tra / chạy tay tác vụ tự động cập nhật:
> ```powershell
> Get-ScheduledTask -TaskName 'GanttAddin_AutoUpdateManifest'
> powershell -ExecutionPolicy Bypass -File '<thư mục cài>\auto-update.ps1'
> ```

⚠ Đây vẫn là add-in **sideload kiểu dev** (không qua Store/M365 admin), nên mỗi
máy cần chạy `install.bat` 1 lần. Sau đó thì không cần cài lại nữa: code chạy
từ GitHub Pages, còn `manifest.xml` được tác vụ hằng ngày tự cập nhật.

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

Chạy `uninstall.bat` (xoá đăng ký registry + tác vụ tự động cập nhật), rồi khởi
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
| `manifest.xml` không tự cập nhật | Kiểm tra tác vụ còn không: `Get-ScheduledTask -TaskName 'GanttAddin_AutoUpdateManifest'`. Nếu đã xoá/di chuyển thư mục cài thì chạy lại `install.bat`. Chạy tay để xem lỗi: `powershell -ExecutionPolicy Bypass -File '<thư mục cài>\auto-update.ps1'` |
| Dữ liệu đã điền trong task pane bị mất | Dữ liệu được lưu **trong chính file Excel** (Office document settings) → phải **lưu file Excel (Ctrl+S)** trước khi đóng. Mở file Excel khác thì task pane trắng là đúng thiết kế |

## Mức độ đã kiểm chứng

| Phần | Trạng thái |
|---|---|
| Logic lịch nghỉ Nhật 2026 (18 ngày, 振替/国民の休日) | ✅ test bằng Node, khớp 18/18 |
| Parse ngày, đếm/ cộng ngày làm việc, tự tính 所要日数, 特別休業日 né ngày làm việc | ✅ test bằng Node |
| Model Gantt (bar bỏ CN/lễ・特別休業日, bật-tắt thứ 7, nhiều tháng, ghi chú ngày đầu) | ✅ test bằng Node |
| Taskpane UI hiển thị, không lỗi console | ✅ test bằng Edge headless thật |
| Server tĩnh (taskpane/js/css/icon, chặn path traversal) | ✅ test bằng curl |
| **Chạy trong Excel thật** (sideload, mở taskpane, nhập liệu, sinh sheet) | ✅ **đã kiểm chứng 2026-09-14 trên Excel desktop thật** (64-bit Windows, Office x86 C2R): sinh sheet đúng 21 hàng × 46 cột cho 2 công việc 4/1–5/15, 曜日 đúng (4/1=水), mũi tên shape có run start/width khớp từng pixel với model Node, 備考 nằm đúng ô ngày làm việc đầu, shade cột CN/lễ (4/29 昭和の日) |
