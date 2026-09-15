# 工程表ツール（m5-gantt-addin）— Add-in Excel cho mục 5 (工程表/Gantt)

Add-in Excel tách riêng từ 施工計画書 Z1.html — **mục 5 (工程表)**. Nhập công việc ngay
trong task pane, sinh Gantt chart sang sheet **ガントチャート** dạng biểu đồ mũi tên
（矢印）vẽ bằng shape, tự tính **lịch nghỉ Nhật** (đúng thuật toán đã kiểm chứng trong
Z1.html: 18/18 ngày lễ 2026, gồm 振替休日・国民の休日).

## Cấu trúc

```
m5-gantt-addin/
├── manifest.xml       ← manifest của add-in (sideload vào Excel qua registry)
├── taskpane.html/.js/.css ← panel điều khiển + logic
├── server.js          ← server tĩnh phục vụ taskpane (port 8080, không cần Python)
├── install.bat         ← cài đặt 1-click (đăng ký + tự bật server khi mở máy)
├── install.cjs         ← phần lõi của install.bat (đăng ký registry + tạo autostart)
├── start-server.bat   ← khởi động server thủ công (double-click)
├── uninstall.cjs      ← gỡ add-in khỏi Excel (và bỏ autostart)
└── assets/            ← icon
```

**Không cần build, không cần webpack, không cần node_modules.**

## Gửi cho đồng nghiệp dùng trên máy khác

Copy nguyên thư mục `m5-gantt-addin/` (USB, chia sẻ file, v.v.) sang máy của đồng
nghiệp — **không cần server chung, mỗi người tự chạy 1 bản trên máy mình**:

1. Cài **Node.js** (bản LTS) từ https://nodejs.org/ nếu máy chưa có.
2. Double-click **`install.bat`** trong thư mục — script sẽ tự:
   - Đăng ký add-in vào Excel (registry, chỉ ảnh hưởng user hiện tại).
   - Tạo 1 file trong thư mục **Startup** của Windows để server tự chạy ngầm
     mỗi khi đăng nhập máy (không cần double-click `start-server.bat` nữa).
   - Khởi động server ngay lập tức.
3. Mở Excel → tab **Home** → nút **工程表ツール** (nhóm 施工計画書) → task pane mở ra.

Muốn gỡ: chạy `node uninstall.cjs` (xóa cả đăng ký registry lẫn autostart).

⚠ Đây là add-in **sideload kiểu dev** (không qua Store/M365 admin), nên mỗi máy
cần tự cài 1 lần như trên; không có bước "cài tập trung" nào thay được việc này
nếu công ty chưa có SharePoint/M365 admin hoặc server nội bộ HTTPS.

## Cách chạy (2 bước, dùng thủ công / máy đang phát triển)

**Bước 1 — khởi động server:** double-click **`start-server.bat`**
(một cửa sổ đen hiện "工程表ツールのサーバを起動しました" — **giữ nguyên cửa sổ này mở**
khi dùng Excel; đóng cửa sổ = task pane không tải được).

**Bước 2 — mở Excel:** add-in đã được sideload vào registry (ID
`dfc3fd23-ff19-4d33-b312-a15d117dd27d` → `manifest.xml` trong thư mục này).
Khởi động Excel, mở file bất kỳ → **Insert → My Add-ins** (hoặc gõ "my add-ins"
vào ô **Tell Me** cạnh tab) → tab **SHARED FOLDER** → chọn
**工程表ツール（5.工程表）** → **Add**. Task pane mở ra bên phải.

> Đăng ký lại registry nếu bị mất (chạy trong PowerShell):
> ```powershell
> New-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer' `
>   -Name 'dfc3fd23-ff19-4d33-b312-a15d117dd27d' `
>   -Value '<đường dẫn đầy đủ tới manifest.xml>' -PropertyType String -Force
> ```
> (Đừng dùng `reg add` trong Git Bash — nó bị biến các switch `/v /t /d /f`
> thành đường dẫn.)

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

Chạy `node uninstall.cjs` trong thư mục này (xóa đăng ký registry **và** file
autostart nếu có), rồi khởi động lại Excel. Muốn dùng lại: chạy `install.bat`
(hoặc lệnh đăng ký PowerShell ở phần trên).

⚠ Nếu di chuyển thư mục này sang chỗ khác, phải đăng ký lại với đường dẫn mới.

## Khắc phục sự cố

| Hiện tượng | Nguyên nhân / cách xử lý |
|---|---|
| `エラー：... Worksheet.delete ... GeneralException (0xA7120001)` khi sinh chart | Sheet cũ không xóa được (sách đang 構成保護 / đang chỉnh sửa ô / sheet duy nhất). Add-in sẽ tự **tái sử dụng sheet cũ** (gỡ merge, xóa shape, xóa dữ liệu) nên thường tự vượt qua; nếu vẫn lỗi: Enter/Esc để thoát chế độ chỉnh sửa ô, hoặc kiểm tra レビュー→シートの保護/ブックの保護 |
| Task pane trắng / "ADD-IN ERROR" | Server chưa chạy → double-click `start-server.bat`. Excel cache add-in cũ: xóa thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` rồi khởi động lại Excel |
| Không tìm thấy add-in trong My Add-ins | Chưa khởi động lại Excel hoàn toàn sau khi sideload; hoặc registry value bị xóa → chạy lại lệnh đăng ký ở trên. Tab đúng là **SHARED FOLDER** |
| Task pane không tải được localhost | Một số máy chặn loopback trong WebView2: chạy `CheckNetIsolation LoopbackExempt -a -n="Microsoft.Win32WebViewHost_cw5n1h2txyewy"` rồi khởi động lại Excel |
| `reg add` báo Invalid syntax | Đang chạy trong Git Bash (switch `/v` bị hiểu là đường dẫn) → dùng PowerShell như ở trên |

## Chuyển sang dùng thật lâu dài (không cần server)

Hiện manifest trỏ tới `http://localhost:8080` (chạy thử). Khi muốn dùng ổn định trên
máy công ty mà không cần chạy server, có 2 lựa chọn:

- **GitHub Pages** (khuyên dùng — công ty đã dùng GitHub): push thư mục này lên repo,
  bật GitHub Pages, rồi sửa trong `manifest.xml` các URL
  `http://localhost:8080/...` thành `https://<user>.github.io/<repo>/...`
  (các chỗ: IconUrl, HighResolutionIconUrl, SupportUrl, AppDomains,
  DefaultSettings/SourceLocation, và 3 bt:Url/bt:Image trong Resources), sideload lại.
- **IIS / server nội bộ công ty**: đặt thư mục vào web server HTTPS của công ty và sửa
  URL tương tự.

## Mức độ đã kiểm chứng

| Phần | Trạng thái |
|---|---|
| Logic lịch nghỉ Nhật 2026 (18 ngày, 振替/国民の休日) | ✅ test bằng Node, khớp 18/18 |
| Parse ngày, đếm/ cộng ngày làm việc, tự tính 所要日数, 特別休業日 né ngày làm việc | ✅ test bằng Node |
| Model Gantt (bar bỏ CN/lễ・特別休業日, bật-tắt thứ 7, nhiều tháng, ghi chú ngày đầu) | ✅ test bằng Node |
| Taskpane UI hiển thị, không lỗi console | ✅ test bằng Edge headless thật |
| Server tĩnh (taskpane/js/css/icon, chặn path traversal) | ✅ test bằng curl |
| **Chạy trong Excel thật** (sideload, mở taskpane, nhập liệu, sinh sheet) | ✅ **đã kiểm chứng 2026-09-14 trên Excel desktop thật** (64-bit Windows, Office x86 C2R): sinh sheet đúng 21 hàng × 46 cột cho 2 công việc 4/1–5/15, 曜日 đúng (4/1=水), mũi tên shape có run start/width khớp từng pixel với model Node, 備考 nằm đúng ô ngày làm việc đầu, shade cột CN/lễ (4/29 昭和の日) |
