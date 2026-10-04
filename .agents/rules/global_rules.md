# ⚡ BORING TASK — GLOBAL AGENT RULES & PROJECT STANDARDS

Tài liệu quy tắc và tiêu chuẩn kỹ thuật bắt buộc dành cho AI Agent và nhà phát triển khi làm việc trên dự án **Boring Task** (`rangercases/claim-helper`).

---

## 🎯 1. ĐỐI TƯỢNG SỬ DỤNG & TRIẾT LÝ SẢN PHẨM

- **Đối tượng người dùng cuối:** Các chị đồng nghiệp văn phòng không rành công nghệ.
- **Tiêu chuẩn trải nghiệm (UX):**
  - Mọi thao tác cài đặt phải là **1-Click** (chỉ cần chạy duy nhất 1 file `Cai_Dat_Boring_Task.bat`).
  - Tuyệt đối không yêu cầu người dùng phải gõ lệnh terminal hay cấu hình đường dẫn phức tạp.
- **Chiến lược bảo mật Antivirus:**
  - Không phân phối file `.exe` đơn thuần vì phần mềm diệt virus công ty sẽ cách ly hoặc xóa bỏ.
  - Luôn khởi chạy ứng dụng trực tiếp bằng trình thông dịch Python chính thống (`pythonw.exe`) trong môi trường máy.

---

## 🏗️ 2. KIẾN TRÚC MÃ NGUỒN & PHÂN QUYỀN GIAO DIỆN (MONOREPO)

- **Một mã nguồn duy nhất (`boring_task.pyw`):**
  - Mọi phân hệ nghiệp vụ (`Purchase Cost Auto-Filled`, `Image Inserter`...) đều nằm chung trong một codebase và một GitHub repository duy nhất.
  - Giúp việc bảo trì, tối ưu giao diện và cập nhật ngầm tự động diễn ra tập trung tại một nơi.
- **Phân quyền giao diện Home theo Role (`config.json`):**
  - Ứng dụng đọc file cấu hình máy `config.json` để ẩn/hiện thẻ module tương ứng:
    - `{"mode": "cost"}`: Chị A (Chỉ hiện thẻ *Purchase Cost Auto-Filled*).
    - `{"mode": "images"}`: Chị B (Chỉ hiện thẻ *Image Inserter / Place in Cell*).
    - `{"mode": "all"}`: Admin / Bạn (Hiện đầy đủ tất cả phân hệ).
  - *(Nếu không có `config.json`, mặc định hiển thị đầy đủ).*
- **Bảo vệ file cấu hình (`.gitignore`):**
  - File `config.json` **BẮT BUỘC PHẢI NẰM TRONG `.gitignore`**, tuyệt đối không đưa lên GitHub để tránh việc cập nhật tự động ghi đè lên cấu hình riêng của từng máy.

---

## ⚠️ 3. BỘ QUY TẮC VÀNG VỀ WINDOWS BATCH (.BAT) & INSTALLER

Đúc kết từ các sự cố thực tế với trình thông dịch `cmd.exe` trên Windows:

1. **BẮT BUỘC ĐỊNH DẠNG XUỐNG DÒNG WINDOWS CRLF (`\r\n`):**
   - File `.bat` nếu vô tình lưu bằng Unix LF (`\n`) sẽ khiến `cmd.exe` đọc ngắt quãng lệnh, nuốt ký tự và **lập tức đóng cửa sổ trong 0.1 giây**.
   - Mọi thao tác sửa hoặc tạo file `.bat` phải bảo đảm định dạng `\r\n`.
2. **TUYỆT ĐỐI KHÔNG DÙNG KHỐI LỆNH NGOẶC LỒNG NHAU `if (...) else (...)` PHỨC TẠP:**
   - Trong `cmd.exe`, ký tự đóng ngoặc `)` xuất hiện trong văn bản tiếng Việt (ví dụ `(Chị A)`, `(Admin)`), URL hoặc mảng PowerShell `@('file1', ...)` sẽ khiến CMD hiểu nhầm là đóng khối lệnh và văng lỗi cú pháp hàng loạt.
   - **Quy tắc:** Luôn dùng kiến trúc **Flat Navigation (`goto :label`)**.
3. **KHÔNG DÙNG LỆNH `timeout` KHI CÓ CHUYỂN HƯỚNG DÒNG LỆNH:**
   - Lệnh `timeout /t X > nul` sẽ gây lỗi fatal `ERROR: Input redirection is not supported` nếu tiến trình chạy không có console tương tác.
   - **Quy tắc:** Luôn dùng `ping 127.0.0.1 -n X >nul` để tạo độ trễ an toàn.
4. **CÀI ĐẶT THƯ VIỆN LUÔN DÙNG CỜ `--user`:**
   - Chạy `pip install --user -r requirements.txt` để cài vào `%APPDATA%`, chống 100% lỗi `PermissionError: [WinError 5] Access is denied` khi máy công ty cài Python trong `C:\Program Files`.
5. **KIỂM TRA PYTHON THẬT SỰ:**
   - Không dựa vào `where python` vì Windows 10/11 có shortcut ảo của Microsoft Store.
   - Luôn kiểm tra bằng `python -c "import sys"` và kiểm tra dự phòng `py -3`.
6. **LƯU TRỮ LOGO VĨNH VIỄN CHO SHORTCUT:**
   - Khi cài đặt, bộ cài tự sao lưu `app_icon.ico` vào `%LOCALAPPDATA%\BoringTask\assets\app_icon.ico`.
   - Shortcut Desktop trỏ icon cố định vào đây, bảo đảm người dùng có di chuyển/xóa thư mục tải về thì **Logo ngoài Desktop không bao giờ bị biến thành hình tờ giấy trắng**.

---

## 🎨 4. TIÊU CHUẨN THIẾT KẾ APPLE UI/UX

- **Bảng màu ấm áp "Soft Ivory Bakery":**
  - Nền cửa sổ: `IVORY = "#f7f2e8"`
  - Thẻ card: `CARD = "#fffdf8"`, viền mảnh `LINE = "#e8e0d0"`, góc bo 16–18px
  - Vùng kéo thả: `SAND = "#f1eadb"`, viền cát `FAINT = "#b3a898"`
  - Chữ chính: `ROAST = "#2b211c"`, chữ phụ: `BROWN = "#6f5d50"`
  - Điểm nhấn chính: `MOSS = "#6b7a5a"`, nền nhấn: `MOSS_SOFT = "#e4e8d8"`
- **Typography:**
  - Tiêu đề & Logo: `Cambria` (Serif cổ điển sang trọng).
  - Nội dung & Nút bấm: `Segoe UI` (Sans-serif sắc nét, hỗ trợ 100% tiếng Việt).
- **Thành phần tương tác:**
  - Nút bấm dạng viên thuốc (Pill button): Bo góc 14px hoặc 24px, hover chuyển màu mượt mà.
  - Thanh tiến trình: Siêu mảnh 3px thanh thoát (`height=3`, track `LINE`, fill `MOSS`).
  - Thẻ Home: Khung 200x168 bo cong 18px, hiệu ứng hover viền xanh rêu nhạy bén.
- **Phòng vệ TkinterDnD:**
  - Luôn bọc `tkdnd._require()` trong khối `try...except`. Nếu máy nào lỗi DLL kéo thả, app vẫn phải mở được và hoạt động 100% thông qua nút chọn file.

---

## 🔒 5. BẢO MẬT & BẢO VỆ DỮ LIỆU CÔNG TY

- **Bảo mật tên doanh nghiệp:** Tuyệt đối không để lộ tên các công ty nội bộ trong bất kỳ tài liệu nào đăng tải lên GitHub (`README.md`, commit message, code comments).
- **Loại trừ dữ liệu Excel:** Toàn bộ file dữ liệu (`.xlsx`, `.xls`, `.csv`, `.pkl`) phải nằm trong `.gitignore` để bảo vệ dữ liệu nghiệp vụ không bị đẩy lên mạng.
- **Quy tắc an toàn thao tác file:** Mọi thao tác sửa hoặc ghi đè file quan trọng phải lập kế hoạch rõ ràng và nhận xác nhận từ người dùng.

---

## 📋 6. QUY TRÌNH BẮT BUỘC: IMPLEMENTATION PLAN & CONFIRMATION

Mục đích: Đảm bảo mọi hành động được lập kế hoạch rõ ràng và có sự xác nhận của người dùng trước khi thực thi.

### Phạm Vi Áp Dụng:
- Tạo/chỉnh sửa file.
- Chạy bash/powershell commands hoặc code.
- Gọi APIs, external tools, hoặc MCP servers.
- Bất kỳ hành động có tác động đến hệ thống/dữ liệu.

### Quy Trình Bắt Buộc:
1. **DESCRIBE THE PLAN:** Ghi rõ chi tiết các bước sẽ thực hiện.
2. **ASK FOR CONFIRMATION:** Yêu cầu người dùng xác nhận bằng câu hỏi:
   `Xác nhận trước khi tôi tiến hành? (Y/N)`
3. **WAIT FOR APPROVAL:** Chỉ thực hiện khi nhận được đồng ý rõ ràng.
4. **EXECUTE:** Thực thi đúng theo plan đã được phê duyệt.

### Ngoại Lệ:
- Câu hỏi thông tin, giải thích, hoặc tư vấn (không cần xác nhận).
- Hành động đã được người dùng xác nhận trước đó trong cùng conversation.

### Nguyên Tắc Cốt Lõi Về Thao Tác File & GitHub:
- ❌ **Không bao giờ tự ý hành động:** MỌI hành động Tạo, Sửa, Ghi đè, Xóa, Gộp file (nhất là trong `D:\OneDrive`) BẮT BUỘC phải lập Plan và hỏi xác nhận.
- ❌ **Dù người dùng có ghi kèm "và đẩy lên github" trong cùng một câu lệnh**, Agent VẪN PHẢI hỏi xác nhận cho phần thao tác file trước, TUYỆT ĐỐI KHÔNG ĐƯỢC tự ý gộp.
- ✅ **Ngoại lệ GitHub chỉ áp dụng khi:**
  - File đã được tạo/sửa xong từ trước và đã có sự đồng ý của người dùng.
  - Người dùng ra lệnh RIÊNG BIỆT chỉ để đẩy git (ví dụ: "đẩy lên git", "push git"). Lúc đó mới không cần trình bày plan cho các lệnh git add/commit/push thuần túy.
- ❌ Khi nghi ngờ hoặc câu lệnh có nhiều hành động -> **LUÔN CHỌN HỎI XÁC NHẬN**.

---

## 🛡️ 7. ANTIGRAVITY IDE — AI AGENT SAFETY RULES (WINDOWS)

### 🚫 Vùng Cấm Tuyệt Đối (Critical Protection Zones - Hard Block):
```text
C:\Windows\*
C:\Program Files\*
C:\Program Files (x86)\*
C:\ProgramData\*
C:\System Volume Information\*
C:\$Recycle.Bin\*
C:\bootmgr
C:\pagefile.sys
C:\hiberfil.sys
C:\Users\[Username]\AppData\*  (Trừ AppData\Local\BoringTask của ứng dụng)
C:\Users\[Username]\Saved Games\*
```
- **Hành động:** Từ chối ngay lập tức nếu có lệnh xóa, di chuyển hoặc can thiệp registry vào các phân vùng trên.

### 🔐 Bảo Vệ Thư Mục OneDrive (`D:\OneDrive`):
- **Yêu cầu xác nhận nghiêm ngặt:**
  - ❌ Xóa bất kỳ file/folder nào trong `D:\OneDrive`.
  - ❌ Di chuyển file ra khỏi `D:\OneDrive`.
  - ❌ Đổi tên file/folder trong `D:\OneDrive`.
  - ❌ Thao tác hàng loạt (Bulk operations).
- **Thao tác an toàn (Safe operations):**
  - ✅ Đọc / Xem file (`view_file`, `read_url_content`).
  - ✅ Copy file từ `D:\OneDrive` sang nơi khác.
  - ✅ Tạo file mới / sửa file theo plan đã được xác nhận.

### 🚫 Các Mẫu Lệnh Nguy Hiểm Bị Chặn Tức Thì (Instant Refuse):
```text
del /s /q ...
rmdir /s /q ...
rd /s ...
move ...\*.* ...
robocopy ... /MIR
powershell Remove-Item ... -Recurse -Force
taskkill /F /IM *
takeown /F ...
icacls ... /grant ...
```
- Tuyệt đối không dùng ký tự đại diện wildcard (`*`) trong các lệnh xóa hoặc di chuyển file.
