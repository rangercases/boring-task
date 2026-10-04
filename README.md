# ⚡ BORING TASK - Desktop Automation Suite

Ứng dụng Desktop tự động hóa các tác vụ lặp đi lặp lại hàng ngày, thiết kế theo phong cách Apple tối giản, tinh tế và dễ dùng cho mọi nhân viên văn phòng.

---

## 🧩 Các Phân Hệ (Modules)

### 1. 📊 Purchase Cost Auto-Filled (Claim & Cost Matcher)
Tự động đối chiếu và điền **Purchase Cost** từ file Master Overview vào các báo cáo khiếu nại (Claim Versus) theo từng quý/tháng:
- **Đối chiếu đa tầng thông minh (Multi-Level Matching)**:
  - Khớp chính xác 100% theo mã sản phẩm (`Customer art No`) và nhà cung cấp (`Casa / CS` hoặc `Nhan Hoang / NH`).
  - Tự động lọc các đơn hàng dị biệt hoặc chỉ giao lẻ 1 kiện (ví dụ: *Only box 1 of 2*).
  - Khớp theo ngày phát hành PO gần nhất với ngày đặt hàng ban đầu của sản phẩm bị khiếu nại (`PostingDate - Days Between Order And Complaint`).
  - Xử lý thông minh các mã rút gọn (Made to Measure - MM) dựa trên phân khúc giá và tỷ giá quy đổi tương đương (`Unit Cost DKK`).
- **Minh bạch số liệu (Transparency)**:
  - Tự động điền cột `Purchase Cost (USD)`.
  - Tự động bổ sung cột `Purchase Cost Note` ghi chú cơ sở tra cứu chi tiết cho từng dòng.
  - Tự động chèn công thức tính tổng `=SUM(...)` ở dòng Total cuối cùng.
- **Xử lý ngầm đa luồng**:
  - Tốc độ đọc ghi nhanh, không đơ lag cửa sổ ngay cả với file Excel dữ liệu lớn (~40MB).

### 2. 🖼️ Image Inserter (Place in Cell)
Công cụ tự động hóa chèn ảnh khiếu nại chất lượng sản phẩm trực tiếp vào ô tính Microsoft Excel chuẩn **Place in Cell**:
- **Tự động nhận diện & Quét ảnh**: Quét các cột link ảnh chuẩn (`Comp. Pic. 1`, `Pic 2`... `Pic 5`).
- **Tải & Nén ảnh HD 16 luồng song song**: Tải siêu tốc từ server Châu Âu, nén chuẩn HD lưu vào thư mục Cache cục bộ.
- **Nhúng Excel COM nội tại ô (Place in Cell)**: Tự động điều chỉnh kích thước dòng/cột tương thích, mở file kết quả ngay sau khi hoàn thành.
- **Kéo & Thả (Drag & Drop)**: Hỗ trợ kéo thả trực tiếp file Excel `.xlsx`, `.xlsm`, `.xls`.

---

## ⚙️ Cấu Hình Phân Quyền Hiển Thị (`config.json`)

Toàn bộ ứng dụng dùng **chung 1 mã nguồn duy nhất**, nhưng giao diện Home sẽ tự động thích ứng hiển thị đúng module theo nhu cầu công việc của từng người máy:

Tạo file `config.json` trong thư mục app (file này nằm trong `.gitignore`, không bao giờ bị ghi đè khi cập nhật):
- **Bản của Bạn (Admin - Hiện tất cả):**
  ```json
  { "mode": "all" }
  ```
- **Bản gửi cho Chị A (Chỉ làm Claim Cost):**
  ```json
  { "mode": "cost" }
  ```
- **Bản gửi cho Chị B (Chỉ làm Chèn Ảnh):**
  ```json
  { "mode": "images" }
  ```
*(Nếu không có file `config.json`, ứng dụng mặc định hiển thị đầy đủ cả 2 phân hệ).*

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng (Dành Cho Máy Công Ty)

### Cách 1: Cài đặt 1-Click (Khuyên Dùng - Tránh Antivirus Chặn)
Vì một số phần mềm diệt virus trong môi trường doanh nghiệp thường tự ý xóa/cách ly các file `.exe` lạ, giải pháp chạy trực tiếp qua môi trường Python chính thức là an toàn và ổn định nhất:

1. Copy/gửi thư mục app sang máy đồng nghiệp (có sẵn `config.json` theo đúng quyền nếu muốn).
2. Nhấp đúp chuột vào file:
   ```text
   Cai_Dat_Boring_Task.bat
   ```
3. File cài đặt sẽ:
   - Tự động kiểm tra môi trường Python.
   - Tự động cài đặt thư viện cần thiết (`openpyxl`, `Pillow`, `pywin32`, `tkinterdnd2`...).
   - Tạo biểu tượng lối tắt **Boring Task** ngoài màn hình Desktop với icon ứng dụng chuẩn.
   - Khởi chạy app ngay lập tức (chạy ngầm không hiện cửa sổ đen CMD).

### 🔄 Cơ chế Tự Động Cập Nhật Ngầm (Silent Auto-Update)
Mỗi khi khởi động **Boring Task**, ứng dụng sẽ tự động chạy tiến trình ngầm kiểm tra mã nguồn mới nhất trên GitHub repository trong vài mili-giây. Nếu bạn đẩy bản cập nhật mới lên GitHub, app trên các máy người dùng sẽ tự động đồng bộ mà người dùng không cần thao tác gì thêm!

---

### Cách 2: Đóng gói thành file EXE độc lập (Tùy chọn)
Nếu máy người dùng được phép chạy file `.exe`:
- Nhấp đúp vào **`build_exe.bat`** để đóng gói toàn bộ ứng dụng thành **`dist/BoringTask.exe`**.

---

## 📁 Cấu Trúc Thư Mục

```text
├── boring_task.pyw           # Ứng dụng chính (Giao diện Apple Home & 2 Phân hệ tích hợp)
├── Cai_Dat_Boring_Task.bat   # Bộ cài đặt 1-click tạo Shortcut Desktop cho đồng nghiệp
├── build_exe.bat             # Script đóng gói thành BoringTask.exe
├── requirements.txt          # Danh sách thư viện phụ thuộc
├── config.json               # Cấu hình phân quyền hiển thị (local, git-ignored)
├── assets/                   # Icon và tài nguyên đồ họa của ứng dụng
├── .gitignore                # Bảo mật dữ liệu nội bộ & cấu hình cá nhân
└── README.md                 # Tài liệu hướng dẫn sử dụng
```
