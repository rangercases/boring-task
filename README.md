# ⚡ BORING TASK - Desktop Automation Suite

Ứng dụng Desktop tự động hóa các tác vụ lặp đi lặp lại hàng ngày, thiết kế theo phong cách Apple tối giản, tinh tế và dễ dùng cho mọi nhân viên văn phòng.

---

## 🧩 Các Phân Hệ (Modules)

### 1. 📊 Purchase Cost Auto-Filled (SOFACOMPANY Claim & Cost Matcher)
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

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng (Dành Cho Máy Công Ty)

### Cách 1: Cài đặt 1-Click (Khuyên Dùng - Tránh Antivirus Chặn)
Vì một số phần mềm diệt virus trong môi trường doanh nghiệp thường tự ý xóa/cách ly các file `.exe` lạ, giải pháp chạy trực tiếp qua môi trường Python chính thức là an toàn và ổn định nhất:

1. Tải thư mục này về máy đồng nghiệp.
2. Nhấp đúp chuột vào file:
   ```text
   Cai_Dat_Boring_Task.bat
   ```
3. File cài đặt sẽ:
   - Tự động kiểm tra môi trường Python (hướng dẫn tải nếu chưa có).
   - Tự động cài đặt các thư viện cần thiết (`openpyxl`, `Pillow`...).
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
├── boring_task.pyw           # Ứng dụng chính (Giao diện Apple Home & Purchase Cost Auto-Filled)
├── Cai_Dat_Boring_Task.bat   # Bộ cài đặt 1-click tạo Shortcut Desktop cho đồng nghiệp
├── build_exe.bat             # Script đóng gói thành BoringTask.exe
├── requirements.txt          # Danh sách thư viện phụ thuộc (openpyxl, Pillow, pyinstaller)
├── assets/                   # Icon và tài nguyên đồ họa của ứng dụng
├── .gitignore                # Bảo mật dữ liệu nội bộ, loại trừ các file Excel (.xlsx)
└── README.md                 # Tài liệu hướng dẫn sử dụng
```
