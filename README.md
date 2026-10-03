# 🛋️ CLAIM HELPER - SOFACOMPANY Claim & Cost Matcher

Ứng dụng Desktop tự động đối chiếu và điền **Purchase Cost** từ file Master Overview vào các báo cáo khiếu nại (Claim Versus) theo từng quý/tháng dành cho SOFACOMPANY.

---

## 🌟 Tính Năng Nổi Bật

1. **Đối chiếu đa tầng thông minh (Multi-Level Matching)**:
   - Khớp chính xác 100% theo mã sản phẩm (`Customer art No`) và nhà cung cấp (`Casa / CS` hoặc `Nhan Hoang / NH`).
   - Tự động lọc các đơn hàng dị biệt hoặc chỉ giao lẻ 1 kiện (ví dụ: *Only box 1 of 2*).
   - Khớp theo ngày phát hành PO gần nhất với ngày đặt hàng ban đầu của sản phẩm bị khiếu nại (`PostingDate - Days Between Order And Complaint`).
   - Xử lý thông minh các mã rút gọn (Made to Measure - MM) dựa trên phân khúc giá và tỷ giá quy đổi tương đương (`Unit Cost DKK`).
2. **Minh bạch số liệu (Transparency)**:
   - Tự động tạo cột `Purchase Cost (USD)`.
   - Tự động tạo cột `Purchase Cost Note` giải thích chi tiết cơ sở chọn giá cho từng dòng.
   - Thêm công thức tính tổng tự động `=SUM(...)` ở dòng cuối cùng (Total row).
3. **Giao diện Windows trực quan**:
   - Dễ sử dụng, hỗ trợ chọn file/thư mục qua hộp thoại và thanh tiến trình thời gian thực.
   - Tích hợp đa luồng (Multi-threading) không làm treo giao diện khi đọc file Excel lớn (~40MB).

---

## 🚀 Hướng Dẫn Sử Dụng

### Cách 1: Khởi chạy trực tiếp (Không có cửa sổ đen CMD)
- **Cách đơn giản nhất:** Click đúp trực tiếp vào file **`claim_helper.pyw`** hoặc file **`Mo_Claim_Helper.vbs`**.
- Giao diện Apple Design sẽ mở lên lập tức mà **hoàn toàn không bật bất kỳ cửa sổ CMD đen nào**.

### Cách 2: Khởi chạy qua dòng lệnh Python
```bash
# Cài đặt thư viện phụ thuộc
pip install -r requirements.txt

# Khởi chạy ứng dụng (chế độ không console)
pythonw claim_helper.pyw
```

### Cách 3: Đóng gói thành file `.exe` cho máy khác dùng độc lập (Không cần cài Python)
1. Click đúp vào file **`build_exe.bat`**.
2. Sau khi biên dịch xong, ứng dụng thực thi sẽ nằm tại:
   ```
   dist\ClaimHelper\ClaimHelper.exe
   ```
3. Bạn có thể nén cả thư mục `dist\ClaimHelper` thành file `.zip` và gửi cho bất kỳ đồng nghiệp nào mở lên dùng ngay (không cần cài Python).

---

## 🔄 Hướng Dẫn Quản Lý & Cập Nhật Từ Xa Qua GitHub

Dự án đã được cấu hình `.gitignore` chuẩn an toàn (ngăn không cho các file dữ liệu nội bộ `.xlsx` bị đẩy lên mạng).

### 1. Đưa mã nguồn lên GitHub lần đầu:
1. Tạo một repository mới trên tài khoản GitHub của bạn (ví dụ đặt tên là `claim-helper`).
2. Mở terminal tại thư mục này và chạy các lệnh:
```bash
git init
git add .
git commit -m "feat: Initial release of Claim Helper v1.0"
git branch -M main
git remote add origin https://github.com/<tai-khoan-cua-ban>/claim-helper.git
git push -u origin main
```

### 2. Cập nhật mã nguồn khi có tính năng mới (Từ xa):
* **Trên máy chỉnh sửa (Push code mới)**:
  ```bash
  git add .
  git commit -m "update: cai tien thuat toan hoac giao dien"
  git push
  ```
* **Trên các máy khác (Tải bản cập nhật về)**:
  ```bash
  git pull
  ```

---

## 📁 Cấu Trúc Thư Mục

```
├── claim_helper.py       # Mã nguồn chính của Desktop App (GUI Tkinter)
├── build_exe.bat         # Script đóng gói 1-click thành ClaimHelper.exe
├── requirements.txt      # Thư viện phụ thuộc (openpyxl, pyinstaller)
├── .gitignore            # Bảo mật dữ liệu công ty, loại trừ file .xlsx
└── README.md             # Hướng dẫn chi tiết
```
