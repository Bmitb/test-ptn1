# 🔬 Pressure Calibration Report Parser AI

Ứng dụng Full-Stack Python xử lý biên bản hiệu chuẩn áp suất chữ viết tay bằng **Google Gemini Vision AI** và tự động ghi/append dữ liệu có cấu trúc vào file **Excel template 2 sheet** bằng `openpyxl`.

---

## 🌟 Tính năng nổi bật

* **🤖 Đọc chữ viết tay tiếng Việt xuất sắc:** Sử dụng **Google Gemini AI** (hỗ trợ `gemini-2.5-flash`, `gemini-2.5-pro`, `gemini-1.5-pro`,...) để nhận diện hình ảnh/PDF phiếu hiệu chuẩn áp suất viết tay hoặc in.
* **🛡️ Cơ chế Kháng lỗi 503 & Rate Limit:** Tự động thử lại (Exponential Backoff) và tự chuyển sang model dự phòng (Model Fallback) khi máy chủ Google quá tải.
* **🔍 Tự động quét Model khả dụng:** Tích hợp tính năng quét toàn bộ model active từ Google API Key của người dùng.
* **✏️ Kiểm tra & Sửa dữ liệu trực tiếp:** Cho phép chỉnh sửa tay trực tiếp trên bảng dữ liệu (Streamlit Interactive Data Editor) trước khi lưu vào Excel.
* **📊 Bảo tồn định dạng Excel:** Sử dụng `openpyxl` nối tiếp dữ liệu vào cuối 2 sheet mà không làm mất định dạng, công thức, font chữ hay khung viền có sẵn của file template `.xlsx`.
* **🎨 Giao diện Dark-Mode cao cấp:** Thiết kế hiện đại, tương phản cao, tối ưu từng khung nhập liệu và nhãn giúp dễ quan sát.

---

## 🛠️ Công nghệ sử dụng

| Thành phần | Thư viện / Công nghệ | Vai trò |
| :--- | :--- | :--- |
| **Giao diện (UI/UX)** | `Streamlit >= 1.35.0` | Khung ứng dụng web tương tác |
| **AI Vision (OCR)** | `google-genai >= 0.7.0` | SDK mới nhất kết nối Google Gemini API |
| **Xử lý Excel** | `openpyxl >= 3.1.2` | Đọc, chỉnh sửa & append dòng vào file `.xlsx` |
| **Xử lý Bảng biểu** | `pandas >= 2.0.0` | Cấu trúc & hiển thị dữ liệu bảng |
| **Chuyển đổi PDF** | `PyMuPDF (fitz) >= 1.24.0` | Render file PDF biên bản thành hình ảnh chất lượng cao |

---

## 📋 Cấu trúc Mapping Dữ liệu Excel (2 Sheets)

Ứng dụng trích xuất và ánh ánh tự động vào 2 Sheet của file template Excel:

### 1️⃣ Sheet 1: Danh mục & Thông tin chung (1 dòng / phiếu)

| Tên Cột Excel | Trường dữ liệu trích xuất | Ví dụ / Mô tả |
| :--- | :--- | :--- |
| **GCN Số** | `gcn_so` | `240815/TB-12` |
| **Mã ID** | `ma_id` | `PI-204` |
| **Mã số nhận dạng** | `ma_id` | `PI-204` |
| **Tên UUT** | `ten_uut` | `Pressure Gauge` / `Áp suất kế` |
| **Khách hàng** | `khach_hang` | `Công ty ABC` |
| **Phiếu YCCV** | `""` | Để trống |
| **Người thực hiện** | `nguoi_thuc_hien` | Nguyễn Văn A |
| **P.pháp HC** | `"DLVN76"` | Mặc định `DLVN76` |
| **Ngày hiệu chuẩn** | `ngay_hc` | `15/08/2026` |
| **Kết quả HC** | `ket_qua` | `OK` (hoặc `FAIL`) |
| **Tem hiệu chuẩn** | `tem_hc` | `TEM-9982` |
| **Ngày HC kế tiếp** | `ngay_ke_tiep` | `15/08/2027` |
| **TB Chuẩn 1** | `tb_chuan_1` | `STD-01` |

### 2️⃣ Sheet 2: Chi tiết các điểm đo (N dòng / phiếu — D1 đến Dn)

| Tên Cột Excel | Trường dữ liệu trích xuất | Ví dụ / Mô tả |
| :--- | :--- | :--- |
| **Mã Phụ** | `{gcn_so}{point_id}` | `240815/TB-12D1` |
| **GCN Số** | `gcn_so` | `240815/TB-12` |
| **Mã QL / Mã ID** | `ma_id` | `PI-204` |
| **Đ.vị** | `don_vi` | `bar` / `MPa` |
| **Min** | `range_min` | `0` (chỉ điền ở dòng D1) |
| **Max** | `range_max` | `100` (chỉ điền ở dòng D1) |
| **Điểm HC** | `point_id` | `D1`, `D2`, `D3`,... |
| **Đơn vị P** | `don_vi` | `bar` |
| **P** | `p_value` | Giá trị áp suất đặt (vd: `0`, `25`, `50`, `75`, `100`) |
| **Đơn vị Chuẩn P** | `don_vi` | `bar` |
| **P c.tăng** | `p_tang` | Số đọc chiều tăng của chuẩn |
| **P c.giảm** | `p_giam` | Số đọc chiều giảm của chuẩn |

---

## 🚀 Hướng dẫn cài đặt & Chạy ứng dụng

### 1. Requirements & Chuẩn bị môi trường
Yêu cầu **Python 3.10+**.

Clone dự án về máy tính:
```bash
git clone https://github.com/Bmitb/test-ptn1.git
cd test-ptn1
```

Cài đặt các thư viện phụ thuộc:
```bash
pip install -r requirements.txt
```

### 2. Cấu hình Google Gemini API Key
Bạn có thể cấu hình API key theo 2 cách:

* **Cách 1 (Khuyên dùng):** Nhập trực tiếp API Key vào ô **Google API Key** ở thanh Sidebar bên trái của ứng dụng Web.
* **Cách 2:** Tạo file `.env` từ file mẫu `.env.example`:
  ```bash
  cp .env.example .env
  ```
  Sau đó điền API key vào file `.env`:
  ```env
  GOOGLE_API_KEY=AIzaSyYourActualKeyHere...
  ```

### 3. Chạy ứng dụng
Khởi chạy ứng dụng Streamlit:
```bash
streamlit run app.py
```

Truy cập giao diện Web tại địa chỉ: `http://localhost:8501`

---

## 📖 Quy trình sử dụng 4 bước trên Web UI

1. **🔑 Cấu hình:** Nhập Google API Key ở Sidebar. Bấm nút **🔍 Lấy danh sách model** để chọn model mạnh nhất (`gemini-2.5-flash` / `gemini-2.5-pro`).
2. **📄 Upload:** Kéo thả file PDF hoặc ảnh chụp phiếu hiệu chuẩn vào Bước 1. (Nếu có file Excel mẫu 2 sheet, upload ở Sidebar).
3. **⚡ Trích xuất AI:** Nhấn nút **Extract Data with AI**. Hệ thống sẽ đọc chữ viết tay và hiển thị 2 bảng dữ liệu xem trước.
4. **💾 Kiểm tra & Xuất Excel:** Chỉnh sửa trực tiếp trên bảng nếu cần, sau đó bấm **Save & Append to Excel** để tải file Excel kết quả về máy.

---

## 📄 Giấy phép & Tác giả
* **Được phát triển bởi:** Senior Full-Stack Developer
* **Phiên bản:** `1.0.0`
