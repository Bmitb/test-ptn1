# 🔬 Pressure Calibration Report Parser AI — Batch & Multi-page PDF Edition

Ứng dụng Full-Stack Python xử lý **hàng loạt** biên bản hiệu chuẩn áp suất chữ viết tay bằng **Google Gemini Vision AI** và tự động ghi/append **toàn bộ dữ liệu** vào file **Excel template 2 sheet** (VILAS 415) chỉ với **1 Click**, sử dụng `openpyxl`.

---

## 🌟 Tính năng nổi bật

* **📄 Multi-page PDF Auto Splitting (MỚI v2.1):** Tự động tách file PDF nhiều trang thành từng trang/biên bản độc lập (`split_pdf_pages`), trích xuất thông tin riêng cho từng thiết bị và tổng hợp đầy đủ vào Excel.
* **🚀 Batch Upload & Processing (MỚI v2.0):** Upload **nhiều file PDF/ảnh cùng lúc**, hệ thống tự động xử lý lần lượt từng file/trang với **thanh tiến trình** (`st.progress`) hiển thị trạng thái theo từng file/trang theo thời gian thực.
* **📋 Bảng tổng hợp thống nhất (MỚI v2.0):** Sau khi trích xuất, **tất cả dữ liệu** của các thiết bị được gộp vào **1 bảng Sheet 1** (mỗi hàng = 1 thiết bị) và **1 bảng Sheet 2** (tất cả điểm đo), có thể chỉnh sửa trực tiếp trước khi lưu.
* **💾 Ghi liên tiếp 1-Click (MỚI v2.0):** Nút "💾 Lưu tất cả vào Excel" ghi **toàn bộ N thiết bị** vào Sheet 1 (N hàng) và Sheet 2 (N × điểm đo, cách nhau 1 dòng trống) chỉ với 1 thao tác.
* **🎯 Căn lề & Định dạng chuẩn VILAS 415 (MỚI v2.2):** Khắc phục triệt để lỗi lệch cột và lệch định dạng màu trên Sheet 2 bằng cách bổ sung cột **Phương pháp HC** (Col 3), căn chỉnh đúng 13 cột và kế thừa dòng mẫu định dạng gốc (Rows 7 & 8).
* **🤖 Đọc chữ viết tay tiếng Việt xuất sắc:** Sử dụng **Google Gemini AI** (hỗ trợ `gemini-2.5-flash`, `gemini-2.5-pro`, `gemini-1.5-pro`,...) để nhận diện hình ảnh/PDF phiếu hiệu chuẩn áp suất viết tay hoặc in.
* **🛡️ Cơ chế Kháng lỗi 503 & Rate Limit:** Tự động thử lại (Exponential Backoff) và tự chuyển sang model dự phòng (Model Fallback) khi máy chủ Google quá tải. Các file lỗi được ghi log riêng, không dừng toàn bộ batch.
* **🔍 Tự động quét Model khả dụng:** Tích hợp tính năng quét toàn bộ model active từ Google API Key của người dùng.
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

## 📋 Cấu trúc Mapping Dữ liệu Excel (2 Sheets - Chuẩn VILAS 415)

Ứng dụng trích xuất và ánh xạ tự động vào 2 Sheet của file template Excel:

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

| Col # | Tên Cột Excel | Trường dữ liệu trích xuất | Ví dụ / Mô tả |
| :---: | :--- | :--- | :--- |
| **1** | **Mã Phụ (Tự động)** | `{gcn_so}{point_id}` | `MCLAB26CN-1.0005D1` |
| **2** | **GCN Số** | `gcn_so` | `MCLAB26CN-1.0005` |
| **3** | **Phương pháp HC** | `phuong_phap_hc` | `DLVN76` |
| **4** | **Mã QL/ Mã ID (Tự động)** | `ma_id` | `MC-05-6-646` |
| **5** | **Đ.vị** | `don_vi` | `bar` |
| **6** | **Min** | `range_min` | `0` (chỉ điền ở dòng D1) |
| **7** | **Max** | `range_max` | `700` (chỉ điền ở dòng D1) |
| **8** | **Điểm hiệu chuẩn** | `point_id` | `D1`, `D2`, `D3`,... |
| **9** | **Đơn vị P** | `don_vi` | `bar` |
| **10** | **P** | `p_value` | Giá trị áp suất đặt (vd: `0`, `100`, `200`) |
| **11** | **Đơn vị Chuẩn P** | `don_vi` | `bar` |
| **12** | **P c.tăng** | `p_tang` | Số đọc chiều tăng của chuẩn (vd: `102.3`) |
| **13** | **P c.giảm** | `p_giam` | Số đọc chiều giảm của chuẩn (vd: `99.5`) |

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

## 📖 Quy trình sử dụng Batch — 4 bước trên Web UI

1. **🔑 Cấu hình:** Nhập Google API Key ở Sidebar. Bấm nút **🔍 Lấy danh sách model** để chọn model mạnh nhất (`gemini-2.5-flash` / `gemini-2.5-pro`).
2. **📄 Upload hàng loạt:** Kéo thả hoặc chọn **nhiều file PDF / ảnh** cùng một lúc vào ô upload ở Bước 1. (Nếu có file Excel mẫu 2 sheet, upload ở Sidebar.)
3. **⚡ Trích xuất AI hàng loạt:** Nhấn nút **🚀 Trích xuất tất cả N file**. Hệ thống lần lượt gọi Gemini API cho từng file, hiển thị thanh tiến trình và log theo thời gian thực.
4. **💾 Kiểm tra & Lưu 1-Click:** Xem bảng tổng hợp tất cả thiết bị, chỉnh sửa ô sai nếu cần, sau đó bấm **💾 Lưu tất cả vào Excel** → toàn bộ dữ liệu được ghi liên tiếp vào file để tải về.

---

## 📄 Giấy phép & Tác giả
* **Được phát triển bởi:** Senior Full-Stack Developer
* **Phiên bản:** `2.2.0` — VILAS 415 Excel Format Edition
* **Tháng:** 08/2026
