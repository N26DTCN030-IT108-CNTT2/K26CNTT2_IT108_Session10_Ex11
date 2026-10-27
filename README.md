# Bài tập 11: Thiết Kế Bộ Hồ Sơ Kỹ Thuật Hoàn Chỉnh Cho Hội Nghị Khoa Học PTIT 2027

## Phần 1. Thiết Kế Kiến Trúc Hệ Thống & Phân Cấp Heading

### 1. Sơ đồ Luồng Dữ Liệu (Mô hình IPO Hệ thống Hồ sơ)

Luồng dữ liệu chuyển hóa giữa các công cụ trong Bộ hồ sơ kỹ thuật:

+-------------------------------------------------------+
|                 DATA NGUỒN (EXCEL)                    |
| Bảng quản lý Báo cáo viên & Danh mục Phòng/Thù lao    |
+-------------------------------------------------------+
|
v  (Trích xuất dữ liệu tổng hợp & số liệu)
+-------------------------------------------------------+
|              TÀI LIỆU HIỂN THỊ THỊ GIÁC               |
+-------------------------------------------------------+
|  (A) WORD (Tài liệu Kế hoạch)   | (B) POWERPOINT      |
|  - Chèn bảng Lịch trình         | - Vẽ Bar/Donut Chart|
|  - Nhúng bảng Cơ cấu giải thưởng| - Slide Khai mạc    |
+-------------------------------------------------------+


- **Thành phần Nguồn (Data Source):** Bảng tính **Excel** (`Hoi_Nghi_KHKH_PTIT_2027_Quan_Ly.xlsx`) giữ vai trò bộ chứa dữ liệu trung tâm (Data Central). Nơi quản lý danh sách báo cáo viên, số liệu thống kê theo Khoa/Viện, và tính toán thù lao.
- **Thành phần Hiển thị (Presentation Layer):**
  - **Word** (`Hoi_Nghi_KHKH_PTIT_2027_Ke_Hoach.docx`): Tiếp nhận dữ liệu lịch trình, danh sách giải thưởng và quy chế để đóng gói thành văn bản hành chính in ấn phát cho Ban Giám khảo đọc trước.
  - **PowerPoint** (`Hoi_Nghi_KHKH_PTIT_2027_Slide.pptx`): Trích xuất các chỉ số tổng hợp (tỷ lệ báo cáo các khoa, tổng số đề tài) vẽ thành biểu đồ trực quan chiếu trực tiếp trước 500 người tham dự.

---

### 2. Thiết kế Cấu trúc Phân cấp Heading (8 Dòng tiêu đề chuẩn H1/H2/H3)

Dưới đây là cây phân cấp tiêu đề chuẩn được gán nhãn Heading Styles trong file Word kế hoạch:

# 1. GIỚI THIỆU TỔNG QUAN (Heading 1)
## 1.1. Mục đích sự kiện (Heading 2)
## 1.2. Phạm vi và quy mô hội nghị (Heading 2)

# 2. CƠ CẤU BAN TỔ CHỨC (Heading 1)
## 2.1. Ban Nội dung & Hội đồng Thẩm định (Heading 2)
### 2.1.1. Tiểu ban Kỹ thuật và CNTT (Heading 3)
### 2.1.2. Tiểu ban Điện tử viễn thông (Heading 3)
## 2.2. Ban Thư ký & Hậu cần (Heading 2)

# 3. LỊCH TRÌNH SỰ KIỆN (Heading 1)
# 4. CƠ CẤU GIẢI THƯỞNG (Heading 1)
# 5. QUY CHẾ VÀ BẢO MẬT BÀI BÁO CÁO (Heading 1)

---

## Phần 2. Mô Tả Chi Tiết Bộ Sản Phẩm

### 1. File Word Kế hoạch Hội nghị (`Hoi_Nghi_KHKH_PTIT_2027_Ke_Hoach.docx`)
- **Phân cấp Heading chuẩn:** Thiết lập 3 cấp tiêu đề (H1, H2, H3) bằng công cụ Styles của Word.
- **Mục lục tự động (Auto TOC):** Tích hợp trường Auto TOC ngay tại Trang 1, tự động thu thập tiêu đề và cập nhật số trang chính xác.
- **Thành phần bảng & danh sách:**
  - Có 1 Bảng biểu lịch trình chi tiết các phiên làm việc (Thời gian, Nội dung, Địa điểm).
  - Có 1 Danh sách có thứ tự (Numbered List) chi tiết cơ cấu giải thưởng (Giải Nhất, Nhì, Ba).

---

### 2. File Excel Quản lý Báo cáo viên (`Hoi_Nghi_KHKH_PTIT_2027_Quan_Ly.xlsx`)
- **Tuân thủ 5 Quy tắc Vàng Database:** Không merge cell, không có dòng/cột trống nằm giữa, hàng tiêu đề duy nhất ở Dòng 1.
- **Hệ thống 2 Sheet độc lập:**
  - `DanhSachBaoCao`: Quản lý danh sách tác giả, khoa, đề tài, phòng trình bày và thù lao.
  - `DanhMucPhong`: Lưu trữ thông tin định mức phòng và thù lao theo mã tác giả.
- **Công thức tra cứu VLOOKUP:** Tra cứu phòng và thù lao tự động bằng công thức bọc hàm chống lỗi:
  ```excel
  =IFERROR(VLOOKUP(TRIM(A2), DanhMucPhong!A$2:C$81, 2, FALSE), "Chưa xếp")
