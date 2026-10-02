# slide_quiz - Hệ Thống Ôn Tập & Trình Chiếu GDQP-AN | Đại Học Đại Nam

> **Nền tảng slide bài giảng điện tử tương tác 2K kết hợp công cụ trắc nghiệm và bộ chuyển đổi xuất file PowerPoint (.pptx) có thể chỉnh sửa 100% (Native Editable PPTX).**

🌐 **Trang Web Trực Tuyến (Live Website):**  
👉 **[https://thehilichurl.github.io/slide_quiz/](https://thehilichurl.github.io/slide_quiz/)**

---

## 🌟 Điểm Nổi Bật (Key Features)

1. **Chuẩn Mực Thiết Kế Đại Nam University ("Đại Nam Style"):**
   - Bộ nhận diện thương hiệu chuẩn xác: Xanh Hoàng Gia (`#003882`), Cam Năng Động (`#EA580C`), Cam Điểm Nhấn Slogan (`#FF7A00`), Xanh Lục Emerald (`#10B981`).
   - Font chữ chuẩn học thuật: `Times New Roman` cho toàn bộ slide và file xuất.
   - Slogan bản quyền nổi bật trên thanh Header: `| HỌC ĐỂ THAY ĐỔI`.

2. **Khung Hình Chiếu Chuẩn 16:9 (1920 × 1080 px):**
   - Canvas cố định tự động co giãn thông minh (`scale`) vừa vặn mọi tỷ lệ màn hình máy tính, máy chiếu hội trường hay màn hình LED.
   - Tối ưu kích cỡ hiển thị: Tiêu đề 42px bold, Chữ cái phương án 35px bold, Nội dung đáp án 33px.

3. **Công Cụ Xuất PowerPoint Soạn Thảo (100% Native Editable PPTX):**
   - Chuẩn hoá trực tiếp giữa các thẻ HTML/CSS DOM với các đối tượng thiết kế nguyên bản của Microsoft PowerPoint:
     - Thẻ card $\rightarrow$ `ROUNDED_RECTANGLE` (Shapes).
     - Tiêu đề, nội dung $\rightarrow$ `TEXT_BOX` (cho phép click đúp sửa từng ký tự, font chữ, màu sắc).
     - Chữ cái đáp án A/B/C/D $\rightarrow$ `OVAL` & `TEXT_BOX`.
     - Ảnh minh hoạ tư liệu $\rightarrow$ `PICTURE` được bo góc vi mô 18px và viền trùng khít tuyệt đối với mép ảnh (100% coincident, không hở viền trắng, không méo oval).
   - Tự động loại bỏ đồng hồ đếm giờ trên bản PPTX, giữ lại chân trang sạch đẹp với liên kết clickable `dainam.edu.vn`.

4. **Nội Dung Học Phần Toàn Diện:**
   - **Bài A3:** Xây dựng nền quốc phòng toàn dân và an ninh nhân dân (10 câu trắc nghiệm trọng tâm + 10 slide phân tích đáp án đúng).
   - **Bài A4:** Chiến tranh nhân dân bảo vệ Tổ quốc Việt Nam XHCN (Cập nhật xung đột Nga - Ukraine, công nghệ cao, Đại hội XIV).

---

## 📂 Cấu Trúc Dự Án (Repository Structure)

```text
slide_quiz/
├── index.html                      # Cổng thông tin (Hub Portal) quản lý bài giảng
├── slide_A3_hoan_thien.html        # Slide trình chiếu tương tác Bài A3 (20 slide)
├── slide_A4_hoan_thien.html        # Slide trình chiếu tương tác Bài A4 (20 slide)
├── assets/
│   ├── styles.css                  # Bảng kiểu CSS chính (Đại Nam Style, HUD, Responsive)
│   ├── html_to_pptx_converter.js   # Bộ chuyển đổi DOM sang PowerPoint có thể soạn thảo
│   ├── dai-nam-logo-ngang.svg      # Logo chính thức Trường Đại học Đại Nam
│   ├── favicon.svg                 # Favicon biểu tượng Đại Nam
│   └── vn-*.jpg, war-*.jpg         # Kho ảnh tư liệu lịch sử, quân sự độ phân giải cao
├── Slide_Agent_Framework/
│   ├── scripts/                    # Scripts tự động hoá xuất PPTX, render kiểm thử
│   ├── templates/                  # Bộ mẫu layout trình chiếu
│   └── docs/                       # Tài liệu thiết kế chuẩn hoá HTML -> PPTX
├── .gitignore                      # Loại trừ file tạm, cache và file binary lớn
└── README.md                       # Tài liệu hướng dẫn sử dụng và triển khai
```

---

## ⌨️ Phím Tắt Trình Chiếu (Keyboard Shortcuts)

| Phím tắt | Thao tác điều khiển |
| :--- | :--- |
| `Space` / `→` / `PageDown` | Chuyển sang slide kế tiếp (Next Slide) |
| `←` / `PageUp` | Quay lại slide trước (Previous Slide) |
| `Home` / `End` | Nhảy về slide đầu tiên / slide cuối cùng |
| `F` | Bật / Tắt chế độ toàn màn hình (Fullscreen) |
| `T` | Hiện / Ẩn đồng hồ đếm ngược trình chiếu |
| `O` | Xem trước đáp án nhanh (Peek Answer) |

---

## 🚀 Triển Khai & Sử Dụng Trực Tuyến (Deployment)

Dự án được xây dựng hoàn toàn bằng công nghệ Web chuẩn (HTML5, Vanilla CSS, JavaScript ES6) không cần build step phức tạp.

1. **GitHub Pages:**
   - Source: nhánh `main`, thư mục root `/`.
   - Trang chủ tự động load file `index.html`.
2. **Khởi chạy cục bộ (Local Server):**
   ```bash
   # Cách 1: Sử dụng Python
   python -m http.server 8000
   
   # Cách 2: Sử dụng Node.js npx
   npx serve .
   ```
   Sau đó mở trình duyệt tại: `http://localhost:8000`.

---

© 2026 **Khoa Giáo Dục Quốc Phòng & An Ninh • Trường Đại Học Đại Nam**.
