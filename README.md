# slide_quiz - Hệ Thống Ôn Tập & Trình Chiếu GDQP-AN | Đại Học Đại Nam

> **Nền tảng slide bài giảng điện tử tương tác 2K kết hợp công cụ trắc nghiệm và bộ chuyển đổi xuất file PowerPoint (.pptx) có thể chỉnh sửa 100% (Native Editable PPTX).**

🌐 **Liên Kết Truy Cập Trực Tuyến (Live Websites - Dành cho Máy tính & Điện thoại):**  
* 🚀 **Chuyên Đề Mới (Slide New):** 👉 **[https://thehilichurl.github.io/slide_quiz/slide_new/](https://thehilichurl.github.io/slide_quiz/slide_new/)** *(Mở trực tiếp trên điện thoại & máy tính)*
* 🏛️ **Cổng Thông Tin Quản Lý (Hub Portal):** 👉 **[https://thehilichurl.github.io/slide_quiz/](https://thehilichurl.github.io/slide_quiz/)**
* 📖 **Bài Giảng Ôn Tập A3:** 👉 **[https://thehilichurl.github.io/slide_quiz/slide_A3_hoan_thien.html](https://thehilichurl.github.io/slide_quiz/slide_A3_hoan_thien.html)**
* 📖 **Bài Giảng Ôn Tập A4:** 👉 **[https://thehilichurl.github.io/slide_quiz/slide_A4_hoan_thien.html](https://thehilichurl.github.io/slide_quiz/slide_A4_hoan_thien.html)**

---

## 🌟 Điểm Nổi Bật (Key Features)

1. **Chuẩn Mực Thiết Kế Đại Nam University ("Đại Nam Style"):**
   - Bộ nhận diện thương hiệu chuẩn xác: Xanh Hoàng Gia (`#003882`), Cam Năng Động (`#EA580C`), Cam Điểm Nhấn Slogan (`#FF7A00`), Xanh Lục Emerald (`#10B981`).
   - Slogan bản quyền nổi bật trên thanh Header: `| HỌC ĐỂ THAY ĐỔI`.
   - Bài giảng mới `slide_new` sử dụng cặp font hiện đại **Plus Jakarta Sans** & **Be Vietnam Pro** với cỡ chữ to rõ, chuẩn chiếu hội trường và đọc tốt trên thiết bị di động.

2. **Khung Hình Chiếu Chuẩn 16:9 (1920 × 1080 px):**
   - Canvas cố định tự động co giãn thông minh (`scale`) vừa vặn mọi tỷ lệ màn hình máy tính, máy chiếu hội trường, màn hình LED hay điện thoại di động.
   - Tối ưu kích cỡ hiển thị: Tiêu đề 42-48px bold, Nội dung 24-28px.

3. **Tích Hợp Sơ Đồ Vector SVG & Hình Ảnh Tư Liệu 100% Việt Nam:**
   - Sơ đồ Thành Cổ Loa 3 vòng xoáy trôn ốc & Nỏ thần Liên Châu.
   - Sơ đồ chiến thuật sông Bạch Đằng: quy luật thủy triều, bãi cọc nhọn bọc sắt.
   - Sơ đồ 4 trụ cột quân sự truyền thống và chu trình "Ngụ binh ư nông".
   - Bảng ma trận đối sánh Binh pháp cổ xưa $\leftrightarrow$ Kiến trúc An toàn thông tin (Zero Trust, Threat Hunting, Honeypot, Cloud DDoS).

4. **Nội Dung Học Phần Toàn Diện:**
   - **Chuyên đề Mới (slide_new):** Truyền thống & Nghệ thuật đánh giặc của ông cha ta (21 slide chuyên sâu, phân tích cội nguồn, 4 trụ cột và liên hệ sinh viên CNTT).
   - **Bài A3:** Xây dựng nền quốc phòng toàn dân và an ninh nhân dân (10 câu trắc nghiệm trọng tâm + 10 slide phân tích đáp án đúng).
   - **Bài A4:** Chiến tranh nhân dân bảo vệ Tổ quốc Việt Nam XHCN (Cập nhật xung đột Nga - Ukraine, công nghệ cao, Đại hội XIV).

---

## 📂 Cấu Trúc Dự Án (Repository Structure)

```text
slide_quiz/
├── index.html                      # Cổng thông tin (Hub Portal) quản lý bài giảng (3 bộ deck)
├── slide_new/                      # [MỚI] Chuyên đề Truyền thống & Nghệ thuật đánh giặc (21 Slide)
│   ├── index.html                  # Giao diện trình chiếu tương tác hiện đại
│   └── assets/                     # Tài nguyên styles.css, sơ đồ SVG và ảnh tư liệu Việt Nam
├── slide_A3_hoan_thien.html        # Slide trình chiếu tương tác Bài A3 (20 slide)
├── slide_A4_hoan_thien.html        # Slide trình chiếu tương tác Bài A4 (20 slide)
├── assets/                         # Thư viện ảnh, kiểu dáng gốc và bộ xuất PPTX
├── Slide_Agent_Framework/          # Công cụ framework hỗ trợ chuyển đổi & render
├── raw.md                          # Tài liệu cốt lõi nội dung chuyên đề
└── README.md                       # Tài liệu hướng dẫn, liên kết trực tuyến và triển khai
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
