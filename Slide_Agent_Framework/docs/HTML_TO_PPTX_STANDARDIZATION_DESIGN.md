# BẢN THIẾT KẾ CHUẨN HOÁ TÍNH NĂNG XUẤT POWERPOINT SOẠN THẢO (EDITABLE PPTX)
## Hệ Thống Chuẩn Hoá Thẻ HTML/CSS Sang Công Cụ Thiết Kế PowerPoint (OpenXML Primitives)
### Dự án: Bài Giảng Điện Tử Đại Học Đại Nam ("Đại Nam Style")

---

## 1. Bối Cảnh & Mục Tiêu Kỹ Thuật

### 1.1 Vấn đề của phương pháp xuất PPTX truyền thống (Ảnh Bitmap 2K)
Trước đây, các công cụ xuất file PowerPoint từ Web Slide chủ yếu sử dụng `html2canvas` để chụp lại toàn bộ màn hình trình duyệt thành một bức ảnh PNG/JPEG lớn, sau đó dán kín slide PowerPoint (`slide.addImage(...)`).
* **Hạn chế nghiêm trọng**:
  - Toàn bộ văn bản bị biến thành điểm ảnh (pixel), **không thể chỉnh sửa**, không thể bôi đen copy, không tìm kiếm được nội dung.
  - Các khối thẻ (cards), huy hiệu (badges), đường viền (borders) bị khóa cứng, không thể di chuyển hoặc thay đổi màu sắc.
  - Hình ảnh minh họa bị dính chặt vào nền slide, không thể thay thế bằng ảnh khác.
  - Dung lượng file PPTX lớn nhưng giá trị tái sử dụng thấp.

### 1.2 Giải pháp đột phá: Kiến trúc chuẩn hoá Đối tượng (DOM-to-PowerPoint Engine)
Thiết kế tính năng xuất PPTX mới dựa trên nguyên lý **Chuẩn hoá ánh xạ 1:1** giữa các thẻ HTML & thuộc tính CSS trong trình duyệt với các công cụ thiết kế đồ họa nguyên bản của Microsoft PowerPoint (Shapes, TextBoxes, Ovals, Picture Objects, Lines, Gradients, Hyperlinks).

* **Kết quả đạt được**:
  - File `.pptx` mở bằng Microsoft PowerPoint, Google Slides, Keynote hoặc WPS Office có thể **chỉnh sửa 100% từng từ, từng khối hình, từng màu sắc, từng bức ảnh**.
  - **Giữ nguyên 100% tỷ lệ, bố cục, khoảng cách, màu sắc thương hiệu và font chữ chuẩn** của thiết kế gốc trên web mà không bị xô lệch.

---

## 2. Hệ Quy Chiếu Hình Học & Chuyển Đổi Tọa Độ (16:9 Projection)

Mọi slide trình chiếu đều được thiết kế trên một không gian tọa độ ảo cố định tỷ lệ 16:9 (`1920 x 1080px`).
Trong Microsoft PowerPoint hiện đại, kích thước chuẩn của slide Widescreen 16:9 là **13.333 inch x 7.5 inch** (tương đương $33.867 \text{ cm} \times 19.05 \text{ cm}$).

### 2.1 Ma trận chuyển đổi tọa độ (Pixels -> Inches)
$$\text{Scale Ratio} = \frac{13.333333 \text{ inch}}{1920 \text{ px}} = \frac{7.5 \text{ inch}}{1080 \text{ px}} = \frac{1}{144} \approx 0.0069444 \text{ inch/px}$$

Như vậy:
$$\text{X}_{\text{inch}} = \frac{\text{X}_{\text{px}}}{144}, \quad \text{Y}_{\text{inch}} = \frac{\text{Y}_{\text{px}}}{144}$$
$$\text{Width}_{\text{inch}} = \frac{\text{Width}_{\text{px}}}{144}, \quad \text{Height}_{\text{inch}} = \frac{\text{Height}_{\text{px}}}{144}$$

### 2.2 Quy đổi kích thước Font chữ (CSS px -> PPTX pt)
Trong thiết kế typography và chuẩn OpenXML PowerPoint:
$$1 \text{ pt} = \frac{1}{72} \text{ inch}$$
Vì $1 \text{ inch} = 144 \text{ virtual px}$, ta có tỷ lệ vàng:
$$1 \text{ pt} = \frac{144}{72} \text{ px} = 2 \text{ px}$$
$$\text{FontSize}_{\text{pt}} = \frac{\text{FontSize}_{\text{px}}}{2}$$

**Bảng đối chiếu kích thước chữ thực tế:**
| Phần tử HTML | Kích thước CSS (px) | Kích thước PowerPoint (pt) | Vai trò trong PowerPoint |
| :--- | :--- | :--- | :--- |
| Tiêu đề Hero (`.hero-main-title`) | 56 px | **28 pt** | Title chính cực kỳ sắc nét |
| Tiêu đề câu hỏi (`.quiz-question-title`) | 42 px | **21 pt** | Tiêu đề câu hỏi chuẩn hội trường |
| Slogan thương hiệu (`.brand-slogan`) | 39 px | **19.5 pt** | Header Display Text |
| Nội dung đáp án (`.opt-text`) | 28 px | **14 pt** | Body Text dễ đọc, tự động ngắt dòng |
| Huy hiệu danh mục/Câu hỏi (`.quiz-badge`) | 27 px | **13.5 pt** | Badge Text in hoa, căn giữa |
| Chữ cái phương án (`.opt-letter`) | 28 px | **14 pt** | Indicator A, B, C, D |
| Thông tin giảng viên (`.footer-center`) | 21 px | **10.5 pt** | Metadata Footer |
| Đồng hồ & Liên kết (`.footer-timer`) | 18 px | **9.0 pt** | Micro Footer Action Tag |

---

## 3. Ma Trận Chuẩn Hoá Thẻ HTML Sang Công Cụ Thiết Kế PowerPoint

| STT | Phần tử HTML / Class CSS | Thuộc tính CSS Trích xuất | Công cụ thiết kế bên PowerPoint (OpenXML Primitive) | Tham số ánh xạ chi tiết |
| :---: | :--- | :--- | :--- | :--- |
| **1** | `<section class="slide">` | `background-color`, `background-image` | `pptx.Slide.background` | Màu nền slide nguyên bản (`#F8FAFC` cho theme sáng hoặc `#050B14` cho theme tối). |
| **2** | `<header class="slide-header">` | `height: 86px`, `background: #003882` | `slide.addShape(RECTANGLE)` | Khối hộp chữ nhật trải dài toàn bộ chiều ngang ($x=0, y=0, w=13.333, h=0.597$), tô màu xanh Cobalt (`#003882`). |
| **3** | Border dưới Header | `border-bottom: 3px solid #EA580C` | `slide.addShape(LINE)` | Đường line kẻ ngang sắc nét tại $y=0.597$, màu cam rực rỡ (`#EA580C`), độ dày 2.5pt. |
| **4** | Slogan Bar (`.brand-slogan::before`) | `width: 8px, height: 42px`, `background: #FF7A00`, `border-radius: 4px` | `slide.addShape(ROUNDED_RECTANGLE)` | Khối chữ nhật bo góc đứng ($w=0.055, h=0.292$), bo cong `rectRadius: 0.5`, tô màu cam sáng (`#FF7A00`). |
| **5** | Slogan Text (`.brand-slogan`) | `font-size: 39px`, `color: #FF7A00`, `font-weight: bold` | `slide.addText("HỌC ĐỂ THAY ĐỔI")` | Khung Textbox độc lập, font `Times New Roman`, màu `#FF7A00`, bold, khoảng cách ký tự `charSpacing: 2`. |
| **6** | Logo Đại Học Đại Nam (`.official-logo-img`) | SVG/PNG, `height: 56px`, `object-fit: contain` | `slide.addImage(...)` | Đối tượng Picture Object độc lập đặt tại góc phải header ($x=11.41, y=0.097, w=1.55, h=0.389$). Người dùng có thể bấm chọn, phóng to/thu nhỏ hoặc thay logo khác. |
| **7** | Huy hiệu Câu hỏi / Đáp án (`.quiz-badge-question`, `.quiz-badge-answer`) | `background-color`, `border-radius: 10px`, `color: #FFFFFF` | `slide.addShape(ROUNDED_RECTANGLE)` + `slide.addText(...)` | Khối Shape bo góc tô nền `#003882` hoặc `#EA580C`, bên trong chứa text in hoa căn giữa dọc/ngang. |
| **8** | Huy hiệu Chủ đề (`.quiz-topic-badge`) | `background: rgba(0,56,130,0.1)`, `border: 1px solid #003882` | `slide.addShape(ROUNDED_RECTANGLE)` + `slide.addText(...)` | Khối Shape viền mỏng xanh `#003882`, nền xanh nhạt, chữ màu xanh `#003882`. |
| **9** | Tiêu đề câu hỏi (`.quiz-question-title`) | `font-size: 42px`, `font-weight: bold`, `color: #0F172A`, `line-height: 1.2` | `slide.addText(titleText)` | Khung Textbox rộng 8.01 inch, tự động xuống dòng (`wrap: true`), khoảng cách dòng `lineSpacingMultiple: 1.2`. Có thể sửa từng chữ. |
| **10** | Thẻ Phương án thường (`.quiz-option-card`) | `background: #FFFFFF`, `border: 1.5px solid #CBD5E1`, `border-radius: 14px` | `slide.addShape(ROUNDED_RECTANGLE)` | Khối Shape bo góc thẻ đáp án ($w=8.01, h=0.514$), bo viền `rectRadius: 0.15`, màu viền xám sáng `#CBD5E1`. |
| **11** | Chữ cái phương án (`.opt-letter`) | `color: #EA580C`, `font-weight: bold`, `font-size: 28px` | `slide.addText("A.")` | Khung Textbox định vị đầu thẻ đáp án, màu cam thương hiệu `#EA580C`. |
| **12** | Chữ cái phương án Đúng (`.opt-letter-circle`) | `width: 44px, height: 44px`, `border-radius: 50%`, `background: #10B981` | `slide.addShape(OVAL)` + `slide.addText("B")` | Khối Hình tròn Oval màu xanh lục Emerald (`#10B981`), bên trong chứa chữ cái in hoa màu trắng căn giữa hoàn hảo. |
| **13** | Nội dung câu trả lời (`.opt-text`) | `font-size: 28px`, `color: #0F172A`, `valign: middle` | `slide.addText(...)` | Khung Textbox chứa lời đáp án, căn lề giữa theo chiều dọc (`valign: middle`), font `Times New Roman`. |
| **14** | Thẻ Phương án Đúng (`.quiz-option-card.correct-answer`) | `background: #ECFDF5`, `border: 2.5px solid #10B981` | `slide.addShape(ROUNDED_RECTANGLE)` | Khối Shape thẻ đáp án đúng với nền xanh ngọc dịu `#ECFDF5`, viền xanh ngọc đậm `#10B981` dày 2.5pt. |
| **15** | Ruy băng ĐÁP ÁN ĐÚNG (`.quiz-correct-tag`) | `background: #10B981`, `color: #FFFFFF`, `border-radius: 9999px` | `slide.addShape(ROUNDED_RECTANGLE)` + `slide.addText(...)` | Huy hiệu pill màu xanh lục nằm góc phải thẻ đáp án, text "ĐÁP ÁN ĐÚNG" 9pt màu trắng bold. |
| **16** | Khung chứa ảnh (`.quiz-visual-card`) | `background: #FFFFFF`, `border: 2.5px solid #10B981` (đáp án) hoặc `#E2E8F0` (câu hỏi) | `slide.addShape(ROUNDED_RECTANGLE)` | Khối Shape khung viền tranh ảnh đặt ở cột phải ($w=4.32, h=5.416$), bo góc mềm mại `rectRadius: 0.05` (~18px chuẩn xác). |
| **17** | Ảnh minh họa (`.quiz-visual-img`) | `object-fit: cover`, `border-radius: 18px` | `slide.addImage(...)` | **Đối tượng hình ảnh nguyên bản trong PowerPoint**, được cắt cúp bo tròn 18px trước bằng Canvas với nền trong suốt (alpha channel), **không dùng `rounding: true` để tránh lỗi biến dạng hình quả trứng/oval**. Người dùng có thể click chuột phải chọn *Change Picture*, cắt cúp, đổi hiệu ứng tranh ảnh tùy ý. |
| **18** | Chân trang (`.slide-footer`) | `height: 62px`, `background: #003882` | `slide.addShape(RECTANGLE)` | Khối chữ nhật footer ở đáy slide ($y=7.069, h=0.431$), màu xanh Cobalt `#003882`. |
| **19** | Đường line trên Chân trang | `border-top: 3px solid #EA580C` | `slide.addShape(LINE)` | Đường line kẻ cam nổi bật tại vị trí $y=7.069$ chạy suốt chiều ngang slide. |
| **20** | Đồng hồ đếm giờ (`.footer-timer-display`) | `display: none` trên bản PowerPoint | **Đã loại bỏ hoàn toàn** | **Không đưa đồng hồ đếm giờ vào file PowerPoint** để bài giảng tĩnh không bị đóng cứng mốc thời gian `00:00`, đáp ứng chuẩn yêu cầu thuyết trình. |
| **21** | Thông tin bản quyền & GV (`.footer-center`) | `color: #FFFFFF`, `color: #FB923C`, `font-weight: bold` | `slide.addText([Run1, Run2, Run3])` | Textbox đa phong cách (Multi-run): "Trường Đại học Đại Nam" (màu cam), " • " (màu trắng), "GV: Đào Bá Công" (màu trắng). |
| **22** | Đường dẫn trường (`.footer-dnu-link`) | `color: #FB923C`, `href="https://dainam.edu.vn"` | `slide.addText(...)` với `hyperlink` | **Liên kết Hyperlink sống trên 1 dòng đơn**, căn phải, độ rộng 1.8 inch chống ngắt dòng lỗi. Khi trình chiếu PowerPoint, người dùng bấm vào dòng chữ này sẽ tự động mở trang web trường! |

---

## 4. Thứ Tự Xếp Lớp Đối Tượng (Z-Index Layering Architecture)

Để đảm bảo trong PowerPoint người dùng có thể nhấp chuột vào bất kỳ phần tử nào mà không bị các phần tử khác che khuất hoặc gây khó khăn khi chỉnh sửa, hệ thống tuân thủ nghiêm ngặt 5 tầng xếp lớp (Z-Order):

```
+-------------------------------------------------------------+
| TẦNG 4: NATIVE TEXTFRAMES & TYPOGRAPHY RUNS                 |
| (Tiêu đề, đáp án, slogan, số trang, siêu liên kết hyperlink)  |
+-------------------------------------------------------------+
| TẦNG 3: HUY HIỆU & RUY BĂNG NỔI                             |
| (Huy hiệu câu hỏi, ruy băng "ĐÁP ÁN ĐÚNG", vòng tròn A, B, C)|
+-------------------------------------------------------------+
| TẦNG 2: PICTURE OBJECTS & LOGO VECTƠ                        |
| (Ảnh tư liệu chiến sự, quân sự, duyệt binh, logo DNU)       |
+-------------------------------------------------------------+
| TẦNG 1: SHAPES & CONTAINERS (KHỐI THẺ NỀN & VIỀN)           |
| (Khung thẻ đáp án, banner header, banner footer, khung ảnh) |
+-------------------------------------------------------------+
| TẦNG 0: NỀN SLIDE (CANVAS BACKGROUND)                       |
| (Màu nền trắng ngà #F8FAFC hoặc nền xanh tối #050B14)        |
+-------------------------------------------------------------+
```

---

## 5. Các Cơ Chế Chống Sai Lệch (Fidelity Protection Mechanisms)

1. **Khử ảnh hưởng co giãn hiển thị (Unscaled Stage Measurement)**:
   Trước khi đo đạc tọa độ hình học, bộ chuyển đổi tạm thời vô hiệu hóa `transform: scale(...)` của `#slide-stage` để lấy giá trị pixel thực tế $1920 \times 1080$ chuẩn xác đến từng $0.1\text{ px}$.
2. **Xử lý chuỗi văn bản an toàn (Cross-visibility Text Extraction)**:
   Sử dụng cơ chế `getCleanText` kết hợp linh hoạt giữa `innerText` và `textContent` giúp bảo toàn 100% nội dung chữ ngay cả khi slide đang ở trạng thái ẩn hoặc đang chạy hiệu ứng chuyển trang CSS.
3. **Bảo toàn chất lượng đồ họa và màu sắc**:
   Bộ phân giải mã màu tự động trích xuất kênh Alpha (độ trong suốt) và chuyển đổi sang chuẩn `transparency` của PowerPoint ($0 - 100\%$), giúp các huy hiệu mờ đục hiển thị mượt mà như trên web.
4. **Hỗ trợ ảnh trong môi trường bảo mật Local File (`file:///`) & Web Server**:
   Kết hợp linh hoạt giữa trực tiếp kết xuất Canvas và Fetch Blob giúp nạp ảnh trơn tru không phụ thuộc vào kết nối mạng.

---

## 6. Hướng Dẫn Sử Dụng

### Cách 1: Xuất trực tiếp trên trình duyệt (Khuyên dùng)
1. Mở bài trình chiếu trên trình duyệt (ví dụ: [slide_A4_hoan_thien.html](file:///c:/Users/Giang/Desktop/slideA3A4/slide_A4_hoan_thien.html)).
2. Nhấp vào nút **"Xuất PPTX (Soạn Thảo)"** màu cam ở thanh điều khiển (HUD) bên dưới.
3. Hộp thoại tiến trình sẽ hiển thị trực quan từng slide được chuẩn hóa. File `.pptx` hoàn chỉnh sẽ tự động tải về máy tính.

### Cách 2: Xuất tự động bằng Python CLI
Chạy lệnh trong terminal:
```bash
# Xuất bài ôn tập A4
python Slide_Agent_Framework/scripts/export_editable_pptx.py --file slide_A4_hoan_thien.html --out ./exports

# Xuất bài ôn tập A3
python Slide_Agent_Framework/scripts/export_editable_pptx.py --file slide_A3_hoan_thien.html --out ./exports
```
