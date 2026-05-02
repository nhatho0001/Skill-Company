# Design Tokens — Tôn Đông Á Report Template

## Slide dimensions

- **Format:** 16:9 widescreen
- **EMU size:** 14,630,400 × 8,229,600 (≈ 16" × 9")
- **Inch size:** 16 × 9
- Khi dùng `python-pptx`: `Inches(16)`, `Inches(9)` hoặc giữ nguyên kích thước template.

## Color palette

| Tên | Hex | Dùng cho |
|---|---|---|
| **Cam chính** | `#ED7D31` | Background slide bìa; nhấn accent |
| **Cam đậm** | `#FF6600` | Line dưới tiêu đề section (NẾU CÓ trong template) |
| **Đỏ tiêu đề** | `#FF0000` / `#FF0101` | Tiêu đề section (VD: "A. KẾT QUẢ CÔNG VIỆC …") |
| **Navy** | `#000099` | Text nhấn mạnh, header cột, số thứ tự |
| **Xanh đậm** | `#0000CC` | Text trên slide bìa (trên nền cam) |
| **Hồng đậm card** | `#E5B2B2` | Viền card TOC / key items |
| **Hồng nhạt card** | `#FFCCCC` / `#FBE5E5` | Background các card nội dung |
| **Trắng** | `#FFFFFF` | Background chính (slide 2–10) |
| **Đen** | `#000000` | Body text |
| **Xám** | `#A5A5A5` | Caption, ghi chú |

> **Quy tắc vàng:** Không đổi màu. Không thêm màu mới. Nếu người dùng yêu cầu đổi màu, cảnh báo rằng điều đó sẽ phá vỡ brand identity.

## Typography

- **Header font:** `Open Sans Bold`
- **Body font:** `Open Sans`
- Fallback: `Arial` (nếu Open Sans không có)

### Size (pt)

| Element | Size | Weight |
|---|---|---|
| Cover title ("BÁO CÁO") | 36–42 | Bold |
| Cover subtitle (kỳ báo cáo) | 20–24 | Bold |
| Cover department | 18 | Bold |
| Slide title (section) | 24–28 | Bold |
| Subtitle / intro paragraph | 14 | Regular |
| Card header | 14–16 | Bold |
| Body text | 11–13 | Regular |
| Number badge (1,2,3…) | 18–22 | Bold |
| Caption / footer | 9–10 | Regular |

## Layout rules

- **Margins:** Tối thiểu 0.4" từ các cạnh slide
- **Logo Tôn Đông Á:** Góc phải trên, cách cạnh ~0.3", size ~0.7" × 0.5" — áp dụng cho MỌI slide nội dung (slide 2 trở đi). Slide bìa đã có logo trong background.
- **Slide title:** Căn trái, đặt gần cạnh trên, màu đỏ `FF0000`
- **Body content:** Căn trái, không căn giữa đoạn văn
- **Card spacing:** Gap giữa các card 0.2–0.3"
- **Không dùng:** Gạch chân tiêu đề màu full-width (AI slop), bóng đổ đậm, gradient loè loẹt

## Iconography

- Emoji được chấp nhận trong tiêu đề (VD: 🛠️, 💡) — đã có trong template gốc
- Icon trong card: dùng shape đơn sắc, không PNG cầu kỳ
- Số thứ tự (1, 2, 3…) đặt trong circle navy `000099`, chữ trắng

## Visual motif

Template gốc dùng các **pill/card hồng nhạt** với border đậm hồng, content navy/đen bên trong. Giữ motif này xuyên suốt — đừng trộn các kiểu card khác nhau trong cùng 1 báo cáo.
