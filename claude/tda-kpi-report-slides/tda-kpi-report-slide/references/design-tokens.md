# Design Tokens — Tôn Đông Á Report Template (v2)

> **Version**: Template v2 (May 2026), 15 master slides, 16:9 cinematic.
> Template gốc đã embed sẵn header trắng + logo + footer trắng (không cần file ảnh ngoài).

## Slide dimensions

- **Format:** 16:9 cinematic
- **EMU size:** `18,288,000 × 10,287,000` (= 20" × 11.25")
- **Inch size:** **20 × 11.25** ⚠️ (KHÁC template v1 vốn 16×9 hoặc 13.33×7.5)
- Khi build từ scratch:
  ```python
  prs.slide_width  = Inches(20)
  prs.slide_height = Inches(11.25)
  ```
- Khi edit template gốc: **giữ nguyên kích thước**, không chỉnh.

## Color palette (v2)

Template v2 dùng tone **đỏ rực + navy đậm** làm chính. Cam đã loại khỏi cover (cover giờ là cam phẳng `#FF6600`, không còn ảnh background phức tạp).

| Tên | Hex | Dùng cho |
|---|---|---|
| **Đỏ chính (RED_PRIMARY)** | `#FF0000` | Title slide (TOC, các section, pending, closing). Tone CHÍNH của brand v2. |
| **Cam cover (ORANGE_COVER)** | `#FF6600` | Background slide bìa (cam phẳng full slide) |
| **Navy đậm (NAVY)** | `#000099` | Body text bold, header card, số thứ tự (1,2,3,4), title cover, label cards |
| **Hồng nhạt card (PINK_CARD)** | `#FBE5E5` (xấp xỉ) | Background các card TOC + 3-col card |
| **Hồng viền (PINK_BORDER)** | `#FF9999` | Viền card (subtle) |
| **Xanh table header (BLUE_TBL_HDR)** | `#4472C4` | Background hàng header của table layout |
| **Xám table even (GRAY_TBL_EVEN)** | `#D9DBE7` | Body row chẵn của table |
| **Xám table odd (GRAY_TBL_ODD)** | `#ECEDF4` | Body row lẻ của table |
| **Trắng** | `#FFFFFF` | Background chính + chữ trên header xanh table |
| **Đen** | `#212121` | Body text |
| **Xám body (GRAY_BODY)** | `#555555` | Subtitle, intro paragraph (dưới title) |

> **Quy tắc vàng:** Không đổi màu. Nếu user yêu cầu đổi màu, cảnh báo rằng sẽ phá vỡ brand identity TDA v2. Template v2 đã chuẩn hoá — đừng pha thêm cam-coral hay xanh dương khác.

```python
# Snippet mặc định (copy-paste)
from pptx.dml.color import RGBColor

RED_PRIMARY   = RGBColor(0xFF, 0x00, 0x00)
ORANGE_COVER  = RGBColor(0xFF, 0x66, 0x00)
NAVY          = RGBColor(0x00, 0x00, 0x99)
PINK_CARD     = RGBColor(0xFB, 0xE5, 0xE5)
PINK_BORDER   = RGBColor(0xFF, 0x99, 0x99)
BLUE_TBL_HDR  = RGBColor(0x44, 0x72, 0xC4)
GRAY_TBL_EVEN = RGBColor(0xD9, 0xDB, 0xE7)
GRAY_TBL_ODD  = RGBColor(0xEC, 0xED, 0xF4)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
BLACK         = RGBColor(0x21, 0x21, 0x21)
GRAY_BODY     = RGBColor(0x55, 0x55, 0x55)
```

## Typography

- **Header & body font:** `Inter` (cả 2 dùng cùng font, weight khác nhau)
- Fallback: `Open Sans`, rồi `Arial`
- KHÔNG đổi sang Calibri/Roboto/Times.

### Size (pt) — đã hiệu chỉnh theo slide 20×11.25

| Element | Size | Weight |
|---|---|---|
| Cover title ("BÁO CÁO") | 54–60 | Bold |
| Cover period (KẾT QUẢ THÁNG…) | 40–44 | Bold |
| Cover department | 32–36 | Bold |
| Slide title (section, đỏ) | 36–44 | Bold |
| Subtitle / intro paragraph | 16–18 | Regular |
| Card header (navy) | 18–22 | Bold |
| Body text | 14–16 | Regular |
| Number badge (1,2,3 trong vòng tròn) | 22–28 | Bold |
| Table cell | 14–16 | Regular (header bold) |
| Closing message | 36–40 | Bold |

> Lý do size lớn hơn v1: slide v2 rộng 20" (vs 13.33" v1) → cần scale font tương ứng để giữ tỷ lệ.

## Layout rules

- **Margins:** Tối thiểu 0.5" từ các cạnh slide (template v2 dùng ~1.09" cho left margin chính)
- **Logo Tôn Đông Á:** Đã embed sẵn ở góc phải trên MỖI slide trong template gốc (kể cả cover). KHÔNG cần `add_picture(logo)` khi dùng Cách A (edit template).
- **Footer URL** (`www.tondonga.com.vn`) — chỉ có ở slide 1 (cover), không có ở slide khác. Đã embed sẵn.
- **Slide title:** Căn trái, đặt gần cạnh trên, màu đỏ `#FF0000`, **giữ ngắn** (≤ 50 ký tự cho 1 dòng, ≤ 80 ký tự cho 2 dòng).
- **Body content:** Căn trái, không căn giữa đoạn văn.
- **Card spacing:** Gap giữa các card 0.2–0.4" (template gốc dùng 0.2" cho TOC, 0.3-0.4" cho 3-col).
- **Không dùng:** Gạch chân tiêu đề màu full-width (AI slop), bóng đổ đậm, gradient loè loẹt.

## Iconography

- Emoji trong tiêu đề được CHẤP NHẬN (template gốc đã có 🛠️ ở slide A, 💡 ở slide D).
- Icon trong card: dùng shape đơn sắc (vòng tròn đỏ, vòng tròn navy có số), KHÔNG PNG cầu kỳ.
- Số thứ tự (1, 2, 3, 4) — template v2 đặt trong **vòng tròn đỏ + chữ trắng** (slide 9 pending), hoặc **navy circle nền hồng** (slide 2 TOC).
- **TUYỆT ĐỐI KHÔNG** dùng emoji trạng thái `✅` `⏳` `🔴` cho header pending — luôn dùng cách tô màu chữ (đỏ = pending, navy = done).

## Visual motif (v2)

- **Card hồng nhạt + viền hồng đậm** — dùng cho TOC, 3-col content (slide 5, 12).
- **Vòng tròn đỏ với số trắng** — dùng cho pending list, timeline.
- **Đường ngang đỏ ngắn (em-dash)** — dùng làm divider trên slide tổng hợp 4-cột (slide 7, ký tự "—").
- **Image card có border-radius nhẹ** — dùng cho 3-card-with-image (slide 8, 12).
- **Giữ motif này xuyên suốt** — đừng trộn các kiểu card khác nhau trong cùng 1 báo cáo.

## So sánh nhanh v1 vs v2

| Aspect | v1 | v2 (hiện tại) |
|---|---|---|
| Slide size | 13.33×7.5 hoặc 16×9 | **20×11.25** |
| Cover background | Ảnh phức tạp (`cover-background.jpg`) | Cam phẳng `#FF6600` (embed in pptx) |
| Logo | File `logo-header.jpg` riêng | Embed sẵn trong template |
| Title color (content) | Navy đậm | **Đỏ `#FF0000`** |
| Số master slides | ~10 | **15** (thêm timeline, table, chart, 3-img) |
| Font | Open Sans | **Inter** |
| Quirk Image 0–4 trên slide A | Có (cần xóa) | **Không có** (template đã sạch) |

> Khi nâng cấp script cũ lên v2: thay tất cả `Inches(13.33)` → `Inches(20)`, tất cả `Inches(7.5)` → `Inches(11.25)`, scale font ~1.5×, đổi RED slot lên làm primary thay vì navy, bỏ `add_logo()` khi dùng Cách A.
