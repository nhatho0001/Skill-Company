# Layout Patterns cho Content Slide (v2)

Reference này định nghĩa **7 layout pattern** cho content slide để chống đơn điệu. Dùng cùng với `building-blocks.md` (snippet cơ bản) và Bước 3b trong `SKILL.md` (decision tree chọn layout).

> **Ràng buộc bất biến** (xem SKILL.md):
> - **Cover, TOC, chapter divider, closing** → giữ nguyên template gốc, KHÔNG vary
> - Logo (đã embed sẵn template), background cover cam phẳng, font Inter, color đỏ-primary/navy → bất biến
> - Các pattern dưới đây **chỉ áp dụng cho content slide**

## Setup chung

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Design tokens v2 (đồng bộ design-tokens.md)
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

FONT_HEAD = "Inter"
FONT_BODY = "Inter"

# Slide size 16:9 cinematic = 20 x 11.25 inch
SLIDE_W = Inches(20)
SLIDE_H = Inches(11.25)
```

Helper `set_text()`, `add_title()` xem `building-blocks.md`.

---

## Pattern 1 — `icon_rows` (Header + Body, full-width, 4-6 rows)

**Khi dùng:** 4–6 item, body 1-2 dòng, không quá quan trọng visual.

**Layout chuẩn template (Slide 3 hiện tại):** Title đỏ trên, body bên trái với header navy + body đen, ảnh lớn bên phải (5"x5"). Mỗi item là 2 textbox liền kề.

**Khi build từ scratch:**
```python
def build_icon_rows(slide, title, items, image_path=None):
    """items: list of {"header": str, "body": str}"""
    # Title đỏ
    title_box = slide.shapes.add_textbox(Inches(0.86), Inches(0.66), Inches(11), Inches(0.9))
    set_text(title_box.text_frame, title, size=40, bold=True, color=RED_PRIMARY)

    # Items - left side (text)
    y = Inches(1.91)
    text_w = Inches(10.88)
    for it in items[:6]:
        # Header navy bold
        head_box = slide.shapes.add_textbox(Inches(1.43), y, text_w, Inches(0.27))
        set_text(head_box.text_frame, it["header"], size=18, bold=True, color=NAVY)

        # Body
        body_box = slide.shapes.add_textbox(Inches(1.43), y + Inches(0.47),
                                            text_w, Inches(0.53))
        set_text(body_box.text_frame, it["body"], size=14, color=BLACK)

        y += Inches(1.31)

    # Right side image (optional, 5x5 inch)
    if image_path:
        slide.shapes.add_picture(image_path,
            Inches(13.02), Inches(2.19), Inches(4.94), Inches(4.94))
```

**Khi pending/đang xử lý:** Tô header sang `RED_PRIMARY` (xem Bước 2b-bis trong SKILL.md).

---

## Pattern 2 — `cards_3col` (3 cột đối xứng có background hồng)

**Khi dùng:** Đúng 3 item ngang hàng (3 hạng mục/3 quy trình/3 phân hệ). Không có ảnh.

**Layout chuẩn template (Slide 5):** 3 card hồng nhạt + viền hồng, header navy + body đen. Card cao 4.36", width 5.94" mỗi card, gap nhỏ.

```python
def build_cards_3col(slide, title, items):
    """items: list of EXACTLY 3 dicts {"header": str, "body": str}"""
    assert len(items) == 3, "cards_3col cần đúng 3 item"

    # Title đỏ (có thể wrap 2 dòng)
    title_box = slide.shapes.add_textbox(Inches(1.09), Inches(2.12),
                                         Inches(17.83), Inches(1.97))
    set_text(title_box.text_frame, title, size=40, bold=True, color=RED_PRIMARY)

    # 3 cards
    card_w = Inches(5.94)
    card_h = Inches(4.36)
    y = Inches(4.71)
    x_starts = [Inches(1.10), Inches(7.03), Inches(12.97)]

    for i, it in enumerate(items):
        x = x_starts[i]
        # Card background
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = PINK_CARD
        card.line.color.rgb = PINK_BORDER
        card.line.width = Pt(1.5)

        # Header (navy bold, top of card)
        head_box = slide.shapes.add_textbox(
            x + Inches(0.31), y + Inches(0.29),
            card_w - Inches(0.62), Inches(0.6))
        set_text(head_box.text_frame, it["header"], size=20, bold=True, color=NAVY)

        # Body
        body_box = slide.shapes.add_textbox(
            x + Inches(0.31), y + Inches(1.0),
            card_w - Inches(0.62), card_h - Inches(1.2))
        set_text(body_box.text_frame, it["body"], size=14, color=BLACK)
```

**Quy tắc:** Header card ≤ 20 ký tự (Quirk 4 trong edit-template.md).

---

## Pattern 3 — `numbered_zigzag_4` (Timeline 4-step zigzag)

**Khi dùng:** 4 mốc/giai đoạn theo thứ tự, kế hoạch nâng cấp/triển khai. Layout chuẩn template (Slide 4).

**Layout:** 4 vòng tròn đỏ có số 1-4 nằm trên 1 đường thẳng đứng giữa slide. Item 1+3 lệch trái, 2+4 lệch phải (zigzag).

```python
def build_numbered_zigzag_4(slide, title, intro, items):
    """items: list of EXACTLY 4 dicts {"header": str, "body": str}"""
    assert len(items) == 4, "numbered_zigzag_4 cần đúng 4 item"

    # Title đỏ
    title_box = slide.shapes.add_textbox(Inches(1.09), Inches(1.0),
                                         Inches(12.97), Inches(0.9))
    set_text(title_box.text_frame, title, size=40, bold=True, color=RED_PRIMARY)

    # Intro
    intro_box = slide.shapes.add_textbox(Inches(1.09), Inches(2.35),
                                         Inches(17.83), Inches(0.67))
    set_text(intro_box.text_frame, intro, size=16, color=NAVY)

    # 4 items - vertical line center at x=10"
    center_x = Inches(10.0)
    y_starts = [Inches(3.45), Inches(5.12), Inches(6.56), Inches(8.01)]

    for i, it in enumerate(items):
        y_num = y_starts[i]
        # Number circle (red)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
            center_x - Inches(0.32), y_num - Inches(0.12),
            Inches(0.64), Inches(0.64))
        circ.fill.solid(); circ.fill.fore_color.rgb = RED_PRIMARY
        circ.line.color.rgb = RED_PRIMARY
        set_text(circ.text_frame, str(i+1),
                 size=22, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        circ.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Text: items 0,2 left; items 1,3 right
        if i % 2 == 0:
            tx_x = Inches(1.09); tw = Inches(7.5); align = PP_ALIGN.RIGHT
        else:
            tx_x = Inches(11.40); tw = Inches(7.5); align = PP_ALIGN.LEFT

        # Header
        head_box = slide.shapes.add_textbox(tx_x, y_num - Inches(0.04),
                                            tw, Inches(0.46))
        set_text(head_box.text_frame, it["header"],
                 size=20, bold=True, color=NAVY, align=align)

        # Body
        body_box = slide.shapes.add_textbox(tx_x, y_num + Inches(0.5),
                                            tw, Inches(1.3))
        set_text(body_box.text_frame, it["body"],
                 size=14, color=BLACK, align=align)
```

**Quy tắc:** Header zigzag ≤ 20 ký tự (Quirk 5).

---

## Pattern 4 — `data_table` (Bảng dữ liệu, số col/row động) ⭐ MỚI

**Khi dùng:**
- So sánh ≥ 3 cột metric (kế hoạch vs thực tế, KPI ngang theo phòng/đơn vị)
- User yêu cầu rõ ràng "làm bảng" / "so sánh dạng bảng"

**Layout chuẩn template (Slide 11):** Tựa đề đỏ + intro xanh + bảng full-width có header xanh `#4472C4`, body alternating xám đậm/xám nhạt, chữ đen bold.

**Số columns / rows quyết định bởi DỮ LIỆU**, không cố định 3×4 như placeholder.

### Cách 1: Edit template (replace placeholder table)

```python
def replace_table_in_slide(slide, headers, rows,
                            x=2.25, y=3.38, w=14.42, max_h=7.5):
    """
    Xóa table placeholder cũ trên slide, thêm table mới dùng data thực.
    headers: list of column names, vd ["Hạng mục", "KPI 1", "KPI 2", "KPI 3"]
    rows: list of list values, mỗi row khớp số cột với headers
    """
    # 1. Xóa table cũ
    for shape in list(slide.shapes):
        if shape.has_table:
            sp = shape._element
            sp.getparent().remove(sp)
            break

    # 2. Tính kích thước
    n_rows = len(rows) + 1  # +1 cho header
    n_cols = len(headers)
    # Auto-height: 0.7" header + 0.5" mỗi body row, capped ở max_h
    h_calc = 0.7 + 0.5 * len(rows)
    h = min(h_calc, max_h)

    # 3. Add table mới
    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = tbl_shape.table

    # 4. Style header row
    for c, hd in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_TBL_HDR
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(hd)
        run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = WHITE
        run.font.name = FONT_HEAD
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 5. Body rows alternating
    for r, row_data in enumerate(rows, start=1):
        bg = GRAY_TBL_EVEN if r % 2 == 1 else GRAY_TBL_ODD
        for c, val in enumerate(row_data):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            tf = cell.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(val)
            run.font.bold = True
            run.font.size = Pt(14)
            run.font.color.rgb = BLACK
            run.font.name = FONT_BODY
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    return tbl_shape
```

### Cách 2: Build mới + cập nhật title

```python
# Replace title của slide 11
replace_shape_by_name(slide11,
    "TextBox X",  # tên textbox title (cần debug shape names trước)
    "BẢNG SO SÁNH KPI THÁNG 10/2025"
)
# Replace subtitle
replace_shape_by_name(slide11, "TextBox Y",
    "So sánh chỉ số hoàn thành công việc giữa các nhóm")
# Replace table
replace_table_in_slide(slide11,
    headers=["Nhóm", "Tổng CV", "Đã xong", "% hoàn tất"],
    rows=[
        ["A. Hạ tầng", 45, 42, "93%"],
        ["B. ERP",      28, 25, "89%"],
        ["C. Eoffice",  18, 17, "94%"],
        ["D. CĐS/AI",   12, 9,  "75%"],
    ]
)
```

### Quy tắc số columns / rows

| Số cột | Dùng cho |
|---|---|
| 2 | KHÔNG khuyến khích — dùng `cards_2col` thay vì table |
| 3 | Tiêu chí + 2 metric (vd: Hạng mục/KPI/Trạng thái) |
| 4 | Tiêu chí + 3 metric (vd: Nhóm/Tổng/Done/% — chuẩn nhất) |
| 5 | Tiêu chí + 4 metric (vd: thêm 1 cột "Pending" hoặc "Ghi chú") |
| 6+ | Cẩn thận — 6 cột × slide 20" rộng = 3.3"/cột (ổn nếu text ngắn). 7+ cột nên rút bỏ hoặc tách 2 bảng. |

| Số rows (không tính header) | Dùng cho |
|---|---|
| 3-5 | Tốt nhất — xem rõ, không tràn |
| 6-10 | OK, height auto-tính, cap ở 7.5" |
| 11-15 | Body row 14pt, có thể chật, cân nhắc tách 2 slide |
| 16+ | KHÔNG khuyến khích — rút gọn data hoặc tách thành nhiều bảng |

### Quy tắc nội dung cell

- **Cell text ngắn**: ưu tiên 1-3 từ hoặc số. Nếu cần câu dài, dùng `icon_rows` thay vì table.
- **Số có đơn vị**: viết liền ("93%", "120 ticket", "2.5h"), không tách dòng.
- **NA/empty**: dùng "—" (em-dash), không để ô trống.
- **Tô màu cell theo trạng thái** (optional): pending → cell màu hồng nhạt, done → giữ alternating mặc định.

---

## Pattern 5 — `image_card_3col` (3 card có ảnh + title + body)

**Khi dùng:** Dự án trọng điểm 3 hạng mục có ảnh minh hoạ; hoặc 3 hoạt động khác có ảnh đại diện. Layout chuẩn template (Slide 8, 12).

**Layout:** 3 ảnh ngang (5.29" × 3.93") trên + title navy giữa + body đen dưới. Ảnh thường là illustration/photo có depth.

```python
def build_image_card_3col(slide, title, items, image_paths=None):
    """
    items: list of EXACTLY 3 {"header": str, "body": str}
    image_paths: list of 3 paths, hoặc None → placeholder navy
    """
    assert len(items) == 3
    image_paths = image_paths or [None] * 3

    # Title đỏ
    title_box = slide.shapes.add_textbox(Inches(1.12), Inches(1.72),
                                         Inches(13), Inches(0.91))
    set_text(title_box.text_frame, title, size=40, bold=True, color=RED_PRIMARY)

    # 3 cards (image + title + body) — same Y
    img_w = Inches(5.29)
    img_h = Inches(3.93)
    y_img = Inches(4.35)
    x_starts = [Inches(1.12), Inches(7.14), Inches(13.59)]

    for i, it in enumerate(items):
        x = x_starts[i]
        # Image hoặc placeholder navy
        if image_paths[i]:
            slide.shapes.add_picture(image_paths[i], x, y_img, img_w, img_h)
        else:
            ph = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, x, y_img, img_w, img_h)
            ph.fill.solid(); ph.fill.fore_color.rgb = NAVY
            ph.line.fill.background()

        # Header navy bold (below image)
        head_box = slide.shapes.add_textbox(
            x + Inches(0.16), Inches(8.74),
            img_w - Inches(0.32), Inches(0.56))
        set_text(head_box.text_frame, it["header"],
                 size=22, bold=True, color=NAVY, align=PP_ALIGN.LEFT)

        # Body
        body_box = slide.shapes.add_textbox(
            x + Inches(0.16), Inches(9.40),
            img_w - Inches(0.32), Inches(1.5))
        set_text(body_box.text_frame, it["body"],
                 size=14, color=BLACK, align=PP_ALIGN.LEFT)
```

**Image strategy:** Skill này không tự sinh ảnh AI. Khi build pattern này:
1. Ảnh do user cung cấp → dùng `img_path`
2. Không có ảnh → placeholder navy (mặc định)
3. Ảnh trong template gốc (slide 8/12 đã có 3 ảnh placeholder Gamma) → giữ nguyên nếu chỉ thay text, hoặc xóa và dùng placeholder navy.

---

## Pattern 6 — `timeline_4_horizontal` (Timeline 4-step ngang dọc)

**Khi dùng:** Lộ trình dài hạn theo năm/quý/giai đoạn (≠ zigzag pattern 3 vốn cho 4 bước cùng kỳ).

**Layout chuẩn template (Slide 10):** 4 ô có icon trong vòng tròn đỏ, line đứt nối qua label, layout dọc (year 1 trên cùng → year 4 dưới cùng), title text lớn ở cột trái.

```python
def build_timeline_4_horizontal(slide, title, items):
    """items: 4 dicts {"label": "Year 01"/"Q1", "header": str, "body": str}"""
    # Title lớn cột trái
    title_box = slide.shapes.add_textbox(Inches(0.73), Inches(4.65),
                                         Inches(5.96), Inches(2.12))
    set_text(title_box.text_frame, title, size=44, bold=True, color=NAVY)

    # 4 step zigzag dọc
    y_starts = [Inches(1.01), Inches(3.44), Inches(6.47), Inches(9.26)]
    label_x = [Inches(7.02), Inches(9.23), Inches(9.23), Inches(7.02)]
    text_x  = [Inches(11.74), Inches(13.33), Inches(13.33), Inches(11.74)]

    for i, it in enumerate(items[:4]):
        y = y_starts[i]
        # Circle icon (red)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
            label_x[i] - Inches(1.6), y - Inches(0.13),
            Inches(1.24), Inches(1.24))
        circ.fill.solid(); circ.fill.fore_color.rgb = RED_PRIMARY
        circ.line.fill.background()

        # Label (Year 01 / Q1...)
        lbl = slide.shapes.add_textbox(label_x[i], y + Inches(0.11),
                                       Inches(2.5), Inches(0.71))
        set_text(lbl.text_frame, it["label"],
                 size=28, bold=True, color=NAVY)

        # Header + Body (right column)
        head_box = slide.shapes.add_textbox(text_x[i], y + Inches(0.05),
                                            Inches(5.96), Inches(0.4))
        set_text(head_box.text_frame, it["header"],
                 size=16, bold=True, color=NAVY)

        body_box = slide.shapes.add_textbox(text_x[i], y + Inches(0.45),
                                            Inches(5.96), Inches(0.6))
        set_text(body_box.text_frame, it["body"],
                 size=13, color=BLACK)
```

---

## Pattern 7 — `chart_with_text` (Chart + text giải thích bên cạnh)

**Khi dùng:** Visualize số liệu định lượng, có data đáng vẽ chart. Layout chuẩn template (Slide 13 = donut left/text right, Slide 14 = column left/text right).

**Layout:** Chart 13.33"×8.89", text panel 4-5" rộng + nhiều paragraph (subtitle ngắn + paragraph dài giải thích).

### Replace chart data trong template

```python
from pptx.chart.data import CategoryChartData

def replace_chart_data(slide, categories, series_data):
    """
    Replace data của chart hiện có trên slide (slide 13 hoặc 14).
    categories: list of category names, vd ["Q1", "Q2", "Q3", "Q4"]
    series_data: dict {series_name: [value1, value2, ...]}, vd {"Sales": [25, 30, 22, 28]}
    """
    for shape in slide.shapes:
        if shape.has_chart:
            chart_data = CategoryChartData()
            chart_data.categories = categories
            for series_name, values in series_data.items():
                chart_data.add_series(series_name, values)
            shape.chart.replace_data(chart_data)
            return shape.chart
    return None

# Example usage cho slide 13 (donut)
replace_chart_data(prs.slides[12],
    categories=["A. Hạ tầng", "B. ERP", "C. Eoffice", "D. CĐS"],
    series_data={"% Hoàn thành": [42, 25, 17, 9]}
)

# Example usage cho slide 14 (column)
replace_chart_data(prs.slides[13],
    categories=["T7", "T8", "T9", "T10"],
    series_data={
        "Đã xong": [30, 35, 28, 42],
        "Pending": [5, 8, 6, 4],
        "Tổng":    [35, 43, 34, 46],
    }
)
```

### Build chart từ scratch

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

def build_chart_with_text(slide, title, subtitle, paragraph,
                          chart_type, categories, series_data,
                          chart_side="right"):
    """
    chart_type: "donut" / "column" / "bar" / "line"
    chart_side: "left" hoặc "right"
    """
    # Vùng chart vs text
    if chart_side == "left":
        ch_x, ch_y, ch_w, ch_h = 0.66, 0.88, 13.33, 8.89
        tx_x = 14.90; tw = 4.37
    else:
        ch_x, ch_y, ch_w, ch_h = 5.33, 0.88, 13.33, 8.89
        tx_x = 0.74; tw = 4.03

    # Title đỏ
    t_box = slide.shapes.add_textbox(Inches(tx_x), Inches(0.73),
                                     Inches(tw), Inches(1.5))
    set_text(t_box.text_frame, title, size=36, bold=True, color=RED_PRIMARY)

    # Subtitle navy
    sub_box = slide.shapes.add_textbox(Inches(tx_x), Inches(2.3),
                                       Inches(tw), Inches(2.5))
    set_text(sub_box.text_frame, subtitle, size=18, bold=True, color=NAVY)

    # Paragraph (long text dưới)
    p_box = slide.shapes.add_textbox(Inches(tx_x), Inches(8.67),
                                     Inches(tw), Inches(1.85))
    set_text(p_box.text_frame, paragraph, size=14, color=NAVY)

    # Chart
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for s_name, values in series_data.items():
        chart_data.add_series(s_name, values)

    type_map = {
        "donut":  XL_CHART_TYPE.DOUGHNUT,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar":    XL_CHART_TYPE.BAR_CLUSTERED,
        "line":   XL_CHART_TYPE.LINE,
        "pie":    XL_CHART_TYPE.PIE,
    }
    ct = type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_shape = slide.shapes.add_chart(ct,
        Inches(ch_x), Inches(ch_y), Inches(ch_w), Inches(ch_h),
        chart_data)
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return chart
```

---

## Pattern 8 — `four_col_summary` (4 cột tổng hợp với dash divider)

**Khi dùng:** Tóm tắt 4 thành tựu/kết quả của kỳ — như slide E template gốc (Slide 7).

**Layout:** 4 cột đối xứng, mỗi cột có dash "—" ở đầu, header navy bold giữa, body đen dưới. Width mỗi cột 4.17", gap 0.4".

```python
def build_four_col_summary(slide, title, items):
    """items: list of EXACTLY 4 {"header": str, "body": str}"""
    assert len(items) == 4

    # Title đỏ
    title_box = slide.shapes.add_textbox(Inches(1.09), Inches(2.69),
                                         Inches(16.67), Inches(0.95))
    set_text(title_box.text_frame, title, size=40, bold=True, color=RED_PRIMARY)

    # 4 columns
    col_w = Inches(4.17)
    x_starts = [Inches(1.09), Inches(5.64), Inches(10.19), Inches(14.75)]

    for i, it in enumerate(items):
        x = x_starts[i]
        # Dash divider
        dash = slide.shapes.add_textbox(x, Inches(4.62), col_w, Inches(0.87))
        set_text(dash.text_frame, "—", size=44, bold=True, color=NAVY)

        # Header navy
        head_box = slide.shapes.add_textbox(
            x + Inches(0.14), Inches(5.86),
            col_w - Inches(0.28), Inches(0.49))
        set_text(head_box.text_frame, it["header"],
                 size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

        # Body
        body_box = slide.shapes.add_textbox(x, Inches(6.45),
                                            col_w, Inches(2.0))
        set_text(body_box.text_frame, it["body"],
                 size=14, color=BLACK, align=PP_ALIGN.CENTER)
```

**Quy tắc:** Header 4-cột ≤ 16 ký tự (Quirk 7).

---

## Decision tree (quick reference từ Bước 3b)

```
Section có data số đáng visualize? ──── yes ──→ chart_with_text (P7)
       │ no
       ▼
So sánh ≥ 3 cột metric? ───────────── yes ──→ data_table (P4) ⭐
       │ no
       ▼
4 mốc thời gian (năm, quý)? ────────── yes ──→ timeline_4_horizontal (P6)
       │ no
       ▼
4 bước/giai đoạn cùng kỳ? ──────────── yes ──→ numbered_zigzag_4 (P3)
       │ no
       ▼
3 dự án có ảnh minh hoạ? ─────────────  yes ──→ image_card_3col (P5)
       │ no
       ▼
Đúng 3 item đối xứng (≤ 80 từ/item)? ─ yes ──→ cards_3col (P2)
       │ no
       ▼
Đúng 4 thành tựu tóm tắt? ──────────── yes ──→ four_col_summary (P8)
       │ no
       ▼
                                              icon_rows (P1, default)
```

## Variation rule

- **Không lặp** cùng pattern quá 2 slide liên tiếp.
- Mỗi báo cáo nên có **ít nhất 2 pattern khác nhau** trong các content slide (chống đơn điệu).
- **Cover (slide 1), TOC (slide 2), closing (slide 15)** không vary — luôn dùng template gốc.
- Pattern 4 (`data_table`) và Pattern 7 (`chart_with_text`) chỉ thêm khi **data thực sự cần** — đừng lạm dụng để filler.

## Anti-patterns (tránh)

- ❌ Dùng `data_table` cho 2 cột → dùng `cards_2col` đẹp hơn (tách từ pattern 2).
- ❌ Dùng `data_table` 16+ rows → tách 2 bảng hoặc rút gọn data.
- ❌ Dùng `cards_3col` cho 2 hoặc 4 item → mất cân đối.
- ❌ Dùng `image_card_3col` không có ảnh & không placeholder → trống nửa slide.
- ❌ Dùng `chart_with_text` cho < 3 data point → dùng `big_stat` (không có trong v2 nhưng có thể custom).
- ❌ Lặp `icon_rows` cho cả 5 content slide → nhàm chán.
- ❌ Đè layout custom lên cover/TOC/closing → vi phạm Ràng buộc bất biến.

## Image strategy (v2)

Skill này **không tự sinh ảnh AI**. Khi dùng pattern có chỗ ảnh:

1. **Ảnh do user cung cấp** — đường dẫn trong `/mnt/user-data/uploads/` hoặc user cho path cụ thể
2. **Giữ ảnh placeholder của template gốc** — slide 8/12 đã có 3 ảnh từ Gamma export; có thể giữ nguyên cho hoạt động khác/dự án dù nội dung khác (vì là illustration generic)
3. **Placeholder shape navy** — fallback nếu xóa ảnh template + không có ảnh user
4. **Đề xuất user upload ảnh** — khi user yêu cầu visual cao mà không có nguồn

Tuyệt đối không tải ảnh từ internet (offline environment + copyright).
