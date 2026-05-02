# Layout Patterns cho Content Slide

Reference này định nghĩa **6 layout pattern** cho content slide để chống đơn điệu. Dùng cùng với `building-blocks.md` (snippet cơ bản) và Bước 3b trong `SKILL.md` (decision tree chọn layout).

> **Ràng buộc bất biến** (xem SKILL.md):
> - **Cover, TOC, chapter divider, closing** → giữ nguyên template gốc, KHÔNG vary
> - Logo, background cover, font Inter, color navy/cam/đỏ → bất biến
> - Các pattern dưới đây **chỉ áp dụng cho content slide**

## Setup chung

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Design tokens (đồng bộ design-tokens.md)
NAVY        = RGBColor(0x00, 0x00, 0x99)
ORANGE      = RGBColor(0xFF, 0x66, 0x00)
RED         = RGBColor(0xFF, 0x00, 0x00)
BLACK       = RGBColor(0x21, 0x21, 0x21)
GRAY_BODY   = RGBColor(0x55, 0x55, 0x55)
BLUE_LIGHT  = RGBColor(0xDC, 0xE7, 0xFB)   # nền card highlight
GRAY_LINE   = RGBColor(0xE0, 0xE0, 0xE0)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

FONT_HEAD   = "Inter"
FONT_BODY   = "Inter"

# Slide size 16:9 = 13.33 x 7.5 inch
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)
```

Các helper `set_text()`, `add_logo()`, `add_title()` giả định đã có (xem `building-blocks.md`).

---

## Pattern 1 — `icon_rows` (Icon + Header + Body, full-width)

**Khi dùng:** 4–6 item, body ngắn (≤ 1 dòng), không có ảnh.

**Layout:** Mỗi row = 1 ô card chiều rộng full slide, bên trái icon trong vòng tròn, bên phải header (bold) + body (nhạt).

```python
def build_icon_rows(slide, title, items):
    """items: list of {"icon": "📊", "header": str, "body": str}"""
    add_title(slide, title)  # title đỏ #FF0000, 32pt, top-left

    y = Inches(1.5)
    row_h = Inches(0.95)
    gap = Inches(0.15)

    for it in items[:6]:
        # Card background
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), y, Inches(12.33), row_h
        )
        card.fill.solid()
        card.fill.fore_color.rgb = BLUE_LIGHT
        card.line.fill.background()

        # Icon (emoji hoặc unicode symbol)
        icon_box = slide.shapes.add_textbox(Inches(0.7), y + Inches(0.18), Inches(0.7), Inches(0.6))
        set_text(icon_box.text_frame, it["icon"], size=24, align=PP_ALIGN.CENTER)

        # Header
        head_box = slide.shapes.add_textbox(Inches(1.5), y + Inches(0.10), Inches(11), Inches(0.4))
        set_text(head_box.text_frame, it["header"], size=15, bold=True, color=NAVY)

        # Body
        body_box = slide.shapes.add_textbox(Inches(1.5), y + Inches(0.48), Inches(11), Inches(0.4))
        set_text(body_box.text_frame, it["body"], size=12, color=GRAY_BODY)

        y += row_h + gap
```

**Khi pending/đang xử lý:** đổi `BLUE_LIGHT` → background trắng + viền đỏ + header màu `RED`.

---

## Pattern 2 — `cards_3col` (3 cột đối xứng)

**Khi dùng:** Đúng 3 item ngang hàng (3 phòng/3 hạng mục/3 giai đoạn). Ảnh tùy chọn.

**Layout:** 3 card đứng chiều rộng bằng nhau, cách đều, có thể có ảnh placeholder phía trên.

```python
def build_cards_3col(slide, title, items, with_image=False, intro=""):
    """items: list of EXACTLY 3 dicts {"icon": str, "header": str, "body": str, "img_path": str?}"""
    assert len(items) == 3, "cards_3col cần đúng 3 item"
    add_title(slide, title)

    if intro:
        intro_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.3), Inches(0.5))
        set_text(intro_box.text_frame, intro, size=13, color=GRAY_BODY)
        y_top = Inches(2.1)
    else:
        y_top = Inches(1.6)

    card_w = Inches(4.0)
    gap = Inches(0.4)
    x_start = (SLIDE_W - 3 * card_w - 2 * gap) / 2

    for i, it in enumerate(items):
        x = x_start + i * (card_w + gap)

        # Image placeholder (nếu có)
        if with_image:
            if it.get("img_path"):
                slide.shapes.add_picture(it["img_path"], x, y_top, width=card_w, height=Inches(2.2))
            else:
                # Placeholder shape gradient navy → light blue
                ph = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_top, card_w, Inches(2.2))
                ph.fill.solid()
                ph.fill.fore_color.rgb = NAVY
                ph.line.fill.background()
                # Icon to ở giữa
                icon_box = slide.shapes.add_textbox(x, y_top + Inches(0.7), card_w, Inches(0.8))
                set_text(icon_box.text_frame, it["icon"], size=48, color=WHITE, align=PP_ALIGN.CENTER)
            y_text = y_top + Inches(2.4)
        else:
            y_text = y_top
            # Icon nhỏ phía trên header
            icon_box = slide.shapes.add_textbox(x, y_text, card_w, Inches(0.6))
            set_text(icon_box.text_frame, it["icon"], size=32, align=PP_ALIGN.CENTER)
            y_text += Inches(0.7)

        # Header
        head_box = slide.shapes.add_textbox(x, y_text, card_w, Inches(0.5))
        set_text(head_box.text_frame, it["header"], size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

        # Body
        body_box = slide.shapes.add_textbox(x, y_text + Inches(0.55), card_w, Inches(2.0))
        set_text(body_box.text_frame, it["body"], size=12, color=GRAY_BODY, align=PP_ALIGN.CENTER)
```

---

## Pattern 3 — `half_bleed_image` (Ảnh nửa trái, text nửa phải)

**Khi dùng:** Slide giới thiệu phòng/chương con; 4–6 item body trung bình; muốn nhịp visual mạnh.

**Layout:** Ảnh chiếm 45% bên trái (full height), text + bullet list bên phải.

```python
def build_half_bleed_image(slide, title, intro, bullets, img_path=None, image_side="left"):
    """
    bullets: list of {"header": str, "body": str}
    img_path: nếu None → placeholder gradient
    image_side: "left" hoặc "right"
    """
    img_w = Inches(5.5)
    if image_side == "left":
        img_x = Inches(0)
        text_x = Inches(6.0)
        text_w = Inches(7.0)
    else:
        img_x = SLIDE_W - img_w
        text_x = Inches(0.5)
        text_w = Inches(7.0)

    # Ảnh hoặc placeholder
    if img_path:
        slide.shapes.add_picture(img_path, img_x, Inches(0), width=img_w, height=SLIDE_H)
    else:
        ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, img_x, Inches(0), img_w, SLIDE_H)
        ph.fill.solid()
        ph.fill.fore_color.rgb = NAVY
        ph.line.fill.background()

    # Title (bên text)
    title_box = slide.shapes.add_textbox(text_x, Inches(0.6), text_w, Inches(1.0))
    set_text(title_box.text_frame, title, size=30, bold=True, color=RED)

    # Intro
    if intro:
        intro_box = slide.shapes.add_textbox(text_x, Inches(1.7), text_w, Inches(0.8))
        set_text(intro_box.text_frame, intro, size=13, color=GRAY_BODY)
        y = Inches(2.6)
    else:
        y = Inches(2.0)

    # Bullets
    for b in bullets[:6]:
        # Bullet marker
        bullet_box = slide.shapes.add_textbox(text_x, y, Inches(0.3), Inches(0.4))
        set_text(bullet_box.text_frame, "▸", size=14, bold=True, color=ORANGE)

        # Header + body inline
        text_box = slide.shapes.add_textbox(text_x + Inches(0.35), y, text_w - Inches(0.35), Inches(0.7))
        tf = text_box.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run_h = p.add_run(); run_h.text = b["header"] + " — "
        run_h.font.bold = True; run_h.font.size = Pt(13); run_h.font.color.rgb = NAVY; run_h.font.name = FONT_BODY
        run_b = p.add_run(); run_b.text = b["body"]
        run_b.font.size = Pt(13); run_b.font.color.rgb = GRAY_BODY; run_b.font.name = FONT_BODY

        y += Inches(0.65)
```

**Image source guidance:**
- Nếu user cung cấp folder ảnh → dùng `img_path`
- Nếu không → dùng placeholder gradient navy (mặc định trên)
- KHÔNG dùng AI image generation (skill không có quyền truy cập internet để gen ảnh)
- Nếu user yêu cầu ảnh thật mà không có nguồn → hỏi rõ trước khi render

---

## Pattern 4 — `big_stat` (Số liệu lớn nổi bật)

**Khi dùng:** 1–3 KPI quan trọng cần highlight; slide tổng quan đầu chương.

**Layout:** 1 hoặc 3 cột, mỗi cột = số to (60–96pt) + label nhỏ + mô tả 1 dòng.

```python
def build_big_stat(slide, title, stats):
    """stats: list of 1-3 dicts {"value": "199", "label": "Kênh hoạt động", "desc": "...", "color": NAVY}"""
    add_title(slide, title)

    n = len(stats)
    assert 1 <= n <= 3, "big_stat dùng 1-3 stats"

    col_w = Inches(12.33) / n
    y_num = Inches(2.5)

    for i, s in enumerate(stats):
        x = Inches(0.5) + i * col_w
        color = s.get("color", NAVY)

        # Số lớn
        num_box = slide.shapes.add_textbox(x, y_num, col_w, Inches(2.0))
        size = 96 if n == 1 else 72
        set_text(num_box.text_frame, str(s["value"]), size=size, bold=True, color=color, align=PP_ALIGN.CENTER)

        # Label
        label_box = slide.shapes.add_textbox(x, y_num + Inches(2.1), col_w, Inches(0.5))
        set_text(label_box.text_frame, s["label"], size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

        # Desc
        if s.get("desc"):
            desc_box = slide.shapes.add_textbox(x, y_num + Inches(2.7), col_w, Inches(1.0))
            set_text(desc_box.text_frame, s["desc"], size=12, color=GRAY_BODY, align=PP_ALIGN.CENTER)

        # Divider giữa các cột (nếu n > 1)
        if n > 1 and i < n - 1:
            line_x = Inches(0.5) + (i + 1) * col_w
            line = slide.shapes.add_connector(1, line_x, y_num + Inches(0.5), line_x, y_num + Inches(2.5))
            line.line.color.rgb = GRAY_LINE
            line.line.width = Pt(1)
```

---

## Pattern 5 — `comparison_table` (Bảng so sánh)

**Khi dùng:** So sánh kế hoạch vs thực tế, tuần này vs tuần trước, ≥ 3 cột tiêu chí.

**Layout:** Bảng full-width với header navy, row xen kẽ trắng/xanh nhạt, không viền dày.

```python
def build_comparison_table(slide, title, headers, rows):
    """
    headers: ["Tiêu chí", "Cột 1", "Cột 2", "Cột 3"]
    rows: [["Hoàn thành", "85%", "90%", "92%"], ...]
    """
    add_title(slide, title)

    n_cols = len(headers)
    n_rows = len(rows) + 1  # +1 cho header

    table_x = Inches(0.5)
    table_y = Inches(1.5)
    table_w = Inches(12.33)
    table_h = Inches(0.6) * n_rows

    table_shape = slide.shapes.add_table(n_rows, n_cols, table_x, table_y, table_w, table_h)
    table = table_shape.table

    # Column widths
    table.columns[0].width = Inches(3.0)
    other_w = (table_w - Inches(3.0)) / (n_cols - 1)
    for c in range(1, n_cols):
        table.columns[c].width = other_w

    # Header row
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        run = p.add_run(); run.text = h
        run.font.bold = True; run.font.size = Pt(13); run.font.color.rgb = WHITE; run.font.name = FONT_HEAD

    # Body rows
    for r, row in enumerate(rows, start=1):
        bg = WHITE if r % 2 == 1 else BLUE_LIGHT
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            tf = cell.text_frame
            tf.clear()
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
            run = p.add_run(); run.text = str(val)
            run.font.size = Pt(12); run.font.color.rgb = BLACK; run.font.name = FONT_BODY
            if c == 0:
                run.font.bold = True
                run.font.color.rgb = NAVY
```

---

## Pattern 6 — `chart_with_insights` (Chart trái + insight phải)

**Khi dùng:** Có data số đáng visualize + cần giải thích/đưa ra hành động cụ thể.

**Layout:** Chart 55% bên trái, bullet insights 40% bên phải. Insights nên là **kết luận**, không phải mô tả lại chart.

```python
def build_chart_with_insights(slide, title, chart_data, insights):
    """
    chart_data: {"type": "bar"|"line", "categories": [...], "series": [{"name": str, "values": [...]}]}
    insights: list of {"icon": str, "header": str, "body": str}
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    add_title(slide, title)

    # Chart bên trái
    chart_data_obj = CategoryChartData()
    chart_data_obj.categories = chart_data["categories"]
    for s in chart_data["series"]:
        chart_data_obj.add_series(s["name"], s["values"])

    chart_type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
    }
    ct = chart_type_map.get(chart_data["type"], XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_shape = slide.shapes.add_chart(
        ct, Inches(0.5), Inches(1.5), Inches(7.0), Inches(5.5),
        chart_data_obj
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False

    # Insights bên phải
    y = Inches(1.7)
    for it in insights[:4]:
        # Icon + header
        head_box = slide.shapes.add_textbox(Inches(7.8), y, Inches(5.2), Inches(0.5))
        tf = head_box.text_frame; tf.clear()
        p = tf.paragraphs[0]
        run_i = p.add_run(); run_i.text = it["icon"] + "  "
        run_i.font.size = Pt(16)
        run_h = p.add_run(); run_h.text = it["header"]
        run_h.font.bold = True; run_h.font.size = Pt(15); run_h.font.color.rgb = NAVY; run_h.font.name = FONT_HEAD

        # Body
        body_box = slide.shapes.add_textbox(Inches(8.1), y + Inches(0.5), Inches(4.9), Inches(0.8))
        set_text(body_box.text_frame, it["body"], size=12, color=GRAY_BODY)

        y += Inches(1.3)
```

---

## Decision tree (quick reference từ Bước 3b)

```
Có chart đáng vẽ? ────── yes ──→ chart_with_insights
       │ no
       ▼
So sánh ≥ 3 cột? ───── yes ──→ comparison_table
       │ no
       ▼
≤ 3 KPI nổi bật? ───── yes ──→ big_stat
       │ no
       ▼
Đúng 3 item đối xứng? ─ yes ──→ cards_3col
       │ no
       ▼
4–6 item, body dài, ─── yes ──→ half_bleed_image
muốn nhịp visual?
       │ no
       ▼
                              icon_rows  (default)
```

## Variation rule

- Không lặp cùng pattern quá 2 slide liên tiếp.
- Mỗi báo cáo nên có **ít nhất 2 pattern khác nhau** trong các content slide (chống đơn điệu).
- Cover, TOC, chapter divider, closing **không vary** — luôn dùng template gốc.

## Anti-patterns (tránh)

- ❌ Dùng `big_stat` cho > 3 số → tràn slide
- ❌ Dùng `cards_3col` cho 2 hoặc 4 item → mất cân đối
- ❌ Dùng `half_bleed_image` không có ảnh & không có placeholder → trống nửa slide
- ❌ Dùng `comparison_table` cho 2 cột (chỉ 2 cột thì dùng 2-col cards đẹp hơn)
- ❌ Lặp `icon_rows` cho cả 5 content slide → nhàm chán, đúng vấn đề skill cũ
- ❌ Đè layout custom lên cover/chapter → vi phạm Ràng buộc bất biến

## Image strategy

Skill này **không tự sinh ảnh AI**. Khi dùng pattern có chỗ ảnh (`cards_3col` with_image, `half_bleed_image`), có 3 lựa chọn theo thứ tự ưu tiên:

1. **Ảnh do user cung cấp** — đường dẫn trong `assets/images/` hoặc user upload kèm
2. **Placeholder shape gradient navy** — mặc định nếu không có ảnh (đã code sẵn trong các snippet)
3. **Đề xuất user upload ảnh** — khi user yêu cầu visual cao mà không có nguồn, hỏi rõ trước khi render

Tuyệt đối không dùng ảnh từ internet không rõ nguồn (vi phạm copyright + offline environment).
