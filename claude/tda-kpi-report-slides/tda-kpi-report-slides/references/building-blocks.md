# Building Blocks — Snippet Layout Mẫu

Các snippet sẵn dùng cho Cách B (build từ đầu bằng `python-pptx`). Mỗi block có code copy-paste được, chỉ cần thay data.

## Setup chung

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Màu chuẩn (từ design-tokens.md)
ORANGE      = RGBColor(0xED, 0x7D, 0x31)
ORANGE_DARK = RGBColor(0xFF, 0x66, 0x00)
RED         = RGBColor(0xFF, 0x00, 0x00)
NAVY        = RGBColor(0x00, 0x00, 0x99)
NAVY_COVER  = RGBColor(0x00, 0x00, 0xCC)
PINK_BORDER = RGBColor(0xE5, 0xB2, 0xB2)
PINK_BG     = RGBColor(0xFF, 0xCC, 0xCC)
PINK_LIGHT  = RGBColor(0xFB, 0xE5, 0xE5)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
BLACK       = RGBColor(0x00, 0x00, 0x00)
GRAY        = RGBColor(0xA5, 0xA5, 0xA5)

FONT_HEADER = "Open Sans"   # sẽ bold
FONT_BODY   = "Open Sans"

# Khởi tạo presentation 16:9
prs = Presentation()
prs.slide_width  = Inches(16)
prs.slide_height = Inches(9)

def blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])  # layout blank

def set_text(tf, text, size=12, bold=False, color=BLACK, font=FONT_BODY, align=PP_ALIGN.LEFT):
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

def add_logo(slide, logo_path):
    """Logo Tôn Đông Á ở góc phải trên, áp dụng mọi slide nội dung."""
    slide.shapes.add_picture(logo_path, Inches(14.8), Inches(0.25),
                              width=Inches(0.9), height=Inches(0.6))
```

---

## Block 1. Slide bìa (Cover)

```python
def build_cover(prs, cover_bg_path, title, period, next_period, department):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Background cam full slide
    slide.shapes.add_picture(cover_bg_path, 0, 0,
                              width=prs.slide_width, height=prs.slide_height)
    # Text overlay (giữa slide, hơi lệch trên)
    tb = slide.shapes.add_textbox(Inches(2.5), Inches(2.8), Inches(11), Inches(3.5))
    tf = tb.text_frame
    tf.clear()              # tránh để lại empty default paragraph
    tf.word_wrap = True

    for i, (txt, sz) in enumerate([
        (title,        36),   # BÁO CÁO
        (period,       24),   # KẾT QUẢ THÁNG 10/2025
        (next_period,  20),   # VÀ KẾ HOẠCH THÁNG 11/2025
        ("",           8),
        (department,   18),   # PHÒNG CÔNG NGHỆ THÔNG TIN
    ]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = txt
        r.font.name = FONT_HEADER
        r.font.size = Pt(sz)
        r.font.bold = True
        r.font.color.rgb = NAVY_COVER
    return slide
```

---

## Block 2. Mục lục (TOC) — 5 card đánh số

```python
def build_toc(prs, logo, section_title, intro, items):
    """items: list[dict(letter='A', title='...', desc='...')]"""
    slide = blank_slide()
    add_logo(slide, logo)

    # Title
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(13), Inches(0.8))
    set_text(t.text_frame, section_title, size=28, bold=True, color=RED, font=FONT_HEADER)

    # Intro paragraph
    intro_tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(14), Inches(0.6))
    set_text(intro_tb.text_frame, intro, size=12, color=BLACK)

    # Card rows (stacked vertically)
    card_h = 1.0
    start_y = 2.1
    gap = 0.2
    for i, it in enumerate(items[:6]):
        y = start_y + i * (card_h + gap)
        # Number circle
        num = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                     Inches(0.6), Inches(y + 0.25),
                                     Inches(0.5), Inches(0.5))
        num.fill.solid(); num.fill.fore_color.rgb = NAVY
        num.line.color.rgb = NAVY
        set_text(num.text_frame, str(i + 1),
                 size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        num.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Card body
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(1.3), Inches(y),
                                      Inches(14), Inches(card_h))
        card.fill.solid(); card.fill.fore_color.rgb = PINK_LIGHT
        card.line.color.rgb = PINK_BORDER
        card.line.width = Pt(1.5)
        tf = card.text_frame
        tf.margin_left = Inches(0.2); tf.margin_top = Inches(0.1)
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = f"{it['letter']}. {it['title']}"
        r1.font.name = FONT_HEADER
        r1.font.size = Pt(14); r1.font.bold = True
        r1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = it['desc']
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = BLACK
    return slide
```

---

## Block 3. Section với icon rows (5–6 item, 1 cột)

Dùng cho Section A kiểu "Hạ tầng CNTT" — nhiều mục nhỏ.

```python
def build_icon_rows(prs, logo, section_title, items):
    """items: list[dict(header='...', body='...')]"""
    slide = blank_slide()
    add_logo(slide, logo)

    # Title
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(14), Inches(0.8))
    set_text(t.text_frame, section_title, size=26, bold=True, color=RED, font=FONT_HEADER)

    # Rows
    row_h = 1.1
    start_y = 1.6
    gap = 0.15
    for i, it in enumerate(items[:6]):
        y = start_y + i * (row_h + gap)
        # Small navy bullet circle
        b = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                   Inches(0.7), Inches(y + 0.35),
                                   Inches(0.35), Inches(0.35))
        b.fill.solid(); b.fill.fore_color.rgb = NAVY
        b.line.fill.background()

        # Text
        tb = slide.shapes.add_textbox(Inches(1.3), Inches(y),
                                      Inches(14), Inches(row_h))
        tf = tb.text_frame; tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = it['header']
        r1.font.name = FONT_HEADER
        r1.font.size = Pt(14); r1.font.bold = True
        r1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = it['body']
        r2.font.name = FONT_BODY; r2.font.size = Pt(11)
        r2.font.color.rgb = BLACK
    return slide
```

---

## Block 4. Timeline 4 bước đánh số (cho Section B kiểu ERP)

```python
def build_timeline_4(prs, logo, section_title, intro, items):
    """items: list[dict(header='...', body='...')], exactly 4 items."""
    slide = blank_slide()
    add_logo(slide, logo)

    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(14), Inches(0.7))
    set_text(t.text_frame, section_title, size=26, bold=True, color=RED, font=FONT_HEADER)

    intro_tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(14), Inches(0.5))
    set_text(intro_tb.text_frame, intro, size=12, color=BLACK)

    # 2x2 grid of numbered items
    col_w = 7.0
    row_h = 2.8
    start_x = 0.8
    start_y = 2.0
    gap_x = 0.4
    gap_y = 0.3

    for i, it in enumerate(items[:4]):
        col = i % 2
        row = i // 2
        x = start_x + col * (col_w + gap_x)
        y = start_y + row * (row_h + gap_y)

        # Number badge (navy circle, white text)
        num = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                     Inches(x), Inches(y + 0.3),
                                     Inches(0.6), Inches(0.6))
        num.fill.solid(); num.fill.fore_color.rgb = NAVY
        num.line.color.rgb = NAVY
        set_text(num.text_frame, str(i + 1),
                 size=18, bold=True, color=WHITE, font=FONT_HEADER,
                 align=PP_ALIGN.CENTER)
        num.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Text block
        tb = slide.shapes.add_textbox(Inches(x + 0.8), Inches(y),
                                      Inches(col_w - 0.8), Inches(row_h))
        tf = tb.text_frame; tf.word_wrap = True

        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = it['header']
        r1.font.name = FONT_HEADER
        r1.font.size = Pt(15); r1.font.bold = True
        r1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = it['body']
        r2.font.name = FONT_BODY; r2.font.size = Pt(11)
        r2.font.color.rgb = BLACK
    return slide
```

---

## Block 5. Cards 3 cột (cho Section C kiểu eOffice)

```python
def build_cards_3col(prs, logo, section_title, cards):
    """cards: list[dict(header='...', body='...')], 3 cards."""
    slide = blank_slide()
    add_logo(slide, logo)

    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(14), Inches(0.8))
    set_text(t.text_frame, section_title, size=26, bold=True, color=RED, font=FONT_HEADER)

    card_w = 4.8
    card_h = 5.5
    start_x = 0.5
    start_y = 2.0
    gap = 0.3

    for i, c in enumerate(cards[:3]):
        x = start_x + i * (card_w + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(start_y),
                                      Inches(card_w), Inches(card_h))
        card.fill.solid(); card.fill.fore_color.rgb = PINK_LIGHT
        card.line.color.rgb = PINK_BORDER
        card.line.width = Pt(1.5)
        tf = card.text_frame
        tf.margin_left = Inches(0.3); tf.margin_right = Inches(0.3)
        tf.margin_top = Inches(0.3)
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = c['header']
        r1.font.name = FONT_HEADER
        r1.font.size = Pt(16); r1.font.bold = True
        r1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        p2.space_before = Pt(10)
        r2 = p2.add_run()
        r2.text = c['body']
        r2.font.name = FONT_BODY; r2.font.size = Pt(12)
        r2.font.color.rgb = BLACK
    return slide
```

---

## Block 6. Slide closing (Trân trọng kính chào)

```python
def build_closing(prs, logo, message="Trân trọng kính chào !"):
    slide = blank_slide()
    add_logo(slide, logo)

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(3.8), Inches(15), Inches(1.2))
    set_text(tb.text_frame, message, size=32, bold=True,
             color=RED, font=FONT_HEADER, align=PP_ALIGN.CENTER)
    return slide
```

---

## Block 7. Chart (chỉ khi thực sự cần)

### Bar chart cơ bản

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

def add_bar_chart(slide, categories, values, series_name="Giá trị",
                  x=1.0, y=2.0, w=10.0, h=5.5, title=None):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x), Inches(y), Inches(w), Inches(h),
        chart_data).chart
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    chart.has_legend = False
    # Tô cam cho các bar
    plot = chart.plots[0]
    for series in plot.series:
        fill = series.format.fill
        fill.solid()
        fill.fore_color.rgb = ORANGE
    return chart
```

### Line chart (so sánh theo thời gian)

```python
def add_line_chart(slide, categories, values, series_name="KPI",
                   x=1.0, y=2.0, w=10.0, h=5.5):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(x), Inches(y), Inches(w), Inches(h),
        chart_data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    return chart
```

---

## Block 8. Table đơn giản

```python
def add_kpi_table(slide, headers, rows, x=1.0, y=2.0, w=14.0, h=4.0):
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Inches(x), Inches(y), Inches(w), Inches(h)
    ).table

    # Header
    for ci, h_text in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        set_text(cell.text_frame, h_text, size=12, bold=True,
                 color=WHITE, font=FONT_HEADER, align=PP_ALIGN.CENTER)

    # Rows
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PINK_LIGHT if ri % 2 == 1 else WHITE
            set_text(cell.text_frame, str(val), size=11, color=BLACK,
                     align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER)
    return table
```

---

## Cách lắp ráp 1 báo cáo hoàn chỉnh

```python
logo = f"{SKILL_DIR}/assets/template/logo-header.jpg"
cover_bg = f"{SKILL_DIR}/assets/template/cover-background.jpg"

build_cover(prs, cover_bg,
            "BÁO CÁO",
            "KẾT QUẢ THÁNG 10/2025",
            "VÀ KẾ HOẠCH THÁNG 11/2025",
            "PHÒNG CÔNG NGHỆ THÔNG TIN")

build_toc(prs, logo,
          "MỤC LỤC: CÁC HOẠT ĐỘNG TRỌNG TÂM TRONG THÁNG",
          "Báo cáo này tập trung vào các kết quả đã đạt được...",
          toc_items)

build_icon_rows(prs, logo,
                "A. KẾT QUẢ CÔNG VIỆC HẠ TẦNG CNTT 🛠️",
                section_a_items)

build_timeline_4(prs, logo,
                 "B. TIẾN ĐỘ CÔNG VIỆC HỆ THỐNG ERP",
                 "Đảm bảo vận hành ổn định...",
                 section_b_items)

# ... các section khác

build_closing(prs, logo)

prs.save("/mnt/user-data/outputs/BaoCao_CNTT_T10-2025.pptx")
```
