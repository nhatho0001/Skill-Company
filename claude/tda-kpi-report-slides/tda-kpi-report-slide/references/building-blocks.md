# Building Blocks — Helper Functions (v2)

Reference này chứa các helper function dùng chung khi build/edit báo cáo. **Khuyến nghị dùng Cách A (edit template)** — chỉ build từ scratch khi yêu cầu khác biệt nhiều so với template gốc.

> Khác biệt v2: Không có hàm `add_logo()` hay `build_cover()` từ scratch nữa, vì template gốc đã embed logo + cover background phẳng. Khi dùng Cách A, chỉ cần thay text trên cover (Quirk 2).

## Setup chung

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from copy import deepcopy
from lxml import etree
from pptx.oxml.ns import qn

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

# Slide dimensions v2
SLIDE_W = Inches(20)
SLIDE_H = Inches(11.25)
```

---

## Helper 1: Set text với formatting

```python
def set_text(tf, text, size=14, bold=False, color=BLACK,
             font=FONT_BODY, align=PP_ALIGN.LEFT):
    """Set text vào textframe. Reset toàn bộ formatting cũ."""
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
```

---

## Helper 2: Debug shape names (CHẠY ĐẦU TIÊN khi gặp template lạ)

```python
def _iter_all_shapes(shapes_collection):
    """Iterate đệ quy qua tất cả shape, kể cả shape trong Group.
    Template Tôn Đông Á v2 có nhiều slide chứa Group of textboxes
    (slide 11 table title, slide 13/14 chart text panels, ...).
    Hàm này được dùng làm helper chung cho tất cả replace/force_color/shrink hàm dưới.
    """
    for shape in shapes_collection:
        yield shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from _iter_all_shapes(shape.shapes)


def debug_slide_shapes(prs, slide_idx=None):
    """In ra shape names + text của 1 slide (hoặc tất cả nếu None).
    LUÔN chạy hàm này TRƯỚC khi sửa template — không skip.
    Recurse vào Group để hiển thị shape con (vì shape title slide 11 nằm trong Group 3).
    """
    indices = [slide_idx] if slide_idx is not None else range(len(prs.slides))
    for i in indices:
        slide = prs.slides[i]
        print(f"\n=== Slide {i+1} ===")
        for shape in _iter_all_shapes(slide.shapes):
            try:
                l = Emu(shape.left).inches
                t = Emu(shape.top).inches
                w = Emu(shape.width).inches
                h = Emu(shape.height).inches
                pos = f"({l:.2f},{t:.2f}) {w:.2f}x{h:.2f}"
            except Exception:
                pos = "(no pos)"
            preview = ""
            if shape.has_text_frame and shape.text_frame.text.strip():
                preview = " | " + shape.text_frame.text.replace("\n", " ⏎ ")[:80]
            elif shape.has_table:
                tbl = shape.table
                preview = f" | [TABLE {len(tbl.rows)}r x {len(tbl.columns)}c]"
            elif shape.has_chart:
                preview = f" | [CHART {shape.chart.chart_type}]"
            print(f"  [{shape.name}] {pos}{preview}")
```

---

## Helper 3: Replace text by shape name (an toàn nhất)

```python
def _iter_all_shapes(shapes_collection):
    """Iterate đệ quy qua tất cả shape, kể cả shape trong Group."""
    for shape in shapes_collection:
        yield shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from _iter_all_shapes(shape.shapes)


def replace_shape_by_name(slide, shape_name, new_text):
    """Thay text của 1 shape có name khớp. Tự động recurse vào Group.
    Trả True nếu thành công.
    """
    for shape in _iter_all_shapes(slide.shapes):
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            p0 = tf.paragraphs[0]
            if p0.runs:
                p0.runs[0].text = new_text
                for r in p0.runs[1:]:
                    r.text = ""
            for p in tf.paragraphs[1:]:
                for r in p.runs:
                    r.text = ""
            return True
    return False
```

---

## Helper 4: Replace text anywhere (khi biết chuỗi cũ duy nhất)

```python
def replace_text_anywhere(slide, old, new):
    """Thay text trong mọi shape của slide chứa chuỗi `old`.
    Tự động recurse vào Group shapes.

    ⚠️ Chỉ match trong 1 paragraph. Nếu `old` trải dài qua nhiều paragraph
    (vd subtitle slide 13/14 chart bị ngắt dòng), dùng `replace_textframe_by_name`.
    """
    found = False
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full = "".join(r.text for r in para.runs)
            if old in full:
                new_full = full.replace(old, new)
                if para.runs:
                    para.runs[0].text = new_full
                    for r in para.runs[1:]:
                        r.text = ""
                found = True
    return found


def replace_textframe_by_name(slide, shape_name, new_text):
    """Replace TOÀN BỘ text frame (gộp nhiều paragraph thành 1).
    Dùng khi text gốc trải qua nhiều paragraph và muốn thay thành 1 đoạn duy nhất.

    Giữ font + format của paragraph + run đầu tiên, xóa các paragraph sau.
    """
    for shape in _iter_all_shapes(slide.shapes):
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            p0 = tf.paragraphs[0]
            # Giữ run 0 của para 0, đổi text
            if p0.runs:
                p0.runs[0].text = new_text
                for r in p0.runs[1:]:
                    r.text = ""
            else:
                # Nếu para 0 không có run → add run mới
                from pptx.util import Pt
                run = p0.add_run()
                run.text = new_text
            # Clear toàn bộ paragraph sau
            for p in tf.paragraphs[1:]:
                for r in p.runs:
                    r.text = ""
            return True
    return False
```

---

## Helper 5: Replace cover period (Quirk 2 — 6 paragraph riêng biệt)

```python
def replace_cover_period(slide_cover, period_current, period_next, department=None):
    """
    Thay 'KẾT QUẢ THÁNG xx/yyyy' và 'VÀ KẾ HOẠCH THÁNG xx/yyyy' trên slide cover.

    Template v2: 6 paragraph riêng biệt trong TextBox 3 (BÁO CÁO / KẾT QUẢ / VÀ KẾ HOẠCH /
    [empty] / PHÒNG…). KHÔNG dùng <a:br/> hack như v1, chỉ cần replace từng paragraph.

    period_current: "10/2025"
    period_next:    "11/2025"
    department:     "PHÒNG CÔNG NGHỆ THÔNG TIN" (nếu None → giữ nguyên)
    """
    for shape in _iter_all_shapes(slide_cover.shapes):
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        if "KẾT QUẢ" not in tf.text and "BÁO CÁO" not in tf.text:
            continue

        for para in tf.paragraphs:
            full = "".join(r.text for r in para.runs)
            if "KẾT QUẢ" in full and para.runs:
                para.runs[0].text = f"KẾT QUẢ THÁNG {period_current}"
                for r in para.runs[1:]:
                    r.text = ""
            elif ("VÀ KẾ HOẠCH" in full or "KẾ HOẠCH" in full) and para.runs:
                para.runs[0].text = f"VÀ KẾ HOẠCH THÁNG {period_next}"
                for r in para.runs[1:]:
                    r.text = ""
            elif department and "PHÒNG" in full and para.runs:
                para.runs[0].text = department
                for r in para.runs[1:]:
                    r.text = ""
        return True
    return False
```

---

## Helper 6: Force header color (đỏ pending / navy done)

```python
def force_header_color(slide, shape_name, is_pending: bool):
    """LUÔN force màu — đỏ nếu pending, navy nếu done.
    QUAN TRỌNG: phải force cả 2 chiều vì template có một số header
    mặc định ĐỎ. Nếu chỉ set màu khi pending, mục done trên slide đó
    vẫn giữ đỏ template → người xem hiểu nhầm.
    """
    target = RED_PRIMARY if is_pending else NAVY
    for shape in _iter_all_shapes(slide.shapes):
        if shape.name != shape_name or not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.text.strip():
                    run.font.color.rgb = target
        return True
    return False
```

---

## Helper 7: Shrink title font nếu quá dài (chống Quirk 3)

```python
def shrink_title_if_long(slide, title_shape_name, max_chars=50, small_size=32):
    """Nếu title > max_chars → giảm font xuống small_size để tránh wrap đè body.
    Áp dụng cho slide 2, 6, 9, 12 nơi title hay quá dài.
    """
    for shape in _iter_all_shapes(slide.shapes):
        if shape.name == title_shape_name and shape.has_text_frame:
            text = shape.text_frame.text
            if len(text) > max_chars:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(small_size)
                return True
    return False
```

---

## Helper 8: Delete / Duplicate slide

```python
def delete_slide(prs, slide_idx):
    """Xóa slide theo index. Sau khi xóa, các slide phía sau lùi index → xóa từ CUỐI lên."""
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    rId = slides_list[slide_idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides_list[slide_idx])


def duplicate_slide(prs, slide_idx):
    """Sao chép 1 slide. Trả về slide mới (append cuối deck)."""
    src_slide = prs.slides[slide_idx]
    blank_layout = src_slide.slide_layout
    new_slide = prs.slides.add_slide(blank_layout)

    # Xóa shapes của blank
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # Copy shapes từ src
    for shape in src_slide.shapes:
        new_el = deepcopy(shape._element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide


def move_slide_to(prs, src_idx, dst_idx):
    """Di chuyển slide src_idx tới vị trí dst_idx (0-indexed)."""
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    src = slides_list[src_idx]
    xml_slides.remove(src)
    xml_slides.insert(dst_idx, src)
```

---

## Helper 9: Replace table với data động (Pattern 4)

```python
def replace_table_in_slide(slide, headers, rows,
                            x=2.25, y=3.38, w=14.42, max_h=7.5):
    """Xóa table cũ + add table mới với data động.
    Xem chi tiết style trong layout-patterns.md → Pattern 4.
    """
    # Xóa table cũ
    for shape in list(slide.shapes):
        if shape.has_table:
            sp = shape._element
            sp.getparent().remove(sp)
            break

    n_rows = len(rows) + 1
    n_cols = len(headers)
    h_calc = 0.7 + 0.5 * len(rows)
    h = min(h_calc, max_h)

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = tbl_shape.table

    # Header
    for c, hd in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_TBL_HDR
        tf = cell.text_frame; tf.clear()
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        run = p.add_run(); run.text = str(hd)
        run.font.bold = True; run.font.size = Pt(15)
        run.font.color.rgb = WHITE; run.font.name = FONT_HEAD
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Body alternating
    for r, row_data in enumerate(rows, start=1):
        bg = GRAY_TBL_EVEN if r % 2 == 1 else GRAY_TBL_ODD
        for c, val in enumerate(row_data):
            cell = tbl.cell(r, c)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg
            tf = cell.text_frame; tf.clear()
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            run = p.add_run(); run.text = str(val)
            run.font.bold = True; run.font.size = Pt(14)
            run.font.color.rgb = BLACK; run.font.name = FONT_BODY
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    return tbl_shape
```

---

## Helper 10: Replace chart data (giữ chart_type cũ của template)

```python
from pptx.chart.data import CategoryChartData

def replace_chart_data(slide, categories, series_data):
    """Replace data của chart hiện có trên slide. Giữ nguyên type (donut/column).
    series_data: dict {series_name: [values]}, vd {"Sales": [25, 30, 22, 28]}.
    """
    for shape in _iter_all_shapes(slide.shapes):
        if shape.has_chart:
            chart_data = CategoryChartData()
            chart_data.categories = categories
            for series_name, values in series_data.items():
                chart_data.add_series(series_name, values)
            shape.chart.replace_data(chart_data)
            return shape.chart
    return None
```

---

## Cách lắp ráp 1 báo cáo hoàn chỉnh (Cách A — Edit Template)

```python
import shutil
from pathlib import Path
from pptx import Presentation

SKILL_DIR = Path("/mnt/skills/user/tda-kpi-report-slides")
TEMPLATE  = SKILL_DIR / "assets/template/report-template.pptx"
OUT_PATH  = Path("/mnt/user-data/outputs/BaoCao_CNTT_T10-2025.pptx")
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

shutil.copy(TEMPLATE, OUT_PATH)
prs = Presentation(str(OUT_PATH))

# 1. Debug TRƯỚC khi sửa
debug_slide_shapes(prs)  # xem mapping shape_name

# 2. Slide 1 — cover
replace_cover_period(prs.slides[0], "10/2025", "11/2025",
                     "PHÒNG CÔNG NGHỆ THÔNG TIN")

# 3. Slide 2 — TOC: replace 5 mục
replace_text_anywhere(prs.slides[1],
    "Báo cáo này tập trung...", "Báo cáo tổng hợp 84/93 công việc...")
# replace_shape_by_name(prs.slides[1], "TextBox 11", "A. Hạ tầng CNTT")
# replace_shape_by_name(prs.slides[1], "TextBox 12", "Kết quả việc Hệ thống Mạng...")
# ...

# 4. Slide 3 — Section A icon_rows
# replace_shape_by_name(prs.slides[2], "TextBox 4", "Hệ thống Mạng & Server")
# replace_shape_by_name(prs.slides[2], "TextBox 5", "Các site hoạt động ổn định...")
# ...

# 5. Slide 11 — Table với data thực
replace_text_anywhere(prs.slides[10], "Table page heading", "BẢNG SO SÁNH KPI")
replace_table_in_slide(prs.slides[10],
    headers=["Nhóm CV", "Tổng", "Đã xong", "% Hoàn thành"],
    rows=[
        ["A. Hạ tầng",  45, 42, "93%"],
        ["B. ERP",      28, 25, "89%"],
        ["C. Eoffice",  18, 17, "94%"],
        ["D. CĐS/AI",   12, 9,  "75%"],
        ["E. Tổng hợp",  6,  6, "100%"],
    ])

# 6. Slide 13/14 — Chart với data thực
replace_chart_data(prs.slides[12],
    categories=["A", "B", "C", "D"],
    series_data={"% Hoàn thành": [42, 25, 17, 9]})

# 7. Xóa slide không dùng (từ cuối lên đầu!)
# delete_slide(prs, 13)  # bỏ chart 2 nếu chỉ cần 1 chart
# delete_slide(prs, 9)   # bỏ timeline placeholder nếu không cần

# 8. Save
prs.save(str(OUT_PATH))
print(f"✓ Saved {OUT_PATH}")
```

---

## Khi nào build từ scratch (Cách B)

Chỉ dùng khi cấu trúc khác biệt mạnh so với template (vd: chỉ làm 3 slide nhanh, hoặc cần custom layout không có trong 7 pattern).

Nếu phải build, **tự re-tạo header trắng + logo + footer** (template gốc embed sẵn nhưng khi build từ scratch mất). Lúc đó cân nhắc copy từ slide layout sẵn có rồi clear shapes thay vì add_slide blank.

```python
# Cách "lai": dùng template làm starting point, nhưng clear shapes của 1 slide
# rồi build mới từ trên đó (giữ được header/logo embedded ở slide master)
prs = Presentation(str(TEMPLATE))
slide = prs.slides[2]  # vd dùng slide 3 làm base
# Xoá hết shapes
for shape in list(slide.shapes):
    sp = shape._element
    sp.getparent().remove(sp)
# Build từ scratch trên slide đã clear
build_icon_rows(slide, "TITLE MỚI", items)
```
