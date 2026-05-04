# Edit Template Workflow v2 (Cách A — Khuyến nghị)

Cách đơn giản nhất để giữ nguyên brand TDA: mở `assets/template/report-template.pptx`, thay text/table/chart, lưu lại. Dùng `python-pptx`.

## ⚠️ Quirks quan trọng của template v2

Template v2 vẫn được xuất từ Gamma.app, kế thừa một số quirk từ v1 + có thêm vài đặc điểm mới. **Đây là 10 điểm cần biết:**

### Quirk 1: Mỗi dòng text là 1 SHAPE riêng

Giống v1: template tách **mỗi dòng text (header, body, số đếm…) thành một shape độc lập** đặt tên `TextBox 3`, `TextBox 4`, `TextBox 5`,... (không dùng `Text 0`, `Text 1` như v1).

**Hệ quả:**
- KHÔNG thể giả định "shape này có cả header + body".
- PHẢI debug shape names trước khi sửa, hoặc match theo `shape.name` cụ thể.

**Cách làm đúng** (luôn chạy đầu tiên):
```python
def debug_slide_shapes(prs, slide_idx):
    slide = prs.slides[slide_idx]
    print(f"\n=== Slide {slide_idx+1} ===")
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            print(f"  [{shape.name}] '{shape.text_frame.text[:90]}'")
        elif shape.has_table:
            tbl = shape.table
            print(f"  [{shape.name}] [TABLE {len(tbl.rows)}r x {len(tbl.columns)}c]")
        elif shape.has_chart:
            print(f"  [{shape.name}] [CHART {shape.chart.chart_type}]")

for i in range(len(prs.slides)):
    debug_slide_shapes(prs, i)
```

### Quirk 2: Slide cover có 6 paragraph riêng biệt — replace từng paragraph

Trên slide bìa (TextBox 3), text được chia thành **6 paragraph riêng biệt**, không phải 1 paragraph với `<a:br/>` như v1.

```
Para 0: '' (empty)
Para 1: 'BÁO CÁO'
Para 2: 'KẾT QUẢ THÁNG 10/2025'
Para 3: 'VÀ KẾ HOẠCH THÁNG 11/2025'
Para 4: '' (empty)
Para 5: 'PHÒNG CÔNG NGHỆ THÔNG TIN'
```

→ **KHÔNG dùng `<a:br/>` hack như v1**. Chỉ cần replace text trong từng paragraph riêng biệt:

**Cách fix:**
```python
def replace_cover_period(slide_cover, period_current, period_next, department=None):
    """Thay kỳ báo cáo + tên phòng trên slide cover.
    Mỗi dòng là 1 paragraph riêng — chỉ cần replace text trong run đầu của mỗi paragraph.
    """
    for shape in slide_cover.shapes:
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

### Quirk 3: Title 2 dòng đè lên content bên dưới (BUG TEMPLATE GỐC)

Đây là vấn đề **mới ở v2**. Khi title của slide section dài quá ~50 ký tự, sẽ wrap xuống dòng 2, **đè vào subtitle/body** đặt ngay dưới (do title textbox được set width hẹp + height cố định).

**Slide có nguy cơ:** 2 (TOC), 4 (B), 6 (D), 7 (E), 9 (Pending), 12 (Others), **11 (Table)**, **13/14 (Chart)**. Quan sát từ render thực tế template gốc:

| Slide | Title gốc | Width title | Có wrap? | Strat |
|---|---|---|---|---|
| 2 | "MỤC LỤC: CÁC HOẠT ĐỘNG TRỌNG TÂM TRONG THÁNG" | 14.89" | YES (2 dòng) — đè intro | Rút title ≤ 38 ký tự |
| 6 | "D. CHUYỂN ĐỔI SỐ (CĐS) & SÁNG KIẾN VỀ AI 💡" | 12.46" | YES — đè body | Rút title ≤ 40 ký tự |
| 7 | "E. TỔNG HỢP CÁC CÔNG VIỆC HOÀN THÀNH" | 16.67" | YES (2 dòng) — không đè | OK |
| 9 | "TỒN ĐỌNG & TRỌNG TÂM THÁNG XX/YYYY" | 15.29" | YES — đè intro | Rút khi tháng 2 chữ số (10, 11, 12) |
| 11 | "Table page heading" | ~6" (trong Group) | YES với title VN dài | **Rút ≤ 25 ký tự** (vd "BẢNG TỔNG HỢP THÁNG 10/2025") |
| 12 | "CÁC HOẠT ĐỘNG KHÁC VÀ DỰ ÁN TRỌNG ĐIỂM THÁNG XX" | 17.83" | YES — đè body | Rút title ≤ 50 ký tự |
| 13 | "Chart page heading" | ~4" (panel trái) | YES với title VN dài | **Rút ≤ 22 ký tự** (vd "PHÂN BỔ KẾT QUẢ T10") |
| 14 | "Chart page heading" | ~4.4" (panel phải) | YES với title VN dài | **Rút ≤ 25 ký tự** |

**Quy tắc bắt buộc khi viết title** thay vào slide section:
- Title 1 dòng: **≤ 50 ký tự** (kể cả emoji + chấm câu)
- Title slide 11/13/14 (panel hẹp): **≤ 22-25 ký tự**
- Title 2 dòng: **≤ 80 ký tự** TỔNG, và content bên dưới phải có top ≥ 2.7" (không đè).
- Nếu title > 50 ký tự, **giảm font xuống 32pt** thay vì 40pt:
  ```python
  for run in shape.text_frame.paragraphs[0].runs:
      run.font.size = Pt(32)
  ```

### Quirk 4: Card header bị wrap đè lên body (TOC + 3-col)

Các card có header textbox với width cố định (vd `2.71"` cho TOC mục C, D). Header > 22 ký tự sẽ wrap xuống dòng 2, đè lên textbox body bên dưới.

**Slide 2 TOC** có 5 card (`TextBox 11, 19, 27, 35, 43`). Nếu tên section > 22 ký tự, header sẽ wrap → đè description.

**Cách fix mỗi khi viết header card:**
- Header card 3-col: **≤ 20 ký tự**
- Header card 4-col: **≤ 16 ký tự**
- Header TOC: **≤ 22 ký tự** (template để width 2.71-4.41" tuỳ mục)
- Header timeline (slide 4 zigzag): **≤ 20 ký tự**

Nếu tên section thực tế dài (vd "C. Eoffice – HCM & Khác" = 22 ký tự — biên), rút bớt: "C. Eoffice & HCM" = 16 ký tự.

### Quirk 5: Slide 4 (timeline) — header text đè body do top fix cứng

Slide 4 dùng layout **timeline 4-step zigzag** (item 1+3 ở trái, 2+4 ở phải). Header và body của mỗi item nằm cách nhau ~0.54", **chỉ đủ cho header 1 dòng**. Nếu header > 25 ký tự (wrap 2 dòng) → đè lên body.

Mapping shape names slide 4:
- Item 1: `TextBox 13` (header) + `TextBox 14` (body)
- Item 2: `TextBox 21` + `TextBox 22`
- Item 3: `TextBox 29` + `TextBox 30`
- Item 4: `TextBox 37` + `TextBox 38`

**Quy tắc:** Header timeline ≤ 20 ký tự (tránh wrap).

### Quirk 6: Slide 6 (D — CĐS & AI) có image lớn bên trái che vùng cột

Slide 6 có `Group 6` ảnh lớn (6"×6") bắt đầu ở `(1.09, 4.05)`. Body cột TRÁI ("Tiếp Cận Công Nghệ AI" — `TextBox 5`) chỉ có không gian rộng 8.65" nhưng cao chỉ 0.41" → text > 90 ký tự sẽ tràn xuống bị ảnh che.

**Quy tắc:** Body slide 6 cột trái: **≤ 90 ký tự**.

Cột phải có 2 sub-card (`TextBox 14+15`, `TextBox 21+22`), không bị giới hạn này.

### Quirk 7: Slide 7 (E — tổng hợp 4 cột) — column textbox width cố định

Slide 7 có 4 column nhỏ (mỗi cột rộng 4.17", body cao 1.58-2.08"). Mỗi cột có 1 dấu "—" (`TextBox 4, 7, 10, 13`) làm divider trên cùng, header (`TextBox 5, 8, 11, 14`) ≤ 3.88" rộng + 0.49" cao, và body (`TextBox 6, 9, 12, 15`).

**Quy tắc:**
- Header 4-cột: **≤ 16 ký tự**
- Body: **≤ 100 ký tự** mỗi cột

### Quirk 8: Slide 11 (TABLE) là placeholder generic — phải build lại từ data

Slide 11 có sẵn 1 `Table 5` (3 cols × 4 rows) với placeholder "Column 1 / content". **Khi user có dữ liệu bảng, phải xóa table cũ và add lại** với số col/row khớp dữ liệu.

**Cách build lại table dynamic** (xem đầy đủ trong `layout-patterns.md` Pattern `data_table`):

```python
def replace_table(slide, headers, rows):
    """Xóa table cũ trên slide, thêm table mới với headers + rows."""
    # 1. Tìm và xóa table cũ
    for shape in list(slide.shapes):
        if shape.has_table:
            sp = shape._element
            sp.getparent().remove(sp)
            break

    # 2. Add table mới (giữ position cũ ~ 2.25, 3.38 và size 14.42 x 4.08)
    from pptx.util import Inches
    n_rows, n_cols = len(rows) + 1, len(headers)
    # Auto-tính height theo số rows: 0.6" header + 0.5" mỗi body row, max 7.5"
    h = min(0.6 + 0.5 * len(rows), 7.5)
    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(2.25), Inches(3.38),
        Inches(14.42), Inches(h)
    )
    # 3. Style + fill cells
    style_table(tbl_shape.table, headers, rows)  # xem layout-patterns.md
```

### Quirk 9: Một số title/subtitle nằm TRONG Group → phải recurse

Trên slide 11 (table), title "Table page heading" và subtitle "Subheading…" nằm **trong `Group 3`** (không phải shape top-level). Tương tự, slide 13/14 chart panel text cũng có thể nằm trong group. Slide 8/12 có 3 image card cũng dùng group.

**Hệ quả:** Nếu chỉ iterate `slide.shapes` thì sẽ KHÔNG thấy shape con, replace text title slide 11/13/14 sẽ fail.

**Cách fix:** Mọi helper iterate shape phải dùng `_iter_all_shapes()` recurse vào group:

```python
def _iter_all_shapes(shapes_collection):
    """Iterate đệ quy qua tất cả shape, kể cả shape trong Group."""
    for shape in shapes_collection:
        yield shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from _iter_all_shapes(shape.shapes)

# Dùng trong replace_shape_by_name, replace_text_anywhere, force_header_color, ...
for shape in _iter_all_shapes(slide.shapes):
    # ... logic match name + replace text
```

Helper `_iter_all_shapes` đã được tích hợp trong tất cả hàm replace của `building-blocks.md` và `scripts/build_example.py`.

### Quirk 10: Subtitle/paragraph chart trải qua NHIỀU paragraph

Slide 13/14 (chart) có subtitle và paragraph dài bị **ngắt thành 2 paragraph riêng** trong cùng 1 textbox:

```
TextBox 8 (subtitle):
  Para 0: 'Subheading that highlights the focus of the charts '
  Para 1: 'and sets context for the data presented. '
TextBox 5 (paragraph):
  Para 0: 'Paragraph that explains the data sources, outlines what the chart illustrates, and provides guidance '
  Para 1: 'on interpreting the visualizations to inform decision-making.'
```

**Hệ quả:** `replace_text_anywhere(slide, "Subheading...presented. ", "X")` sẽ fail vì `old` cross 2 paragraph mà helper chỉ check trong 1 paragraph.

**Cách fix:** Dùng `replace_textframe_by_name(slide, shape_name, new_text)` — clear toàn bộ text frame và viết lại 1 đoạn duy nhất:

```python
replace_textframe_by_name(s13, "TextBox 8",
    "Tỷ trọng % hoàn thành theo nhóm CV chính của Phòng CNTT.")
replace_textframe_by_name(s13, "TextBox 5",
    "Dữ liệu tổng hợp từ tháng 10/2025, phản ánh tỷ trọng hoàn thành...")
```

**Mapping shape names slide 13 (donut, chart bên phải):**
- `Group 6` (icon decoration trái-trên)
- `TextBox 5` — paragraph dưới cùng (3-4 dòng dài)
- `TextBox 7` — title đỏ "Chart page heading"
- `TextBox 8` — subtitle navy "Subheading that highlights..."
- `Chart 11` — donut chart

**Mapping shape names slide 14 (column, chart bên trái) tương tự** nhưng chart panel ở phải.

---

## Cấu trúc 15 slide template gốc

| # | Loại | Layout | Ghi chú |
|---|---|---|---|
| **1** | Cover | Cam phẳng + title + period + dept | Quirk 2 áp dụng |
| **2** | TOC | 5 card đánh số dọc | Quirk 4 (header wrap) |
| **3** | Section A | Icon rows (5 item) + ảnh phải lớn | Layout chính cho section nhiều mục nhỏ |
| **4** | Section B | Timeline 4 bước zigzag | Quirk 5 (header overflow) |
| **5** | Section C | 3 card ngang (cards_3col) | Layout chuẩn 3 cột đối xứng |
| **6** | Section D | 1 ảnh lớn trái + 2 cards phải | Quirk 6 (image overlap) |
| **7** | Section E | 4 cột nhỏ với dash divider | Quirk 7 (header ≤16) |
| **8** | "Page heading" 3 project | 3 card image+title+body | **Placeholder generic** — dùng cho dự án trọng điểm |
| **9** | Pending | 4 ô đánh số có vòng tròn đỏ | Slide tồn đọng & trọng tâm kỳ tới |
| **10** | Timeline page | 4 step có icon + line đứt | **Placeholder generic** — dùng cho lộ trình theo thời gian |
| **11** | Table | Bảng 3 cols × 4 rows generic | **Placeholder — phải build lại** (Quirk 8) |
| **12** | "Other" 3 project | 3 card image+title+body | Tương tự slide 8, dùng cho hoạt động khác |
| **13** | Chart 1 (donut, image trái) | Donut chart bên phải, text trái | Đã có chart sẵn — replace_chart_data() |
| **14** | Chart 2 (column, image phải) | Column chart bên trái, text phải | Tương tự slide 13 |
| **15** | Closing | "Trân trọng kính chào !" giữa slide | Title nhỏ, layout đơn giản |

> **Đừng ngại xóa slide không cần** — gọi `delete_slide()`. Đừng ngại nhân slide cần thiết — gọi `duplicate_slide()`. Số slide cuối cùng phụ thuộc content (xem Bước 3b trong SKILL.md).

---

## Quy trình sửa chuẩn

### Bước 1. Copy + mở template

```python
import shutil
from pptx import Presentation
from pathlib import Path

SKILL_DIR = Path("/mnt/skills/user/tda-kpi-report-slides")
template_src = SKILL_DIR / "assets/template/report-template.pptx"
out_path = Path("/mnt/user-data/outputs/BaoCao_<dept>_<period>.pptx")
out_path.parent.mkdir(parents=True, exist_ok=True)
shutil.copy(template_src, out_path)
prs = Presentation(str(out_path))
```

### Bước 2. **DEBUG shape names TRƯỚC KHI sửa** (bắt buộc)

Chạy `debug_slide_shapes()` trên cả 15 slide → ghi lại mapping `shape.name` → ý nghĩa. Skip bước này = sửa nhầm shape.

### Bước 3. Thay text theo thứ tự đúng

1. Cover (slide 1) — period + department, dùng Quirk 2 fix
2. TOC (slide 2) — intro + 5 card title/desc
3. Sections A/B/C/D/E (slide 3-7) — title + body + header (theo thứ tự body trước header để tránh nhầm match)
4. Slides phụ (8, 10, 11, 12) — chỉ dùng nếu cần
5. Pending (slide 9) — title + intro + 4 item
6. Charts (13, 14) — replace data (xem `building-blocks.md` → Replace Chart Data)
7. Closing (slide 15) — giữ nguyên

### Bước 4. Xóa slide / nhân slide

```python
def delete_slide(prs, slide_idx):
    """Xóa slide theo index. Xóa từ CUỐI lên ĐẦU để tránh lệch index."""
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    rId = slides_list[slide_idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides_list[slide_idx])

def duplicate_slide(prs, slide_idx):
    """Sao chép 1 slide và append vào cuối. Trả về slide mới.
    Hữu ích khi 1 section có > 6 item, cần tách thành 2 slide cùng layout.
    """
    from copy import deepcopy
    from pptx.oxml.ns import qn

    src_slide = prs.slides[slide_idx]
    blank_layout = src_slide.slide_layout
    new_slide = prs.slides.add_slide(blank_layout)

    # Xóa hết shapes của blank
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # Copy shapes từ src
    for shape in src_slide.shapes:
        new_el = deepcopy(shape._element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide
```

**⚠️** Sau khi xóa slide index `k`, các slide phía sau lùi index. **Xóa từ cuối lên đầu** hoặc xử lý content trước rồi xóa.

### Bước 5. Save

```python
prs.save(str(out_path))
```

---

## Helper functions chuẩn (copy-paste được)

```python
from pptx.dml.color import RGBColor
from pptx.util import Pt

RED  = RGBColor(0xFF, 0x00, 0x00)
NAVY = RGBColor(0x00, 0x00, 0x99)


def replace_text_anywhere(slide, old, new):
    """Thay text trong mọi shape. Dùng khi biết rõ chuỗi cũ + duy nhất."""
    for shape in slide.shapes:
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


def replace_shape_by_name(slide, shape_name, new_text):
    """Thay text của shape theo name. An toàn nhất."""
    for shape in slide.shapes:
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


def force_header_color(slide, shape_name, is_pending: bool):
    """LUÔN force màu — đỏ nếu pending, navy nếu done.
    QUAN TRỌNG: phải force cả 2 chiều vì template có một số header
    mặc định ĐỎ. Nếu chỉ set màu khi pending, mục done trên slide đó
    vẫn giữ đỏ template → người xem hiểu nhầm.
    """
    target = RED if is_pending else NAVY
    for shape in slide.shapes:
        if shape.name != shape_name or not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.text.strip():
                    run.font.color.rgb = target
        return True
    return False


def shrink_title_if_long(slide, title_shape_name, max_chars=50, small_size=32):
    """Nếu title > max_chars → giảm font xuống small_size để không wrap đè body."""
    for shape in slide.shapes:
        if shape.name == title_shape_name and shape.has_text_frame:
            text = shape.text_frame.text
            if len(text) > max_chars:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(small_size)


def delete_slide(prs, slide_idx):
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    rId = slides_list[slide_idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides_list[slide_idx])
```

---

## Troubleshooting

| Triệu chứng | Nguyên nhân | Fix |
|---|---|---|
| 2 dòng cover dính ("04/2026VÀ") | Replace nhầm chỗ — text v2 nằm 2 paragraph riêng | Quirk 2 — `replace_cover_period()` mới (no `<a:br/>` hack) |
| Title đè lên intro/body | Title quá dài, wrap 2 dòng | Quirk 3 — rút ≤ 50 ký tự hoặc `shrink_title_if_long()` |
| Header card đè description | Header > 22 ký tự, width cố định | Quirk 4 — rút header card ≤ 20 ký tự |
| Header timeline đè body | Header > 25 ký tự | Quirk 5 — rút ≤ 20 ký tự |
| Body slide 6 bị ảnh che | Body > 90 ký tự, overflow xuống ảnh | Quirk 6 — rút ≤ 90 ký tự |
| Header column slide 7 dính | Header > 16 ký tự, 4 cột chật | Quirk 7 — rút ≤ 16 ký tự |
| Table render generic "Column 1 / content" | Slide 11 chưa replace table | Quirk 8 — `replace_table()` với data thực |
| Body cũ còn vết sau replace | Paragraph 2, 3 chưa clear | Clear tất cả runs sau para đầu |
| Slide thừa sau khi xóa | Index sai sau khi xóa trước đó | Xóa từ cuối lên đầu |
| Title slide đè logo (vd "BẢNG TỔNG HỢP CÔNG VIỆC THÁNG 10/2025") | Title quá dài, wrap khỏi width title | Quirk 3 — `shrink_title_if_long(s, "TextBox name", max_chars=40, small_size=36)` hoặc rút gọn title |
| Title slide 11/13/14 không thay được | Shape nằm trong Group, helper cũ không recurse | Quirk 9 — dùng helper có `_iter_all_shapes()` (đã tích hợp) |
| Subtitle chart vẫn còn placeholder gốc | Text trải qua 2 paragraph, `replace_text_anywhere` không match | Quirk 10 — dùng `replace_textframe_by_name(slide, shape_name, new_text)` |
| Header pending đỏ nhưng item đã done | Template default đỏ, không force navy | `force_header_color(s, name, is_pending=False)` |
