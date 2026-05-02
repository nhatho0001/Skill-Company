# Edit Template Workflow (Cách A — Khuyến nghị)

Cách đơn giản nhất để giữ nguyên brand: mở template gốc, thay text, lưu lại. Dùng `python-pptx`.

## ⚠️ Quirks quan trọng của template Tôn Đông Á

Template gốc được xuất từ Gamma.app, có **7 điểm đặc biệt** cần chú ý:

### Quirk 1: Header và body là 2 SHAPE riêng biệt

Không giống PowerPoint truyền thống (1 textbox chứa cả header + body), template này tách **mỗi dòng text thành một shape độc lập**, đặt tên `Text 0`, `Text 1`, `Text 2`,...

**Hệ quả:**
- KHÔNG thể dùng logic "shape này chứa header, paragraph kế tiếp là body"
- PHẢI match theo `shape.name` hoặc quét tuần tự và biết shape nào là header/body

**Cách làm đúng:**
```python
# Debug: in ra tất cả shape + text để biết name tương ứng
for shape in slide.shapes:
    if shape.has_text_frame and shape.text_frame.text.strip():
        print(f"[{shape.name}] '{shape.text_frame.text[:80]}'")
```
→ Từ output, ghi lại: "Text 6 = header 1, Text 7 = body 1, Text 12 = header 2, Text 13 = body 2,..."

### Quirk 2: Multi-run paragraph với line break

Trên slide bìa, "KẾT QUẢ THÁNG 10/2025" và "VÀ KẾ HOẠCH THÁNG 11/2025" nằm trong **CÙNG 1 paragraph** nhưng có `<a:br/>` ngăn cách. Khi thay text vào run đầu, mặc định mất line break → 2 dòng bị dính.

**Cách fix:**
```python
from copy import deepcopy
from lxml import etree
from pptx.oxml.ns import qn

for para in shape.text_frame.paragraphs:
    full = "".join(r.text for r in para.runs)
    if "KẾT QUẢ THÁNG" in full:
        para.runs[0].text = "KẾT QUẢ THÁNG 04/2026"
        for r in para.runs[1:]:
            r.text = ""
        p_xml = para._p
        etree.SubElement(p_xml, qn("a:br"))
        new_r = deepcopy(para.runs[0]._r)
        new_r.find(qn("a:t")).text = "VÀ KẾ HOẠCH THÁNG 05/2026"
        p_xml.append(new_r)
```

### Quirk 3: Textbox width cố định → text dài bị overflow

Các card/header có width cố định (ví dụ `3.10"`). Header text > ~20 ký tự sẽ bị cắt hoặc tràn sang cột kế bên.

**Quy tắc:**
- Header card 3-cột: **≤ 20 ký tự**
- Header card 4-cột: **≤ 16 ký tự**
- Body: có thể dài vì text tự wrap

### Quirk 4: Text-based replace có thể thay NHẦM

Ví dụ: nếu đổi header "Nâng cấp ERP" → "Website TDA.LA" TRƯỚC, rồi gọi `replace_text("Hợp đồng dự án Nâng cấp ERP...", ...)`, sẽ không match vì "Nâng cấp ERP" đã biến mất.

**Cách phòng tránh:**
- Thay body TRƯỚC header, hoặc
- Dùng `replace_shape_by_name()` với shape name cụ thể

### Quirk 5: Icon Image của template có thể bị lỗi reference

Trên Slide 3 (Section A), template có 5 icon mũi tên "→" (`Image 0` đến `Image 4`) đứng trước mỗi mục. Các icon này là **picture shape không có embed thực tế** — sẽ render thành ô vuông trống hoặc nét đứt trong PowerPoint.

**Cách fix (nên làm mặc định):**
```python
# Xoá toàn bộ icon Image 0-4 trên slide 3
s3 = prs.slides[2]
for shape in list(s3.shapes):
    if shape.name in ("Image 0", "Image 1", "Image 2", "Image 3", "Image 4"):
        sp = shape._element
        sp.getparent().remove(sp)

# Thay bằng bullet "▸" prepend vào header (Text 1, 3, 5, 7, 9)
for name in ("Text 1", "Text 3", "Text 5", "Text 7", "Text 9"):
    for shape in s3.shapes:
        if shape.name == name and shape.has_text_frame:
            p0 = shape.text_frame.paragraphs[0]
            if p0.runs and not p0.runs[0].text.startswith("▸"):
                p0.runs[0].text = f"▸ {p0.runs[0].text}"
```

**⚠️ Image 5 KHÁC** — đây là ảnh data center lớn (3.95"), giữ nguyên.

### Quirk 6: Slide 3 — body width bị giới hạn bởi ảnh data center bên phải

Image 5 (ảnh data center) bắt đầu tại `left = 9.67in`, chiếm cả vùng phải của slide từ top ~1.35in. Body text của 5 mục trên slide 3 (`Text 2/4/6/8/10`) đã được set width = 8.7in để không tràn sang ảnh, nhưng **text vẫn có thể bị che nếu nội dung dài quá 1 dòng và wrap không xuống dòng đúng**.

**Quy tắc nội dung cho slide 3:**
- Body mỗi mục: **≤ 130 ký tự** (≈ 2 dòng tại font 13pt)
- Tránh dùng câu dài liệt kê >5 thành phần — tách thành 2 mục hoặc rút gọn

### Quirk 7 (đã fix sẵn trong template): Font size slide 3

Template gốc xuất từ Gamma có body slide 3 = **9.5pt** (quá nhỏ so với các slide khác đều ≥12pt). Đã được fix sẵn trong template:
- Header (`Text 1/3/5/7/9`): **15pt** bold
- Body (`Text 2/4/6/8/10`): **13pt** regular
- Vị trí: re-stack với row height 1.05in, top bắt đầu từ 1.5in
- Icon mũi tên (`Image 0–4`) đã align lại theo top mới của header

→ **Không cần** chỉnh font size slide 3 nữa khi build báo cáo. Nếu thấy slide 3 vẫn nhỏ ≤10pt → template đã bị revert, kiểm tra lại.

---

## Cấu trúc template gốc

| # | Nội dung | Layout |
|---|---|---|
| 1 | **Cover** | Background cam toàn slide |
| 2 | **Mục lục** | 5 card số 1–5 dọc |
| 3 | **Section A** — Hạ tầng CNTT | Icon rows (5 item) + ảnh phải |
| 4 | **Section B** — ERP | Timeline 4 bước đánh số |
| 5 | **Section C** — eOffice / HCM | 3 card ngang |
| 6 | **Section D** — CĐS & AI | 2 cột + ảnh trái (**thường XÓA**) |
| 7 | **Section E** — Tổng hợp | 4 cột nhỏ |
| 8 | **Pending** | 4 card đánh số có circle |
| 9 | **Kiến nghị** | 3 card có ảnh |
| 10 | **Closing** | Tối giản |

---

## Quy trình sửa chuẩn

### Bước 1. Copy + mở template

```python
import shutil
from pptx import Presentation
from pathlib import Path

SKILL_DIR = Path("/home/claude/tda-kpi-report-slides")
template_src = SKILL_DIR / "assets/template/report-template.pptx"
out_path = Path("/mnt/user-data/outputs/BaoCao_<kỳ>.pptx")
shutil.copy(template_src, out_path)
prs = Presentation(str(out_path))
```

### Bước 2. **DEBUG shape names TRƯỚC KHI sửa** (bắt buộc)

```python
for i, slide in enumerate(prs.slides):
    print(f"\n=== Slide {i+1} ===")
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            print(f"  [{shape.name}] '{shape.text_frame.text[:90]}'")
```

### Bước 3. Thay text theo thứ tự đúng

1. Title slide
2. Subtitle/intro paragraph
3. Body trước
4. Header sau (hoặc dùng shape name cho cả header + body)

### Bước 4. Xóa slide thừa

```python
def delete_slide(prs, slide_idx):
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    rId = slides_list[slide_idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides_list[slide_idx])
```

**⚠️** Sau khi xóa slide index `k`, các slide phía sau lùi index. Xử lý các slide TRƯỚC khi xóa, hoặc xóa TỪ CUỐI lên đầu.

### Bước 5. Save

```python
prs.save(str(out_path))
```

---

## Helper functions chuẩn (copy-paste được)

```python
def replace_text_anywhere(slide, old, new):
    """Thay text trong mọi shape. Dùng khi biết rõ chuỗi cũ."""
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
| 2 dòng text dính ("04/2026VÀ") | Multi-run paragraph mất `<a:br/>` | Quirk 2 snippet |
| Text bị cut ("Long An" → "Lon") | Header dài hơn width box | Rút gọn ≤ 20 ký tự |
| 2 header cột liền nhau dính | Text tràn sang cột kế | Giống trên, hoặc giảm font 1-2pt |
| Body không đổi dù đã replace | Text cũ không khớp | Dùng `replace_shape_by_name` |
| Body cũ còn vết sau khi thay | Paragraph 2, 3 chưa clear | Clear tất cả runs sau para đầu |
| Slide thừa sau khi xóa | Index sai sau khi xóa trước đó | Xóa từ cuối lên đầu |
