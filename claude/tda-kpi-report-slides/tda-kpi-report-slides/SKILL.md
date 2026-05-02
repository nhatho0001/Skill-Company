---
name: tda-kpi-report-slides
description: "Tạo slide báo cáo định kỳ (tuần/tháng/năm) bằng tiếng Việt từ dữ liệu thô Excel/CSV/JSON/TXT theo template Tôn Đông Á. Hỗ trợ nhiều phòng ban (CNTT, Kế toán, HCM, Kinh doanh, QLK...) — tự đoán phòng từ prefix cột 'Loại công việc' hoặc theo yêu cầu của user. Hãy dùng skill này mỗi khi người dùng nhắc đến 'báo cáo tuần/tháng/năm', 'KPI report', 'slide tổng hợp', 'báo cáo vận hành', 'làm slide từ dữ liệu', 'báo cáo phòng …' (bất kỳ phòng nào), hoặc upload file dữ liệu (xlsx, csv, json, txt) kèm yêu cầu tổng hợp/trình bày/làm slide — kể cả khi người dùng không nói rõ chữ 'skill' hay 'template'. Luôn dùng skill này khi đầu ra mong muốn là file .pptx báo cáo KPI/vận hành bằng tiếng Việt."
---

# TDA KPI Report Slides

Skill tạo slide báo cáo định kỳ (tuần / tháng / năm) cho phòng ban, tập trung vào KPI & vận hành. Đầu ra là file `.pptx` theo đúng bộ nhận diện của Tôn Đông Á (màu cam / xanh navy, logo, font Open Sans).

## Khi nào dùng skill này

Trigger khi người dùng:
- Yêu cầu "làm báo cáo tuần/tháng/năm", "tổng hợp KPI", "slide báo cáo phòng …"
- Upload 1 hoặc nhiều file dữ liệu thô (`.xlsx`, `.csv`, `.json`, `.txt`) và muốn trình bày lại
- Nhắc đến "slide Tôn Đông Á", "template phòng CNTT", "báo cáo vận hành"
- Bất kỳ yêu cầu nào kết thúc bằng một file `.pptx` báo cáo định kỳ bằng tiếng Việt

## Pipeline tổng thể

```
Dữ liệu thô  →  Phân tích  →  Tổng hợp  →  Render slide  →  QA  →  .pptx
  (Excel/CSV/                    (KPI, nổi bật,  (theo template
   JSON/TXT)                      tồn đọng)      Tôn Đông Á)
```

Chạy theo 6 bước dưới đây; đừng nhảy bước.

---

## Bước 1. Thu thập đầu vào & làm rõ kỳ báo cáo

Kiểm tra người dùng đã cung cấp:

1. **File dữ liệu**: Ở `/mnt/user-data/uploads/`. Dùng `ls /mnt/user-data/uploads/` để liệt kê.
2. **Kỳ báo cáo**: Tuần thứ mấy? Tháng nào? Năm nào? Áp dụng thứ tự:
   - Nếu người dùng **đã nói rõ** (VD: "báo cáo tháng 10/2025") → dùng ngay.
   - Nếu không, **tự suy luận** từ cột ngày/tháng trong dữ liệu (max date → xác định tuần/tháng/năm gần nhất).
   - Nếu vẫn không rõ → **hỏi người dùng**, chỉ 1 câu ngắn, kèm giả định mặc định.
3. **Phòng ban**: Skill này hỗ trợ báo cáo cho **nhiều phòng ban khác nhau** (CNTT, Kế toán, HCM, Kinh doanh, v.v.). Quy tắc xác định:

   **Quy tắc bắt buộc:**
   - Nếu yêu cầu của user **đã nói rõ** phòng ("báo cáo P.CNTT", "báo cáo cho P.Kế toán", "báo cáo phòng HCM"…) → dùng ngay, không hỏi.
   - Nếu yêu cầu **KHÔNG nói rõ** phòng → **PHẢI HỎI user** bằng `ask_user_input_v0`. **TUYỆT ĐỐI KHÔNG suy luận từ data** (kể cả khi data chỉ có 1 phòng duy nhất hoặc cột `Phòng ban` có giá trị) — báo cáo cấp phòng/ban là tài liệu chính thức, không được đoán.

   **Cách hỏi (khi user không nói rõ):**
   ```python
   ask_user_input_v0(questions=[{
       "question": "Bạn muốn làm báo cáo cho phòng ban nào?",
       "options": ["P.CNTT", "P.Kế toán", "P.HC-NS", "P.Kinh doanh", "Khác — tôi sẽ nhập"],
       "type": "single_select"
   }])
   ```
   - Nếu user chọn "Khác — tôi sẽ nhập" → đợi user trả lời tên phòng cụ thể trong message tiếp theo, rồi mới tiếp tục.
   - Có thể tuỳ chỉnh danh sách option theo các phòng phổ biến trong tổ chức user, nhưng **luôn có option "Khác"** để mở.

   **Sau khi đã có tên phòng**, filter dữ liệu theo prefix cột `Loại công việc`:
   ```python
   DEPT_PREFIXES = {
       "CNTT": ["CNTT-", "Phần mềm EOffice"],     # P.CNTT
       "KT":   ["KT-", "KE-"],                     # P.Kế toán
       "HCM":  ["HCM-", "HC-"],                    # P.HC-NS
       "KD":   ["KD-"],                            # P.Kinh doanh
       "QLK":  ["QLK-"],                           # P.Quản lý kho
       # ... mở rộng theo nhu cầu
   }

   GENERIC_TYPES = ["Xử lý công việc quy trình", "Dự án", "Báo Cáo tuần/tháng/quý/năm",
                    "Xây dựng quy trình, ISO"]  # áp dụng cho mọi phòng

   def filter_by_dept(df, dept_code):
       prefixes = DEPT_PREFIXES.get(dept_code, [])
       def matches(loai):
           if pd.isna(loai): return False
           if any(loai.startswith(p) for p in prefixes): return True
           if loai in GENERIC_TYPES: return True
           return False
       return df[df["Loại công việc"].apply(matches)].copy()
   ```

   **Tên hiển thị trên slide bìa** (đầy đủ, dùng cho `cover.department`):
   ```python
   DEPT_FULLNAME = {
       "CNTT": "PHÒNG CÔNG NGHỆ THÔNG TIN",
       "KT":   "PHÒNG KẾ TOÁN",
       "HCM":  "PHÒNG HÀNH CHÍNH - NHÂN SỰ",
       "KD":   "PHÒNG KINH DOANH",
       "QLK":  "PHÒNG QUẢN LÝ KHO",
   }
   ```

Ghi lại các thông tin này — sẽ dùng nhiều lần về sau.

---

## Bước 1b. Phân nhóm Loại công việc → Sections (generic per phòng)

Sau khi filter dữ liệu theo phòng, gom `Loại công việc` thành 4-5 **section/nhóm logic** để render thành slide A, B, C, D, E. Mỗi phòng có cách gom khác nhau — KHÔNG hard-code mapping CNTT cho phòng khác.

**Mapping mặc định cho P.CNTT** (đã chuẩn):
```python
CNTT_GROUPS = {
    "A. Hạ tầng & Vận hành": [
        "CNTT-HT-Check in hằng ngày", "CNTT-HT-Thiết bị CNTT", "CNTT-HT-Firewall",
        "CNTT-HT-Hạ tầng mạng CNTT", "CNTT-HT-Fileserver", "CNTT-HT-Inactive tài khoản",
        "CNTT-HT-Hệ thống camera", "CNTT-HT-Bảo trì thiết bị CNTT",
        "CNTT-HT-Việc cần phối hợp PB/NCC", "CNTT-HT-Hệ thống trạm cân",
        "CNTT-HT-Gia hạn dịch vụ CNTT",
    ],
    "B. Hệ thống ERP": [
        "CNTT-ERP-Hỗ Trợ", "CNTT-ERP-Hoạt động", "CNTT-ERP-Database",
        "CNTT-ERP-Change request", "CNTT-ERP-Setup", "CNTT-ERP-Phát triển",
    ],
    "C. Hỗ trợ quy trình & PB": [
        "Xử lý công việc quy trình", "Xây dựng quy trình, ISO", "CNTT-PR mua hàng",
        "CNTT-Hỗ trợ Eoffice/HCM", "Phần mềm EOffice, HCM",
    ],
    "D. Dự án & Nghiên cứu": [
        "Dự án", "CNTT-Nghiên cứu giải pháp",
    ],
    "E. Khác": [
        "Báo Cáo tuần/tháng/quý/năm", "CNTT-Trainning",
    ],
}
```

**Cho phòng khác**: Khi gặp data của phòng mới, **AUTO-GENERATE mapping** bằng cách:
1. Lấy danh sách `Loại công việc` unique của phòng đó
2. Gom theo prefix sau dấu `-` đầu tiên (vd `KT-Sổ cái`, `KT-Báo cáo` → nhóm "Sổ sách & Báo cáo")
3. Gom các generic type (`Xử lý công việc quy trình`, `Dự án`, ...) thành section riêng "Hỗ trợ quy trình"
4. Tối đa **5 section** — nếu nhiều hơn, gộp các nhóm ít item lại thành "Khác"
5. **Hỏi user xác nhận mapping** nếu có ≥ 3 nhóm chưa rõ ngữ cảnh, kèm preview tên section đề xuất.

---

## Bước 2. Đọc & phân tích dữ liệu

Đọc từng file dữ liệu bằng công cụ phù hợp:

| Loại file | Công cụ |
|----|----|
| `.xlsx`, `.xls` | `pandas.read_excel()` (cài: `pip install openpyxl pandas --break-system-packages`) |
| `.csv`, `.tsv` | `pandas.read_csv()` |
| `.json` | `json.load()` + chuyển sang DataFrame nếu tabular |
| `.txt` | Đọc raw, quan sát xem có bảng không |

**Luôn bắt đầu bằng thăm dò:**
```python
import pandas as pd
df = pd.read_excel(path, sheet_name=None)  # tất cả sheet
for name, sheet in df.items():
    print(name, sheet.shape, sheet.columns.tolist())
    print(sheet.head(3))
```

### Bước 2a. **DROP CỘT NHIỄU** (BẮT BUỘC trước khi phân tích)

Trước khi phân tích, **luôn luôn** drop các cột metadata không cần cho báo cáo. Đây là danh sách chuẩn cho file thống kê công việc:

```python
COLUMNS_TO_DROP = [
    "Lý do hủy", "Ngày tạo", "Lý do bỏ hoàn tất", "Ngày hoàn tất",
    "Đơn vị", "Chủ trì", "Phòng ban", "Đơn vị.1", "Liên hệ",
    "Phối hợp", "Để biết", "Hạng mục", "Đvt",
    "Số Ticket liên quan", "Khối lượng", "Người tạo",
    "Người hủy", "Ngày hủy", "Người hoàn tất", "Ngày Ticket liên quan"
]
# Chỉ drop cột thực sự tồn tại để tránh KeyError
existing_drops = [c for c in COLUMNS_TO_DROP if c in df.columns]
df = df.drop(columns=existing_drops)
print(f"Đã drop {len(existing_drops)} cột nhiễu. Cột còn lại: {df.columns.tolist()}")
```

**11 cột cốt lõi còn lại** (mong muốn) sau khi drop:
- **Thời gian:** `Từ ngày`, `Đến ngày` → tính duration & xác định kỳ báo cáo
- **Nội dung:** `Tiêu đề`, `Mô tả`, `Kết quả` → hiểu công việc làm gì
- **Phân loại:** `Dự án`, `Loại công việc` → nhóm theo nhóm
- **Ưu tiên:** `Khẩn cấp`, `Quan trọng` → tính score ưu tiên
- **Trạng thái:** `Đã hoàn tất`, `Đã hủy` → tách done/pending

### Bước 2b. Quy tắc nội dung khi viết slide

**Tuyệt đối KHÔNG** ghi tên người chủ trì, người tạo, người liên quan trong slide output (VD: không viết "A. Ngọ chủ trì", "do anh Quốc thực hiện"). Báo cáo cấp phòng/ban chỉ nói về **công việc, dự án, hệ thống, đơn vị**, không nói về cá nhân.

Nếu báo cáo cần nhắc đến chủ trì, dùng cụm chung như: "đội Hạ tầng", "đội ERP", "P.CNTT", "BP.UDCNTT" — KHÔNG tên riêng.

**KHÔNG ghi thời gian thực hiện** (như "(24-26/04)", "3 ngày (10-13/04)", "24h", "7 ngày", "22 ngày") trong phần mô tả công việc trên slide. Thời gian là metadata nội bộ, không phải nội dung công việc. Slide chỉ nói **CÔNG VIỆC LÀM GÌ + TRẠNG THÁI** — không cần thời gian. Dùng "Đã hoàn tất:", "Đang triển khai:", "Chưa hoàn tất —" để diễn tả trạng thái.

Ví dụ:
- ❌ "3 ngày (10-13/04). Đã hoàn tất: xử lý PR tồn lỗi..."
- ✅ "Đã hoàn tất: xử lý PR tồn lỗi, đảm bảo tính toàn vẹn hệ thống Purchase Request."

- ❌ "Dự án 7 ngày (17-24/04) — chủ trì A. Sang. Đang trong quá trình..."
- ✅ "Dự án hạ tầng đang trong quá trình trình duyệt thiết bị, chưa hoàn tất."

### Bước 2b-bis. **Đánh dấu trạng thái cho công việc chưa hoàn tất** (BẮT BUỘC)

Người xem chỉ cần biết **mục nào CHƯA xong** để chú ý. Mục đã xong giữ trang trọng, không cần đánh dấu.

| Trạng thái | Cách đánh dấu |
|------------|---------------|
| Đã hoàn tất (`Đã hoàn tất == True`) | **Tiêu đề giữ nguyên màu mặc định** (navy `#000099` của template). Không thêm icon, không tô màu. |
| Đang triển khai / chưa hoàn tất (`Đã hoàn tất == False`) | **Tô màu ĐỎ** (`#FF0000` hoặc `RGBColor(0xC0, 0x00, 0x00)`) cho toàn bộ chữ tiêu đề. Không thêm icon. |

**Quy tắc áp dụng:**
- Chỉ tô màu **HEADER**, không tô body
- Nếu 1 item gộp nhiều CV mà có ít nhất 1 CV chưa xong → tô đỏ (ưu tiên cảnh báo)
- Section "Tồn đọng & Trọng tâm kỳ tới" và "Hoạt động khác (kế hoạch)" → tô đỏ toàn bộ header (vì là việc tương lai/chưa làm)
- KHÔNG dùng icon `✅`, `⏳`, `🔴` ở đầu header — chỉ dùng cách tô màu

**⚠️ TUYỆT ĐỐI KHÔNG dùng icon emoji** cho trạng thái — đã thử `⏳` và `🔴` nhưng render không nhất quán hoặc gây rối thị giác. Tô màu chữ là cách sạch nhất.

**Cách implement bằng python-pptx:**
```python
from pptx.dml.color import RGBColor

RED  = RGBColor(0xFF, 0x00, 0x00)  # đỏ cảnh báo (pending)
NAVY = RGBColor(0x00, 0x00, 0x99)  # navy mặc định (done)

def force_header_color(slide, shape_name, is_pending: bool):
    """LUÔN force màu — đỏ nếu pending, navy nếu done.
    QUAN TRỌNG: phải force cả 2 chiều vì template có một số header
    mặc định ĐỎ (slide 6 Text 1, 4 và toàn bộ slide 9). Nếu chỉ set
    màu khi pending, các mục done trên các slide này vẫn giữ đỏ → sai.
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

# Sử dụng: cả mục done và pending đều phải gọi force_header_color()
force_header_color(s3, "Text 1", is_pending=False)  # → navy
force_header_color(s3, "Text 5", is_pending=True)   # → đỏ
force_header_color(s6, "Text 1", is_pending=False)  # FORCE navy (template default đỏ)
force_header_color(s9, "Text 1", is_pending=True)   # đỏ (template default đỏ — trùng)
```

**⚠️ Quirk quan trọng**: Template gốc có một số header mặc định màu đỏ — cụ thể slide 6 (`Text 1`, `Text 4`) và **toàn bộ slide 9** (`Text 1`, `Text 3`, `Text 5`). Nếu mục trên các slide này thực sự ĐÃ HOÀN TẤT, phải chủ động set NAVY (không thể bỏ qua), nếu không sẽ giữ màu đỏ template → người xem hiểu nhầm.

Ví dụ:
- ❌ "🔴 Trang bị thiết bị mạng" (navy) — có icon, không tô màu
- ❌ "✅ Check-in hệ thống hằng ngày" (navy) — có icon thừa
- ✅ "Trang bị thiết bị mạng" (**màu ĐỎ**) — không icon, tô đỏ
- ✅ "Check-in hệ thống hằng ngày" (navy) — không icon, không tô

### Bước 2c. Xác định cột & dữ liệu cần dùng

Sau khi drop, xác định:
- Cột thời gian (để chia theo kỳ)
- Cột phân loại (để nhóm)
- Cột Khẩn cấp / Quan trọng (để tính priority score)
- Dữ liệu thiếu / bất thường (cảnh báo ngay, đừng fill bậy)

### Bước 2d. **Tính priority score** cho mỗi công việc

Đây là trái tim của skill — quyết định CV nào lên slide chính.

```python
df["Duration_hours"] = (df["Đến ngày"] - df["Từ ngày"]).dt.total_seconds() / 3600

def priority_score(row):
    score = 0
    # 1. Quan trọng — ưu tiên cao nhất
    if row.get("Quan trọng") == "Quan trọng":
        score += 100
    # 2. Thuộc dự án/quy trình cụ thể (cột Dự án có giá trị)
    if pd.notna(row.get("Dự án")) and str(row.get("Dự án")).strip():
        score += 50
    # 3. Thời gian thực hiện dài
    dur = row.get("Duration_hours", 0)
    if pd.notna(dur):
        if dur > 24:    score += 30
        elif dur > 8:   score += 15
    # 4. Khẩn cấp
    if row.get("Khẩn cấp") == "Khẩn cấp":
        score += 20
    return score

df["Score"] = df.apply(priority_score, axis=1)
```

**Quy tắc chọn CV cho slide:**
- Mỗi section chỉ chọn **3–5 CV có Score cao nhất** (không liệt kê tất cả)
- CV có Score = 0 (việc thường ngày, ngắn, không quan trọng) chỉ được **đếm** trong tổng số, không lên slide riêng

---

## Bước 3. Tổng hợp thành cấu trúc báo cáo

Từ dữ liệu đã phân tích, tạo ra một **report outline** dạng Python dict/JSON. Đây là bước suy luận quan trọng nhất, không được lười.

Cấu trúc mặc định (phỏng theo template Tôn Đông Á):

```python
report = {
    "cover": {
        "title": "BÁO CÁO",
        "period": "KẾT QUẢ THÁNG 10/2025",     # điền đúng kỳ
        "next_period": "VÀ KẾ HOẠCH THÁNG 11/2025",  # nếu có
        "department": "PHÒNG CÔNG NGHỆ THÔNG TIN",
    },
    "toc": [
        {"letter": "A", "title": "…", "desc": "…"},
        # 3–6 mục
    ],
    "sections": [
        {
            "letter": "A",
            "title": "KẾT QUẢ CÔNG VIỆC …",
            "layout": "icon_rows",   # hoặc "cards_3col", "timeline_4step", "table"
            "items": [
                {"header": "…", "body": "…"},
                # …
            ],
            "chart": None,            # hoặc spec chart (xem Bước 4)
        },
        # …
    ],
    "pending": {                      # tồn đọng / trọng tâm kỳ tới
        "title": "TỒN ĐỌNG & TRỌNG TÂM THÁNG …",
        "items": [{"num": 1, "header": "…", "body": "…"}, …]
    },
    "others": {…},                    # hoạt động khác (optional)
    "closing": {"message": "Trân trọng kính chào !"}
}
```

**Nguyên tắc tổng hợp:**
- Mỗi section 3–6 item, mỗi item header ngắn (≤ 8 từ), body 1–2 câu.
- Ngôn ngữ **tiếng Việt**, trang trọng, dùng danh từ hành động ("Hoàn thành …", "Triển khai …", "Ký duyệt …").
- Số liệu đi kèm ngữ cảnh (VD: "199/205 kênh hoạt động", không phải chỉ "199").
- Không bịa: nếu data không nói, để trống hoặc bỏ mục đó.

---

## Bước 4. Quyết định chart (chỉ khi cần)

**Chỉ thêm chart nếu** dữ liệu có 1 trong các đặc điểm:
- So sánh ≥ 3 mốc thời gian → line chart
- So sánh ≥ 3 hạng mục có giá trị số → bar chart
- Phân bổ tổng thể → pie/donut (tối đa 5 slice)
- Tỷ lệ hoàn thành / KPI đơn lẻ → progress bar / big stat

**Không thêm chart khi:**
- Dữ liệu định tính (trạng thái, mô tả)
- Chỉ 1–2 điểm dữ liệu
- Người dùng đã có bảng rõ ràng

Chart nhúng dùng `pptxgenjs` (xem `references/building-blocks.md` phần Charts).

---

## Bước 5. Render file .pptx

**KHÔNG viết code from-scratch.** Có 2 cách, ưu tiên Cách A:

### Cách A — Edit template (mặc định, giữ nguyên brand)

Dùng khi cấu trúc báo cáo **gần giống template gốc**. Đây là cách chính, nên dùng luôn.

**Luôn luôn bắt đầu bằng script debug shape names** — template được xuất từ Gamma.app nên mỗi dòng text là một shape riêng (`Text 0`, `Text 1`, `Text 2`...), không phải paragraph trong cùng textbox.

```python
for i, slide in enumerate(prs.slides):
    print(f"\n=== Slide {i+1} ===")
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            print(f"  [{shape.name}] '{shape.text_frame.text[:90]}'")
```

**Các quirks quan trọng** (xem chi tiết tại `references/edit-template.md`):
1. Header & body là 2 shape riêng biệt → match theo `shape.name`, không theo paragraph
2. Multi-run paragraph với line break cần thêm `<a:br/>` XML khi thay text
3. Header box width cố định → rút gọn ≤ 20 ký tự (3-cột) hoặc ≤ 16 ký tự (4-cột)
4. Thay **body trước header** khi dùng text-based replace (tránh nhầm chuỗi)
5. Icon `Image 0–4` trên Slide 3 không có embed → xóa & thay bằng bullet "▸"
6. Slide 3 body width bị giới hạn bởi ảnh data center → giữ body ≤ 130 ký tự
7. Slide 3 font size đã được fix sẵn (header 15pt, body 13pt) — không cần chỉnh thêm

Template `scripts/build_example.py` có đầy đủ helper functions copy-paste được.

Nếu báo cáo không cần một section nào đó (VD: không có nội dung CĐS/AI thì xóa slide D), dùng `delete_slide()`. Xóa từ cuối lên đầu để tránh lệch index.

### Cách B — Build từ pptxgenjs theo design system

Chỉ dùng khi cấu trúc **khác biệt nhiều** (VD: người dùng chỉ muốn 3 slide nhanh, hoặc có nhiều chart).

Tuân thủ design tokens trong `references/design-tokens.md`. Phải dùng đúng:
- Màu (cam `#ED7D31`, đỏ `#FF0000`, navy `#000099`)
- Font Open Sans / Open Sans Bold
- Logo (`assets/template/logo-header.jpg` góc phải trên)
- Background cover (`assets/template/cover-background.jpg`) cho slide bìa

Xem các snippet mẫu tại `references/building-blocks.md`.

### Output path
Lưu file cuối cùng vào `/mnt/user-data/outputs/` với tên mô tả đầy đủ, VD:
`BaoCao_CNTT_T10-2025_KeHoach_T11-2025.pptx`

---

## Bước 6. QA (bắt buộc)

Sau khi tạo xong, **luôn** chạy QA:

```bash
# 1. Text check
extract-text /mnt/user-data/outputs/<file>.pptx

# 2. Check placeholder còn sót
extract-text /mnt/user-data/outputs/<file>.pptx | grep -iE "\bx{3,}\b|lorem|\[insert|TODO|<.*>"

# 3. Visual check
cd /home/claude
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf /mnt/user-data/outputs/<file>.pptx
pdftoppm -jpeg -r 120 <file>.pdf slide
ls slide-*.jpg
```

Xem lại từng slide bằng `view` tool. Kiểm tra:
- [ ] Logo Tôn Đông Á ở góc phải trên MỖI slide (trừ slide bìa đã có sẵn trong background)
- [ ] Font Open Sans, size slide title ≥ 28pt, body 12–16pt
- [ ] Không text tràn khỏi box
- [ ] Không chồng chéo (overlap)
- [ ] Màu cam `ED7D31` / đỏ `FF0000` / navy `000099` đúng theo template
- [ ] Tiếng Việt có dấu, không lỗi font (ô vuông, "???")
- [ ] Ngày tháng, kỳ báo cáo đã điền đúng (không còn "tháng 10/2025" nếu user yêu cầu tháng 11)

**Sửa tối đa 1 vòng.** Lỗi nhỏ về pixel thì bỏ qua.

---

## Present file

Cuối cùng, dùng `present_files` với đường dẫn file `.pptx` đã lưu ở `/mnt/user-data/outputs/`.

---

## Reference files

- `references/design-tokens.md` — Bộ màu, font, size, spacing chính thức
- `references/edit-template.md` — **Chi tiết cách sửa template + 7 quirks quan trọng**
- `references/building-blocks.md` — Snippet layout (icon rows, cards 3-col, timeline, table, chart)
- `scripts/build_example.py` — **Script mẫu copy-paste được**, đã áp dụng đúng Bước 2a/2b/2d
- `assets/template/report-template.pptx` — File template gốc
- `assets/template/logo-header.jpg` — Logo Tôn Đông Á
- `assets/template/cover-background.jpg` — Background cam trang bìa

## Phụ thuộc

```bash
pip install python-pptx pandas openpyxl --break-system-packages
# Tùy chọn: npm install -g pptxgenjs  (nếu build từ đầu bằng JS)
```

LibreOffice + pdftoppm đã có sẵn trong môi trường.
