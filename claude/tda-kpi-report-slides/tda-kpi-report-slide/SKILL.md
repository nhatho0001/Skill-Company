---
name: tda-kpi-report-slides
description: "Tạo slide báo cáo định kỳ (tuần/tháng/năm) bằng tiếng Việt từ dữ liệu thô Excel/CSV/JSON/TXT theo template Tôn Đông Á v2. Hỗ trợ nhiều phòng ban (CNTT, Kế toán, HCM, Kinh doanh, QLK...) — tự đoán phòng từ prefix cột 'Loại công việc' hoặc theo yêu cầu của user. Hãy dùng skill này mỗi khi người dùng nhắc đến 'báo cáo tuần/tháng/năm', 'KPI report', 'slide tổng hợp', 'báo cáo vận hành', 'làm slide từ dữ liệu', 'báo cáo phòng …' (bất kỳ phòng nào), hoặc upload file dữ liệu (xlsx, csv, json, txt) kèm yêu cầu tổng hợp/trình bày/làm slide — kể cả khi người dùng không nói rõ chữ 'skill' hay 'template'. Luôn dùng skill này khi đầu ra mong muốn là file .pptx báo cáo KPI/vận hành bằng tiếng Việt."
---

# TDA KPI Report Slides (v2)

Skill tạo slide báo cáo định kỳ (tuần / tháng / năm) cho phòng ban, tập trung vào KPI & vận hành. Đầu ra là file `.pptx` theo bộ nhận diện Tôn Đông Á v2 (đỏ chính + navy + cam cover, font Inter, slide 20×11.25 inch).

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
   JSON/TXT)                      tồn đọng)      Tôn Đông Á v2)
```

Chạy theo các bước dưới đây; đừng nhảy bước. Bước 3b cho phép số slide động theo content.

---

## Ràng buộc bất biến (KHÔNG được vi phạm)

Đây là **brand identity** Tôn Đông Á v2 — không thay đổi trong bất kỳ tình huống nào:

| Element | Quy tắc |
|---|---|
| **Template gốc** | Chỉ dùng `assets/template/report-template.pptx` (15 slide v2). Không dùng template ngoài. |
| **Logo + footer** | Đã embed sẵn trong template — KHÔNG cần `add_picture(logo)` khi dùng Cách A. KHÔNG xóa các shape header/footer trắng có sẵn ở góc trên/dưới. |
| **Slide size** | 20 × 11.25 inch (16:9 cinematic) — không resize |
| **Font** | `Inter` (header + body). Không đổi sang Arial/Calibri/Roboto |
| **Color cover** | Cam phẳng `#FF6600` + title navy `#000099` + dept navy — đã embed |
| **Color content** | Title đỏ `#FF0000`, header card navy `#000099`, body đen `#212121`. **Không tráo đổi** |
| **Slide 1 (cover)** | Layout cố định: title navy, period navy, department navy. Chỉ thay text qua `replace_cover_period()` (Quirk 2 — 6 paragraph riêng) |
| **Slide 15 (closing)** | "Trân trọng kính chào !" đỏ, layout đơn giản — chỉ thay message nếu user yêu cầu |

**Cho phép vary** (ở Bước 3b, Bước 5):
- Layout của các content slide (icon_rows, cards_3col, numbered_zigzag_4, image_card_3col, data_table, timeline_4_horizontal, four_col_summary, chart_with_text)
- **Số lượng slide** (theo content density, có cap)
- **Số columns / rows trong table** (tùy data, xem Pattern 4)
- Image minh họa trong content slide (dùng ảnh user upload, hoặc placeholder navy, hoặc giữ ảnh template)
- Replace chart data với số liệu thực

Nếu vi phạm Ràng buộc bất biến → fail QA, render lại.

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

Trước khi phân tích, **luôn luôn** drop các cột metadata không cần cho báo cáo:

```python
COLUMNS_TO_DROP = [
    "Lý do hủy", "Ngày tạo", "Lý do bỏ hoàn tất", "Ngày hoàn tất",
    "Đơn vị", "Chủ trì", "Phòng ban", "Đơn vị.1", "Liên hệ",
    "Phối hợp", "Để biết", "Hạng mục", "Đvt",
    "Số Ticket liên quan", "Khối lượng", "Người tạo",
    "Người hủy", "Ngày hủy", "Người hoàn tất", "Ngày Ticket liên quan"
]
existing_drops = [c for c in COLUMNS_TO_DROP if c in df.columns]
df = df.drop(columns=existing_drops)
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
| Đã hoàn tất (`Đã hoàn tất == True`) | **Tiêu đề navy** (`#000099`). Không icon, không tô màu khác. |
| Đang triển khai / chưa hoàn tất (`Đã hoàn tất == False`) | **Tô màu ĐỎ** (`#FF0000` / `RGBColor(0xFF, 0x00, 0x00)`) cho toàn bộ chữ tiêu đề. Không icon. |

**Quy tắc áp dụng:**
- Chỉ tô màu **HEADER**, không tô body
- Nếu 1 item gộp nhiều CV mà có ít nhất 1 CV chưa xong → tô đỏ (ưu tiên cảnh báo)
- Slide 9 "Tồn đọng & Trọng tâm kỳ tới" và slide 12 "Hoạt động khác (kế hoạch)" → tô đỏ toàn bộ header (vì là việc tương lai/chưa làm)
- KHÔNG dùng icon `✅`, `⏳`, `🔴` ở đầu header — chỉ dùng cách tô màu

**⚠️ TUYỆT ĐỐI KHÔNG dùng icon emoji** cho trạng thái — render không nhất quán hoặc gây rối thị giác. Tô màu chữ là cách sạch nhất.

**Cách implement bằng python-pptx:**
```python
from pptx.dml.color import RGBColor

RED  = RGBColor(0xFF, 0x00, 0x00)  # đỏ pending
NAVY = RGBColor(0x00, 0x00, 0x99)  # navy done

def force_header_color(slide, shape_name, is_pending: bool):
    """LUÔN force màu — đỏ nếu pending, navy nếu done.
    Phải force cả 2 chiều vì template có một số header mặc định ĐỎ.
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
```

**⚠️ Quirk quan trọng**: Template gốc v2 có một số header mặc định màu đỏ (slide 9 toàn bộ pending, slide 12 placeholder). Nếu mục trên các slide này thực sự ĐÃ HOÀN TẤT, phải chủ động set NAVY (không thể bỏ qua), nếu không sẽ giữ màu đỏ template → người xem hiểu nhầm.

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
- Mỗi section chọn **3–5 CV có Score cao nhất** (không liệt kê tất cả)
- CV có Score = 0 (việc thường ngày, ngắn, không quan trọng) chỉ được **đếm** trong tổng số, không lên slide riêng

---

## Bước 3. Tổng hợp thành cấu trúc báo cáo

Từ dữ liệu đã phân tích, tạo ra một **report outline** dạng Python dict/JSON. Đây là bước suy luận quan trọng nhất, không được lười.

Cấu trúc mặc định (phỏng theo template Tôn Đông Á v2):

```python
report = {
    "cover": {
        "title": "BÁO CÁO",
        "period_current": "10/2025",     # tháng hiện tại
        "period_next":    "11/2025",     # tháng kế tiếp
        "department":     "PHÒNG CÔNG NGHỆ THÔNG TIN",
    },
    "toc": [
        {"letter": "A", "title": "…", "desc": "…"},
        # 3–5 mục
    ],
    "sections": [
        {
            "letter": "A",
            "title": "KẾT QUẢ CÔNG VIỆC …",
            "layout": "icon_rows",   # hoặc "cards_3col", "numbered_zigzag_4", "data_table", ...
            "items": [
                {"header": "…", "body": "…", "is_pending": False},
                # …
            ],
            "chart": None,  # hoặc spec chart (xem Bước 4)
            "table": None,  # hoặc {"headers": [...], "rows": [[...]]} (xem Pattern 4)
        },
        # …
    ],
    "pending": {                      # tồn đọng / trọng tâm kỳ tới
        "title": "TỒN ĐỌNG & TRỌNG TÂM THÁNG …",
        "items": [{"header": "…", "body": "…"}, …]
    },
    "others": {…},                    # hoạt động khác (optional)
    "closing": {"message": "Trân trọng kính chào !"}
}
```

**Nguyên tắc tổng hợp:**
- Mỗi section 3–6 item, mỗi item header **ngắn** (≤ 20 ký tự cho 3-col / ≤ 16 cho 4-col), body 1–2 câu (≤ 130 ký tự cho slide có ảnh, ≤ 200 cho slide thường).
- Ngôn ngữ **tiếng Việt**, trang trọng, dùng danh từ hành động ("Hoàn thành …", "Triển khai …", "Ký duyệt …").
- Số liệu đi kèm ngữ cảnh (VD: "199/205 kênh hoạt động", không phải chỉ "199").
- Không bịa: nếu data không nói, để trống hoặc bỏ mục đó.

---

## Bước 3b. Quyết định số slide & layout (động theo dữ liệu, có cap)

**Đừng cố định số slide bằng số chương.** Số slide phụ thuộc **content density**: mỗi section có thể bung thành 1, 2, hoặc 3 slide tùy lượng item & độ dài text.

### B3b.1 — Rule tách slide theo content density

Với mỗi section trong `report["sections"]`, đếm số item và chiều dài body để quyết định:

| Số item trong section | Body trung bình | Số slide | Layout gợi ý |
|---|---|---|---|
| 1–3 item | ≤ 1 dòng | 1 slide | `icon_rows` hoặc `cards_3col` (nếu đủ 3) |
| 4–6 item | ≤ 1 dòng | 1 slide | `icon_rows` |
| 4–6 item | 2–3 dòng | 1 slide | `icon_rows` (slide A có ảnh phải) hoặc `cards_3col` |
| 7–10 item | bất kỳ | 2 slide | Chia theme: kết quả nổi bật / kết quả thường |
| > 10 item | bất kỳ | 2–3 slide | Chia theo priority hoặc theme con |
| Có data số đáng visualize | — | +1 slide | `chart_with_text` (slide 13/14 template) |
| So sánh ≥ 3 cột metric | — | +1 slide | `data_table` (slide 11 template) ⭐ |
| 4 mốc thời gian (năm/quý) | — | +1 slide | `timeline_4_horizontal` (slide 10 template) |
| 3 dự án có ảnh minh hoạ | — | +1 slide | `image_card_3col` (slide 8/12 template) |

### B3b.2 — Slide bắt buộc & slide tùy chọn

**Bắt buộc** (luôn có):
- 1 slide cover (slide 1)
- 1 slide TOC (slide 2)
- 1 slide content cho mỗi section chính (slide 3-7 tùy số section)
- 1 slide pending (slide 9)
- 1 slide closing (slide 15)

**Tùy chọn** (thêm khi data đủ để có ý nghĩa):
- Slide 8 / 12 (`image_card_3col`) — khi có 3 dự án trọng điểm có ảnh
- Slide 10 (`timeline_4_horizontal`) — khi có lộ trình theo năm/quý
- Slide 11 (`data_table`) — khi có so sánh ≥ 3 cột metric hoặc user yêu cầu rõ "làm bảng"
- Slide 13 / 14 (`chart_with_text`) — khi có data số đáng visualize

### B3b.3 — Cap & quy trình khi vượt cap

- **Sàn:** ≥ 5 slide (cover + TOC + 1 content + pending + closing)
- **Cap mềm:** ≤ 15 slide
- **Tỷ lệ vàng:** 60-70% là content slide chính, 20% slide phụ (table, chart, timeline, image), 10-20% slide đặc biệt

**Quy trình tách slide theo cap:**

1. Đếm tổng số item trong tất cả sections + pending + others.
2. **Tách tự động**:
   - Nếu 1 section có > 6 item → tách section đó thành 2 slide cùng layout (theo priority score: 5 item top + N item còn lại).
   - Nếu vẫn ≤ 15 slide → OK, render.
3. **Khi vượt cap 15 slide**: HỎI user trước khi tiếp tục:
   ```python
   ask_user_input_v0(questions=[{
       "question": f"Báo cáo có {N_total} mục, dự kiến cần {N_slides} slide (>15). Bạn muốn:",
       "options": [
           "Giữ nguyên — tạo đầy đủ slide (có thể dài)",
           "Rút gọn xuống 15 slide — chỉ giữ top priority mỗi section",
           "Tách 2 báo cáo: phần 1 (≤15 slide) + phần 2 (phụ lục)"
       ],
       "type": "single_select"
   }])
   ```
   - **Không tự rút gọn** khi user chưa quyết — báo cáo cấp phòng cần user kiểm soát nội dung được giữ lại.

### B3b.4 — Gán layout cho từng section

Sau khi quyết định số slide, mở rộng `report["sections"]` thành `report["slides"]` với layout cụ thể. Ví dụ:

```python
# Trước (output Bước 3)
{"letter": "A", "title": "KẾT QUẢ CNTT", "items": [...8 items...]}

# Sau Bước 3b (cho content density cao)
slides = [
    {"type": "content", "layout": "icon_rows",  "title": "A. Hệ thống & Hạ tầng (1/2)",  "items": items_top5},
    {"type": "content", "layout": "icon_rows",  "title": "A. Hệ thống & Hạ tầng (2/2)",  "items": items_remaining},
]
```

Pattern library đầy đủ + code snippet xem **`references/layout-patterns.md`**.

### B3b.5 — Decision tree chọn layout

```
Section có data số đáng visualize? ──── yes ──→ chart_with_text
       │ no
       ▼
So sánh ≥ 3 cột metric, hoặc user yêu cầu "bảng"? ─ yes ──→ data_table ⭐
       │ no
       ▼
4 mốc thời gian (năm, quý)? ────────── yes ──→ timeline_4_horizontal
       │ no
       ▼
4 bước/giai đoạn cùng kỳ? ──────────── yes ──→ numbered_zigzag_4
       │ no
       ▼
3 dự án có ảnh minh hoạ? ───────────── yes ──→ image_card_3col
       │ no
       ▼
Đúng 3 item đối xứng (≤ 80 từ/item)? ─ yes ──→ cards_3col
       │ no
       ▼
Đúng 4 thành tựu tóm tắt? ──────────── yes ──→ four_col_summary
       │ no
       ▼
                                              icon_rows (default)
```

### B3b.6 — Variation rule (chống đơn điệu)

**Không lặp cùng 1 layout quá 2 slide liên tiếp.** Nếu Bước 3b.4 ra 3 section liên tiếp đều `icon_rows`, đổi 1 trong 3 thành `cards_3col` hoặc `image_card_3col` để báo cáo có nhịp.

Cover, TOC, closing → **giữ template gốc**, không vary (Ràng buộc bất biến).

---

## Bước 4. Quyết định chart (chỉ khi cần)

**Chỉ thêm chart nếu** dữ liệu có 1 trong các đặc điểm:
- So sánh ≥ 3 mốc thời gian → line / column chart (slide 14 template)
- Phân bổ tổng thể → pie/donut, tối đa 5 slice (slide 13 template)
- Tỷ lệ hoàn thành / KPI đơn lẻ → progress bar / big stat
- Tỷ trọng giữa các nhóm CV → donut chart (slide 13)

**Không thêm chart khi:**
- Dữ liệu định tính (trạng thái, mô tả)
- Chỉ 1–2 điểm dữ liệu
- Người dùng đã có bảng rõ ràng

Template v2 có sẵn slide 13 (donut) và slide 14 (column-clustered). Dùng `replace_chart_data()` (helper trong `building-blocks.md`) để thay data, **không cần build chart từ scratch**.

---

## Bước 4b. Quyết định table (Pattern 4)

**Khi nào dùng `data_table`** — slide 11 template:
1. **So sánh ≥ 3 cột metric** (vd: "Hạng mục / Tổng / Đã xong / % hoàn thành") — đây là use case chuẩn.
2. **User yêu cầu rõ ràng** ("làm bảng", "so sánh dạng bảng", "table KPI", v.v.) — dùng ngay, không hỏi lại.

**Số columns / rows quyết định bởi DATA**, không cố định 3×4 như placeholder. Tham khảo Pattern 4 trong `layout-patterns.md`:

```python
replace_table_in_slide(slide11,
    headers=["Nhóm CV", "Tổng", "Đã xong", "% Hoàn thành"],
    rows=[
        ["A. Hạ tầng",  45, 42, "93%"],
        ["B. ERP",      28, 25, "89%"],
        # ... rows tùy data
    ])
```

| Số cột | Quy tắc |
|---|---|
| 2 cột | KHÔNG dùng table — dùng `cards_2col` |
| 3-5 cột | Tốt nhất, slide 20" rộng đủ |
| 6+ cột | Cẩn thận — text mỗi cell phải ngắn, hoặc tách 2 bảng |

| Số rows | Quy tắc |
|---|---|
| 3-5 | Tốt nhất |
| 6-10 | OK, height auto-tính |
| 11+ | Tách 2 bảng / 2 slide |

---

## Bước 5. Render file .pptx

**KHÔNG viết code from-scratch.** Có 2 cách, ưu tiên Cách A:

### Cách A — Edit template (mặc định, giữ nguyên brand)

Dùng khi cấu trúc báo cáo **gần giống template gốc**. Đây là cách chính, nên dùng luôn.

**Luôn luôn bắt đầu bằng `debug_slide_shapes()`** — template được xuất từ Gamma.app nên mỗi dòng text là một shape riêng (`TextBox 3`, `TextBox 4`, …), không phải paragraph trong cùng textbox.

```python
def debug_slide_shapes(prs, slide_idx=None):
    indices = [slide_idx] if slide_idx is not None else range(len(prs.slides))
    for i in indices:
        slide = prs.slides[i]
        print(f"\n=== Slide {i+1} ===")
        for shape in slide.shapes:
            preview = ""
            if shape.has_text_frame and shape.text_frame.text.strip():
                preview = " | " + shape.text_frame.text[:80]
            elif shape.has_table:
                preview = f" | [TABLE {len(shape.table.rows)}r x {len(shape.table.columns)}c]"
            elif shape.has_chart:
                preview = " | [CHART]"
            print(f"  [{shape.name}]{preview}")

debug_slide_shapes(prs)  # CHẠY TRƯỚC khi sửa
```

**Các quirks quan trọng** (xem chi tiết tại `references/edit-template.md`):
1. Mỗi dòng text là 1 shape riêng → match theo `shape.name`
2. Cover (slide 1): mỗi dòng là 1 paragraph riêng — dùng `replace_cover_period()` để replace từng paragraph (không hack `<a:br/>` như v1)
3. **Title đa dòng có thể đè body bên dưới** — dùng `shrink_title_if_long()` cho slide 2, 6, 9, 12
4. Header card width cố định → ≤ 20 ký tự (3-cột) / ≤ 16 ký tự (4-cột)
5. Slide 4 timeline header ≤ 20 ký tự (tránh wrap)
6. Slide 6 body cột trái ≤ 90 ký tự (image che)
7. Slide 7 header column ≤ 16 ký tự
8. **Slide 11 (Table) là placeholder** — phải dùng `replace_table_in_slide()` với data thực
9. **Một số title nằm trong Group** (slide 11/13/14) — mọi helper phải recurse vào Group bằng `_iter_all_shapes()`

Template `scripts/build_example.py` có đầy đủ helper functions copy-paste được, đã áp dụng các quirk trên.

**Xóa slide không dùng** (vd: nếu không có nội dung CĐS/AI → xóa slide D, không có chart → xóa slide 13/14):
```python
delete_slide(prs, slide_idx)  # xóa từ CUỐI lên ĐẦU để tránh lệch index
```

### Cách B — Build từ pptxgenjs hoặc python-pptx từ đầu

Chỉ dùng khi cấu trúc **khác biệt nhiều** (VD: người dùng chỉ muốn 3 slide nhanh, hoặc layout custom không có trong template).

Tuân thủ design tokens trong `references/design-tokens.md`. Phải dùng đúng:
- Slide size 20×11.25 inch
- Màu (đỏ `#FF0000` primary, navy `#000099`, cam cover `#FF6600`)
- Font Inter
- KHÔNG cần file logo ngoài (template gốc embed sẵn — nếu build từ scratch thì cần re-tạo header/footer trắng, nhưng tốt hơn là dùng "Cách lai" — copy 1 slide template, clear shapes, rồi build trên đó để giữ logo embedded ở slide layout).

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
extract-text /mnt/user-data/outputs/<file>.pptx | grep -iE "lorem|column 1|content|page heading|TODO|\bx{3,}\b|\[insert"

# 3. Visual check
cd /home/claude
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf /mnt/user-data/outputs/<file>.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 100 <file>.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

Xem lại từng slide bằng `view` tool. Kiểm tra:
- [ ] Logo Tôn Đông Á ở góc phải trên MỖI slide (đã embed sẵn — chỉ verify còn nguyên)
- [ ] Footer URL `tondonga.com.vn` chỉ ở slide 1 cover
- [ ] Font Inter (hoặc fallback Open Sans), title ≥ 36pt, body 14–16pt
- [ ] Không text tràn khỏi box, không title 2 dòng đè body (quirk 3, 4, 5)
- [ ] Không chồng chéo (overlap) header card với description
- [ ] Màu đỏ `FF0000` cho title section, navy `000099` cho header card, **đỏ ≠ navy** đúng theo trạng thái pending/done
- [ ] Tiếng Việt có dấu, không lỗi font (ô vuông, "???")
- [ ] Ngày tháng, kỳ báo cáo đã điền đúng (không còn "tháng 10/2025" nếu user yêu cầu tháng 11)
- [ ] Slide table (nếu có): không còn placeholder "Column 1 / content / page heading"
- [ ] Slide chart (nếu có): data đã thay, không còn "Sales / 1st Qtr / 2nd Qtr"

**Sửa tối đa 1 vòng.** Lỗi nhỏ về pixel thì bỏ qua.

---

## Present file

Cuối cùng, dùng `present_files` với đường dẫn file `.pptx` đã lưu ở `/mnt/user-data/outputs/`.

---

## Reference files

- `references/design-tokens.md` — Bộ màu, font, size, spacing chính thức v2
- `references/edit-template.md` — **Chi tiết cách sửa template + 8 quirks quan trọng + 15-slide map**
- `references/layout-patterns.md` — **Pattern library 7 layout content slide** (Bước 3b), bao gồm `data_table` mới
- `references/building-blocks.md` — Helper functions chuẩn (replace_cover_period, replace_table_in_slide, replace_chart_data, force_header_color, ...)
- `scripts/build_example.py` — **Script mẫu copy-paste được**, đã áp dụng đúng Bước 2a/2b/2d/3b
- `assets/template/report-template.pptx` — File template gốc v2 (15 slide, 20×11.25in, đã embed logo + footer)

## Phụ thuộc

```bash
pip install python-pptx pandas openpyxl --break-system-packages
```

LibreOffice + pdftoppm đã có sẵn trong môi trường.

---

## Changelog

### v2 (May 2026)
- Slide size đổi từ 13.33×7.5 → **20×11.25 inch**
- Cover background: ảnh phức tạp → **cam phẳng `#FF6600` (embed)**
- Logo + footer: file ngoài → **embed sẵn trong template**
- Title content: navy → **đỏ `#FF0000` primary**
- Số master slides: 10 → **15** (thêm timeline, **table dynamic**, 2 chart, 3-image cards)
- Font: Open Sans → **Inter**
- **Pattern mới `data_table`** — số col/row tùy data, không cố định 3×4
- **Cap 15 slide + ask user khi vượt** thay vì tự cắt
- Xóa file asset `cover-background.jpg` + `logo-header.jpg` (không cần nữa)
