# TDA KPI Report Slides — Hướng dẫn triển khai

Skill **`tda-kpi-report-slides`** (v4.1) dành cho Claude AI trong Cowork mode, giúp tự động tạo slide báo cáo định kỳ (tuần / tháng / năm) bằng tiếng Việt theo đúng bộ nhận diện thương hiệu Tôn Đông Á.

---

## 1. Skill này dùng để làm gì?

- Đọc dữ liệu công việc từ file qua **MCP `file-process`** → phân tích → xuất file PowerPoint (`.pptx`) báo cáo định kỳ.
- Hỗ trợ nhiều phòng ban: **CNTT, Kế toán, HC-NS, Kinh doanh, Quản lý kho, …**
- Đảm bảo 100% slide xuất ra đúng màu sắc, logo, font chữ chuẩn của Tôn Đông Á (đỏ `#FF0000` · navy `#000099` · cam bìa `#FF6600` · font Segoe UI).
- Rút ngắn thời gian làm báo cáo định kỳ từ **vài giờ xuống còn vài phút**.

---

## 2. Yêu cầu trước khi cài đặt

| Yêu cầu | Ghi chú |
|---|---|
| Ứng dụng Claude Desktop | Phiên bản Cowork mode (có hỗ trợ Skills & MCP) |
| MCP `file-process` đã kết nối | Skill đọc dữ liệu qua `file-process:read_handle_document_report` — bắt buộc |
| File skill (`.skill`) | `tda-kpi-report-slides.skill` (file zip đổi đuôi) |
| File dữ liệu nguồn | Excel/CSV với 13 cột chuẩn (xem mục 6) — đặt ở vị trí MCP có thể đọc được |

---

## 3. Cấu trúc thư mục skill

```
tda-kpi-report-slides/
├── SKILL.md                         # Hướng dẫn pipeline 6 bước cho Claude
├── README.md                        # File này
├── assets/
│   └── brand/
│       ├── bg-cover.jpg             # Nền bìa (logo + footer bake sẵn)
│       ├── bg-content.jpg           # Nền slide nội dung (logo bake sẵn)
│       ├── logo.png                 # Logo rời (tùy chọn)
│       └── icons/                   # 22 icon trắng (tùy chọn)
├── references/
│   └── brand-constraints.md        # Hợp đồng brand: màu, font, vùng an toàn
└── scripts/
    ├── helpers.py                   # Logic phân tích dữ liệu + validators
    └── brand_kit.js                 # Brand kit pptxgenjs (primitive ép brand)
```

---

## 4. Hướng dẫn cài đặt

### Bước 1. Tải file skill về máy

Truy cập repo và tải file `tda-kpi-report-slides.skill` về máy bằng cách bấm vào nút **Download** (biểu tượng mũi tên xuống) ở góc phải:

![Tải file skill từ GitHub](docs/images/01-github-download.png)

> **Lưu ý:** Giữ nguyên định dạng `.skill`, **không giải nén**.

---

### Bước 2. Mở mục Customize trong Claude

Đăng nhập vào ứng dụng Claude Desktop, sau đó bấm vào biểu tượng **vali (Customize)** ở thanh sidebar bên trái:

![Vào Customize trong Claude](docs/images/02-claude-customize-icon.png)

---

### Bước 3. Chọn "Create new skills"

Trong trang Customize Claude, chọn mục **Create new skills — Teach Claude your processes, team norms, and expertise**:

![Chọn Create new skills](docs/images/03-create-new-skills.png)

---

### Bước 4. Upload file skill

Trong trang Skills, bấm dấu **`+`** ở góc trên → chọn **`Create skill`** → chọn **`Upload a skill`**, sau đó chọn file `tda-kpi-report-slides.skill`:

![Upload skill](docs/images/04-upload-skill.png)

Sau khi upload thành công, skill sẽ hiện trong danh sách **Personal skills**.

> **Lưu ý:** Nhớ **bật toggle** kích hoạt skill để Claude có thể sử dụng được.

---

## 5. Cách sử dụng skill

Skill hoạt động theo **pipeline 6 bước tự động**. Người dùng chỉ cần gõ yêu cầu — Claude sẽ tự đọc dữ liệu qua MCP, phân tích và xuất slide.

### Yêu cầu ví dụ

> *"Làm báo cáo tuần 2/6–8/6 cho phòng CNTT."*

> *"Giúp tôi làm slide báo cáo tháng 5 cho phòng Kế toán."*

> *"Làm slide báo cáo năm 2025 cho phòng HC-NS."*

Nếu không nói rõ phòng ban, Claude sẽ hỏi lại trước khi tiến hành.

### Các cụm từ kích hoạt skill phổ biến

- *"Báo cáo tuần / tháng / năm"*
- *"KPI report"*
- *"Slide tổng hợp"*
- *"Báo cáo vận hành"*
- *"Làm slide từ dữ liệu"*
- *"Báo cáo phòng [tên phòng]"*

---

## 6. Định dạng file dữ liệu đầu vào

Skill đọc dữ liệu qua MCP `file-process:read_handle_document_report`. File phải ở vị trí mà MCP có thể truy cập.

### 13 cột chuẩn (bắt buộc)

| Cột | Kiểu | Ghi chú |
|---|---|---|
| `Từ ngày` | Date | Ngày bắt đầu công việc |
| `Đến ngày` | Date | Ngày kết thúc công việc |
| `Tiêu đề` | Text | Tên công việc |
| `Mô tả` | Text | Mô tả chi tiết |
| `Kết quả` | Text | Kết quả đạt được |
| `Dự án` | Text | Tên dự án (nếu có) |
| `Loại công việc` | Text | Prefix theo phòng (vd `CNTT-HT-*`, `KT-*`) |
| `Khẩn cấp` | Boolean | True/False |
| `Quan trọng` | Boolean | True/False |
| `Đã hoàn tất` | Boolean | True/False |
| `Đã hủy` | Boolean | True/False |
| `Duration_hours` | Number | Số giờ thực hiện |
| `Score` | Number | Điểm ưu tiên (tự tính hoặc để sẵn) |

> **Lưu ý quan trọng:** Skill **KHÔNG** nhận file upload trực tiếp từ chat. Dữ liệu phải được đọc qua MCP `file-process`. Nếu MCP chưa kết nối hoặc file thiếu cột → Claude sẽ dừng và báo lỗi cụ thể.

---

## 7. Phòng ban được hỗ trợ

| Mã | Tên đầy đủ trên bìa | Prefix `Loại công việc` |
|---|---|---|
| CNTT | PHÒNG CÔNG NGHỆ THÔNG TIN | `CNTT-HT-*`, `CNTT-ERP-*`, `CNTT-PR-*`… |
| KT | PHÒNG KẾ TOÁN | `KT-*` |
| HCM | PHÒNG HÀNH CHÍNH - NHÂN SỰ | `HCM-*` |
| KD | PHÒNG KINH DOANH | `KD-*` |
| QLK | PHÒNG QUẢN LÝ KHO | `QLK-*` |

Phòng ban chưa có mapping → Claude tự sinh từ prefix data thực tế và hỏi xác nhận trước khi xuất slide.

---

## 8. Quy trình tự động (Pipeline 6 bước)

Claude thực hiện tuần tự, không nhảy bước:

| Bước | Nội dung |
|---|---|
| **1** | Xác định kỳ báo cáo (tuần/tháng/năm) và phòng ban |
| **2** | Đọc dữ liệu qua MCP, parse 13 cột, tính Score ưu tiên |
| **2b** | Phân nhóm `Loại công việc` → 4–5 section logic |
| **3** | Tổng hợp outline (`report_data`): viết body câu liền mạch, đánh giá + định hướng |
| **4** | Xuất `report_data.json` qua validators (kiểm tra brand + nội dung) |
| **5** | Render slide bằng pptxgenjs (skill `pptx` lo layout tự do, brand_kit.js ép brand) |
| **6** | QA visual từng slide → xuất file `.pptx` |

---

## 9. Ràng buộc brand (không thay đổi được)

| Element | Quy tắc |
|---|---|
| Kích thước slide | 20 × 11.25 inch |
| Nền + logo + footer | Từ `bg-cover.jpg` / `bg-content.jpg` — KHÔNG tự vẽ logo/footer |
| Font | Segoe UI (toàn bộ — kể cả tiếng Việt) |
| Màu tiêu đề | Đỏ `#FF0000` |
| Màu header | Navy `#000099` = đã hoàn tất · Đỏ = chưa hoàn tất |
| Cấm | Emoji trạng thái (`✅⏳🔴`), tên cá nhân, slide chỉ có text thuần |

---

## 10. Xử lý sự cố thường gặp

| Vấn đề | Nguyên nhân | Cách khắc phục |
|---|---|---|
| "MCP fail / thiếu cột" | File dữ liệu thiếu cột hoặc MCP chưa kết nối | Kiểm tra kết nối MCP `file-process`; đảm bảo file có đủ 13 cột |
| Slide bị lỗi font (ô vuông) | Xem trước trên Linux không có Segoe UI | Mở file `.pptx` bằng PowerPoint trên Windows — chữ sẽ hiển thị đúng |
| Skill không tự kích hoạt | Toggle skill chưa bật | Vào Customize → Skills → bật toggle ở góc phải |
| Claude hỏi phòng ban | Yêu cầu không nêu rõ phòng | Ghi rõ phòng trong câu yêu cầu, hoặc chọn trong hộp hội thoại |
| Báo cáo thiếu nhóm công việc | Prefix `Loại công việc` chưa match mapping | Kiểm tra file data; Claude sẽ hỏi xác nhận nếu prefix lạ |
| Định hướng lặp lại "Trọng tâm" | Lỗi logic phân biệt pending/direction | Báo Phòng CNTT để cập nhật `helpers.py` |

---

## 11. Bảo trì và cập nhật

- **Phiên bản hiện tại:** 4.1 (tháng 6/2026)
- **Đơn vị quản lý:** Phòng Công nghệ thông tin — Công ty Cổ phần Tôn Đông Á
- **Đầu mối hỗ trợ:** Liên hệ Phòng CNTT qua EOffice hoặc email nội bộ

Khi có thay đổi template (logo mới, đổi màu, layout mới), Phòng CNTT cập nhật file `.skill` mới và thông báo qua kênh nội bộ. Người dùng upload lại file mới — Claude tự thay phiên bản cũ.

---

## 12. Tài liệu tham khảo

- `SKILL.md` — Pipeline 6 bước đầy đủ, ràng buộc brand, quy tắc viết body
- `references/brand-constraints.md` — Hợp đồng brand chi tiết (màu, font, vùng an toàn)
- `scripts/helpers.py` — Logic phân tích, validators, mapping phòng ban
- `scripts/brand_kit.js` — Brand kit pptxgenjs (primitive ép ràng buộc)

---

*Tài liệu được biên soạn bởi Phòng CNTT — Tôn Đông Á.*
