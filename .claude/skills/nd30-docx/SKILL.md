---
name: nd30-docx
description: "Tạo, kiểm tra và format văn bản hành chính chuẩn NĐ30/2020/NĐ-CP. Hỗ trợ tất cả 29 loại văn bản: Quyết định, Công văn, Tờ trình, Thông báo, Nghị quyết, Chỉ thị, Kế hoạch, Báo cáo, Biên bản, Giấy mời, Giấy giới thiệu, Giấy nghỉ phép, Công điện... Tự động sinh file .docx đúng thể thức. Kích hoạt khi: soạn thảo văn bản hành chính, tạo tờ trình/công văn/thông báo, kiểm tra format NĐ30."
---

# SKILL: Tạo & Kiểm tra Văn bản Hành chính theo NĐ30

## Vai trò
Bạn là "Chuyên gia Văn thư Lưu trữ & Pháp chế cấp cao", am hiểu sâu sắc quy định về thể thức và kỹ thuật trình bày văn bản hành chính theo **Nghị định 30/2020/NĐ-CP**.

> **Full spec**: Xem `AGENT.md` ở project root cho toàn bộ quy tắc thể thức, code patterns, và reference files.

---

## 3 Chế độ hoạt động

### Chế độ 1: Tạo mới văn bản
Quy trình: Thu thập thông tin → Soạn thảo nội dung → Tạo JSON → Chạy script → Kiểm tra kết quả

### Chế độ 2: Kiểm tra format
```bash
python scripts/check_nd30.py --check <file.docx>
```
Kiểm tra: khổ giấy, lề, font, quốc hiệu, tiêu ngữ, số ký hiệu, ngày tháng, nơi nhận.

### Chế độ 3: Format lại văn bản
Dùng skill `docx` để unpack → sửa XML → repack.

---

## Bảng mapping 29 loại văn bản

### Nhóm A — Văn bản có tên loại (dùng `generate_van_ban_co_ten_loai.js` hoặc wrapper)

| Loại | Viết tắt | Wrapper riêng | Ghi chú |
|------|---------|---------------|---------|
| Quyết định | QĐ | `generate_quyet_dinh.js` | Điều/Khoản, kết thúc `./.` |
| Nghị quyết | NQ | `generate_nghi_quyet.js` | Tương tự QĐ |
| Chỉ thị | CT | `generate_chi_thi.js` | Không Kính gửi |
| Thông báo | TB | `generate_thong_bao.js` | Không Kính gửi |
| Kế hoạch | KH | `generate_ke_hoach.js` | Mục đích/nhiệm vụ/tổ chức |
| Báo cáo | BC | `generate_bao_cao.js` | Có thể Kính gửi |
| Hướng dẫn | HD | — | set `ten_loai_van_ban` |
| Chương trình | CTr | — | set `ten_loai_van_ban` |
| QC/QĐ con/TC/PA/DA/HĐ/BGN/BTT/GUQ/PG/PC/PB/TCV | — | — | Generic nhóm A |

### Nhóm B–G — Văn bản đặc thù

| Nhóm | Loại | Script | Ghi chú |
|------|------|--------|---------|
| B | Công văn (CV) | `generate_cong_van.js` | Không tên loại, V/v, có Kính gửi |
| B | Công điện (CD) | `generate_cong_dien.js` | Như CV + tên "CÔNG ĐIỆN" |
| C | Tờ trình (TTr) | `generate_to_trinh.js` | Kính gửi cấp trên |
| D | Giấy mời (GM) | `generate_giay_moi.js` | Kính mời, thời gian, địa điểm |
| E | Giấy giới thiệu (GGT) | `generate_giay_gioi_thieu.js` | Người được giới thiệu |
| F | Biên bản (BB) | `generate_bien_ban.js` | Dual signature |
| G | Giấy nghỉ phép (GNP) | `generate_giay_nghi_phep.js` | Thời gian nghỉ, lý do |

---

## Quy trình tạo văn bản

### Bước 0: Thu thập thông tin

| STT | Thông tin | Chi tiết |
|-----|----------|----------|
| 1 | Loại văn bản | Tờ trình, công văn, quyết định, thông báo... |
| 2 | Cơ quan ban hành | Tên cơ quan chủ quản + cơ quan ban hành |
| 3 | Số ký hiệu | Số: .../...-... |
| 4 | Địa danh, ngày tháng | Hà Nội, ngày... tháng... năm... |
| 5 | Căn cứ pháp lý | Luật, Nghị định, Quyết định liên quan |
| 6 | Nội dung chính | Mục đích, yêu cầu, phương án |
| 7 | Nơi nhận | Danh sách nhận |
| 8 | Người ký | Chức danh + họ tên |

### Bước 1: Soạn thảo
- Xác định loại văn bản → chọn nhóm layout (A–G)
- Soạn thảo nội dung hành chính chuẩn mực
- Tự động bổ sung căn cứ pháp lý nếu User không cung cấp

### Bước 2: Tạo JSON
Xuất dữ liệu vào file `.json`. Xem `AGENT.md` cho schema chi tiết từng nhóm.

### Bước 3: Chạy script
```bash
cd .claude/skills/nd30-docx
node scripts/generate_<loai>.js --input data.json --output output.docx
```

### Bước 4: Kiểm tra
```bash
python scripts/check_nd30.py --check output.docx
```

---

## Quy tắc thể thức tóm tắt

> Chi tiết: `AGENT.md` + `references/formatting-rules.md`

### Thiết lập trang
- Khổ A4 (210×297mm), lề: trái 30mm, phải 15mm, trên/dưới 20mm
- Font: Times New Roman, Unicode, màu đen

### Cỡ chữ từng thành phần

| Thành phần | Cỡ (pt) | Kiểu |
|-----------|---------|------|
| Quốc hiệu | 12–13 | IN HOA, đứng, **đậm** |
| Tiêu ngữ | 13–14 | In thường, **đứng, đậm** (KHÔNG nghiêng) |
| Tên CQ chủ quản | 12–13 | IN HOA, đứng, **không đậm** |
| Tên CQ ban hành | 12–13 | IN HOA, đứng, **đậm** |
| Số, ký hiệu | 13 | Đứng |
| Địa danh, ngày tháng | 13–14 | *Nghiêng* (thành phần DUY NHẤT in nghiêng) |
| Tên loại VB | 13–14 | IN HOA, đứng, **đậm**, căn giữa |
| Trích yếu (có tên loại) | 13–14 | **Đậm**, căn giữa |
| Trích yếu (công văn V/v) | 12–13 | Đứng, căn giữa |
| Căn cứ ban hành | 13–14 | *Nghiêng*, cuối dòng `;`, dòng cuối `.` |
| Nội dung | 13–14 | Đứng, đều 2 lề, lùi đầu dòng 1cm |
| Nơi nhận (label) | 12 | *Nghiêng*, **đậm** |
| Nơi nhận (list) | 11 | Đứng |

### Khoảng cách
- Đoạn: tối thiểu 6pt (120 DXA)
- Dòng: đơn đến 1.5 lines
- Lùi đầu dòng: 1cm (~567 DXA) hoặc 1.27cm (~720 DXA)

### Layout footer
- Nơi nhận (trái, 55%) ngang hàng với khối ký (phải, 45%) trong bảng 2 cột ẩn viền
- KHÔNG xếp nơi nhận phía dưới khối ký

---

## Pre-export Checklist (18 mục)

1. ☐ Font Times New Roman toàn bộ
2. ☐ Khổ A4, lề đúng (trái 30mm, phải 15mm, trên/dưới 20mm)
3. ☐ Quốc hiệu IN HOA, đậm, 12–13pt
4. ☐ Tiêu ngữ in thường, đậm, KHÔNG nghiêng
5. ☐ Tên CQ chủ quản IN HOA, không đậm
6. ☐ Tên CQ ban hành IN HOA, đậm
7. ☐ Số ký hiệu đúng format (`Số: XX/YYY-ZZZ`)
8. ☐ Ngày tháng in nghiêng, ngày/tháng < 10 có số 0
9. ☐ Tên loại VB IN HOA, đậm, căn giữa
10. ☐ Trích yếu đúng vị trí (dưới tên loại hoặc V/v)
11. ☐ Căn cứ in nghiêng, đúng dấu câu (`;` cuối dòng, `.` dòng cuối)
12. ☐ Nội dung căn đều 2 lề, lùi đầu dòng
13. ☐ Nơi nhận: label nghiêng đậm 12pt, list 11pt
14. ☐ Khối ký: chức danh IN HOA đậm, tên không học hàm/học vị
15. ☐ Footer: nơi nhận bên trái, ký bên phải
16. ☐ Không có khoảng trắng thừa
17. ☐ Kết thúc bằng `./.` (nếu QĐ, NQ)
18. ☐ Số trang từ trang 2

---

## Tài liệu tham khảo trong skill

| File | Nội dung |
|------|----------|
| `references/formatting-rules.md` | Quy tắc format chi tiết (Phụ lục I) |
| `references/checklist.md` | Checklist kiểm tra 11 mục A–K |
| `references/layout-diagram.md` | Sơ đồ layout A4 (ASCII art) |
| `references/document-types.md` | Cấu trúc 11 loại VB |
| `references/abbreviations.md` | Bảng viết tắt |
| `references/quy_tac_the_thuc.md` | Quy tắc thể thức (không dấu) |
| `references/quy_tac_viet_hoa.md` | Quy tắc viết hoa (Phụ lục II) |
| `references/phan_quyen_ky.md` | TM/KT/Q/TL/TUQ |
| `references/huong_dan_bo_cuc_noi_dung.md` | Bố cục nội dung |
| `references/bang_viet_tat_29_loai.md` | 29 loại VB + script |

## Templates mẫu trong `assets/`
- 6 mẫu Python: `mau_*.docx` (từ `create_vbhc.py`)
- 12 template Node.js: `Template_*.docx` (từ generators)
