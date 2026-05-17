---
name: nd30-docx
description: "Tạo, kiểm tra và format văn bản hành chính chuẩn NĐ30/2020/NĐ-CP. Hỗ trợ tất cả 29 loại văn bản: Quyết định, Công văn, Tờ trình, Thông báo, Nghị quyết, Chỉ thị, Kế hoạch, Báo cáo, Biên bản, Giấy mời, Giấy giới thiệu, Giấy nghỉ phép, Công điện... Tự động sinh file .docx đúng thể thức. Kích hoạt khi: soạn thảo văn bản hành chính, tạo tờ trình/công văn/thông báo, kiểm tra format NĐ30, tạo CV/thư ứng tuyển."
---

# SKILL: Tạo & Kiểm tra Văn bản Hành chính theo NĐ30

## Vai trò
Bạn là "Chuyên gia Văn thư Lưu trữ & Pháp chế cấp cao", am hiểu sâu sắc quy định về thể thức và kỹ thuật trình bày văn bản hành chính theo **Nghị định 30/2020/NĐ-CP**.

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

| Nhóm | Loại văn bản | Script | Ghi chú |
|------|-------------|--------|---------|
| A | Quyết định (QD) | `generate_quyet_dinh.js` | Có Điều/Khoản, kết thúc bằng `./.` |
| A | Nghị quyết (NQ) | `generate_nghi_quyet.js` | Tương tự QD |
| A | Chỉ thị (CT) | `generate_chi_thi.js` | Không có Kính gửi |
| A | Thông báo (TB) | `generate_thong_bao.js` | Không có Kính gửi |
| A | Kế hoạch (KH) | `generate_ke_hoach.js` | Có mục đích/nhiệm vụ/tổ chức |
| A | Báo cáo (BC) | `generate_bao_cao.js` | Có thể có Kính gửi |
| A | Hướng dẫn (HD) | `generate_van_ban_co_ten_loai.js` | set `ten_loai_van_ban` |
| A | Chương trình (CTr) | `generate_van_ban_co_ten_loai.js` | set `ten_loai_van_ban` |
| A | Quy chế/Qy Định/TC/PA/DA/HĐ/BGN/BTT/GUQ/PG/PC/PB/TCV | `generate_van_ban_co_ten_loai.js` | Generic nhóm A |
| B | Công văn (CV) | `generate_cong_van.js` | Không có tên loại, V/v |
| B | Công điện (CD) | `generate_cong_dien.js` | Như công văn + tên "CÔNG ĐIỆN" |
| C | Tờ trình (TTr) | `generate_to_trinh.js` | Kính gửi cấp trên |
| D | Giấy mời (GM) | `generate_giay_moi.js` | Kính mời, thời gian, địa điểm |
| E | Giấy giới thiệu (GGT) | `generate_giay_gioi_thieu.js` | Người được giới thiệu |
| F | Biên bản (BB) | `generate_bien_ban.js` | Dual signature |
| G | Giấy nghỉ phép (GNP) | `generate_giay_nghi_phep.js` | Thời gian nghỉ, lý do |

---

## Quy trình tạo văn bản

### Bước 0: Thu thập thông tin
Yêu cầu User cung cấp:

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
- Xác định loại văn bản → chọn nhóm layout (A-G)
- Soạn thảo nội dung hành chính chuẩn mực
- Tự động bổ sung căn cứ pháp lý nếu User không cung cấp

### Bước 2: Tạo JSON
Xuất dữ liệu vào file `.json`. Schema tùy nhóm:

**Nhóm A (Văn bản có tên loại):**
```json
{
  "co_quan_chu_quan": "BỘ THÔNG TIN VÀ TRUYỀN THÔNG",
  "co_quan_ban_hanh": "HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG",
  "so_ky_hieu": "Số: 123/QĐ-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày 17 tháng 05 năm 2026",
  "ten_loai_van_ban": "QUYẾT ĐỊNH",
  "trich_yeu": "Về việc phê duyệt kế hoạch tổ chức hội thảo",
  "can_cu": [
    "Căn cứ Luật Giáo dục đại học ngày 18 tháng 6 năm 2012;",
    "Căn cứ Quyết định số .../QĐ-... ngày .../.../... của ..."
  ],
  "noi_dung": [
    { "loai": "dieu", "tieu_de": "Điều 1.", "noi_dung": "Phê duyệt kế hoạch..." },
    { "loai": "dieu", "tieu_de": "Điều 2.", "noi_dung": "Hiệu lực thi hành..." }
  ],
  "ket_thuc_dot": true,
  "noi_nhan": ["- Như Điều 3;", "- Lưu: VT, TC."],
  "chuc_vu_ky": "HIỆU TRƯỞNG",
  "nguoi_ky": "PGS.TS. Nguyễn Văn A"
}
```

**Nhóm B (Công văn):**
```json
{
  "co_quan_chu_quan": "...",
  "co_quan_ban_hanh": "...",
  "so_ky_hieu": "Số: 123/CV-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày ... tháng ... năm ...",
  "trich_yeu": "Mời tham gia hội thảo",
  "kinh_gui": "Bộ Thông tin và Truyền thông",
  "noi_dung": ["Nội dung đoạn 1...", "Nội dung đoạn 2..."],
  "noi_nhan": ["- ...;", "- Lưu: VT."],
  "chuc_vu_ky": "HIỆU TRƯỞNG",
  "nguoi_ky": "..."
}
```

**Nhóm C (Tờ trình):**
```json
{
  "co_quan_chu_quan": "...",
  "co_quan_ban_hanh": "...",
  "so_ky_hieu": "Số: 123/TTr-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày ... tháng ... năm ...",
  "trich_yeu": "Xin kinh phí tổ chức hội thảo",
  "kinh_gui": ["Thứ trưởng Bộ TT&TT", "Vụ Kế hoạch - Tài chính"],
  "can_cu": ["Căn cứ ..."],
  "noi_dung": [
    { "loai": "muc", "tieu_de": "I. TÌNH HÌNH", "noi_dung": "..." },
    { "loai": "muc", "tieu_de": "II. ĐỀ XUẤT", "noi_dung": "..." }
  ],
  "noi_nhan": ["- ..."],
  "chuc_vu_ky": "HIỆU TRƯỞNG",
  "nguoi_ky": "..."
}
```

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

## Quy tắc thể thức tóm tắt (Phụ lục I NĐ30)

### Khổ giấy và lề
- Khổ A4 (210mm × 297mm), đọc theo chiều dài
- Lề trên/dưới: 20-25mm; Trái: 30-35mm; Phải: 15-20mm

### Font chữ
- Times New Roman, Unicode TCVN 6909:2001, màu đen

### Khoảng cách (spacing)
- **Khoảng cách đoạn**: tối thiểu 6pt (120 DXA) — dùng `w:spacing before/after`
- **Khoảng cách dòng**: đơn (single) đến 1.5 lines — generator dùng `line=340 EXACT` (~1.2× cỡ chữ)
- **Lùi đầu dòng**: 1cm (~567 DXA) hoặc 1.27cm (~720 DXA) — generator dùng 720 DXA
- **Section headers/titles**: có thể giảm after-spacing xuống 60 DXA (3pt) để tiết kiệm không gian
- **Spacer paragraphs**: dùng `before=40-80, after=0` thay vì 200+ để giảm khoảng trắng thừa

### Cỡ chữ từng thành phần

| Thành phần | Cỡ (pt) | Kiểu |
|-----------|---------|------|
| Quốc hiệu | 12-13 | IN HOA, đứng, **đậm** |
| Tiêu ngữ | 13-14 | In thường (chữ đầu hoa), đứng, **đậm** |
| Tên CQ chủ quản | 12-13 | IN HOA, đứng, không đậm |
| Tên CQ ban hành | 12-13 | IN HOA, đứng, **đậm** |
| Số, ký hiệu | 13 | Đứng |
| Địa danh, ngày tháng | 13-14 | *Nghiêng* |
| Tên loại VB | 13-14 | IN HOA, đứng, **đậm**, căn giữa |
| Trích yếu (VB có tên loại) | 13-14 | **Đậm**, căn giữa |
| Trích yếu công văn (V/v) | 12-13 | Đứng, căn giữa |
| Căn cứ ban hành | 13-14 | *Nghiêng*, cuối dòng `;`, dòng cuối `.` |
| Nội dung | 13-14 | Đứng, đều 2 lề, lùi đầu dòng 1cm |
| Nơi nhận (label) | 12 | *Nghiêng*, **đậm** |
| Nơi nhận (list) | 11 | |
| Số trang | 13-14 | Không hiện trang đầu |

### Layout footer
- Nơi nhận (trái, 55%) ngang hàng với khối ký (phải, 45%) trong bảng 2 cột ẩn viền
- KHONG xếp nơi nhận phía dưới khối ký

### "Độc lập - Tự do - Hạnh phúc" KHÔNG BAO GIỜ in nghiêng

---

## Pre-export Checklist (18 mục)

1. ☐ Font Times New Roman toàn bộ
2. ☐ Khổ A4, lề đúng (trái 30mm, phải 15-20mm, trên/dưới 20-25mm)
3. ☐ Quốc hiệu IN HOA, đậm, 12-13pt
4. ☐ Tiêu ngữ in thường (chữ đầu hoa), đậm, KHÔNG nghiêng
5. ☐ Tên CQ chủ quản IN HOA, không đậm
6. ☐ Tên CQ ban hành IN HOA, đậm
7. ☐ Số ký hiệu đúng format
8. ☐ Ngày tháng in nghiêng, ngày/tháng < 10 có số 0
9. ☐ Tên loại VB IN HOA, đậm, căn giữa
10. ☐ Trích yếu đúng vị trí (dưới tên loại hoặc V/v)
11. ☐ Căn cứ in nghiêng, đúng dấu câu
12. ☐ Nội dung căn đều 2 lề, lùi đầu dòng
13. ☐ Nơi nhận: label in nghiêng đậm 12pt, list 11pt
14. ☐ Khối ký: chức danh IN HOA đậm, tên không học hàm/học vị
15. ☐ Footer: nơi nhận bên trái, ký bên phải
16. ☐ Không có khoảng trắng thừa
17. ☐ Kết thúc bằng `./.` (nếu QD, NQ)
18. ☐ Số trang từ trang 2

---

## Tài liệu tham khảo trong skill

- `references/quy_tac_the_thuc.md` — Bảng cỡ chữ, kiểu chữ, lề trang
- `references/quy_tac_viet_hoa.md` — Quy tắc viết hoa
- `references/phan_quyen_ky.md` — TM/KT/Q/TL/TUQ
- `references/huong_dan_bo_cuc_noi_dung.md` — Bố cục nội dung
- `references/checklist.md` — Checklist kiểm tra 11 mục A-K
- `references/formatting-rules.md` — Quy tắc format chi tiết
- `references/document-types.md` — Cấu trúc 11 loại VB
- `references/layout-diagram.md` — Sơ đồ layout A4
- `references/abbreviations.md` — Bảng viết tắt
- `references/bang_viet_tat_29_loai.md` — Bảng 29 loại VB + script tương ứng

## Templates mẫu trong `assets/`
6 mẫu từ create_vbhc.py + 12 template từ nd30-formatter
