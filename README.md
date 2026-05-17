# nd30-guide

Công cụ tạo và kiểm tra văn bản hành chính theo **Nghị định 30/2020/NĐ-CP** (NĐ30), hỗ trợ bởi Claude Code skills.

## Tổng quan

Dự án bao gồm bộ Claude Code skills để tự động sinh file `.docx` đúng thể thức văn bản hành chính, kiểm tra tuân thủ NĐ30, và tạo hồ sơ ứng tuyển (CV, cover letter, email).

### Tình huống thực tế

Dự án sử dụng tình huống: **"Nền tảng AI hỗ trợ bảo tồn và giảng dạy tiếng Jrai (mở rộng sang Bahnar) cho chuyển đổi số tỉnh Gia Lai"** tại Học viện Công nghệ Bưu chính Viễn thông (PTIT).

## Cấu trúc

```
.claude/skills/
  nd30-docx/           # Skill chính: tạo & kiểm tra VBHC theo NĐ30
    SKILL.md           # Hướng dẫn sử dụng skill
    scripts/           # Generator scripts (Node.js) + checker (Python)
    references/        # Quy tắc thể thức, checklist, layout
    assets/            # 18 template .docx mẫu
    examples/          # Ví dụ JSON data
  cv-email/            # Skill tạo CV, cover letter, email
    SKILL.md
    references/        # Cấu trúc CV, ATS tips, lỗi thường gặp
    templates/         # JSON schema
    examples/          # Ví dụ JSON
  docx/                # Skill thao tác file .docx
  nd30-formatter/      # Web app format VBHC (Cloudflare Workers)
  nd30-document-drafter/ # Draft VBHC từ nội dung

scripts/
  rewrite_admin_docs.py   # Tạo lại 3 VBHC (tờ trình, công văn, thông báo)
  compile_report.py       # Compile báo cáo SKD1103 vào 1 file .docx
  create_cover_letter_email.py  # Tạo cover letter + email
  fix_cv.py               # Sửa CV (font, lề)
  fix_admin_docs.py       # Sửa VBHC (header, ngày tháng)

output/                # Thư mục output (gitignored)
```

## Sử dụng

### Kiểm tra NĐ30

```bash
python .claude/skills/nd30-docx/scripts/check_nd30.py --check output/file.docx
```

### Tạo VBHC bằng generator

```bash
cd .claude/skills/nd30-docx
node scripts/generate_to_trinh.js --input data.json --output to_trinh.docx
node scripts/generate_cong_van.js --input data.json --output cong_van.docx
node scripts/generate_thong_bao.js --input data.json --output thong_bao.docx
```

### Tạo lại VBHC bằng Python

```bash
python scripts/rewrite_admin_docs.py
```

### Compile báo cáo

```bash
python scripts/compile_report.py
```

## Quy tắc thể thức chính (NĐ30)

| Thành phần | Cỡ (pt) | Kiểu |
|-----------|---------|------|
| Quốc hiệu | 12-13 | IN HOA, đứng, đậm |
| Tiêu ngữ | 13-14 | In thường, đứng, đậm (KHÔNG nghiêng) |
| Tên CQ chủ quản | 12-13 | IN HOA, đứng, không đậm |
| Tên CQ ban hành | 12-13 | IN HOA, đứng, đậm |
| Số, ký hiệu | 13 | Đứng |
| Địa danh, ngày tháng | 13-14 | *Nghiêng* |
| Nội dung | 13-14 | Đứng, đều 2 lề, lùi đầu dòng 1cm |
| Nơi nhận (label) | 12 | *Nghiêng*, đậm |
| Nơi nhận (list) | 11 | Đứng |

**Lề trang:** Trái 30mm, phải 15mm, trên/dưới 20mm. Khổ A4.

**Khoảng cách:** Đoạn tối thiểu 6pt, dòng đơn đến 1.5 lines.

## Tài liệu tham khảo

- [Nghị định 30/2020/NĐ-CP](https://vanban.chinhphu.vn/default.aspx?pageid=27160&docid=204173)
- Phụ lục I: Thể thức và kỹ thuật trình bày VBHC
- Giáo trình SKD1103 — Kỹ năng tạo lập văn bản tiếng Việt, PTIT
