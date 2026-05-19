# AGENT.md — NĐ30 Document Generation Guide

Comprehensive reference for any agent working with Vietnamese administrative documents (văn bản hành chính) under **Nghị định 30/2020/NĐ-CP**.

---

## Project Overview

This project automates creation and validation of NĐ30-compliant `.docx` files. It includes:
- **Skills** (`.claude/skills/`) — specialized generators and checkers
- **Scripts** (`scripts/`) — standalone Python pipelines for report compilation, document splitting, CV/cover letter generation
- **Output** (`output/`) — generated `.docx` files

---

## NĐ30 Formatting Spec (Phụ lục I)

### Page Setup

| Property | Value | DXA (for python-docx) |
|----------|-------|-----------------------|
| Paper | A4 (210 × 297 mm) | width: 11906, height: 16838 |
| Margin top | 20–25 mm | 1134–1418 |
| Margin bottom | 20–25 mm | 1134–1418 |
| Margin left | 30–35 mm | 1701–1985 |
| Margin right | 15–20 mm | 851–1134 |
| Default font | Times New Roman, Unicode, black | — |

### 9 Mandatory Components (Thể thức)

| # | Component | Size | Style | Position |
|---|-----------|------|-------|----------|
| 1a | Quốc hiệu | 12–13pt | **IN HOA, đứng, đậm** | Top-right header table |
| 1b | Tiêu ngữ | 13–14pt | In thường, **đứng, đậm** (KHÔNG nghiêng) | Below QH, with horizontal rule |
| 2 | Tên CQ chủ quản | 12–13pt | IN HOA, đứng, **không đậm** | Top-left header table |
| 2 | Tên CQ ban hành | 12–13pt | IN HOA, **đứng, đậm** | Below CQ chủ quản |
| 3 | Số, ký hiệu | 13pt | Đứng | Below header, centered |
| 4 | Địa danh + ngày | 13–14pt | In thường, **nghiêng** (ONLY italic element) | Same line as Ô 3, right-aligned |
| 5a | Tên loại VB | 13–14pt | IN HOA, **đứng, đậm**, centered | Below Ô 3/4 (for docs with name type) |
| 5b | V/v (Công văn) | 12–13pt | In thường, đứng | Below Ô 3/4 (no name type) |
| 6 | Nội dung | 13–14pt | Đứng, justify, indent 1cm | Main body |
| 7a | Quyền hạn + Chức vụ | 13–14pt | IN HOA, **đứng, đậm** | Footer right |
| 7b | Họ tên người ký | 13–14pt | In thường, **đứng, đậm** (NO học hàm/học vị) | Below 7a |
| 8 | Dấu cơ quan | — | Red, overlaps 1/3 signature left | — |
| 9a | Kính gửi | 13–14pt | Đứng | Before content (CV, TTr, BC only) |
| 9b | Nơi nhận | 12pt label (nghiêng, đậm), 11pt list | Đứng | Footer left, parallel to signature |

### Spacing Rules

| Element | Value |
|---------|-------|
| Paragraph spacing | min 6pt (120 DXA) before/after |
| Line spacing | single to 1.5 lines |
| First-line indent | 1cm (~567 DXA) or 1.27cm (~720 DXA) |
| Content table headers | IN HOA, đậm, 12–13pt, centered, repeat on new page |

### Header/Footer Layout (2-column invisible tables)

```
HEADER TABLE (50%–50%):
┌─────────────────────┬─────────────────────────────┐
│ LEFT: Ô 2           │ RIGHT: Ô 1                  │
│ Tên CQ chủ quản     │ Quốc hiệu (IN HOA, đậm)     │
│ Tên CQ ban hành (đậm)│ Tiêu ngữ (đứng, đậm)        │
│    ___________       │ ___________________________  │
└─────────────────────┴─────────────────────────────┘

FOOTER TABLE (55%–45%):
┌─────────────────────┬─────────────────────────────┐
│ LEFT: Ô 9b           │ RIGHT: Ô 7a                 │
│ Nơi nhận:            │ TM./KT./Q. CHỨC VỤ          │
│ - ...;               │     [Chữ ký]                │
│ - Lưu: VT, VP.       │     Họ tên (Ô 7b)           │
│ (Ô 12: PL.(300))     │     [Dấu] (Ô 8)             │
└─────────────────────┴─────────────────────────────┘
```

### Document Type Groups

| Group | Types | Has Name Type | Has Kính gửi | Script |
|-------|-------|---------------|-------------|--------|
| A | QĐ, NQ, CT, TB, KH, BC, HD, CTr, QC, QĐ con, TC, PA, DA, HĐ, BGN, BTT, GUQ, PG, PC, PB, TCV | Yes | No (some BC: Yes) | `generate_van_ban_co_ten_loai.js` or specific wrapper |
| B | Công văn, Công điện | No (V/v) | Yes | `generate_cong_van.js` / `generate_cong_dien.js` |
| C | Tờ trình | Yes (TỜ TRÌNH) | Yes (cấp trên) | `generate_to_trinh.js` |
| D | Giấy mời | — | — | `generate_giay_moi.js` |
| E | Giấy giới thiệu | — | — | `generate_giay_gioi_thieu.js` |
| F | Biên bản | — | — | `generate_bien_ban.js` |
| G | Giấy nghỉ phép | — | — | `generate_giay_nghi_phep.js` |

### Common Mistakes

| # | Error | Correct |
|---|-------|---------|
| 1 | Tiêu ngữ in nghiêng | **Đứng, đậm** — NEVER italic |
| 2 | CQ chủ quản in đậm | **Không đậm** |
| 3 | Địa danh+ngày đứng | **Nghiêng** — ONLY italic element |
| 4 | Số 5 (thiếu 0) | `Số: 05/...` |
| 5 | Ngày 3 tháng 1 | `ngày 03 tháng 01` |
| 6 | Font Arial/Cambria | **Times New Roman** |
| 7 | Ghi "TS.", "ThS." trước tên ký | **Không ghi** học hàm/học vị |
| 8 | Nơi nhận cỡ 13 | Label 12pt, list 11pt |
| 9 | Lề trái 25mm | **30–35 mm** |
| 10 | Thiếu `./.` | Cuối QĐ, NQ: `./.` |

---

## Available Tools

### Skills (`.claude/skills/`)

| Skill | Purpose | Trigger |
|-------|---------|---------|
| `nd30-docx` | Create/check NĐ30 admin docs (29 types) | Soạn thảo VBHC, tạo tờ trình/công văn/thông báo |
| `cv-email` | CV, cover letter, email generation | Tạo CV, thư ứng tuyển, email |
| `docx` | Generic .docx manipulation | Read/edit/manipulate .docx files |
| `nd30-document-drafter` | Draft VBHC from content | Draft from natural language |
| `skill-stvb` | Alternative admin doc skill | — |

### Python Scripts (`scripts/`)

| Script | Purpose |
|--------|---------|
| `compile_report.py` | Compile full SKD1103 report from multiple .docx sources |
| `split_vbhc.py` | Split combined VBHC into ToTrinh, CongVan, ThongBao |
| `split_cover_email.py` | Split combined cover letter + email into 2 files |
| `rewrite_admin_docs.py` | Regenerate 3 admin docs from scratch |
| `create_cover_letter_email.py` | Generate cover letter + email .docx |
| `rewrite_cv_header.py` | Generate CV with 2-col header (80:20 ratio) |
| `extract_cv_from_report.py` | Extract CV section from compiled report |
| `create_theory.py` | Generate theory section content |
| `fix_cv.py` | Fix CV formatting (font, margins) |
| `fix_admin_docs.py` | Fix admin docs (header, dates) |

### Node.js Generators (`.claude/skills/nd30-docx/scripts/`)

| Script | Document Type |
|--------|--------------|
| `generate_van_ban_co_ten_loai.js` | Generic Group A (QĐ, NQ, CT, TB, HD, CTr, etc.) |
| `generate_quyet_dinh.js` | Quyết định |
| `generate_cong_van.js` | Công văn |
| `generate_to_trinh.js` | Tờ trình |
| `generate_ke_hoach.js` | Kế hoạch |
| `generate_bien_ban.js` | Biên bản |
| `generate_giay_moi.js` | Giấy mời |
| `generate_giay_gioi_thieu.js` | Giấy giới thiệu |
| `generate_giay_nghi_phep.js` | Giấy nghỉ phép |
| `generate_cong_dien.js` | Công điện |
| `check_nd30.py` | Compliance checker |

### Usage

```bash
# Generate a document
cd .claude/skills/nd30-docx
node scripts/generate_cong_van.js --input data.json --output cong_van.docx

# Check compliance
python .claude/skills/nd30-docx/scripts/check_nd30.py --check output/file.docx

# Compile report
python scripts/compile_report.py

# Split VBHC
python scripts/split_vbhc.py
```

---

## Code Patterns (python-docx)

### Page Setup

```python
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

for section in doc.sections:
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)

style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(13)
```

### Font Setting (critical for Vietnamese)

python-docx does NOT reliably set East Asian fonts. Always use raw XML:

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_font(paragraph, name='Times New Roman', size=Pt(13), bold=False, italic=False):
    for run in paragraph.runs:
        run.font.name = name
        run.font.size = size
        run.bold = bold
        run.italic = italic
        r = run._element
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            r.insert(0, rPr)
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), name)
        rFonts.set(qn('w:hAnsi'), name)
        rFonts.set(qn('w:eastAsia'), name)
```

### Paragraph Builder

```python
def add_para(doc, text, bold=False, italic=False, size=Pt(13), align=None,
             first_line_indent=None, space_after=None, line_spacing=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = size
    run.bold = bold
    run.italic = italic
    if align:
        p.alignment = align
    pf = p.paragraph_format
    if first_line_indent is not None:
        # EMU to twips: val / 914400 * 1440
        pPr = p._element.get_or_add_pPr()
        ind = OxmlElement('w:ind')
        ind.set(qn('w:firstLine'), str(int(first_line_indent / 914400 * 1440)))
        pPr.append(ind)
    if space_after is not None:
        pf.space_after = Pt(space_after)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    return p
```

### Heading Builder

```python
def add_heading_styled(doc, text, level=1, space_before=None, space_after=None):
    sizes = {1: Pt(16), 2: Pt(14), 3: Pt(13)}
    p = doc.add_heading(text, level=level)
    set_font(p, size=sizes.get(level, Pt(13)), bold=True)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p
```

### Table Utilities

```python
from docx.shared import Cm

def set_col_widths(table, widths_cm):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(widths_cm[i])

def set_cell_padding(table, top=0, bottom=100, left=80, right=80):
    """Set cell margins/padding (in twips)."""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)
            tcMar = OxmlElement('w:tcMar')
            for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
                el = OxmlElement(f'w:{side}')
                el.set(qn('w:w'), str(val))
                el.set(qn('w:type'), 'dxa')
                tcMar.append(el)
            old = tcPr.find(qn('w:tcMar'))
            if old is not None:
                tcPr.remove(old)
            tcPr.append(tcMar)

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    borders = OxmlElement('w:tblBorders')
    for bname in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{bname}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        borders.append(b)
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(borders)
```

### Page Break Strategies

```python
# Strategy 1: Attach PB to last non-bookmark paragraph
def add_page_break_to_last_para(doc):
    body = doc.element.body
    for elem in reversed(list(body)):
        if elem.tag == qn('w:p'):
            has_bm = any(b.get(qn('w:name')) for b in elem.iter(qn('w:bookmarkStart')))
            if has_bm:
                continue
            runs = elem.findall(qn('w:r'))
            run = runs[-1] if runs else OxmlElement('w:r')
            if not runs:
                elem.append(run)
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            run.append(br)
            return

# Strategy 2: Add PB after a specific element
def add_page_break_after(elem):
    if elem.tag == qn('w:p'):
        runs = elem.findall(qn('w:r'))
        run = runs[-1] if runs else OxmlElement('w:r')
        if not runs:
            elem.append(run)
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        run.append(br)
    elif elem.tag == qn('w:tbl'):
        p = OxmlElement('w:p')
        run = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        run.append(br)
        p.append(run)
        elem.addnext(p)
```

### Bookmark-Based Embedding

```python
_bookmark_counter = 100

def add_insertion_point(doc, name):
    global _bookmark_counter
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    _bookmark_counter += 1
    bm_id = str(_bookmark_counter)
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), bm_id)
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), bm_id)
    p._element.insert(0, start)
    p._element.append(end)
    return p

def insert_docx_after_bookmark(target_doc, source_path, bookmark_name):
    source = Document(source_path)
    body = target_doc.element.body
    marker = None
    for bm in body.iter(qn('w:bookmarkStart')):
        if bm.get(qn('w:name')) == bookmark_name:
            marker = bm
            break
    if marker is None:
        sect_pr = body.find(qn('w:sectPr'))
        insert_after = sect_pr
    else:
        insert_after = marker.getparent()
    last_inserted = None
    for element in source.element.body:
        if element.tag == qn('w:sectPr'):
            continue
        new_elem = deepcopy(element)
        insert_after.addnext(new_elem)
        insert_after = new_elem
        last_inserted = new_elem
    return last_inserted
```

### Cleanup: Remove Empty PB Paragraphs

```python
# After building all sections, clean up stray empty PB paragraphs
body = doc.element.body
for elem in list(body):
    if elem.tag == qn('w:p'):
        has_text = len(list(elem.iter(qn('w:t')))) > 0
        has_pb = len(list(elem.iter(qn('w:br')))) > 0
        if not has_text and has_pb:
            next_sib = elem.getnext()
            if next_sib is not None:
                # Transfer PB to next sibling
                if next_sib.tag == qn('w:p'):
                    runs = next_sib.findall(qn('w:r'))
                    run = runs[-1] if runs else OxmlElement('w:r')
                    if not runs:
                        next_sib.append(run)
                    br = OxmlElement('w:br')
                    br.set(qn('w:type'), 'page')
                    run.append(br)
                elif next_sib.tag == qn('w:tbl'):
                    p = OxmlElement('w:p')
                    run = OxmlElement('w:r')
                    br = OxmlElement('w:br')
                    br.set(qn('w:type'), 'page')
                    run.append(br)
                    p.append(run)
                    next_sib.addnext(p)
            body.remove(elem)
```

### Horizontal Rule (Vietnamese: đường kẻ ngang)

```python
def add_horizontal_rule(doc):
    p = doc.add_paragraph()
    pPr = p._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p._element.insert(0, pPr)
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)
```

---

## Reference Files

All located in `.claude/skills/nd30-docx/references/`:

| File | Content |
|------|---------|
| `formatting-rules.md` | Full Phụ lục I spec: 9 components, sizes, styles |
| `checklist.md` | 11-section validation checklist (A–K) |
| `layout-diagram.md` | ASCII art A4 layout with all 14 boxes |
| `document-types.md` | Structure for 11 document types with sample layouts |
| `abbreviations.md` | Document type abbreviations, signing authority codes |
| `quy_tac_the_thuc.md` | Formatting rules (Vietnamese, no diacritics) |
| `quy_tac_viet_hoa.md` | Capitalization rules (Phụ lục II) |
| `phan_quyen_ky.md` | TM/KT/Q/TL/TUQ signing authority details |
| `huong_dan_bo_cuc_noi_dung.md` | Content structure for QĐ and KH |
| `mau_can_cu_theo_linh_vuc.md` | Sample legal bases by field |
| `bang_viet_tat_29_loai.md` | All 29 document types with script mapping |

---

## JSON Schema for Generators

### Group A (Văn bản có tên loại)

```json
{
  "co_quan_chu_quan": "BỘ THÔNG TIN VÀ TRUYỀN THÔNG",
  "co_quan_ban_hanh": "HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG",
  "so_ky_hieu": "Số: 123/QĐ-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày 17 tháng 05 năm 2026",
  "ten_loai_van_ban": "QUYẾT ĐỊNH",
  "trich_yeu": "Về việc phê duyệt kế hoạch",
  "can_cu": ["Căn cứ Luật ...;", "Căn cứ NĐ ..."],
  "noi_dung": [
    { "loai": "dieu", "tieu_de": "Điều 1.", "noi_dung": "..." },
    { "loai": "dieu", "tieu_de": "Điều 2.", "noi_dung": "..." }
  ],
  "ket_thuc_dot": true,
  "noi_nhan": ["- Như Điều 3;", "- Lưu: VT."],
  "chuc_vu_ky": "HIỆU TRƯỞNG",
  "nguoi_ky": "PGS.TS. Nguyễn Văn A"
}
```

### Group B (Công văn)

```json
{
  "co_quan_chu_quan": "...",
  "co_quan_ban_hanh": "...",
  "so_ky_hieu": "Số: 123/CV-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày ... tháng ... năm ...",
  "trich_yeu": "Mời tham gia hội thảo",
  "kinh_gui": "Bộ Thông tin và Truyền thông",
  "noi_dung": ["Đoạn 1...", "Đoạn 2..."],
  "noi_nhan": ["- ...;", "- Lưu: VT."],
  "chuc_vu_ky": "HIỆU TRƯỞNG",
  "nguoi_ky": "..."
}
```

### Group C (Tờ trình)

```json
{
  "co_quan_chu_quan": "...",
  "co_quan_ban_hanh": "...",
  "so_ky_hieu": "Số: 123/TTr-HV",
  "dia_danh": "Hà Nội",
  "ngay_thang": "ngày ... tháng ... năm ...",
  "trich_yeu": "Xin kinh phí tổ chức hội thảo",
  "kinh_gui": ["Thứ trưởng Bộ TT&TT", "Vụ KHTC"],
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

---

## Workflow

### Create a Document

1. Collect info from user (type, agency, number, date, content, signer, recipients)
2. Select document group (A/B/C/D/E/F/G)
3. Draft content in administrative style
4. Generate JSON data file
5. Run appropriate generator script
6. Run `check_nd30.py` to validate

### Check Compliance

```bash
python .claude/skills/nd30-docx/scripts/check_nd30.py --check output/file.docx
```

Checks: paper size, margins, font, QH/TN text, document number, date format, Kính gửi formatting, separator lines, line spacing, Nơi nhận label.

### Recompile Report

```bash
python scripts/split_vbhc.py          # Split VBHC into 3 files
python scripts/compile_report.py       # Compile full report
python scripts/extract_cv_from_report.py  # Extract CV (optional)
python scripts/compile_report.py       # Final recompile
```
