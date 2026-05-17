#!/usr/bin/env python3
"""
Tạo lại 3 văn bản hành chính (Tờ trình, Công văn, Thông báo) từ đầu
với đúng format theo common.js và các generator scripts.
Spacing đã giảm: body=6pt (minimum spec), section headers/titles tighter.
"""
from docx import Document
from docx.shared import Pt, Mm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# CONSTANTS từ common.js / generators
# ============================================================
PAGE_W = 11906  # A4 width DXA
PAGE_H = 16838  # A4 height DXA
MARGIN_TOP = 1134
MARGIN_RIGHT = 851
MARGIN_BOTTOM = 1134
MARGIN_LEFT = 1701

COL_LEFT = 3500   # Header left col DXA
COL_RIGHT = 5571  # Header right col DXA

# Font sizes (half-points)
SZ_QH = 26       # Quốc hiệu: 13pt
SZ_TN = 28       # Tiêu ngữ: 14pt
SZ_CQ = 26       # Cơ quan: 13pt
SZ_SO = 26       # Số: 13pt
SZ_DIA_DANH = 28 # Địa danh: 14pt
SZ_TEN_LOAI = 28 # Tên loại: 14pt
SZ_BODY = 28      # Nội dung: 14pt
SZ_TRICH_YEU = 28 # Trích yếu: 14pt
SZ_TRICH_YEU_CV = 24 # Trích yếu CV V/v: 12pt
SZ_NOI_NHAN_TITLE = 24  # "Nơi nhận:" 12pt
SZ_NOI_NHAN_ITEM = 22   # Items: 11pt
SZ_CHUC_VU = 28   # Chức vụ: 14pt
SZ_CHU_KY = 28     # Chữ ký: 14pt

# Body spacing: 6pt minimum per spec
SP_BODY = 120  # 6pt
SP_LINE = 340  # EXACT line spacing

# Tighter spacing for headers/titles
SP_SECTION_BEFORE = 120  # 6pt
SP_SECTION_AFTER = 60    # 3pt
SP_TITLE_BEFORE = 100    # 5pt
SP_TITLE_AFTER = 60      # 3pt
SP_SUBTITLE_BEFORE = 40  # 2pt
SP_SUBTITLE_AFTER = 80   # 4pt

INDENT_FIRST = 720  # ~1.27cm

# ============================================================
# HELPERS
# ============================================================

def set_run_font(run, size_half_pt, bold=False, italic=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_half_pt / 2)
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
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def add_para(container, text, size=SZ_BODY, bold=False, italic=False,
             alignment=WD_ALIGN_PARAGRAPH.LEFT, no_indent=False,
             spacing_before=None, spacing_after=None, line_spacing=None):
    if hasattr(container, 'add_paragraph'):
        para = container.add_paragraph()
    else:
        if container.paragraphs:
            para = container.paragraphs[0] if not container.paragraphs[0].text else container.add_paragraph()
        else:
            para = container.add_paragraph()

    para.alignment = alignment

    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)

    spacing = OxmlElement('w:spacing')
    sb = spacing_before if spacing_before is not None else SP_BODY
    sa = spacing_after if spacing_after is not None else SP_BODY
    sl = line_spacing if line_spacing is not None else SP_LINE
    spacing.set(qn('w:before'), str(sb))
    spacing.set(qn('w:after'), str(sa))
    spacing.set(qn('w:line'), str(sl))
    spacing.set(qn('w:lineRule'), 'exact')
    old_sp = pPr.find(qn('w:spacing'))
    if old_sp is not None:
        pPr.remove(old_sp)
    pPr.append(spacing)

    if not no_indent:
        ind = OxmlElement('w:ind')
        ind.set(qn('w:firstLine'), str(INDENT_FIRST))
    else:
        ind = OxmlElement('w:ind')
        ind.set(qn('w:firstLine'), '0')
    old_ind = pPr.find(qn('w:ind'))
    if old_ind is not None:
        pPr.remove(old_ind)
    pPr.append(ind)

    run = para.add_run(text)
    set_run_font(run, size, bold, italic)
    return para

def add_spacer(container, before=40, after=0):
    return add_para(container, '', no_indent=True,
                    spacing_before=before, spacing_after=after)

def top_border_paragraph(container, indent_left=1100, indent_right=1100):
    para = container.add_paragraph()
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)

    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '20')
    spacing.set(qn('w:after'), '0')
    pPr.append(spacing)

    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(indent_left))
    ind.set(qn('w:right'), str(indent_right))
    pPr.append(ind)

    pBdr = OxmlElement('w:pBdr')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '2')
    top.set(qn('w:color'), '000000')
    top.set(qn('w:space'), '1')
    pBdr.append(top)
    pPr.append(pBdr)
    return para

def set_cell_no_borders(cell):
    tc = cell._element
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    tcBorders = OxmlElement('w:tcBorders')
    for bn in ['top', 'bottom', 'left', 'right']:
        b = OxmlElement(f'w:{bn}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:color'), 'FFFFFF')
        b.set(qn('w:space'), '0')
        tcBorders.append(b)
    old = tcPr.find(qn('w:tcBorders'))
    if old is not None:
        tcPr.remove(old)
    tcPr.append(tcBorders)

def set_cell_width(cell, width_dxa):
    tc = cell._element
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(width_dxa))
    tcW.set(qn('w:type'), 'dxa')

def set_table_width(table, width_pct=100):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(width_pct))
    tblW.set(qn('w:type'), 'pct')

# ============================================================
# HEADER TABLE
# ============================================================

def create_header_table(doc, co_quan_chu_quan, co_quan_ban_hanh, so_ky_hieu,
                        dia_danh="Hà Nội", ngay_thang=", ngày 05 tháng 05 năm 2025",
                        show_trich_yeu_cv=False, trich_yeu_cv=""):
    table = doc.add_table(rows=2, cols=2)
    set_table_width(table, 100)

    left_top = table.rows[0].cells[0]
    right_top = table.rows[0].cells[1]
    set_cell_width(left_top, COL_LEFT)
    set_cell_width(right_top, COL_RIGHT)
    set_cell_no_borders(left_top)
    set_cell_no_borders(right_top)

    if co_quan_chu_quan:
        add_para(left_top, co_quan_chu_quan, size=SZ_CQ,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)
    add_para(left_top, co_quan_ban_hanh, size=SZ_CQ, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)
    top_border_paragraph(left_top, indent_left=1350, indent_right=1350)

    add_para(right_top, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=SZ_QH, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)
    add_para(right_top, "Độc lập - Tự do - Hạnh phúc", size=SZ_TN, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)
    top_border_paragraph(right_top, indent_left=1100, indent_right=1100)

    left_bot = table.rows[1].cells[0]
    right_bot = table.rows[1].cells[1]
    set_cell_width(left_bot, COL_LEFT)
    set_cell_width(right_bot, COL_RIGHT)
    set_cell_no_borders(left_bot)
    set_cell_no_borders(right_bot)

    add_para(left_bot, so_ky_hieu, size=SZ_SO,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)

    if show_trich_yeu_cv and trich_yeu_cv:
        add_para(left_bot, trich_yeu_cv, size=SZ_TRICH_YEU_CV,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
                 spacing_before=60, spacing_after=0)

    add_para(right_bot, f"{dia_danh}{ngay_thang}", size=SZ_DIA_DANH, italic=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)

    return table

# ============================================================
# FOOTER TABLE
# ============================================================

def create_footer_table(doc, noi_nhan_items, chuc_vu_ky, nguoi_ky,
                        cap_ky="", chuc_vu_cap_tren=""):
    table = doc.add_table(rows=1, cols=2)
    set_table_width(table, 100)

    left = table.rows[0].cells[0]
    right = table.rows[0].cells[1]
    set_cell_width(left, 5500)
    set_cell_width(right, 4500)
    set_cell_no_borders(left)
    set_cell_no_borders(right)

    add_para(left, "Nơi nhận:", size=SZ_NOI_NHAN_TITLE, bold=True, italic=True,
             no_indent=True, spacing_before=60, spacing_after=30)
    for item in noi_nhan_items:
        add_para(left, item, size=SZ_NOI_NHAN_ITEM,
                 no_indent=True, spacing_before=0, spacing_after=10)

    valid_prefixes = ["KT", "TL", "TUQ", "TM", "Q"]
    if cap_ky.upper() in valid_prefixes and chuc_vu_cap_tren:
        add_para(right, f"{cap_ky.upper()}. {chuc_vu_cap_tren}", size=SZ_CHUC_VU,
                 bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
                 spacing_before=0, spacing_after=30)

    add_para(right, chuc_vu_ky, size=SZ_CHUC_VU, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=0, spacing_after=30)
    add_para(right, "(Ký, ghi rõ họ tên)", size=SZ_CHU_KY, italic=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=20, spacing_after=200)
    add_para(right, nguoi_ky, size=SZ_CHU_KY, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=0, spacing_after=60)

    return table

# ============================================================
# TỜ TRÌNH
# ============================================================

def create_to_trinh(doc):
    create_header_table(
        doc,
        co_quan_chu_quan="BỘ THÔNG TIN VÀ TRUYỀN THÔNG",
        co_quan_ban_hanh="HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG",
        so_ky_hieu="Số: 01/2025/TTr-PTIT",
        dia_danh="Hà Nội",
        ngay_thang=", ngày 05 tháng 05 năm 2025"
    )

    add_spacer(doc, before=80, after=0)

    add_para(doc, "TỜ TRÌNH", size=SZ_TEN_LOAI, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_TITLE_BEFORE, spacing_after=SP_TITLE_AFTER)

    add_para(doc, "V/v đề nghị phê duyệt và cấp kinh phí đề tài khoa học và công nghệ",
             size=SZ_TRICH_YEU, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SUBTITLE_BEFORE, spacing_after=SP_SUBTITLE_AFTER)

    add_para(doc, "Kính gửi: Ban Giám đốc Học viện Công nghệ Bưu chính Viễn thông",
             size=SZ_BODY, italic=True, no_indent=True,
             spacing_before=60, spacing_after=80)

    add_para(doc, "Căn cứ Luật Khoa học và Công nghệ ngày 18/6/2013;",
             size=SZ_BODY, italic=True, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư;",
             size=SZ_BODY, italic=True, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Căn cứ Quy chế quản lý khoa học và công nghệ của Học viện Công nghệ Bưu chính Viễn thông;",
             size=SZ_BODY, italic=True, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_spacer(doc, before=30, after=30)

    add_para(doc, "I. SỰ CẦN THIẾT VÀ CĂN CỨ ĐỀ XUẤT",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Gia Lai là tỉnh có tỷ lệ đồng bào dân tộc thiểu số chiếm 44,7% dân số, trong đó người Jrai và Bahnar chiếm tỷ lệ lớn. Tiếng Jrai là ngôn ngữ được sử dụng rộng rãi trong sinh hoạt cộng đồng nhưng đang đối mặt với nguy cơ mai một do quá trình đô thị hóa và thiếu hụt tài liệu giảng dạy chuẩn hóa. Việc ứng dụng trí tuệ nhân tạo (AI) vào bảo tồn và giảng dạy tiếng Jrai không chỉ giúp lưu giữ di sản ngôn ngữ mà còn thúc đẩy chuyển đổi số trong giáo dục tại tỉnh Gia Lai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Căn cứ Nghị quyết số 57-NQ/TW ngày 22/12/2024 của Bộ Chính trị về đột phá phát triển khoa học, công nghệ, đổi mới sáng tạo và chuyển đổi số quốc gia; Căn cứ Quyết định số 1131/QĐ-BTTTT ngày 15/4/2025 của Bộ Thông tin và Truyền thông về Kế hoạch chuyển đổi số ngành TT&TT năm 2025, Khoa Công nghệ Thông tin – PTIT đề xuất thực hiện đề tài nghiên cứu ứng dụng AI phục vụ bảo tồn và giảng dạy tiếng Jrai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "II. NỘI DUNG ĐỀ TÀI",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Tên đề tài: Nền tảng AI hỗ trợ bảo tồn và dạy học tiếng Jrai (định hướng mở rộng tiếng Bahnar) phục vụ chuyển đổi số tỉnh Gia Lai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Loại hình: Đề tài nghiên cứu ứng dụng và phát triển công nghệ.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Thời gian thực hiện: 12 tháng kể từ ngày được phê duyệt.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Đơn vị chủ trì: Khoa Công nghệ Thông tin – Học viện Công nghệ Bưu chính Viễn thông.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Chủ nhiệm đề tài: TS. Phạm Vũ Minh Tú.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "III. DỰ KIẾN KINH PHÍ",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Tổng kinh phí đề nghị phê duyệt: 2.100.000.000 đồng (Hai tỷ một trăm triệu đồng) trong thời gian 12 tháng, từ nguồn ngân sách khoa học và công nghệ của Học viện.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "IV. ĐỀ NGHỊ",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Khoa Công nghệ Thông tin kính đề nghị Ban Giám đốc Học viện Công nghệ Bưu chính Viễn thông xem xét, phê duyệt đề tài khoa học và công nghệ nêu trên và cấp kinh phí để triển khai theo quy định.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_spacer(doc, before=80, after=0)

    create_footer_table(
        doc,
        noi_nhan_items=[
            "- Như kính gửi (để phê duyệt);",
            "- Phòng Khoa học, Công nghệ và HTQT (để phối hợp);",
            "- Lưu: VT, PTIT."
        ],
        chuc_vu_ky="TRƯỞNG KHOA CÔNG NGHỆ THÔNG TIN",
        nguoi_ky="TS. Nguyễn Minh Tuấn"
    )

# ============================================================
# CÔNG VĂN
# ============================================================

def create_cong_van(doc):
    create_header_table(
        doc,
        co_quan_chu_quan="",
        co_quan_ban_hanh="HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG",
        so_ky_hieu="Số: 02/2025/CV-PTIT",
        show_trich_yeu_cv=True,
        trich_yeu_cv="V/v đề nghị phối hợp triển khai đề tài khoa học và công nghệ",
        dia_danh="Hà Nội",
        ngay_thang=", ngày 12 tháng 05 năm 2025"
    )

    add_spacer(doc, before=40, after=40)

    add_para(doc, "Kính gửi: Sở Khoa học và Công nghệ tỉnh Gia Lai",
             size=SZ_BODY, no_indent=True,
             spacing_before=40, spacing_after=60)

    add_spacer(doc, before=30, after=30)

    paragraphs_text = [
        "Học viện Công nghệ Bưu chính Viễn thông (PTIT) đang triển khai đề tài nghiên cứu khoa học và công nghệ: \"Nền tảng AI hỗ trợ bảo tồn và dạy học tiếng Jrai (định hướng mở rộng tiếng Bahnar) phục vụ chuyển đổi số tỉnh Gia Lai\", do TS. Phạm Vũ Minh Tú làm chủ nhiệm, thuộc Khoa Công nghệ Thông tin.",
        "Đây là đề tài gắn trực tiếp với địa bàn tỉnh Gia Lai, nhằm xây dựng nền tảng AI phục vụ bảo tồn ngôn ngữ và thúc đẩy chuyển đổi số trong giáo dục. Đề tài bao gồm: thu thập và chuẩn hóa ngữ liệu tiếng Jrai; phát triển mô hình AI nhận diện phát âm, hội thoại tự động, chuyển văn bản thành giọng nói; xây dựng ứng dụng di động và học liệu số cho học sinh, giáo viên và cộng đồng.",
        "Học viện Công nghệ Bưu chính Viễn thông trân trọng đề nghị Sở Khoa học và Công nghệ tỉnh Gia Lai xem xét phối hợp trong các nội dung sau:",
    ]
    for text in paragraphs_text:
        add_para(doc, text, size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    list_items = [
        "1. Hỗ trợ kết nối với các cơ sở giáo dục, cộng đồng dân tộc và cơ quan quản lý tại địa phương phục vụ thu thập ngữ liệu và thử nghiệm nền tảng.",
        "2. Tham gia góp ý định hướng chính sách và đánh giá kết quả thử nghiệm nền tảng AI tiếng Jrai tại tỉnh Gia Lai.",
        "3. Xem xét hỗ trợ kinh phí phối hợp trong khuôn khổ chương trình khoa học và công nghệ của tỉnh theo thẩm quyền.",
    ]
    for item in list_items:
        add_para(doc, item, size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "Học viện Công nghệ Bưu chính Viễn thông rất mong nhận được phản hồi từ Quý Sở. Đầu mối liên hệ: TS. Phạm Vũ Minh Tú – ĐT: 090 459 2738, Email: tuptm@ptit.edu.vn.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_spacer(doc, before=40, after=0)

    create_footer_table(
        doc,
        noi_nhan_items=[
            "- Như kính gửi;",
            "- Phòng KH,CN&HTQT (để theo dõi);",
            "- Lưu: VT, PTIT."
        ],
        cap_ky="KT",
        chuc_vu_cap_tren="GIÁM ĐỐC",
        chuc_vu_ky="PHÓ GIÁM ĐỐC",
        nguoi_ky="PGS.TS. Trần Quang Minh"
    )

# ============================================================
# THÔNG BÁO
# ============================================================

def create_thong_bao(doc):
    create_header_table(
        doc,
        co_quan_chu_quan="",
        co_quan_ban_hanh="HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG",
        so_ky_hieu="Số: 03/2025/TB-PTIT",
        dia_danh="Hà Nội",
        ngay_thang=", ngày 20 tháng 05 năm 2025"
    )

    add_spacer(doc, before=80, after=0)

    add_para(doc, "THÔNG BÁO", size=SZ_TEN_LOAI, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_TITLE_BEFORE, spacing_after=SP_TITLE_AFTER)

    add_para(doc, "V/v Triển khai đề tài khoa học và công nghệ",
             size=SZ_TRICH_YEU, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SUBTITLE_BEFORE, spacing_after=SP_SUBTITLE_AFTER)

    add_para(doc, "Căn cứ Tờ trình số 01/2025/TTr-PTIT ngày 05/5/2025 của Khoa Công nghệ Thông tin về việc đề nghị phê duyệt và cấp kinh phí đề tài;",
             size=SZ_BODY, italic=True, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "Căn cứ Kế hoạch khoa học và công nghệ năm 2025 của Học viện Công nghệ Bưu chính Viễn thông;",
             size=SZ_BODY, italic=True, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_spacer(doc, before=30, after=30)

    add_para(doc, "Học viện Công nghệ Bưu chính Viễn thông thông báo về kế hoạch triển khai đề tài khoa học và công nghệ như sau:",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "1. Tên đề tài",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Nền tảng AI hỗ trợ bảo tồn và dạy học tiếng Jrai (định hướng mở rộng tiếng Bahnar) phục vụ chuyển đổi số tỉnh Gia Lai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "2. Đơn vị và cá nhân phụ trách",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "- Đơn vị chủ trì: Khoa Công nghệ Thông tin – Học viện Công nghệ Bưu chính Viễn thông.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Chủ nhiệm đề tài: TS. Phạm Vũ Minh Tú – ĐT: 090 459 2738.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Đơn vị phối hợp: Phòng Khoa học, Công nghệ và HTQT; Sở Khoa học và Công nghệ tỉnh Gia Lai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "3. Thời gian và kế hoạch triển khai",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "- Thời gian thực hiện: 12 tháng, từ tháng 6/2025 đến tháng 5/2026.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Tháng 6–8/2025: Khảo sát thực trạng; xây dựng và chuẩn hóa kho ngữ liệu số tiếng Jrai.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Tháng 9–12/2025: Phát triển các mô-đun AI (nhận diện phát âm, hội thoại, chuyển văn bản thành giọng nói).",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Tháng 1–3/2026: Xây dựng ứng dụng di động; tích hợp mô-đun AI và học liệu số.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc, "- Tháng 4–5/2026: Thử nghiệm tại cơ sở giáo dục tại tỉnh Gia Lai; hoàn thiện và nghiệm thu.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "4. Kinh phí thực hiện",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Tổng kinh phí được phê duyệt: 2.100.000.000 đồng (Hai tỷ một trăm triệu đồng).",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc, "5. Yêu cầu",
             size=SZ_BODY, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True,
             spacing_before=SP_SECTION_BEFORE, spacing_after=SP_SECTION_AFTER)
    add_para(doc, "Các đơn vị liên quan thuộc Học viện phối hợp chặt chẽ với Khoa Công nghệ Thông tin trong suốt quá trình triển khai. Mọi thắc mắc xin liên hệ Phòng Khoa học, Công nghệ và HTQT.",
             size=SZ_BODY, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_spacer(doc, before=80, after=0)
    add_para(doc, "./.", size=SZ_BODY, bold=True,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, no_indent=True)

    add_spacer(doc, before=80, after=0)

    create_footer_table(
        doc,
        noi_nhan_items=[
            "- Các Khoa, Phòng, Ban liên quan (để thực hiện);",
            "- Sở KH&CN tỉnh Gia Lai (để phối hợp);",
            "- Lưu: VT, PTIT."
        ],
        cap_ky="KT",
        chuc_vu_cap_tren="GIÁM ĐỐC",
        chuc_vu_ky="PHÓ GIÁM ĐỐC",
        nguoi_ky="PGS.TS. Trần Quang Minh"
    )

# ============================================================
# MAIN
# ============================================================

def main():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Emu(PAGE_W * 635)
    section.page_height = Emu(PAGE_H * 635)
    section.top_margin = Emu(MARGIN_TOP * 635)
    section.bottom_margin = Emu(MARGIN_BOTTOM * 635)
    section.left_margin = Emu(MARGIN_LEFT * 635)
    section.right_margin = Emu(MARGIN_RIGHT * 635)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        style.element.append(rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    create_to_trinh(doc)
    doc.add_page_break()

    create_cong_van(doc)
    doc.add_page_break()

    create_thong_bao(doc)

    output_path = 'output/VanBanHanhChinh_PTIT_Rewritten.docx'
    doc.save(output_path)
    print(f'Đã tạo: {output_path}')

if __name__ == '__main__':
    main()
