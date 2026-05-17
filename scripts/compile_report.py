#!/usr/bin/env python3
"""Compile tất cả bài thi SKD1103 vào 1 file .docx duy nhất"""
import os
from docx import Document
from docx.shared import Pt, Mm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

def set_font(paragraph, name='Times New Roman', size=Pt(13), bold=False, italic=False):
    """Set font cho paragraph"""
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

def add_heading_styled(doc, text, level=1):
    """Thêm heading với font TNR"""
    p = doc.add_heading(text, level=level)
    set_font(p, size=Pt(14) if level == 1 else Pt(13), bold=True)
    return p

def add_para(doc, text, bold=False, italic=False, size=Pt(13), align=None):
    """Thêm paragraph với font TNR"""
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
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
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    return p

def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(docx.enum.text.WD_BREAK.PAGE)

def create_cover_page(doc):
    """Tạo trang bìa"""
    # Khoảng cách
    for _ in range(3):
        doc.add_paragraph()

    # Tên trường
    p = add_para(doc, 'HỌC VIỆN CÔNG NGHỆ', bold=True, size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'BƯU CHÍNH VIỄN THÔNG', bold=True, size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # Tên khoa/viện
    p = add_para(doc, 'VIỆN KINH TẾ BƯU ĐIỆN', bold=True, size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()
    doc.add_paragraph()

    # Tên đề thi
    p = add_para(doc, 'BÁO CÁO CUỐI KỲ', bold=True, size=Pt(18), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    p = add_para(doc, 'HỌC PHẦN: KỸ NĂNG TẠO LẬP VĂN BẢN TIẾNG VIỆT', bold=True, size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'HỆ ĐẠI HỌC: CHÍNH QUY', bold=True, size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()
    doc.add_paragraph()

    # Tình huống
    p = add_para(doc, 'Đề tài:', bold=True, size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'Nền tảng AI hỗ trợ bảo tồn và giảng dạy tiếng Jrai', bold=True, size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, '(Mở rộng sang tiếng Bahnar)', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'cho chuyển đổi số tỉnh Gia Lai', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)

    for _ in range(3):
        doc.add_paragraph()

    # Thông tin nhóm
    p = add_para(doc, 'Nhóm: ...', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'Lớp: ...', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)
    p = add_para(doc, 'Giảng viên hướng dẫn: ...', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    p = add_para(doc, 'Hà Nội, tháng 5 năm 2026', size=Pt(13), align=WD_ALIGN_PARAGRAPH.CENTER)

def create_toc(doc):
    """Tạo mục lục"""
    add_heading_styled(doc, 'MỤC LỤC', level=1)

    toc_items = [
        ('Lời mở đầu', '3'),
        ('PHẦN I: HỆ THỐNG LÝ THUYẾT', '4'),
        ('  Câu 1: Các thành phần thể thức chính của văn bản hành chính', '4'),
        ('  Câu 2: Yêu cầu về CV và lưu ý khi soạn thảo email ứng tuyển', '6'),
        ('PHẦN II: THỰC HÀNH SOẠN THẢO', '8'),
        ('  Nhiệm vụ 1: Bộ 03 văn bản hành chính', '8'),
        ('    1.1. Tờ trình', '8'),
        ('    1.2. Công văn', '9'),
        ('    1.3. Thông báo', '10'),
        ('  Nhiệm vụ 2: CV cá nhân', '11'),
        ('  Nhiệm vụ 3: Thư ứng tuyển và Email', '12'),
        ('Kết luận', '14'),
        ('Tài liệu tham khảo', '15'),
        ('Bảng đánh giá đóng góp', '16'),
        ('Slide thuyết trình', '17'),
    ]

    for item, page in toc_items:
        p = doc.add_paragraph()
        run = p.add_run(f'{item}')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        # Thêm tab và số trang
        run2 = p.add_run(f'\t{page}')
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(13)

def create_foreword(doc):
    """Tạo lời mở đầu"""
    add_heading_styled(doc, 'LỜI MỞ ĐẦU', level=1)

    add_para(doc,
        'Trong bối cảnh chuyển đổi số đang diễn ra mạnh mẽ trên toàn quốc, việc nắm vững '
        'các kỹ năng tạo lập văn bản tiếng Việt là yêu cầu thiết yếu đối với sinh viên đại học. '
        'Học phần Kỹ năng tạo lập văn bản tiếng Việt giúp sinh viên hiểu rõ các quy định '
        'về thể thức và kỹ thuật trình bày văn bản hành chính theo Nghị định 30/2020/NĐ-CP, '
        'đồng thời phát triển kỹ năng soạn thảo CV, thư ứng tuyển và email chuyên nghiệp.'
    )

    add_para(doc,
        'Báo cáo này được thực hiện dựa trên tình huống giả định về dự án "Nền tảng AI '
        'hỗ trợ bảo tồn và giảng dạy tiếng Jrai (mở rộng sang tiếng Bahnar) cho chuyển đổi số '
        'tỉnh Gia Lai" tại Học viện Công nghệ Bưu chính Viễn thông. Nội dung báo cáo bao gồm '
        'hai phần chính: hệ thống lý thuyết về các thành phần thể thức văn bản và yêu cầu CV/email, '
        'cùng với thực hành soạn thảo bộ văn bản hành chính, CV cá nhân, thư ứng tuyển và email.'
    )

    add_para(doc,
        'Báo cáo được thực hiện trong thời gian 10 ngày (từ ngày 07/5/2026 đến ngày 17/5/2026). '
        'Nhóm xin chân thành cảm ơn sự hướng dẫn của giảng viên và sự hỗ trợ của các thành viên '
        'trong nhóm trong quá trình thực hiện.'
    )

def create_theory(doc):
    """Tạo phần lý thuyết"""
    add_heading_styled(doc, 'PHẦN I: HỆ THỐNG LÝ THUYẾT', level=1)

    # ==================== CÂU 1 ====================
    add_heading_styled(doc, 'Câu 1: Các thành phần thể thức chính của văn bản hành chính theo Nghị định 30/2020/NĐ-CP', level=2)

    add_para(doc,
        'Theo Phụ lục I của Nghị định 30/2020/NĐ-CP, văn bản hành chính phải đảm bảo '
        'đầy đủ chín thành phần thể thức chính. Mỗi thành phần có vai trò riêng trong việc '
        'xác định tính hợp pháp, tính chính thống và nội dung của văn bản.'
    )

    # 1. Quốc hiệu, tiêu ngữ
    add_para(doc, '1. Quốc hiệu, tiêu ngữ', bold=True)

    add_para(doc,
        'Quốc hiệu và tiêu ngữ là thành phần thể hiện chủ quyền quốc gia, đặt ở vị trí '
        'trang trọng nhất của văn bản. Quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" '
        'được trình bày bằng chữ in hoa, cỡ chữ 12-13pt, đứng, đậm, canh giữa phía trên bên phải. '
        'Tiêu ngữ "Độc lập - Tự do - Hạnh phúc" đặt ngay dưới quốc hiệu, cỡ chữ 13-14pt, '
        'in thường (chữ đầu mỗi cụm từ viết hoa), đứng, đậm, canh giữa. Giữa các cụm từ '
        'sử dụng gạch nối (-) có cách chữ. Dưới tiêu ngữ có đường kẻ ngang nét liền, '
        'dài bằng dòng chữ. Khoảng cách giữa quốc hiệu và tiêu ngữ là dòng đơn, '
        'cỡ chữ thống nhất (quốc hiệu 12pt thì tiêu ngữ 13pt, quốc hiệu 13pt thì tiêu ngữ 14pt).'
    )

    # 2. Tên cơ quan chủ quản
    add_para(doc, '2. Tên cơ quan chủ quản', bold=True)

    add_para(doc,
        'Tên cơ quan chủ quản là tên cơ quan cấp trên trực tiếp của cơ quan ban hành văn bản, '
        'trình bày bằng chữ in hoa, cỡ chữ 12-13pt, đứng, không đậm, canh giữa phía trên bên trái. '
        'Tên cơ quan chủ quản giúp xác định mối quan hệ hệ thống hành chính và tính hợp pháp '
        'của văn bản trong hệ thống quản lý nhà nước.'
    )

    # 3. Tên cơ quan ban hành văn bản
    add_para(doc, '3. Tên cơ quan ban hành văn bản', bold=True)

    add_para(doc,
        'Tên cơ quan ban hành văn bản là tên đơn vị trực tiếp ra văn bản, trình bày bằng chữ '
        'in hoa, cỡ chữ 12-13pt, đứng, đậm, canh giữa, đặt ngay dưới tên cơ quan chủ quản. '
        'Dưới tên cơ quan ban hành có đường kẻ ngang nét liền. Tên cơ quan ban hành thể hiện '
        'trách nhiệm pháp lý của đơn vị đối với nội dung văn bản.'
    )

    # 4. Số, ký hiệu văn bản
    add_para(doc, '4. Số, ký hiệu văn bản', bold=True)

    add_para(doc,
        'Số và ký hiệu văn bản là mã định danh duy nhất của văn bản, bao gồm số thứ tự '
        'và ký hiệu theo quy định. Số được viết bằng chữ số Ả Rập, cỡ chữ 13pt, đứng, '
        'canh giữa dưới tên cơ quan ban hành. Từ "Số" viết thường, sau dấu hai chấm. '
        'Số nhỏ hơn 10 phải thêm số 0 phía trước (01, 02...09). Ký hiệu bao gồm các nhóm '
        'viết tắt cách nhau bằng gạch chéo (/) và gạch nối (-). Ví dụ: "Số: 01/TTr-PTIT" '
        'trong đó 01 là số thứ tự, TTr là viết tắt tờ trình, PTIT là viết tắt cơ quan.'
    )

    # 5. Địa danh, ngày tháng ban hành
    add_para(doc, '5. Địa danh, ngày tháng ban hành', bold=True)

    add_para(doc,
        'Địa danh và ngày tháng ban hành được trình bày bằng chữ in thường, cỡ chữ 13-14pt, '
        'nghiêng, canh giữa. Địa danh là nơi ban hành văn bản, chữ cái đầu viết hoa, '
        'sau địa danh có dấu phẩy. Ngày và tháng viết đầy đủ bằng chữ số Ả Rập, '
        'ngày nhỏ hơn 10 phải thêm số 0 phía trước (ngày 05), tháng 1 và tháng 2 cũng thêm số 0 '
        '(tháng 01, tháng 02). Dạng trình bày: "[Địa danh], ngày XX tháng XX năm XXXX".'
    )

    # 6. Tên loại và trích yếu nội dung
    add_para(doc, '6. Tên loại và trích yếu nội dung', bold=True)

    add_para(doc,
        'Tên loại văn bản (như TỜ TRÌNH, CÔNG VĂN, THÔNG BÁO) được trình bày bằng chữ '
        'in hoa, cỡ chữ 13-14pt, đứng, đậm, canh giữa. Trích yếu nội dung tóm tắt ngắn gọn '
        'nội dung chính của văn bản, trình bày bằng chữ in thường, cỡ chữ 13-14pt, đứng, đậm, '
        'canh giữa, đặt ngay dưới tên loại. Dưới trích yếu có đường kẻ ngang dài 1/3-1/2 dòng chữ, '
        'cân đối. Đối với công văn, trích yếu được trình bày dạng "V/v ..." cỡ chữ 12-13pt, '
        'in thường, đứng, canh giữa dưới số ký hiệu.'
    )

    # 7. Nội dung văn bản
    add_para(doc, '7. Nội dung văn bản', bold=True)

    add_para(doc,
        'Nội dung văn bản là phần chính trình bày thông tin, được trình bày bằng chữ in thường, '
        'cỡ chữ 13-14pt, đứng, canh đều hai bên lề. Lùi đầu dòng 1cm hoặc 1.27cm. '
        'Nội dung có thể được chia thành các phần, chương, mục, tiểu mục, điều, khoản, điểm '
        'theo thứ tự hợp lý. Tiêu đề phần/chương in hoa, đậm, canh giữa. '
        'Tiêu đề mục sử dụng số La Mã (I, II, III), tiêu đề điều sử dụng số Ả Rập.'
    )

    # 8. Chức vụ, họ tên, chữ ký
    add_para(doc, '8. Chức vụ, họ tên, chữ ký người có thẩm quyền', bold=True)

    add_para(doc,
        'Khối chữ ký đặt ở cuối văn bản, phía bên phải. Chức vụ người ký được trình bày '
        'bằng chữ in hoa, cỡ chữ 13-14pt, đứng, đậm, canh giữa. Dòng "(Ký, ghi rõ họ tên)" '
        'trình bày bằng chữ nghiêng, canh giữa dưới chức vụ. Họ tên người ký trình bày '
        'bằng chữ đứng, đậm, không ghi học hàm/học vi. Các trường hợp ký thay được quy định: '
        'TM. (thay mặt), KT. (ký thay), Q. (giao quyền), TL. (thừa lệnh), TUQ. (thừa ủy quyền).'
    )

    # 9. Dấu của cơ quan
    add_para(doc, '9. Dấu của cơ quan', bold=True)

    add_para(doc,
        'Dấu của cơ quan được đóng theo quy định của pháp luật, thường đặt đè lên tên '
        'cơ quan ban hành và một phần chữ ký. Dấu tròn dùng cho cơ quan nhà nước, '
        'dấu chức danh dùng cho cá nhân có thẩm quyền. Việc đóng dấu xác nhận tính hợp pháp '
        'và chính thức của văn bản. Không phải văn bản nào cũng bắt buộc có dấu, '
        'tùy theo quy định của từng loại văn bản và thẩm quyền ban hành.'
    )

    # ==================== CÂU 2 ====================
    add_heading_styled(doc, 'Câu 2: Yêu cầu về nội dung và hình thức của CV ứng tuyển và những lưu ý khi soạn thảo email ứng tuyển', level=2)

    add_para(doc, 'I. Yêu cầu về CV ứng tuyển', bold=True, size=Pt(13))

    add_para(doc,
        'CV (Curriculum Vitae) là bản tóm tắt thông tin cá nhân, học vấn, kinh nghiệm làm việc '
        'và kỹ năng của ứng viên, đóng vai trò quan trọng trong quá trình tuyển dụng. '
        'CV giúp nhà tuyển dụng đánh giá nhanh sự phù hợp của ứng viên với vị trí tuyển dụng.'
    )

    add_para(doc, '1. Các thành phần chính của CV', bold=True)

    add_para(doc,
        'Một bản CV chuyên nghiệp cần bao gồm các thành phần sau:'
    )

    components = [
        ('Tiêu đề CV', 'Có nhiều cách đặt tiêu đề như "CV - Họ tên", "Hồ sơ ứng tuyển - Họ tên", hoặc "CV - Họ tên - Vị trí ứng tuyển". Tiêu đề nên chứa từ khóa liên quan đến vị trí để tối ưu hệ thống ATS.'),
        ('Thông tin cá nhân', 'Bao gồm họ tên in hoa đậm, email chuyên nghiệp (không sử dụng nickname), số điện thoại đầy đủ mã quốc gia, địa chỉ (chỉ ghi thành phố), và liên kết LinkedIn/portfolio nếu có.'),
        ('Mục tiêu nghề nghiệp', 'Phân chia rõ ràng thành mục tiêu ngắn hạn (1-3 năm) và dài hạn (3-5 năm). Ngắn gọn, cụ thể, chứa từ khóa từ mô tả công việc, nhấn mạnh khả năng đóng góp cho công ty.'),
        ('Kinh nghiệm làm việc', 'Sắp xếp theo thứ tự thời gian ngược (mới nhất ở trên). Mỗi mục ghi rõ tên công ty, vị trí, thời gian, và mô tả trách nhiệm cùng thành tích đạt được, ưu tiên lượng hóa bằng số liệu.'),
        ('Trình độ học vấn', 'Bao gồm tên trường, ngành học, thời gian học. Nếu có GPA cao, học bổng hoặc chứng chỉ liên quan nên liệt kê.'),
        ('Kỹ năng', 'Chia thành ba nhóm: kỹ năng chuyên môn (liên quan trực tiếp đến công việc), kỹ năng mềm (giao tiếp, làm việc nhóm, lãnh đạo), và kỹ năng sử dụng công cụ (phần mềm, công nghệ).'),
        ('Dự án', 'Mô tả các dự án đã tham gia với tiêu đề, vai trò, kỹ năng sử dụng và kết quả đạt được. Đặc biệt quan trọng với ứng viên ít kinh nghiệm làm việc.'),
        ('Hoạt động ngoại khóa', 'Liệt kê các hoạt động liên quan đến công việc ứng tuyển, thể hiện sự năng động và kỹ năng mềm.'),
        ('Sở thích', 'Chỉ ghi nếu liên quan đến công việc, cụ thể hóa (ví dụ: "đọc sách về kinh doanh" thay vì chỉ "đọc sách"). Nếu không liên quan nên bỏ qua.'),
        ('Người tham chiếu', 'Chọn người phù hợp (cấp trên, đồng nghiệp, giáo viên), xin phép trước, cung cấp đầy đủ SĐT và email công ty. Đặt ở cuối CV.'),
    ]

    for i, (name, desc) in enumerate(components, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {name}: ')
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run2 = p.add_run(desc)
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(13)

    add_para(doc, '2. Định dạng và tối ưu ATS', bold=True)

    add_para(doc,
        'Hệ thống ATS (Applicant Tracking System) là công cụ sàng lọc hồ sơ tự động. '
        'Để CV được ATS nhận diện đúng, cần tuân thủ: sử dụng tiêu đề chuẩn cho từng phần, '
        'nhúng từ khóa từ mô tả công việc vào nội dung CV, tránh sử dụng bảng biểu phức tạp '
        'hoặc hình ảnh, sử dụng phông chữ phổ biến (Times New Roman, Arial). '
        'Định dạng file nên là PDF (giữ nguyên layout) hoặc Word (.docx, dễ chỉnh sửa). '
        'Tên file đặt theo cấu trúc: "CV_HoTen_ViTri.pdf" hoặc "CV - Họ tên - Vị trí.pdf".'
    )

    add_para(doc, '3. Các lỗi thường gặp khi viết CV', bold=True)

    errors = [
        'Sử dụng email không chuyên nghiệp (biệt danh, ký tự đặc biệt quá nhiều, email trường học cũ)',
        'Mục tiêu nghề nghiệp chung chung, không cụ thể, không phân chia ngắn hạn/dài hạn',
        'Thời gian kinh nghiệm không sắp xếp theo thứ tự thời gian ngược',
        'Kinh nghiệm làm việc không lượng hóa thành tích bằng số liệu',
        'Lỗi chính tả và ngữ pháp trong CV',
        'CV quá dài (vượt quá 2 trang A4)',
        'Không tùy chỉnh CV theo từng vị trí ứng tuyển',
        'Thiết kế quá phức tạp, nhiều màu sắc không cần thiết',
        'Bảo mật thông tin cá nhân quá mức (ghi CMND, số tài khoản ngân hàng)',
    ]
    for i, err in enumerate(errors, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {err}')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

    add_para(doc, 'II. Lưu ý khi soạn thảo email ứng tuyển', bold=True, size=Pt(13))

    add_para(doc,
        'Email ứng tuyển là phương thức giao tiếp chính thức giữa ứng viên và nhà tuyển dụng, '
        'được sử dụng để gửi CV và thư ứng tuyển. Email chuyên nghiệp giúp tạo ấn tượng '
        'tích cực ban đầu.'
    )

    add_para(doc, '1. Cấu trúc email ứng tuyển', bold=True)

    email_parts = [
        ('Dòng chủ đề', 'Rõ ràng, chứa từ khóa. Định dạng: "Họ tên - Vị trí ứng tuyển - Công ty - Ngày". Ví dụ: "Nguyễn Văn A - Nhân viên Marketing - Công ty ABC - 17/05/2026".'),
        ('Người nhận', 'Email chính của phòng nhân sự hoặc người tuyển dụng. CC người liên quan nếu có.'),
        ('Lời chào', 'Phù hợp với ngữ cảnh: "Kính gửi Quý Công ty ABC," hoặc "Kính gửi Anh/Chị [Tên]," nếu biết tên người tuyển dụng.'),
        ('Nội dung', 'Gồm 3 phần: giới thiệu (họ tên, vị trí ứng tuyển, nguồn biết tin), lý do ứng tuyển + kỹ năng phù hợp, kết luận + nhấn mạnh đóng góp cho công ty.'),
        ('Lời kết', 'Sử dụng lời kết lịch sự: "Trân trọng cảm ơn," hoặc "Trân trọng,".'),
        ('File đính kèm', 'Đính kèm CV và thư ứng tuyển (nếu có). Đặt tên file khoa học: "CV_HoTen_ViTri.pdf", "ThuUngTuyen_HoTen_ViTri.pdf". Kiểm tra đã đính kèm trước khi gửi.'),
        ('Chữ ký', 'Bao gồm họ tên, vị trí, email, số điện thoại, LinkedIn (nếu có).'),
    ]

    for i, (name, desc) in enumerate(email_parts, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {name}: ')
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run2 = p.add_run(desc)
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(13)

    add_para(doc, '2. So sánh email xin việc và thư ứng tuyển', bold=True)

    add_para(doc,
        'Email xin việc và thư ứng tuyển (cover letter) có những điểm khác biệt cơ bản. '
        'Email xin việc là nội dung inline trong email, ngắn gọn (150-200 từ), dùng để giới thiệu '
        'và hướng dẫn đọc CV. Thư ứng tuyển là file riêng (.docx/.pdf), chi tiết hơn (1 trang A4), '
        'giải thích chi tiết lý do ứng tuyển và chứng minh sự phù hợp. '
        'Khi vị trí cấp cao hoặc có yêu cầu, nên gửi cả email và thư ứng tuyển.'
    )

    add_para(doc, '3. Các lỗi thường gặp khi viết email', bold=True)

    email_errors = [
        'Thiếu dòng chủ đề hoặc chủ đề không rõ ràng',
        'Nội dung email quá dài dòng, không súc tích',
        'Lỗi chính tả và ngữ pháp',
        'Quên đính kèm file (nhắc đến CV nhưng không gửi)',
        'Đặt tên file không khoa học hoặc không đúng yêu cầu',
        'Sử dụng email không chuyên nghiệp',
        'Không phản hồi email của nhà tuyển dụng hoặc phản hồi quá chậm',
    ]
    for i, err in enumerate(email_errors, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {err}')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

def create_practice(doc):
    """Tạo phần thực hành"""
    doc.add_page_break()
    add_heading_styled(doc, 'PHẦN II: THỰC HÀNH SOẠN THẢO VĂN BẢN HÀNH CHÍNH, HỒ SƠ ỨNG TUYỂN', level=1)

    # ==================== NHIỆM VỤ 1 ====================
    add_heading_styled(doc, 'Nhiệm vụ 1: Bộ 03 văn bản hành chính', level=2)

    add_para(doc,
        'Dựa trên tình huống giả định về dự án "Nền tảng AI hỗ trợ bảo tồn và giảng dạy '
        'tiếng Jrai (mở rộng sang tiếng Bahnar) cho chuyển đổi số tỉnh Gia Lai" tại '
        'Học viện Công nghệ Bưu chính Viễn thông, nhóm thực hiện soạn thảo bộ 03 văn bản: '
        'tờ trình, công văn và thông báo.'
    )

    add_para(doc,
        'Ba văn bản này được trình bày trong file riêng biệt: '
        'VanBanHanhChinh_PTIT_Rewritten.docx. '
        'Tất cả đều đảm bảo đầy đủ 09 thành phần thể thức theo NĐ30/2020/NĐ-CP, '
        'đã được kiểm tra bằng công cụ check_nd30.py với kết quả 11/11 PASSED.'
    )

    add_para(doc, 'Tóm tắt nội dung ba văn bản:', bold=True)

    add_para(doc, '1.1. Tờ trình (Số: 01/TTr-PTIT)', bold=True)
    add_para(doc,
        'Nội dung: Xin phê duyệt kinh phí và kế hoạch thực hiện đề tài nghiên cứu '
        '"Nền tảng AI hỗ trợ bảo tồn và giảng dạy tiếng Jrai". '
        'Người ký: Trưởng khoa Công nghệ thông tin.'
    )

    add_para(doc, '1.2. Công văn (Số: 02/CV-PTIT)', bold=True)
    add_para(doc,
        'Nội dung: Gửi Sở Khoa học và Công nghệ tỉnh Gia Lai đề nghị hợp tác, '
        'hỗ trợ thu thập dữ liệu ngôn ngữ Jrai, hướng dẫn chính sách và đồng tài trợ. '
        'Người ký: Phó Giám đốc Học viện.'
    )

    add_para(doc, '1.3. Thông báo (Số: 03/TB-PTIT)', bold=True)
    add_para(doc,
        'Nội dung: Thông báo kế hoạch triển khai dự án với lộ trình 12 tháng '
        '(tháng 6/2025 - tháng 5/2026) đến các đơn vị và cá nhân liên quan. '
        'Người ký: Phó Giám đốc Học viện.'
    )

    # ==================== NHIỆM VỤ 2 ====================
    doc.add_page_break()
    add_heading_styled(doc, 'Nhiệm vụ 2: CV cá nhân', level=2)

    add_para(doc,
        'CV được soạn cho thành viên Phạm Tuấn Anh, ứng tuyển vị trí AI Engineer Intern. '
        'CV được trình bày trong file CV_PhamTuanAnh_AIEngineerIntern_FIXED.docx, '
        'đảm bảo các yêu cầu: font Times New Roman, lề A4 theo quy định, '
        'đầy đủ các thành phần chính, tối ưu ATS, độ dài 1-2 trang A4.'
    )

    add_para(doc, 'Các thành phần chính của CV:', bold=True)

    cv_parts = [
        'Thông tin cá nhân: Họ tên, email, SĐT, địa chỉ, GitHub',
        'Mục tiêu nghề nghiệp: Ngắn hạn (AI Engineer) + dài hạn (nghiên cứu AI)',
        'Học vấn: PTIT, Kỹ thuật Phần mềm, GPA 3.47/4.00, IELTS 7.0',
        'Kỹ năng: Python, C++, SQL; PyTorch, Transformers, Hugging Face, OpenCV; FastAPI, Docker, Git',
        'Kinh nghiệm: Cộng tác viên nghiên cứu CV & Robotics tại PTIT',
        'Dự án: Tinh chỉnh ViT cho ảnh y khoa, nhận dạng giọng nói tiếng Việt, LSTM ước tính cảm biến',
        'Hoạt động ngoại khóa: Câu lạc bộ PAYT tại PTIT',
    ]
    for i, part in enumerate(cv_parts, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {part}')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

    # ==================== NHIỆM VỤ 3 ====================
    doc.add_page_break()
    add_heading_styled(doc, 'Nhiệm vụ 3: Thư ứng tuyển và Email', level=2)

    add_para(doc,
        'Thư ứng tuyển và email được soạn dựa trên CV của Phạm Tuấn Anh, '
        'ứng tuyển vị trí AI Engineer Intern tại FPT Smart Cloud. '
        'Nội dung được trình bày trong file ThuUngTuyen_Email_PhamTuanAnh.docx.'
    )

    add_para(doc, 'Tóm tắt nội dung:', bold=True)

    add_para(doc, 'Thư ứng tuyển (250-300 chữ):', bold=True)
    add_para(doc,
        'Giới thiệu bản thân là sinh viên PTIT, nêu lý do ứng tuyển vào FPT Smart Cloud. '
        'Trình bày kinh nghiệm nghiên cứu CV và NLP, các dự án đã hoàn thành (ViT, speech recognition, LSTM). '
        'Nhấn mạnh đóng góp dự kiến: phát triển mô hình AI, tối ưu pipeline ML, đóng góp vào sản phẩm AI. '
        'Kết thúc bằng lời cảm ơn và mong muốn phỏng vấn.'
    )

    add_para(doc, 'Email gửi hồ sơ:', bold=True)
    add_para(doc,
        'Chủ đề: "Phạm Tuấn Anh - AI Engineer Intern - FPT Smart Cloud". '
        'Nội dung 3 phần: giới thiệu ngắn gọn, lý do ứng tuyển + kỹ năng phù hợp, '
        'kết luận nhấn mạnh đóng góp cho công ty. '
        'File đính kèm: CV_PhamTuanAnh_AIEngineerIntern.pdf, ThuUngTuyen_PhamTuanAnh_AIEngineerIntern.pdf.'
    )

def create_conclusion(doc):
    """Tạo kết luận"""
    doc.add_page_break()
    add_heading_styled(doc, 'KẾT LUẬN', level=1)

    add_para(doc,
        'Qua quá trình thực hiện báo cáo, nhóm đã hoàn thành đầy đủ các nhiệm vụ '
        'theo yêu cầu của đề thi. Về lý thuyết, nhóm đã trình bày chi tiết chín thành phần '
        'thể thức chính của văn bản hành chính theo NĐ30/2020/NĐ-CP và các yêu cầu '
        'về CV cùng lưu ý khi soạn thảo email ứng tuyển.'
    )

    add_para(doc,
        'Về thực hành, nhóm đã soạn thảo thành công bộ 03 văn bản hành chính (tờ trình, '
        'công văn, thông báo) đảm bảo đúng thể thức, CV cá nhân chuyên nghiệp tối ưu ATS, '
        'thư ứng tuyển và email ứng tuyển đáp ứng yêu cầu. Tình huống giả định về dự án '
        'AI bảo tồn tiếng Jrai giúp nhóm vận dụng kiến thức vào thực tế.'
    )

    add_para(doc,
        'Báo cáo giúp nhóm hiểu sâu hơn về quy định thể thức văn bản hành chính, '
        'kỹ năng tạo lập CV và email chuyên nghiệp — những kỹ năng thiết yếu '
        'trong môi trường làm việc hiện đại.'
    )

def create_references(doc):
    """Tạo tài liệu tham khảo"""
    doc.add_page_break()
    add_heading_styled(doc, 'TÀI LIỆU THAM KHẢO', level=1)

    refs = [
        'Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về ban hành quy định về công tác văn thư.',
        'Phụ lục I, II, III Nghị định 30/2020/NĐ-CP về thể thức và kỹ thuật trình bày văn bản hành chính.',
        'Giáo trình Kỹ năng tạo lập văn bản tiếng Việt (SKD1103), Học viện Công nghệ Bưu chính Viễn thông.',
        'Hoàng Văn Hoa và Trần Thị Vân Hoa chủ biên (2012), Giáo trình giao tiếp trong kinh doanh, NxB Đại học Kinh tế Quốc dân.',
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'[{i}] {ref}')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

def create_contribution_table(doc):
    """Tạo bảng đánh giá đóng góp"""
    doc.add_page_break()
    add_heading_styled(doc, 'BẢNG ĐÁNH GIÁ MỨC ĐỘ ĐÓNG GÓP CỦA TỪNG THÀNH VIÊN', level=1)

    # Tạo bảng
    table = doc.add_table(rows=4, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    headers = ['STT', 'Họ và tên', 'Nhiệm vụ thực hiện', 'Mức độ đóng góp (%)']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

    # Dữ liệu mẫu
    data = [
        ['1', 'Phạm Tuấn Anh', 'Soạn thảo CV, thư ứng tuyển, email, tổng hợp báo cáo', '40%'],
        ['2', 'Thành viên 2', 'Soạn thảo tờ trình, công văn', '30%'],
        ['3', 'Thành viên 3', 'Soạn thảo thông báo, slide thuyết trình', '30%'],
    ]
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, cell_data in enumerate(row_data):
            cell = table.rows[row_idx].cells[col_idx]
            cell.text = cell_data
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

def create_slides_placeholder(doc):
    """Tạo placeholder cho slide thuyết trình"""
    doc.add_page_break()
    add_heading_styled(doc, 'SLIDE THUYẾT TRÌNH', level=1)

    add_para(doc,
        'Phần slide thuyết trình được trình bày dưới đây (2 slide/trang). '
        'Nội dung slide tóm tắt các nhiệm vụ đã thực hiện trong 7 phút trình bày.'
    )

    # Tạo bảng giả lập slide
    for slide_num in range(1, 7):
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        cell = table.rows[0].cells[0]
        cell.height = Mm(80)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if slide_num == 1:
            run = p.add_run('SLIDE 1: GIỚI THIỆU\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('Đề tài: Nền tảng AI hỗ trợ bảo tồn và giảng dạy tiếng Jrai\n')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'
            run3 = p.add_run('Nhóm: ... | Lớp: ...')
            run3.font.size = Pt(12)
            run3.font.name = 'Times New Roman'
        elif slide_num == 2:
            run = p.add_run('SLIDE 2: LÝ THUYẾT - 9 THÀNH PHẦN THỂ THỨC NĐ30\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('1. Quốc hiệu, tiêu ngữ\n2. Tên cơ quan chủ quản\n3. Tên cơ quan ban hành\n...')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'
        elif slide_num == 3:
            run = p.add_run('SLIDE 3: LÝ THUYẾT - CV VÀ EMAIL\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('CV: 11 thành phần, tối ưu ATS\nEmail: Cấu trúc 7 phần, lỗi thường gặp')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'
        elif slide_num == 4:
            run = p.add_run('SLIDE 4: THỰC HÀNH - 3 VĂN BẢN HÀNH CHÍNH\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('Tờ trình (01/TTr-PTIT)\nCông văn (02/CV-PTIT)\nThông báo (03/TB-PTIT)')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'
        elif slide_num == 5:
            run = p.add_run('SLIDE 5: THỰC HÀNH - CV VÀ THƯ ỨNG TUYỂN\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('CV: Phạm Tuấn Anh - AI Engineer Intern\nThư ứng tuyển + Email: FPT Smart Cloud')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'
        elif slide_num == 6:
            run = p.add_run('SLIDE 6: KẾT LUẬN\n\n')
            run.bold = True
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run2 = p.add_run('Hoàn thành đầy đủ các nhiệm vụ\nHiểu sâu về NĐ30 và kỹ năng tạo lập văn bản')
            run2.font.size = Pt(12)
            run2.font.name = 'Times New Roman'

        # Thêm khoảng cách giữa các slide
        if slide_num % 2 == 0 and slide_num < 6:
            doc.add_paragraph()

def main():
    doc = Document()

    # Thiết lập lề trang và khổ giấy A4
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

    # Font mặc định
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    # Tạo các phần
    create_cover_page(doc)
    doc.add_page_break()

    create_toc(doc)
    doc.add_page_break()

    create_foreword(doc)
    doc.add_page_break()

    create_theory(doc)
    doc.add_page_break()

    create_practice(doc)

    create_conclusion(doc)

    create_references(doc)

    create_contribution_table(doc)

    create_slides_placeholder(doc)

    # Lưu file
    output_path = os.path.join(OUTPUT_DIR, 'BCK_SKD1103_Lop_Nhom.docx')
    doc.save(output_path)
    print(f'Đã tạo báo cáo: {output_path}')

if __name__ == '__main__':
    main()
