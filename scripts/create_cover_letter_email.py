#!/usr/bin/env python3
"""Tạo thư ứng tuyển (cover letter) + email cho Phạm Tuấn Anh"""
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_cover_letter_email(output_path):
    doc = Document()

    # Thiết lập lề trang
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(20)

    # Font mặc định
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    # ==================== COVER LETTER ====================
    # Tiêu đề
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('THƯ ỨNG TUYỂN')
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'

    # Thông tin người gửi
    sender = doc.add_paragraph()
    sender.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sender.add_run('Phạm Tuấn Anh').bold = True
    sender.add_run('\nEmail: phamtuananhai@gmail.com')
    sender.add_run('\nSĐT: 0912 345 678')
    sender.add_run('\nGitHub: github.com/phamtuananhai')

    # Thông tin người nhận
    receiver = doc.add_paragraph()
    receiver.add_run('Phòng Nhân sự')
    receiver.add_run('\nCông ty FPT Smart Cloud')
    receiver.add_run('\nFPT Tower, Phạm Văn Bạch, Cầu Giấy, Hà Nội')

    # Ngày
    date_para = doc.add_paragraph()
    date_para.add_run('Hà Nội, ngày 17 tháng 05 năm 2026')

    # Tiêu đề thư
    subject = doc.add_paragraph()
    subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subject.add_run('Ứng tuyển vị trí AI Engineer Intern')
    run.bold = True
    run.font.size = Pt(13)

    # Lời chào
    greeting = doc.add_paragraph()
    greeting.add_run('Kính gửi Quý Công ty FPT Smart Cloud,')

    # Mở đầu
    intro = doc.add_paragraph()
    intro.add_run(
        'Tôi là Phạm Tuấn Anh, sinh viên năm thứ ba chuyên ngành Kỹ thuật Phần mềm '
        'trường Đại học Công nghệ Viễn thông, Viện Đào tạo Tài năng PTIT. '
        'Tôi viết thư này để bày tỏ mong muốn được ứng tuyển vào vị trí AI Engineer Intern '
        'tại FPT Smart Cloud. Qua tìm hiểu, tôi nhận thấy công ty đang có những bước phát triển '
        'ấn tượng trong lĩnh vực điện toán đám mây và trí tuệ nhân tạo, và tôi rất mong muốn '
        'được đóng góp vào sự phát triển đó.'
    )

    # Thân thư - Kinh nghiệm + Kỹ năng
    body1 = doc.add_paragraph()
    body1.add_run(
        'Trong quá trình học tập và nghiên cứu, tôi đã tích lũy được kinh nghiệm đáng kể '
        'trong lĩnh vực Computer Vision và Natural Language Processing. Tôi hiện đang là '
        'cộng tác viên nghiên cứu tại PTIT, tham gia các dự án về thị giác máy tính và robotics. '
        'Tôi đã hoàn thành nhiều dự án thực tế, bao gồm: tinh chỉnh mô hình Vision Transformer (ViT) '
        'cho bài toán phân loại ảnh y khoa với độ chính xác 88.58%; xây dựng pipeline nhận dạng giọng nói '
        'cho ngôn ngữ ít tài nguyên (tiếng Việt phương ngữ), giảm 29% lỗi tương đối so với Whisper large-v2; '
        'và phát triển mô hình LSTM ước tính dữ liệu cảm biến với RMSE 0.0166.'
    )

    body2 = doc.add_paragraph()
    body2.add_run(
        'Tôi thành thạo Python, C++ và các framework AI/ML phổ biến như PyTorch, Transformers, '
        'Hugging Face, OpenCV. Bên cạnh đó, tôi có kinh nghiệm với các kỹ thuật tối ưu mô hình '
        'như LoRA/Adapter fine-tuning và Retrieval-Augmented Generation (RAG). '
        'Tôi cũng quen thuộc với các công cụ DevOps như FastAPI, Docker, Linux và Git.'
    )

    # Đóng góp dự kiến
    contribution_title = doc.add_paragraph()
    run = contribution_title.add_run('Đóng góp dự kiến:')
    run.bold = True

    contributions = [
        'Phát triển và tối ưu các mô hình AI/ML cho sản phẩm cloud của FPT',
        'Áp dụng kinh nghiệm Computer Vision và NLP vào các dự án thực tế của công ty',
        'Đóng góp vào nghiên cứu và phát triển công nghệ AI tiên tiến',
        'Hỗ trợ đội ngũ kỹ thuật trong việc triển khai và đánh giá mô hình'
    ]
    for c in contributions:
        p = doc.add_paragraph(c, style='List Bullet')

    # Kết thúc
    closing = doc.add_paragraph()
    closing.add_run(
        'Tôi tin rằng sự nhiệt huyết, tinh thần cầu tiến và các kỹ năng chuyên môn '
        'sẽ giúp tôi trở thành một thực tập sinh mang lại giá trị cho FPT Smart Cloud. '
        'Tôi rất mong có cơ hội được trao đổi thêm với Quý công ty về khả năng đóng góp '
        'của mình qua buổi phỏng vấn.'
    )

    thanks = doc.add_paragraph()
    thanks.add_run('Tôi xin chân thành cảm ơn Quý công ty đã dành thời gian xem xét hồ sơ của tôi.')

    # Ký tên
    sign = doc.add_paragraph()
    sign.add_run('Trân trọng,').bold = False
    sign.add_run('\n').bold = False
    sign.add_run('Phạm Tuấn Anh').bold = True

    # ==================== PAGE BREAK ====================
    doc.add_page_break()

    # ==================== EMAIL ====================
    email_title = doc.add_paragraph()
    email_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = email_title.add_run('NỘI DUNG EMAIL GỬI HỒ SƠ')
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'

    # Chủ đề
    subject_email = doc.add_paragraph()
    subject_email.add_run('Chủ đề: ').bold = True
    subject_email.add_run('Phạm Tuấn Anh - AI Engineer Intern - FPT Smart Cloud')

    # Người nhận
    to = doc.add_paragraph()
    to.add_run('Gửi: ').bold = True
    to.add_run('Phòng Nhân sự FPT Smart Cloud (hr@fptsmartcloud.com)')

    # Nội dung email
    email_body = doc.add_paragraph()
    email_body.add_run('Kính gửi Quý Công ty FPT Smart Cloud,\n\n')

    email_body.add_run(
        'Tôi là Phạm Tuấn Anh, sinh viên năm thứ ba chuyên ngành Kỹ thuật Phần mềm '
        'trường Đại học Công nghệ Viễn thông. Tôi viết email này để ứng tuyển vào vị trí '
        'AI Engineer Intern tại FPT Smart Cloud.\n\n'
    )

    email_body.add_run(
        'Với kinh nghiệm nghiên cứu trong lĩnh vực Computer Vision và NLP, cùng với kỹ năng '
        'Python, PyTorch, Transformers, tôi tin rằng mình phù hợp với yêu cầu của vị trí này. '
        'Tôi đã hoàn thành nhiều dự án thực tế, bao gồm tinh chỉnh mô hình ViT cho phân loại ảnh y khoa '
        'và xây dựng pipeline nhận dạng giọng nói cho ngôn ngữ ít tài nguyên.\n\n'
    )

    email_body.add_run(
        'Tôi tin rằng mình có thể đóng góp vào mục tiêu phát triển sản phẩm AI cloud '
        'của FPT Smart Cloud thông qua các kỹ năng chuyên môn và tinh thần học hỏi. '
        'Rất mong có cơ hội được trao đổi thêm tại buổi phỏng vấn.\n\n'
    )

    email_body.add_run('Trân trọng cảm ơn,\n')
    email_body.add_run('Phạm Tuấn Anh').bold = True
    email_body.add_run('\nEmail: phamtuananhai@gmail.com')
    email_body.add_run('\nSĐT: 0912 345 678')

    # File đính kèm
    attachment = doc.add_paragraph()
    attachment.add_run('\nFile đính kèm:').bold = True
    attachment.add_run('\n1. CV_PhamTuanAnh_AIEngineerIntern.pdf')
    attachment.add_run('\n2. ThuUngTuyen_PhamTuanAnh_AIEngineerIntern.pdf')

    doc.save(output_path)
    print(f"Đã tạo thư ứng tuyển + email: {output_path}")

if __name__ == '__main__':
    create_cover_letter_email('output/ThuUngTuyen_Email_PhamTuanAnh.docx')
