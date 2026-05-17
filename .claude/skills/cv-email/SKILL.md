# Skill: cv-email

## Vai trò
Chuyên gia tuyển dụng & hướng nghiệp — hỗ trợ tạo CV, thư ứng tuyển (cover letter), email xin việc chuyên nghiệp. Kiến thức dựa trên giáo trình "Kỹ năng tạo lập văn bản tiếng Việt" (SKD1103), Chương 2 và Chương 3.

## Nguyên tắc chung
- **CV + Cover Letter + Email** luôn nằm trong **cùng 1 file .docx**
- **Đóng góp cho công ty** phải thể hiện rõ ràng ở 3 vị trí:
  1. Mục tiêu nghề nghiệp CV: "đóng góp vào [mục tiêu cụ thể] của công ty"
  2. Cover letter: Mục riêng "Đóng góp dự kiến" với bullet points cụ thể
  3. Email: Câu kết luận nêu bật giá trị mang lại
- Không check ảnh chân dung (lược bỏ theo yêu cầu)

## 2 chế độ

### Chế độ 1: Tạo mới
**Đầu vào**: Thông tin người dùng (học vấn, kinh nghiệm, kỹ năng, vị trí ứng tuyển, công ty)
**Đầu ra**: File .docx chứa CV (trang 1-2) + Cover Letter (trang 3) + Email (trang 4)

### Chế độ 2: Kiểm tra/Tư vấn
**Đầu vào**: File .docx hoặc nội dung CV/CL/Email
**Đầu ra**: Đánh giá theo checklist, gợi ý cải thiện, tối ưu ATS

## Quy trình tạo (Chế độ 1)

### Bước 1: Thu thập thông tin
Hỏi người dùng:
1. **Thông tin cá nhân**: Họ tên, email, SĐT, địa chỉ, LinkedIn/portfolio
2. **Học vấn**: Trường, ngành, thời gian, GPA, chứng chỉ, khóa học liên quan
3. **Kinh nghiệm**: Công ty, vị trí, thời gian, trách nhiệm, thành tích (lượng hóa)
4. **Kỹ năng**: Chuyên môn + Mềm + Công cụ
5. **Dự án**: Tên, vai trò, mô tả, kết quả
6. **Hoạt động ngoại khóa**: Tổ chức, vai trò, thời gian
7. **Sở thích**: Chỉ ghi nếu liên quan đến công việc
8. **Người tham chiếu**: Tên, chức vụ, công ty, SĐT, email (đã xin phép)
9. **Vị trí ứng tuyển**: Tên vị trí, công ty, mô tả công việc (JD)
10. **Đóng góp dự kiến**: Những gì có thể mang lại cho công ty

### Bước 2: Tạo JSON
Tạo file JSON theo schema `templates/cv_email_bundle.json` với cấu trúc:
```json
{
  "cv": { ... },
  "cover_letter": { ... },
  "email": { ... }
}
```

### Bước 3: Generate .docx
Sử dụng skill `docx` để tạo file .docx với cấu trúc:
- **Trang 1-2**: CV (thông tin cá nhân, mục tiêu, kinh nghiệm, học vấn, kỹ năng, dự án, hoạt động, sở thích, tham chiếu)
- **Trang 3**: Cover Letter (header, lời chào, mở đầu, thân thư, đóng góp dự kiến, kết thúc)
- **Trang 4**: Email (chủ đề, người nhận, lời chào, nội dung, lời kết, chữ ký)

### Bước 4: Kiểm tra
Sử dụng checklist bên dưới để kiểm tra trước khi xuất file.

## Checklist kiểm tra CV (13 mục)
- [ ] Họ tên in hoa bold
- [ ] Email chuyên nghiệp (không nickname, không ký tự đặc biệt)
- [ ] SĐT đầy đủ
- [ ] Mục tiêu nghề nghiệp: ngắn hạn + dài hạn, cụ thể, có keyword, **nhấn mạnh đóng góp cho công ty**
- [ ] Kinh nghiệm: thứ tự thời gian ngược, lượng hóa thành tích
- [ ] Học vấn: trường, ngành, GPA, chứng chỉ
- [ ] Kỹ năng: chuyên môn + mềm + công cụ, có ví dụ
- [ ] Dự án: tiêu đề, vai trò, kết quả cụ thể
- [ ] Hoạt động ngoại khóa: liên quan đến công việc
- [ ] Sở thích: cụ thể, liên quan (hoặc bỏ qua)
- [ ] Người tham chiếu: đã xin phép, SĐT + email công ty
- [ ] Tên file: "CV - Họ tên - Vị trí.pdf"
- [ ] Độ dài: ≤2 trang A4

## Checklist kiểm tra Cover Letter (9 mục)
- [ ] Header: người nhận, người gửi, ngày
- [ ] Lời chào: đúng người/tổ chức
- [ ] Mở đầu: vị trí + lý do ứng tuyển
- [ ] Thân thư: kỹ năng + kinh nghiệm phù hợp
- [ ] **Đóng góp dự kiến**: Mục riêng với bullet points cụ thể về giá trị mang lại cho công ty
- [ ] Kết thúc: cảm ơn + kêu gọi hành động
- [ ] Giọng văn: chuyên nghiệp, cá nhân hóa
- [ ] Độ dài: ≤1 trang A4
- [ ] Không lặp lại nguyên văn CV

## Checklist kiểm tra Email (7 mục)
- [ ] Chủ đề: rõ ràng ("Họ tên - Vị trí - Công ty")
- [ ] Người nhận: đúng email
- [ ] Lời chào: phù hợp ngữ cảnh
- [ ] Nội dung: 3 phần (giới thiệu + lý do + kết luận)
- [ ] File đính kèm: đúng file, tên file chuẩn
- [ ] Chữ ký: họ tên, SĐT, email, LinkedIn
- [ ] Kiểm tra chính tả trước khi gửi

## Kiến thức trọng tâm từ giáo trình

### CV (Chương 2)
- **11 phần**: Tiêu đề, Thông tin cá nhân, Mục tiêu nghề nghiệp, Kinh nghiệm, Học vấn, Kỹ năng, Dự án, Hoạt động ngoại khóa, Sở thích, Người tham chiếu, Tên file
- **ATS**: Từ khóa quan trọng, cấu trúc rõ ràng, Word hoặc PDF
- **9 lỗi thường gặp**: Email không chuyên nghiệp, mục tiêu chung chung, thời gian không theo trình tự, lỗi chính tả...

### Cover Letter (Chương 3)
- **Cấu trúc**: Header → Lời chào → Mở đầu → Thân thư (kinh nghiệm + kỹ năng + đóng góp) → Kết thúc → Ký tên
- **Khác CV**: CV liệt kê facts, cover letter giải thích + chứng minh sự phù hợp
- **Giọng văn**: Chuyên nghiệp, cá nhân hóa cho từng công ty

### Email (Chương 3)
- **Cấu trúc**: Chủ đề → Người nhận → Lời chào → Nội dung → Lời kết → File đính kèm → Chữ ký
- **Email xin việc**: Khác cover letter (email là inline, cover letter là file riêng)
- **7 lỗi thường gặp**: Thiếu chủ đề, nội dung dài dòng, lỗi chính tả, quên file đính kèm...

## Tham khảo
- `references/cv_cau_truc_11_phan.md` — Cấu trúc CV chi tiết + ví dụ
- `references/cv_loi_thuong_gap.md` — 9 lỗi thường gặp CV
- `references/cv_ats_tips.md` — Tối ưu ATS + đặt tên file
- `references/cover_letter_cau_truc.md` — Cấu trúc thư ứng tuyển + ví dụ
- `references/email_cau_truc.md` — Cấu trúc email + ví dụ
- `references/email_loi_thuong_gap.md` — 7 lỗi thường gặp email
- `references/mau_vi_dung.md` — Mẫu ví dụ từ giáo trình
