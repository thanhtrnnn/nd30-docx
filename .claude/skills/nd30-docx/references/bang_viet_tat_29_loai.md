# Bảng viết tắt 29 loại văn bản hành chính theo NĐ30

| STT | Viết tắt | Tên loại văn bản | Nhóm layout | Có tên loại | Có Kính gửi | Script |
|-----|----------|-------------------|-------------|-------------|-------------|--------|
| 1 | NQ | Nghị quyết | A | Có | Không | generate_nghi_quyet.js |
| 2 | QD | Quyết định | A | Có | Không | generate_quyet_dinh.js |
| 3 | CT | Chỉ thị | A | Có | Không | generate_chi_thi.js |
| 4 | QC | Quy chế | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 5 | QĐi | Quy định | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 6 | TC | Thông cáo | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 7 | TB | Thông báo | A | Có | Không | generate_thong_bao.js |
| 8 | HD | Hướng dẫn | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 9 | CTr | Chương trình | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 10 | KH | Kế hoạch | A | Có | Không | generate_ke_hoach.js |
| 11 | PA | Phương án | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 12 | DA | Đề án | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 13 | DA | Dự án | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 14 | BC | Báo cáo | A | Có | Có (tùy) | generate_bao_cao.js |
| 15 | BB | Biên bản | F | Có | Không | generate_bien_ban.js |
| 16 | TTr | Tờ trình | C | Có | Có | generate_to_trinh.js |
| 17 | HĐ | Hợp đồng | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 18 | CV | Công văn | B | Không | Có | generate_cong_van.js |
| 19 | CD | Công điện | B | Không | Có | generate_cong_dien.js |
| 20 | BGN | Biên bản ghi nhớ | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 21 | BTT | Biên bản thỏa thuận | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 22 | GUQ | Giấy ủy quyền | A | Có | Không | generate_van_ban_co_ten_loai.js |
| 23 | GM | Giấy mời | D | Có | Có | generate_giay_moi.js |
| 24 | GGT | Giấy giới thiệu | E | Có | Có | generate_giay_gioi_thieu.js |
| 25 | GNP | Giấy nghỉ phép | G | Có | Có | generate_giay_nghi_phep.js |
| 26 | PG | Phiếu gửi | H | Có | Không | generate_van_ban_co_ten_loai.js |
| 27 | PC | Phiếu chuyển | H | Có | Không | generate_van_ban_co_ten_loai.js |
| 28 | PB | Phiếu báo | H | Có | Không | generate_van_ban_co_ten_loai.js |
| 29 | TCV | Thư công | H | Có | Có | generate_van_ban_co_ten_loai.js |

## Nhóm layout

- **Nhóm A**: Văn bản có tên loại + trích yếu (Mẫu 1.4 NĐ30). Phần lớn 29 loại thuộc nhóm này.
- **Nhóm B**: Công văn / Công điện (Mẫu 1.5/1.6). Không có tên loại; trích yếu "V/v ..." dưới số ký hiệu.
- **Nhóm C**: Tờ trình. Có tên loại + Kính gửi cấp trên + nội dung đề xuất.
- **Nhóm D**: Giấy mời (Mẫu 1.7). Kính mời, thời gian, địa điểm.
- **Nhóm E**: Giấy giới thiệu (Mẫu 1.8). Người được giới thiệu, nhiệm vụ.
- **Nhóm F**: Biên bản (Mẫu 1.9). Thời gian, địa điểm, dual signature.
- **Nhóm G**: Giấy nghỉ phép (Mẫu 1.10). Thời gian nghỉ, lý do.
- **Nhóm H**: Các loại khác (Phiếu gửi/chuyển/báo, Thư công).

## Số ký hiệu

- **VB có tên loại**: `Số: XX/YYY-ZZZ` (XX = số thứ tự, YYY = mã loại, ZZZ = mã cơ quan)
- **Công văn**: `Số: XX/ZZZ-WWW` (XX = số thứ tự, ZZZ = mã cơ quan, WW = mã loại CV)
- Số < 10 phải có số 0 phía trước (vd: 01, 02, ..., 09)
