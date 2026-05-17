# SKILL: Tao van ban hanh chinh theo ND30

## Muc tieu

- Sinh file `.docx` cho cac loai: cong van, quyet dinh, ke hoach.
- Tuan thu the thuc va ky thuat trinh bay theo Phu luc I Nghi dinh 30/2020/ND-CP.
- Tuan thu quy tac viet hoa theo Phu luc II Nghi dinh 30/2020/ND-CP.
- Dung cac script Node.js trong thu muc `scripts/`.
- Phu hop boi canh van hanh chinh quyen dia phuong 2 cap (neu nguoi dung yeu cau).
- Bat buoc tra cuu van ban phap ly tuong ung truoc khi viet noi dung.

## Nguyen tac bat buoc

1. Khong duoc soan van ban neu chua co danh sach can cu phap ly da tra cuu.
2. Moi van ban phai co toi thieu 2-3 can cu: 1 van ban goc + 1 van ban huong dan/van ban dia phuong (neu co).
3. Khi co yeu to "chinh quyen 2 cap", uu tien tra cuu van ban cap Trung uong + van ban trien khai cap tinh.
4. Neu khong xac minh duoc hieu luc van ban, phai gan co "CAN_XAC_MINH_HIEU_LUC" trong ghi chu.
5. Truoc khi tao JSON cuoi, agent phai brainstorm sau va lap de cuong day du.
6. Mac dinh mo hinh 2 cap KHONG con cap huyen; khong dua "UBND huyen" vao co quan chu quan, noi nhan, hoac tham chieu mac dinh.
7. Khi soan QUYET DINH, bat buoc co du 3 dieu cot loi: Dieu 1 (noi dung ban hanh), Dieu 2 (hieu luc + thay the/huy bo neu co), Dieu 3 (trach nhiem thi hanh).
8. Khi soan KE HOACH, bat buoc du 4 muc I-II-III-IV (Muc dich/yeu cau, Nhiem vu-giai phap, Kinh phi, To chuc thuc hien).
9. Neu nguon tham khao web chi truy xuat duoc metadata do redirect PDF, phai danh dau `CAN_XAC_MINH_NOI_DUNG_GOC` trong `ghi_chu_tra_cuu`.

## Quy tac the thuc bat buoc (Phu luc I ND30)

### Kho giay va le trang
- Kho A4 (210mm x 297mm), doc theo chieu dai.
- Le tren/duoi: 20-25mm; Le trai: 30-35mm; Le phai: 15-20mm.

### Font chu
- Times New Roman, Unicode TCVN 6909:2001, mau den.

### Co chu tung thanh phan
- Quoc hieu: 12-13pt, IN HOA, dam.
- Tieu ngu: 13-14pt, in thuong (dau viet hoa), dam.
- Ten CQ chu quan: 12-13pt, IN HOA, khong dam.
- Ten CQ ban hanh: 12-13pt, IN HOA, dam.
- So, ky hieu: 13pt.
- Dia danh, ngay thang: 13-14pt, nghieng. Ngay < 10 va thang 1, 2 phai them so 0 phia truoc.
- Ten loai VB: 13-14pt, IN HOA, dam, canh giua.
- Trich yeu (VB co ten loai): 13-14pt, dam, canh giua.
- Trich yeu cong van ("V/v"): 12-13pt, canh giua duoi so ky hieu.
- Can cu ban hanh: 13-14pt, nghieng, cuoi dong ";", dong cuoi ".".
- Noi dung: 13-14pt, canh deu 2 le, lui dau dong 1cm/1.27cm.
- Noi nhan (label): 12pt, nghieng, dam.
- Noi nhan (list): 11pt.
- So trang: 13-14pt, khong hien thi trang dau.

### Layout footer
- Noi nhan (trai, 55%) ngang hang voi khoi ky (phai, 45%) trong table 2 cot an vien.
- KHONG xep tuong tu (noi nhan phia duoi khoi ky).

### Viet hoa
- Tham chieu `references/quy_tac_viet_hoa.md` (Phu luc II ND30).

## Deep Brainstorm Protocol (bat buoc)

1. Nhan dien dung loai van ban va muc tieu quan ly nha nuoc cua van ban.
2. Brainstorm it nhat 2 phuong an bo cuc noi dung, chon 1 phuong an toi uu.
3. Lap de cuong day du: mo dau, can cu, muc tieu, nhiem vu/giai phap, to chuc thuc hien, hieu luc/che do bao cao (tuy loai van ban).
4. Kiem tra logic tham quyen: co quan ban hanh, cap ky, doi tuong thi hanh.
5. Kiem tra dau cau, chinh ta, viet hoa, tinh ro rang va kha nang thi hanh.

## Tieu chuan van ban day du

- Co du phan can cu phap ly (ghi day du ten loai, so, ky hieu, CQ ban hanh, ngay, trich yeu).
- Co du phan noi dung chinh theo loai van ban (Dieu 1-2-3 hoac chuong/muc).
- Co phan to chuc thuc hien/doi tuong chiu trach nhiem.
- Co khoi noi nhan va khoi ky dung tham quyen.
- Co ghi chu tra cuu phap ly trong du lieu dau vao.

## Quy trinh agent

1. Thu thap thong tin dau vao (loai van ban, co quan, nguoi ky, muc dich).
2. Tra cuu van ban phap ly tuong ung theo chu de; lap bang tham chieu.
3. Kiem tra boi canh chinh quyen 2 cap va cap tham quyen ban hanh.
4. Xac dinh cap ky (TM/KT/Q/TL/TUQ) va khoi ky.
5. Tao file JSON du lieu.
6. Chay script phu hop de sinh file `.docx`.
7. Kiem tra ket qua va tra file cho nguoi dung kem danh sach can cu da dung.

## Dau ra bat buoc truoc khi sinh docx

- Muc `can_cu` (array) phai duoc dien day du.
- Muc `ghi_chu_tra_cuu` (array) de ghi nguon tra cuu, ngay tra cuu, tinh trang hieu luc.
- Neu van ban lien quan to chuc bo may dia phuong: them `ghi_chu_2_cap`.
- QUYET DINH:
  - Phai co `chu_the_ban_hanh` (vd: "UY BAN NHAN DAN XA AN NINH").
  - Phai co `theo_de_nghi` (vd: "Theo de nghi cua ... tai To trinh so ...").
  - Phai kiem tra du Dieu 1, Dieu 2, Dieu 3.
- KE HOACH: moi nhiem vu chinh phai co moc thoi gian va don vi chu tri/phoi hop (neu co).

## Cac loai van ban hanh chinh theo ND30

ND30 quy dinh 29 loai van ban hanh chinh:
Nghi quyet (ca biet), Quyet dinh (ca biet), Chi thi, Quy che, Quy dinh, Thong cao, Thong bao, Huong dan, Chuong trinh, Ke hoach, Phuong an, De an, Du an, Bao cao, Bien ban, To trinh, Hop dong, Cong van, Cong dien, Ban ghi nho, Ban cam ket, Ban thoa thuan, Giay uy quyen, Giay moi, Giay gioi thieu, Giay nghi phep, Phieu gui, Phieu chuyen, Thu cong.

Skill hien tai ho tro sinh docx cho: Cong van, Quyet dinh, Ke hoach.

## Lenh mac dinh

- Cong van: `node scripts/generate_cong_van.js --input data.json --output cong_van.docx`
- Quyet dinh: `node scripts/generate_quyet_dinh.js --input data.json --output quyet_dinh.docx`
- Ke hoach: `node scripts/generate_ke_hoach.js --input data.json --output ke_hoach.docx`

## Tai lieu tham chieu trong skill

- `references/quy_tac_the_thuc.md` — Bang co chu, kieu chu, le trang day du theo Phu luc I ND30.
- `references/quy_tac_viet_hoa.md` — Quy tac viet hoa theo Phu luc II ND30.
- `references/phan_quyen_ky.md` — TM/KT/Q/TL/TUQ va trinh bay khoi ky.
- `references/tra_cuu_van_ban_bat_buoc.md`
- `references/chinh_quyen_2_cap.md`
- `references/mau_can_cu_theo_linh_vuc.md`
- `references/huong_dan_bo_cuc_noi_dung.md`
- `references/benchmark_10_mau_web_va_bai_hoc.md`
