# Skill The Thuc VB ND30

Skill tao file `.docx` van ban hanh chinh theo huong dan ND30.

## Luu y chuan phap ly

- Ve the thuc trinh bay van ban hanh chinh, skill nay dang bam theo Nghi dinh 30/2020/ND-CP.
- Neu don vi cua ban co bo quy dinh noi bo goi la "ND35", co the map them rule vao `references/` va mo rong script.

## Chinh sach su dung moi (bat buoc)

- Truoc khi soan, phai tra cuu van ban phap ly tuong ung voi tung loai van ban.
- Khi co boi canh chinh quyen dia phuong 2 cap, phai doi chieu tham quyen cap ban hanh va can cu cap Trung uong + dia phuong.
- Ket qua tra cuu phai duoc ghi vao `can_cu` va `ghi_chu_tra_cuu` trong file JSON dau vao.
- Bat buoc brainstorm sau (de cuong + bo cuc + kiem tra tham quyen) truoc khi xuat ban cuoi.
- Van ban dau ra phai day du cau truc, khong xuat ban nhap hoac ban thieu muc.

Xem them:
- `references/tra_cuu_van_ban_bat_buoc.md`
- `references/chinh_quyen_2_cap.md`
- `references/mau_can_cu_theo_linh_vuc.md`
- `references/benchmark_10_mau_web_va_bai_hoc.md`
- `examples/mau_can_cu_templates.json`

## Nang cap chat luong sau benchmark 10 mau

- Quyet dinh bat buoc co du Dieu 1-2-3, trong do Dieu 2 phai neu hieu luc va xu ly van ban cu (neu co).
- Quyet dinh uu tien co dong `Theo de nghi ... tai To trinh so ...` truoc phan `QUYET DINH:`.
- Ke hoach bat buoc du 4 muc I-II-III-IV va moi nhiem vu chinh co moc thoi gian + don vi chu tri.
- Neu URL tra cuu chuyen huong PDF khong trich duoc toan van, ghi ro trong `ghi_chu_tra_cuu`: `CAN_XAC_MINH_NOI_DUNG_GOC`.

## Cai dat

```bash
cd .agents/skill/Skill_The_Thuc_VB_ND30
npm install
```

## Sinh cong van

```bash
node scripts/generate_cong_van.js --input data.json --output cong_van.docx
```

## Sinh quyet dinh

```bash
node scripts/generate_quyet_dinh.js --input data.json --output quyet_dinh.docx
```

## Kiem tra nhanh (da xac thuc)

```bash
node scripts/generate_cong_van.js --input examples/cong_van.json --output examples/out_cong_van.docx
node scripts/generate_quyet_dinh.js --input examples/quyet_dinh.json --output examples/out_quyet_dinh.docx
```

## Mau bo sung truong tra cuu

```json
{
  "can_cu": [
    "Van ban 1...",
    "Van ban 2..."
  ],
  "ghi_chu_tra_cuu": [
    "Nguon: ...; ngay tra cuu: ...; tinh trang: con hieu luc",
    "Nguon: ...; ngay tra cuu: ...; tinh trang: can xac minh"
  ],
  "ghi_chu_2_cap": "Da doi chieu tham quyen ky va cap ban hanh theo boi canh chinh quyen 2 cap"
}
```

## JSON mau cong van

```json
{
  "loai_van_ban": "cong_van",
  "co_quan_chu_quan": "BO TAI CHINH",
  "co_quan_ban_hanh": "CUC THUE TP. HA NOI",
  "don_vi_soan_thao": "CT-HNi",
  "dia_danh": "Ha Noi",
  "trich_yeu": "V/v huong dan ke khai thue TNCN nam 2026",
  "kinh_gui": [
    "Cac Chi cuc Thue khu vuc",
    "Phong Thue thu nhap ca nhan"
  ],
  "noi_dung": "Noi dung cong van...",
  "cap_ky": "KT",
  "chuc_vu_cap_tren": "CUC TRUONG",
  "chuc_vu_ky": "PHO CUC TRUONG",
  "nguoi_ky": "Nguyen Van An",
  "noi_nhan": [
    "- Nhu tren;",
    "- Luu: VT"
  ]
}
```
