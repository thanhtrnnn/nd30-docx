const fs = require("fs");
const {
  AlignmentType, BorderStyle, Document, Packer, Paragraph,
  Table, TableCell, TableRow, TextRun, WidthType, LineRuleType,
} = require("docx");

function parseCliArgs(argv) {
  const args = { input: "data.json", output: "giay_moi.docx" };
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === "--input" && argv[i + 1]) args.input = argv[i + 1];
    if (argv[i] === "--output" && argv[i + 1]) args.output = argv[i + 1];
  }
  return args;
}

function readJsonFile(p) {
  const abs = require("path").resolve(process.cwd(), p);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  return JSON.parse(fs.readFileSync(abs, "utf8"));
}

async function writeDocx(doc, p) {
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(require("path").resolve(process.cwd(), p), buf);
}

function noBorders() {
  return {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };
}

function paraText(text, opts = {}) {
  const indentSpec = opts.indent !== undefined ? opts.indent : opts.noIndent ? { firstLine: 0 } : { firstLine: 720 };
  return new Paragraph({
    alignment: opts.alignment || AlignmentType.LEFT,
    spacing: opts.spacing || { before: 120, after: 120, line: 340, lineRule: LineRuleType.EXACT },
    indent: indentSpec,
    children: [new TextRun({ text: text || "", size: opts.size || 28, bold: Boolean(opts.bold), italics: Boolean(opts.italics), font: "Times New Roman" })],
  });
}

function topBorderParagraph(indent = 1100) {
  return new Paragraph({
    spacing: { before: 20, after: 0 },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: "000000", space: 1 } },
    indent: { left: indent, right: indent },
  });
}

function formatNgayThang(ngayThang) {
  if (!ngayThang) return ", ngày ... tháng ... năm ...";
  return String(ngayThang).replace(/ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})/, (_, d, m, y) => {
    return `ngày ${d.length === 1 ? "0" + d : d} tháng ${m.length === 1 ? "0" + m : m} năm ${y}`;
  });
}

async function generateGiayMoi(jsonPath, outputPath) {
  const data = readJsonFile(jsonPath);
  const required = ["co_quan_chu_quan", "co_quan_ban_hanh", "kinh_gui", "chuc_vu_ky", "nguoi_ky"];
  const missing = required.filter(k => !data[k]);
  if (missing.length) throw new Error(`Missing: ${missing.join(", ")}`);

  const diaDanh = data.dia_danh || "Hà Nội";
  const soKyHieu = data.so_ky_hieu || "Số: ..../GM-...";
  const ngayThang = formatNgayThang(data.ngay_thang);

  const children = [];

  // Header
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [
        new TableCell({ width: { size: 3500, type: WidthType.DXA }, borders: noBorders(), children: [
          paraText(data.co_quan_chu_quan, { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
          paraText(data.co_quan_ban_hanh, { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
          topBorderParagraph(1350),
        ]}),
        new TableCell({ width: { size: 5571, type: WidthType.DXA }, borders: noBorders(), children: [
          paraText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
          paraText("Độc lập - Tự do - Hạnh phúc", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true }),
          topBorderParagraph(1100),
        ]}),
      ]}),
      new TableRow({ children: [
        new TableCell({ width: { size: 3500, type: WidthType.DXA }, borders: noBorders(), children: [
          paraText(soKyHieu, { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
        ]}),
        new TableCell({ width: { size: 5571, type: WidthType.DXA }, borders: noBorders(), children: [
          paraText(`${diaDanh}${ngayThang}`, { alignment: AlignmentType.CENTER, noIndent: true, size: 28, italics: true }),
        ]}),
      ]}),
    ],
  }));

  children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));

  // GIẤY MỜI
  children.push(paraText("GIẤY MỜI", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 200, after: 120 } }));

  // Kính mời
  const kinhGui = Array.isArray(data.kinh_gui) ? data.kinh_gui.join("; ") : data.kinh_gui;
  children.push(paraText(`Kính mời: ${kinhGui}`, { alignment: AlignmentType.LEFT, noIndent: true, size: 28, italics: true, spacing: { before: 120, after: 200 } }));

  // Nội dung mời
  if (data.noi_dung_moi) {
    children.push(paraText(data.noi_dung_moi, { alignment: AlignmentType.JUSTIFIED, size: 28 }));
  }

  // Thời gian
  if (data.thoi_gian) {
    children.push(paraText(`Thời gian: ${data.thoi_gian}`, { alignment: AlignmentType.JUSTIFIED, size: 28, bold: true }));
  }

  // Địa điểm
  if (data.dia_diem) {
    children.push(paraText(`Địa điểm: ${data.dia_diem}`, { alignment: AlignmentType.JUSTIFIED, size: 28, bold: true }));
  }

  // Nội dung chi tiết
  if (Array.isArray(data.noi_dung)) {
    data.noi_dung.forEach(block => {
      if (typeof block === "string") {
        block.split("\n").forEach(line => {
          children.push(paraText(line, { alignment: AlignmentType.JUSTIFIED, size: 28 }));
        });
      }
    });
  }

  // Trân trọng kính mời
  children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));
  children.push(paraText("Trân trọng kính mời!", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, italics: true }));

  children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));

  // Footer
  const noiNhan = Array.isArray(data.noi_nhan) && data.noi_nhan.length ? data.noi_nhan : ["- Như trên;", "- Lưu: VT."];
  const noiNhanChildren = [
    paraText("Nơi nhận:", { noIndent: true, size: 24, bold: true, italics: true, spacing: { before: 120, after: 60 } }),
    ...noiNhan.map(item => paraText(item, { noIndent: true, size: 22, spacing: { before: 0, after: 20 } })),
  ];

  const capKy = (data.cap_ky || "").toUpperCase();
  const sigLines = [];
  if (["KT", "TL", "TUQ", "TM", "Q"].includes(capKy) && data.chuc_vu_cap_tren) sigLines.push(`${capKy}. ${data.chuc_vu_cap_tren}`);
  sigLines.push(data.chuc_vu_ky || "NGƯỜI KÝ");
  const sigChildren = [
    ...sigLines.map(l => paraText(l, { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 0, after: 60 } })),
    paraText("(Ký, ghi rõ họ tên)", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, italics: true, spacing: { before: 40, after: 500 } }),
    paraText(data.nguoi_ky || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 0, after: 120 } }),
  ];

  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: 55, type: WidthType.PERCENTAGE }, borders: noBorders(), children: noiNhanChildren }),
      new TableCell({ width: { size: 45, type: WidthType.PERCENTAGE }, borders: noBorders(), children: sigChildren }),
    ]})],
  }));

  const doc = new Document({
    styles: { default: { document: { run: { font: "Times New Roman", size: 28 }, paragraph: { indent: { firstLine: 0 } } } } },
    sections: [{
      properties: { page: { margin: { top: 1134, right: 851, bottom: 1134, left: 1701 }, size: { width: 11906, height: 16838 } } },
      children,
    }],
  });

  await writeDocx(doc, outputPath);
  console.log(`Generated: ${outputPath}`);
}

const args = parseCliArgs(process.argv);
generateGiayMoi(args.input, args.output).catch(err => { console.error("Error:", err.message); process.exit(1); });
