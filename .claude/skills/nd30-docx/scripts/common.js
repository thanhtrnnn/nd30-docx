const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  Packer,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  LineRuleType,
} = require("docx");

const PAGE = {
  marginTop: 1134,
  marginRight: 851,
  marginBottom: 1134,
  marginLeft: 1701,
};

function parseCliArgs(argv) {
  const args = { input: "data.json", output: "output.docx" };
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === "--input" && argv[i + 1]) args.input = argv[i + 1];
    if (argv[i] === "--output" && argv[i + 1]) args.output = argv[i + 1];
  }
  return args;
}

function readJsonFile(inputPath) {
  const absPath = path.resolve(process.cwd(), inputPath);
  if (!fs.existsSync(absPath)) {
    throw new Error(`Input file not found: ${absPath}`);
  }
  return JSON.parse(fs.readFileSync(absPath, "utf8"));
}

async function writeDocx(doc, outputPath) {
  const buffer = await Packer.toBuffer(doc);
  const absPath = path.resolve(process.cwd(), outputPath);
  fs.writeFileSync(absPath, buffer);
  return absPath;
}

function paraText(text, opts = {}) {
  const indentSpec = opts.indent !== undefined
    ? opts.indent
    : opts.noIndent
      ? { firstLine: 0 }
      : { firstLine: 720 };

  return new Paragraph({
    alignment: opts.alignment || AlignmentType.LEFT,
    spacing: opts.spacing || { before: 120, after: 120, line: 340, lineRule: LineRuleType.EXACT },
    indent: indentSpec,
    children: [
      new TextRun({
        text: text || "",
        size: opts.size || 28,
        bold: Boolean(opts.bold),
        italics: Boolean(opts.italics),
        font: "Times New Roman",
      }),
    ],
  });
}

function ensureEndingPunctuation(text, fallback = ".") {
  const value = String(text || "").trim();
  if (!value) return "";
  if (/[\.;:!?]$/.test(value)) return value;
  return `${value}${fallback}`;
}

function topBorderParagraph(indent = 1100) {
  return new Paragraph({
    spacing: { before: 20, after: 0 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 2, color: "000000", space: 1 },
    },
    indent: { left: indent, right: indent },
  });
}

function headerTable(data, options = {}) {
  const diaDanh = data.dia_danh || "Hà Nội";
  const soKyHieu = data.so_ky_hieu || "Số: ...../.....";
  const ngayThang = formatNgayThang(data.ngay_thang);

  const leftTop = [
    paraText(data.co_quan_chu_quan || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
    paraText(data.co_quan_ban_hanh || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
    topBorderParagraph(1350),
  ];

  const rightTop = [
    paraText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
    paraText("Độc lập - Tự do - Hạnh phúc", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true }),
    topBorderParagraph(1100),
  ];

  const leftBottom = [
    paraText(soKyHieu, { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
  ];

  if (!options.decision) {
    leftBottom.push(paraText(data.trich_yeu || "V/v ...", { alignment: AlignmentType.CENTER, noIndent: true, size: 24, spacing: { before: 60, after: 0 } }));
  }

  const rightBottom = [
    paraText(`${diaDanh}${ngayThang}`, { alignment: AlignmentType.CENTER, noIndent: true, size: 28, italics: true }),
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 3500, type: WidthType.DXA },
            borders: noBorders(),
            children: leftTop,
          }),
          new TableCell({
            width: { size: 5571, type: WidthType.DXA },
            borders: noBorders(),
            children: rightTop,
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 3500, type: WidthType.DXA },
            borders: noBorders(),
            children: leftBottom,
          }),
          new TableCell({
            width: { size: 5571, type: WidthType.DXA },
            borders: noBorders(),
            children: rightBottom,
          }),
        ],
      }),
    ],
  });
}

function noBorders() {
  return {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };
}

function signatureBlock(data) {
  const lines = [];
  const capKy = (data.cap_ky || "").toUpperCase();
  const validPrefixes = ["KT", "TL", "TUQ", "TM", "Q"];

  if (validPrefixes.includes(capKy) && data.chuc_vu_cap_tren) {
    lines.push(`${capKy}. ${data.chuc_vu_cap_tren}`);
  }

  lines.push(data.chuc_vu_ky || "NGƯỜI KÝ");

  const result = lines.map((line) =>
    paraText(line, { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 0, after: 60 } })
  );

  result.push(paraText("(Ký, ghi rõ họ tên)", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, italics: true, spacing: { before: 40, after: 500 } }));
  result.push(paraText(data.nguoi_ky || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 0, after: 120 } }));

  return result;
}

function noiNhanItems(data) {
  const noiNhan = Array.isArray(data.noi_nhan) && data.noi_nhan.length
    ? data.noi_nhan
    : ["- Như trên;", "- Lưu: VT."];

  const result = [
    paraText("Nơi nhận:", { noIndent: true, size: 24, bold: true, italics: true, spacing: { before: 120, after: 60 } }),
  ];
  noiNhan.forEach((item) => {
    result.push(paraText(item, { noIndent: true, size: 22, spacing: { before: 0, after: 20 } }));
  });
  return result;
}

function footerTable(data) {
  const leftFooter = noiNhanItems(data);
  const rightFooter = signatureBlock(data);

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: noBorders(),
            children: leftFooter,
          }),
          new TableCell({
            width: { size: 45, type: WidthType.PERCENTAGE },
            borders: noBorders(),
            children: rightFooter,
          }),
        ],
      }),
    ],
  });
}

function formatNgayThang(ngayThang) {
  if (!ngayThang) return ", ngày ... tháng ... năm ...";
  return String(ngayThang).replace(
    /ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})/,
    (_, d, m, y) => {
      const day = d.length === 1 ? `0${d}` : d;
      const month = m.length === 1 ? `0${m}` : m;
      return `ngày ${day} tháng ${month} năm ${y}`;
    }
  );
}

function buildDocument(children) {
  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Times New Roman",
            size: 28,
          },
          paragraph: { indent: { firstLine: 0 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: PAGE.marginTop,
              right: PAGE.marginRight,
              bottom: PAGE.marginBottom,
              left: PAGE.marginLeft,
            },
          },
        },
        children,
      },
    ],
  });
}

function validateRequired(data, fields) {
  const missing = fields.filter((key) => !data[key]);
  if (missing.length) {
    throw new Error(`Missing required fields: ${missing.join(", ")}`);
  }
}

module.exports = {
  AlignmentType,
  Table,
  TableCell,
  TableRow,
  WidthType,
  buildDocument,
  ensureEndingPunctuation,
  footerTable,
  formatNgayThang,
  headerTable,
  noBorders,
  noiNhanItems,
  paraText,
  parseCliArgs,
  readJsonFile,
  signatureBlock,
  validateRequired,
  writeDocx,
  topBorderParagraph,
};
