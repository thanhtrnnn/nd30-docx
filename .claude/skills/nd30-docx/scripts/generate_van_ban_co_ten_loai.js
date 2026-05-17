const fs = require("fs");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  LineRuleType,
} = require("docx");

// --- CLI ---
function parseCliArgs(argv) {
  const args = { input: "data.json", output: "output.docx" };
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === "--input" && argv[i + 1]) args.input = argv[i + 1];
    if (argv[i] === "--output" && argv[i + 1]) args.output = argv[i + 1];
  }
  return args;
}

function readJsonFile(inputPath) {
  const absPath = require("path").resolve(process.cwd(), inputPath);
  if (!fs.existsSync(absPath)) throw new Error(`Input file not found: ${absPath}`);
  return JSON.parse(fs.readFileSync(absPath, "utf8"));
}

async function writeDocx(doc, outputPath) {
  const buffer = await Packer.toBuffer(doc);
  const absPath = require("path").resolve(process.cwd(), outputPath);
  fs.writeFileSync(absPath, buffer);
  return absPath;
}

function validateRequired(data, fields) {
  const missing = fields.filter((key) => !data[key]);
  if (missing.length) throw new Error(`Missing required fields: ${missing.join(", ")}`);
}

// --- Helpers ---
function noBorders() {
  return {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };
}

function paraText(text, opts = {}) {
  const indentSpec = opts.indent !== undefined
    ? opts.indent
    : opts.noIndent ? { firstLine: 0 } : { firstLine: 720 };
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

function topBorderParagraph(indent = 1100) {
  return new Paragraph({
    spacing: { before: 20, after: 0 },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: "000000", space: 1 } },
    indent: { left: indent, right: indent },
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

function ensureEndingPunctuation(text, fallback = ".") {
  const value = String(text || "").trim();
  if (!value) return "";
  if (/[\.;:!?]$/.test(value)) return value;
  return `${value}${fallback}`;
}

// --- Header Table ---
function headerTable(data) {
  const diaDanh = data.dia_danh || "Hà Nội";
  const soKyHieu = data.so_ky_hieu || "Số: ...../.....";
  const ngayThang = formatNgayThang(data.ngay_thang);

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 3500, type: WidthType.DXA },
            borders: noBorders(),
            children: [
              paraText(data.co_quan_chu_quan || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
              paraText(data.co_quan_ban_hanh || "", { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
              topBorderParagraph(1350),
            ],
          }),
          new TableCell({
            width: { size: 5571, type: WidthType.DXA },
            borders: noBorders(),
            children: [
              paraText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { alignment: AlignmentType.CENTER, noIndent: true, size: 26, bold: true }),
              paraText("Độc lập - Tự do - Hạnh phúc", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true }),
              topBorderParagraph(1100),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 3500, type: WidthType.DXA },
            borders: noBorders(),
            children: [
              paraText(soKyHieu, { alignment: AlignmentType.CENTER, noIndent: true, size: 26 }),
            ],
          }),
          new TableCell({
            width: { size: 5571, type: WidthType.DXA },
            borders: noBorders(),
            children: [
              paraText(`${diaDanh}${ngayThang}`, { alignment: AlignmentType.CENTER, noIndent: true, size: 28, italics: true }),
            ],
          }),
        ],
      }),
    ],
  });
}

// --- Footer Table ---
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
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            borders: noBorders(),
            children: noiNhanItems(data),
          }),
          new TableCell({
            width: { size: 45, type: WidthType.PERCENTAGE },
            borders: noBorders(),
            children: signatureBlock(data),
          }),
        ],
      }),
    ],
  });
}

// --- Main Generator ---
async function generateVanBanCoTenLoai(jsonPath, outputPath) {
  const data = readJsonFile(jsonPath);
  validateRequired(data, ["co_quan_chu_quan", "co_quan_ban_hanh", "ten_loai_van_ban", "noi_dung", "chuc_vu_ky", "nguoi_ky"]);

  const children = [];

  // Header
  children.push(headerTable(data));
  children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));

  // Tên loại văn bản
  children.push(paraText((data.ten_loai_van_ban || "").toUpperCase(), {
    alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 200, after: 120 },
  }));

  // Trích yếu
  if (data.trich_yeu) {
    children.push(paraText(data.trich_yeu, {
      alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 60, after: 200 },
    }));
  }

  // Căn cứ
  if (Array.isArray(data.can_cu) && data.can_cu.length) {
    data.can_cu.forEach((cc) => {
      children.push(paraText(cc, { alignment: AlignmentType.JUSTIFIED, size: 28, italics: true }));
    });
    children.push(new Paragraph({ spacing: { before: 60, after: 60 } }));
  }

  // Nội dung
  if (Array.isArray(data.noi_dung)) {
    data.noi_dung.forEach((block) => {
      if (typeof block === "string") {
        // Split by newlines
        block.split("\n").forEach((line) => {
          children.push(paraText(ensureEndingPunctuation(line), { alignment: AlignmentType.JUSTIFIED, size: 28 }));
        });
      } else if (block.loai === "dieu") {
        // Điều: bold title + content
        const pChildren = [];
        if (block.tieu_de) {
          pChildren.push(new TextRun({ text: block.tieu_de + " ", font: "Times New Roman", size: 28, bold: true }));
        }
        if (block.noi_dung) {
          pChildren.push(new TextRun({ text: ensureEndingPunctuation(block.noi_dung), font: "Times New Roman", size: 28 }));
        }
        children.push(new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { firstLine: 720 },
          spacing: { before: 120, after: 120, line: 340, lineRule: LineRuleType.EXACT },
          children: pChildren,
        }));
      } else if (block.loai === "muc") {
        // Mục: uppercase bold centered
        children.push(paraText((block.tieu_de || "").toUpperCase(), {
          alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true, spacing: { before: 200, after: 120 },
        }));
        if (block.noi_dung) {
          children.push(paraText(ensureEndingPunctuation(block.noi_dung), { alignment: AlignmentType.JUSTIFIED, size: 28 }));
        }
      } else if (block.loai === "khoan") {
        children.push(paraText(ensureEndingPunctuation(block.noi_dung || ""), { alignment: AlignmentType.JUSTIFIED, size: 28 }));
      }
    });
  } else if (typeof data.noi_dung === "string") {
    data.noi_dung.split("\n").forEach((line) => {
      children.push(paraText(ensureEndingPunctuation(line), { alignment: AlignmentType.JUSTIFIED, size: 28 }));
    });
  }

  // Dấu ./.
  if (data.ket_thuc_dot) {
    children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));
    children.push(paraText("./.", { alignment: AlignmentType.CENTER, noIndent: true, size: 28, bold: true }));
  }

  children.push(new Paragraph({ spacing: { before: 200, after: 0 } }));

  // Footer
  children.push(footerTable(data));

  // Build document
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Times New Roman", size: 28 },
          paragraph: { indent: { firstLine: 0 } },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 1134, right: 851, bottom: 1134, left: 1701 },
          size: { width: 11906, height: 16838 },
        },
      },
      children,
    }],
  });

  await writeDocx(doc, outputPath);
  console.log(`Generated: ${outputPath}`);
}

// --- CLI ---
const args = parseCliArgs(process.argv);
generateVanBanCoTenLoai(args.input, args.output).catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
