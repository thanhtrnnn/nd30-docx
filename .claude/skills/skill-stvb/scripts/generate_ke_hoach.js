const {
  AlignmentType,
  buildDocument,
  ensureEndingPunctuation,
  footerTable,
  headerTable,
  paraText,
  parseCliArgs,
  readJsonFile,
  topBorderParagraph,
  validateRequired,
  writeDocx,
} = require("./common");

function addLines(children, lines, opts = {}) {
  (lines || []).forEach((line) => {
    const normalized = opts.keepRaw ? String(line || "") : ensureEndingPunctuation(line);
    children.push(
      paraText(normalized, {
        alignment: opts.alignment || AlignmentType.JUSTIFIED,
        noIndent: Boolean(opts.noIndent),
        bold: Boolean(opts.bold),
        italics: Boolean(opts.italics),
        size: opts.size || 28,
        spacing: opts.spacing,
      })
    );
  });
}

async function main() {
  try {
    const args = parseCliArgs(process.argv);
    const data = readJsonFile(args.input);

    validateRequired(data, [
      "co_quan_chu_quan",
      "co_quan_ban_hanh",
      "trich_yeu",
      "noi_dung",
      "chuc_vu_ky",
      "nguoi_ky",
    ]);

    const children = [];
    children.push(headerTable(data, { decision: true }));

    children.push(
      paraText("KẾ HOẠCH", {
        bold: true,
        alignment: AlignmentType.CENTER,
        noIndent: true,
        spacing: { before: 220, after: 80 },
      })
    );
    children.push(
      paraText(data.trich_yeu, {
        bold: true,
        alignment: AlignmentType.CENTER,
        noIndent: true,
        spacing: { before: 0, after: 20 },
      })
    );
    children.push(topBorderParagraph(1700));

    (data.noi_dung || []).forEach((section) => {
      children.push(
        paraText(section.tieu_de || "", {
          bold: true,
          noIndent: true,
          spacing: { before: 120, after: 60 },
        })
      );

      addLines(children, section.doan_van, {
        alignment: AlignmentType.JUSTIFIED,
      });

      addLines(children, section.gach_dau_dong, {
        noIndent: true,
      });
    });

    children.push(paraText("", { noIndent: true, spacing: { before: 60, after: 60 } }));
    children.push(footerTable(data));

    const doc = buildDocument(children);
    const out = await writeDocx(doc, args.output);
    console.log(`Generated ke hoach: ${out}`);
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

main();
