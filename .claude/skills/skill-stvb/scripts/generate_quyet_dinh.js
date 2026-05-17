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

function buildNoiDungQuyetDinh(noiDung) {
  const items = [];

  if (Array.isArray(noiDung)) {
    noiDung.forEach((dieu) => {
      const tieuDe = dieu.tieu_de || "Điều ...";
      items.push(paraText(tieuDe, { bold: true, noIndent: true, spacing: { before: 80, after: 40 } }));
      items.push(paraText(ensureEndingPunctuation(dieu.noi_dung || ""), { alignment: AlignmentType.JUSTIFIED }));
    });
    return items;
  }

  const lines = String(noiDung || "")
    .split(/\n+/)
    .map((x) => x.trim())
    .filter(Boolean);

  lines.forEach((line) => {
    const isDieu = /^Đi[eề]u\s+\d+/i.test(line) || /^Dieu\s+\d+/i.test(line);
    items.push(paraText(isDieu ? line : ensureEndingPunctuation(line), { alignment: AlignmentType.JUSTIFIED, bold: isDieu, noIndent: isDieu }));
  });

  return items;
}

async function main() {
  try {
    const args = parseCliArgs(process.argv);
    const data = readJsonFile(args.input);

    validateRequired(data, [
      "co_quan_chu_quan",
      "co_quan_ban_hanh",
      "trich_yeu",
      "can_cu",
      "noi_dung",
      "chuc_vu_ky",
      "nguoi_ky",
    ]);

    const children = [];
    children.push(headerTable(data, { decision: true }));

    children.push(paraText("QUYẾT ĐỊNH", { bold: true, alignment: AlignmentType.CENTER, noIndent: true, size: 28, spacing: { before: 240, after: 80 } }));
    children.push(paraText(data.trich_yeu, { bold: true, alignment: AlignmentType.CENTER, noIndent: true, spacing: { before: 0, after: 20 } }));
    children.push(topBorderParagraph(1700));

    if (data.chu_the_ban_hanh) {
      children.push(paraText(data.chu_the_ban_hanh, { bold: true, alignment: AlignmentType.CENTER, noIndent: true, size: 28, spacing: { before: 160, after: 80 } }));
    }

    children.push(paraText("", { noIndent: true, spacing: { before: 80, after: 80 } }));

    const canCu = Array.isArray(data.can_cu) ? data.can_cu : [String(data.can_cu)];
    canCu.forEach((item, idx) => {
      const suffix = idx === canCu.length - 1 ? "." : ";";
      const text = item.startsWith("Căn cứ") ? `${item}${suffix}` : `Căn cứ ${item}${suffix}`;
      children.push(paraText(text, { italics: true, noIndent: true }));
    });

    if (data.theo_de_nghi) {
      children.push(paraText(ensureEndingPunctuation(data.theo_de_nghi), { italics: true, noIndent: true, spacing: { before: 60, after: 60 } }));
    }

    children.push(paraText("QUYẾT ĐỊNH:", { bold: true, alignment: AlignmentType.CENTER, noIndent: true, spacing: { before: 120, after: 120 } }));

    buildNoiDungQuyetDinh(data.noi_dung).forEach((p) => children.push(p));

    children.push(paraText("", { noIndent: true, spacing: { before: 80, after: 80 } }));
    children.push(footerTable(data));

    const doc = buildDocument(children);
    const out = await writeDocx(doc, args.output);
    console.log(`Generated quyet dinh: ${out}`);
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

main();
