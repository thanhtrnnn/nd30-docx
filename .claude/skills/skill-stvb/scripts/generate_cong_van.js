const {
  AlignmentType,
  buildDocument,
  ensureEndingPunctuation,
  footerTable,
  headerTable,
  paraText,
  parseCliArgs,
  readJsonFile,
  validateRequired,
  writeDocx,
} = require("./common");

async function main() {
  try {
    const args = parseCliArgs(process.argv);
    const data = readJsonFile(args.input);

    validateRequired(data, [
      "co_quan_chu_quan",
      "co_quan_ban_hanh",
      "trich_yeu",
      "kinh_gui",
      "noi_dung",
      "chuc_vu_ky",
      "nguoi_ky",
    ]);

    const children = [];
    children.push(headerTable(data, { decision: false }));

    children.push(paraText("", { noIndent: true, spacing: { before: 80, after: 80 } }));

    const kinhGui = Array.isArray(data.kinh_gui) ? data.kinh_gui : [String(data.kinh_gui)];
    if (kinhGui.length === 1) {
      children.push(paraText(`Kính gửi: ${kinhGui[0]}`, { noIndent: true, spacing: { before: 80, after: 80 } }));
    } else {
      children.push(paraText("Kính gửi:", { noIndent: true, spacing: { before: 80, after: 30 } }));
      kinhGui.forEach((item, idx) => {
        const suffix = idx === kinhGui.length - 1 ? "." : ";";
        const text = item.startsWith("-") ? ensureEndingPunctuation(item, suffix) : `- ${ensureEndingPunctuation(item, suffix)}`;
        children.push(paraText(text, { noIndent: true, spacing: { before: 0, after: 30 } }));
      });
    }

    children.push(paraText("", { noIndent: true, spacing: { before: 60, after: 60 } }));

    const contentLines = String(data.noi_dung)
      .split(/\n+/)
      .map((x) => x.trim())
      .filter(Boolean);

    contentLines.forEach((line) => {
      children.push(paraText(ensureEndingPunctuation(line), { alignment: AlignmentType.JUSTIFIED }));
    });

    children.push(paraText("", { noIndent: true, spacing: { before: 80, after: 80 } }));
    children.push(footerTable(data));

    const doc = buildDocument(children);
    const out = await writeDocx(doc, args.output);
    console.log(`Generated cong van: ${out}`);
  } catch (error) {
    console.error(error.message || error);
    process.exit(1);
  }
}

main();
