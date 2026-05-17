// Báo cáo: Nhóm A - wrapper (có thể có Kính gửi)
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

function parseCliArgs(argv) {
  const args = { input: "data.json", output: "bao_cao.docx" };
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === "--input" && argv[i + 1]) args.input = argv[i + 1];
    if (argv[i] === "--output" && argv[i + 1]) args.output = argv[i + 1];
  }
  return args;
}

const args = parseCliArgs(process.argv);
const absInput = path.resolve(process.cwd(), args.input);
const data = JSON.parse(fs.readFileSync(absInput, "utf8"));

if (!data.ten_loai_van_ban) data.ten_loai_van_ban = "BÁO CÁO";

const tmpFile = path.join(path.dirname(absInput), "_tmp_bao_cao.json");
fs.writeFileSync(tmpFile, JSON.stringify(data, null, 2));

try {
  const script = path.join(__dirname, "generate_van_ban_co_ten_loai.js");
  execSync(`node "${script}" --input "${tmpFile}" --output "${args.output}"`, { stdio: "inherit" });
} finally {
  if (fs.existsSync(tmpFile)) fs.unlinkSync(tmpFile);
}
