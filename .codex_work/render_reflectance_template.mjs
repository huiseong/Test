import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = String.raw`C:\Users\lunarianmetal\Desktop\Test\outputs\reflectance-thickness-template-20261007\반사율_극값_두께_계산_템플릿.xlsx`;
const outputDir = String.raw`C:\Users\lunarianmetal\Desktop\Test\.codex_work\reflectance_previews`;
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const previews = [
  ["결과", "A1:N46"],
  ["설정", "A1:G38"],
  ["Nb2O5 입력", "A1:BD20"],
  ["2층 입력", "A1:BD20"],
  ["굴절률", "A1:D46"],
];
for (const [sheetName, range] of previews) {
  const blob = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  const fileName = `${sheetName.replaceAll(/[^a-zA-Z0-9가-힣_-]/g, "_")}.png`;
  await fs.writeFile(`${outputDir}\\${fileName}`, new Uint8Array(await blob.arrayBuffer()));
  console.log(`PREVIEW\t${sheetName}\t${outputDir}\\${fileName}`);
}
