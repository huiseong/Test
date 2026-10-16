import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = [
  {
    key: "measurement",
    path: String.raw`C:\Users\lunarianmetal\Documents\신진 필름개발팀\01. 설비 균일도, 마그넷 조정\2호기 균일도\20230919_RTR#2_Nb2O5, SiO2 균일도\20230919_RTR#2_Nb2O5, SiO2 균일도.xlsx`,
  },
  {
    key: "refractive",
    path: String.raw`C:\Users\lunarianmetal\Documents\신진 필름개발팀\01. 설비 균일도, 마그넷 조정\굴절률.xlsx`,
  },
];

const outDir = String.raw`C:\Users\lunarianmetal\Desktop\Test\.codex_work\thinfilm_source_previews`;
await fs.mkdir(outDir, { recursive: true });

for (const file of files) {
  const input = await FileBlob.load(file.path);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const inspection = await workbook.inspect({
    kind: "workbook,sheet,table,drawing,formula",
    maxChars: 12000,
    tableMaxRows: 8,
    tableMaxCols: 12,
    options: { maxResults: 80 },
  });
  console.log(`===== ${file.key} =====`);
  console.log(inspection.ndjson);
  const sheets = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
  const names = [...sheets.ndjson.matchAll(/"name":"([^"]+)"/g)].map((m) => m[1]);
  for (const sheetName of names) {
    const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 0.8, format: "png" });
    const safeName = sheetName.replaceAll(/[^a-zA-Z0-9가-힣_-]/g, "_");
    await fs.writeFile(`${outDir}/${file.key}_${safeName}.png`, new Uint8Array(await preview.arrayBuffer()));
    console.log(`PREVIEW\t${file.key}\t${sheetName}\t${outDir}/${file.key}_${safeName}.png`);
  }
}
