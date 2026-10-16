import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const taskDir = path.dirname(new URL(import.meta.url).pathname.replace(/^\/(?=[A-Za-z]:)/, ""));
const inputPath = path.resolve(taskDir, "..", "..", "실습 데이터.xlsx");
const mode = process.argv[2] ?? "inspect";

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

if (mode === "inspect") {
  const summary = await workbook.inspect({
    kind: "workbook,sheet,table,definedName,drawing",
    include: "id,name,range,values,formulas",
    maxChars: 16000,
    tableMaxRows: 12,
    tableMaxCols: 30,
    tableMaxCellChars: 120,
  });
  console.log(summary.ndjson);

  const firstSheet = workbook.worksheets.getItemAt(0);
  const region = await workbook.inspect({
    kind: "region,formula,computedStyle",
    sheetId: firstSheet.name,
    range: "A1:AZ40",
    include: "values,formulas",
    maxChars: 24000,
    tableMaxRows: 40,
    tableMaxCols: 52,
    options: { maxResults: 300 },
  });
  console.log(region.ndjson);

  const preview = await workbook.render({
    sheetName: firstSheet.name,
    autoCrop: "all",
    scale: 1.5,
    format: "png",
  });
  const previewPath = path.join(taskDir, "source_preview.png");
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
  console.log(JSON.stringify({ previewPath, firstSheet: firstSheet.name }));
}
