import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = String.raw`C:\Users\lunarianmetal\Desktop\Test\outputs\reflectance-thickness-template-20261007\반사율_극값_두께_계산_템플릿.xlsx`;
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));

const sheetCheck = await workbook.inspect({ kind: "sheet,drawing", include: "id,name", maxChars: 8000 });
console.log("===== SHEETS AND DRAWINGS =====");
console.log(sheetCheck.ndjson);

const keyRanges = [
  ["결과", "A2:O24"],
  ["설정", "A2:G38"],
  ["Nb2O5 입력", "AU5:BD18"],
  ["2층 입력", "AU5:BD18"],
  ["굴절률", "A2:D46"],
  ["TMM 계산표", "A2:I12"],
];
for (const [sheetName, range] of keyRanges) {
  const check = await workbook.inspect({
    kind: "table",
    range: `${sheetName}!${range}`,
    include: "values,formulas",
    tableMaxRows: sheetName === "굴절률" ? 12 : 40,
    tableMaxCols: 14,
    maxChars: 16000,
  });
  console.log(`===== ${sheetName}!${range} =====`);
  console.log(check.ndjson);
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "saved workbook formula error scan",
});
console.log("===== ERRORS =====");
console.log(errors.ndjson);

const nb = workbook.worksheets.getItem("Nb2O5 입력");
const settings = workbook.worksheets.getItem("설정");
const results = workbook.worksheets.getItem("결과");
const before = Number(results.getRange("G12").values[0][0]);
nb.getRange("Z6").values = [[1]];
workbook.recalculate();
const after = Number(results.getRange("G12").values[0][0]);
console.log(`RECALC_TEST\tbefore=${before}\tafter=${after}`);
if (!Number.isFinite(before) || !Number.isFinite(after) || Math.abs(before - after) < 0.01) {
  throw new Error("Representative input-change recalculation test did not change Nb2O5 thickness as expected.");
}

const sioBefore = Number(results.getRange("N12").values[0][0]);
settings.getRange("C16").values = [[8]];
workbook.recalculate();
const sioAfter = Number(results.getRange("N12").values[0][0]);
console.log(`TMM_RECALC_TEST\tbefore=${sioBefore}\tafter=${sioAfter}`);
if (!Number.isFinite(sioBefore) || !Number.isFinite(sioAfter) || Math.abs(sioBefore - sioAfter) < 0.1) {
  throw new Error("Representative Nb2O5 underlayer change did not update the SiO2 TMM thickness as expected.");
}
