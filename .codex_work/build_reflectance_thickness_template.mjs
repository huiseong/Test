import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const measurementPath = String.raw`C:\Users\lunarianmetal\Documents\신진 필름개발팀\01. 설비 균일도, 마그넷 조정\2호기 균일도\20230919_RTR#2_Nb2O5, SiO2 균일도\20230919_RTR#2_Nb2O5, SiO2 균일도.xlsx`;
const refractivePath = String.raw`C:\Users\lunarianmetal\Documents\신진 필름개발팀\01. 설비 균일도, 마그넷 조정\굴절률.xlsx`;
const outputDir = String.raw`C:\Users\lunarianmetal\Desktop\Test\outputs\reflectance-thickness-template-20261007`;
const outputPath = `${outputDir}\\반사율_극값_두께_계산_템플릿.xlsx`;

const fontName = "Arial";
const colors = {
  navy: "#1F4E78",
  blue: "#5B9BD5",
  lightBlue: "#D9EAF7",
  paleBlue: "#EAF3F8",
  input: "#FFF2CC",
  inputDark: "#FFD966",
  gray: "#F2F2F2",
  midGray: "#D9E1F2",
  line: "#D0D7DE",
  text: "#1F2937",
  muted: "#5F6B7A",
  redFill: "#FCE8E6",
  redText: "#B91C1C",
  amberFill: "#FFF4CE",
  amberText: "#8A5A00",
  greenFill: "#E2F0D9",
  greenText: "#38761D",
};

function setFont(range, size = 10, options = {}) {
  range.format.font = {
    name: fontName,
    size,
    color: options.color ?? colors.text,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
  };
  range.format.verticalAlignment = "center";
}

function styleTitle(sheet, endCol) {
  const title = sheet.getRange(`A2:${endCol}2`);
  setFont(title, 15, { bold: true, color: colors.navy });
  title.format.rowHeight = 24;
  const rule = sheet.getRange(`A3:${endCol}3`);
  rule.format.borders = { bottom: { style: "thin", color: colors.blue } };
}

function styleHeader(range) {
  range.format.fill = colors.navy;
  range.format.font = { name: fontName, size: 10, bold: true, color: "#FFFFFF" };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.borders = {
    insideVertical: { style: "thin", color: "#FFFFFF" },
    bottom: { style: "thin", color: colors.navy },
  };
}

function styleSection(range) {
  range.format.fill = colors.lightBlue;
  range.format.font = { name: fontName, size: 10, bold: true, color: colors.navy };
  range.format.borders = { preset: "outside", style: "thin", color: colors.blue };
  range.format.verticalAlignment = "center";
}

function styleBody(range) {
  setFont(range, 10);
  range.format.borders = {
    insideHorizontal: { style: "thin", color: colors.line },
    bottom: { style: "thin", color: colors.line },
  };
}

function addStatusFormatting(range) {
  range.conditionalFormats.add("containsText", {
    text: "오류",
    format: { fill: colors.redFill, font: { color: colors.redText, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "필요",
    format: { fill: colors.redFill, font: { color: colors.redText, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "확인",
    format: { fill: colors.amberFill, font: { color: colors.amberText, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "복수",
    format: { fill: colors.amberFill, font: { color: colors.amberText, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "완료",
    format: { fill: colors.greenFill, font: { color: colors.greenText } },
  });
}

function minValueFormula(row, settingsRow) {
  return `=IF(OR('설정'!$E$${settingsRow}>'설정'!$F$${settingsRow},COUNT(H${row}:AT${row})=0),"",MIN(BE${row}:CQ${row}))`;
}

function maxValueFormula(row, settingsRow) {
  return `=IF(OR('설정'!$E$${settingsRow}>'설정'!$F$${settingsRow},COUNT(H${row}:AT${row})=0),"",MAX(BE${row}:CQ${row}))`;
}

function extremumWavelengthFormula(row, settingsRow, extremeCell) {
  return `=IF(${extremeCell}${row}="","",SUMIFS($BE$5:$CQ$5,BE${row}:CQ${row},${extremeCell}${row})/COUNTIFS(BE${row}:CQ${row},${extremeCell}${row}))`;
}

function refractiveFormula(row, nColumn) {
  return `=IF(AZ${row}="","",IF(OR(AZ${row}<'굴절률'!$A$6,AZ${row}>'굴절률'!$A$44),"",IF(COUNTIFS('굴절률'!$A$6:$A$44,AZ${row})>0,INDEX('굴절률'!$${nColumn}$6:$${nColumn}$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,0)),INDEX('굴절률'!$${nColumn}$6:$${nColumn}$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1))+(AZ${row}-INDEX('굴절률'!$A$6:$A$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1)))*(INDEX('굴절률'!$${nColumn}$6:$${nColumn}$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1)+1)-INDEX('굴절률'!$${nColumn}$6:$${nColumn}$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1)))/(INDEX('굴절률'!$A$6:$A$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1)+1)-INDEX('굴절률'!$A$6:$A$44,MATCH(AZ${row},'굴절률'!$A$6:$A$44,1))))))`;
}

function tmmThicknessFormula(row, positionRow) {
  const nb = `'설정'!$C$${positionRow}`;
  const wavelengthGrid = `'TMM 계산표'!$A$6:$A$3806`;
  const nbHeaders = `'TMM 계산표'!$B$5:$I$5`;
  const thicknessTable = `'TMM 계산표'!$B$6:$I$3806`;
  const wavelengthRow = `MATCH(AZ${row},${wavelengthGrid},1)`;
  const exactNbColumn = `MATCH(${nb},${nbHeaders},0)`;
  const lowerNbColumn = `MATCH(${nb},${nbHeaders},1)`;
  const lowerThickness = `INDEX(${thicknessTable},${wavelengthRow},${lowerNbColumn})`;
  const upperThickness = `INDEX(${thicknessTable},${wavelengthRow},${lowerNbColumn}+1)`;
  const lowerNb = `INDEX(${nbHeaders},1,${lowerNbColumn})`;
  const upperNb = `INDEX(${nbHeaders},1,${lowerNbColumn}+1)`;
  return `=IF(OR(AZ${row}="",${nb}="",${nb}<8,${nb}>15),"",IF(COUNTIFS(${nbHeaders},${nb})>0,INDEX(${thicknessTable},${wavelengthRow},${exactNbColumn}),${lowerThickness}+(${nb}-${lowerNb})*(${upperThickness}-${lowerThickness})/(${upperNb}-${lowerNb})))`;
}

function statusFormula(row, settingsRow, stackApproximation = false, positionRow = null) {
  const nbGuard = stackApproximation
    ? `IF(OR('설정'!$C$${positionRow}="",'설정'!$C$${positionRow}<8,'설정'!$C$${positionRow}>15),"Nb 두께 입력 확인",`
    : "";
  const nbClose = stackApproximation ? ")" : "";
  return `=IF(COUNT(H${row}:AT${row})=0,"입력 필요",IF(COUNT(H${row}:AT${row})<>39,"데이터 점검",${nbGuard}IF(OR('설정'!$E$${settingsRow}>'설정'!$F$${settingsRow},'설정'!$E$${settingsRow}<360,'설정'!$F$${settingsRow}>740),"검색 범위 오류",IF(AY${row}<>"","수동 파장",IF(OR(AZ${row}='설정'!$E$${settingsRow},AZ${row}='설정'!$F$${settingsRow}),"경계 극값 확인","완료")))${nbClose}))`;
}

function createInputSheet(sheet, options) {
  const { title, sourceRows, settingsRow, nColumn, stackApproximation } = options;
  sheet.showGridLines = false;
  sheet.tabColor = stackApproximation ? "#70AD47" : colors.blue;
  sheet.getRange("A2").values = [[title]];
  styleTitle(sheet, "BD");
  sheet.getRange("A3").values = [["CM-36dG 원본 13개 행(A:AT)을 노란색 영역에 붙여넣으면 극값과 두께가 자동 계산됩니다."]];
  setFont(sheet.getRange("A3:BD3"), 9, { italic: true, color: colors.muted });

  const wavelengths = Array.from({ length: 39 }, (_, i) => 360 + i * 10);
  const rawHeaders = [
    "시료", "Data Name", "L*(D65)", "a*(D65)", "b*(D65)", "Y(D65)", "Group Traits",
    ...wavelengths,
  ];
  const helperHeaders = [
    "최소 R(%)", "최소 λ(nm)", "최대 R(%)", "최대 λ(nm)", "수동 λ(nm)",
    "사용 λ(nm)", stackApproximation ? "SiO2 n" : "굴절률 n", stackApproximation ? "하부 Nb2O5(nm)" : "cosθf", stackApproximation ? "SiO2 TMM 두께(nm)" : "Nb2O5 추정 두께(nm)", "상태",
  ];
  sheet.getRange("A5:BD5").values = [[...rawHeaders, ...helperHeaders]];
  styleHeader(sheet.getRange("A5:BD5"));
  sheet.getRange("H5:AT5").format.numberFormat = `0"nm"`;
  sheet.getRange("A6:AT18").values = sourceRows;
  sheet.getRange("A6:AT18").format.fill = colors.input;
  styleBody(sheet.getRange("A6:BD18"));
  sheet.getRange("A6:BD18").format.rowHeight = 20;
  sheet.getRange("A6:A18").format.horizontalAlignment = "center";
  sheet.getRange("C6:F18").format.numberFormat = "0.00";
  sheet.getRange("H6:AT18").format.numberFormat = "0.00";
  sheet.getRange("AU6:AX18").format.numberFormat = "0.00";
  sheet.getRange("AY6:AY18").format.fill = colors.input;
  sheet.getRange("AY6:AZ18").format.numberFormat = "0.0";
  sheet.getRange("BA6:BB18").format.numberFormat = "0.00000";
  sheet.getRange("BC6:BC18").format.numberFormat = "0.00";
  sheet.getRange("AU6:BD18").format.horizontalAlignment = "right";
  sheet.getRange("BD6:BD18").format.horizontalAlignment = "left";

  sheet.getRange("BE5:CQ5").values = [wavelengths];
  sheet.getRange("BE5:CQ18").format.fill = colors.gray;
  setFont(sheet.getRange("BE5:CQ18"), 8, { color: colors.muted });
  sheet.getRange("BE5:CQ18").format.numberFormat = "0.00";
  sheet.getRange("BE6").formulas = [[`=IF(AND(BE$5>='설정'!$E$${settingsRow},BE$5<='설정'!$F$${settingsRow}),H6,"")`]];
  sheet.getRange("BE6:CQ6").fillRight();
  sheet.getRange("BE6:CQ18").fillDown();

  const row = 6;
  sheet.getRange(`AU${row}:AX${row}`).formulas = [[
    minValueFormula(row, settingsRow),
    extremumWavelengthFormula(row, settingsRow, "AU"),
    maxValueFormula(row, settingsRow),
    extremumWavelengthFormula(row, settingsRow, "AW"),
  ]];
  sheet.getRange("AU6:AX18").fillDown();
  sheet.getRange(`AZ${row}:BD${row}`).formulas = [[
    `=IF(AY${row}<>"",AY${row},IF('설정'!$C$${settingsRow}="최소",AV${row},AX${row}))`,
    refractiveFormula(row, nColumn),
    stackApproximation
      ? `='설정'!$C$${10 + row}`
      : `=IF(BA${row}="","",COS(ASIN(SIN(RADIANS('설정'!$G$${settingsRow}))/BA${row})))`,
    stackApproximation
      ? tmmThicknessFormula(row, 10 + row)
      : `=IF(OR(AZ${row}="",BA${row}="",BB${row}=""),"",IF('설정'!$C$${settingsRow}="최소",'설정'!$D$${settingsRow}*AZ${row}/(2*BA${row}*BB${row}),(2*'설정'!$D$${settingsRow}-1)*AZ${row}/(4*BA${row}*BB${row})))`,
    statusFormula(row, settingsRow, stackApproximation, 10 + row),
  ]];
  sheet.getRange("AZ6:BD18").fillDown();

  sheet.getRange("AY6:AY18").dataValidation = {
    rule: { type: "whole", operator: "between", formula1: 360, formula2: 740 },
  };
  addStatusFormatting(sheet.getRange("BD6:BD18"));

  sheet.getRange("A20").values = [["참고: 원본의 13-1 재측정 행은 13개 본 위치 결과에서 제외했습니다."]];
  setFont(sheet.getRange("A20:BD20"), 9, { italic: true, color: colors.muted });

  sheet.getRange("A1:A20").format.columnWidth = 8;
  sheet.getRange("B1:B20").format.columnWidth = 21;
  sheet.getRange("C1:F20").format.columnWidth = 10;
  sheet.getRange("G1:G20").format.columnWidth = 12;
  sheet.getRange("H1:AT20").format.columnWidth = 8;
  sheet.getRange("AU1:AX20").format.columnWidth = 12;
  sheet.getRange("AY1:AZ20").format.columnWidth = 12;
  sheet.getRange("BA1:BC20").format.columnWidth = 14;
  sheet.getRange("BD1:BD20").format.columnWidth = 18;
  sheet.getRange("BE1:CQ20").format.columnWidth = 3;
  sheet.freezePanes.freezeRows(5);
  sheet.freezePanes.freezeColumns(7);
}

function interpolate(xs, ys, x) {
  const exactIndex = xs.indexOf(x);
  if (exactIndex >= 0) return ys[exactIndex];
  let lower = 0;
  while (lower + 1 < xs.length && xs[lower + 1] < x) lower += 1;
  const t = (x - xs[lower]) / (xs[lower + 1] - xs[lower]);
  return ys[lower] + t * (ys[lower + 1] - ys[lower]);
}

function tmmReflectance(wavelength, sioThickness, nbThickness) {
  const n0 = 1;
  const n1 = interpolate(wavelengths, sioIndex, wavelength);
  const n2 = interpolate(wavelengths, nbIndex, wavelength);
  const n3 = interpolate(wavelengths, petIndex, wavelength);
  const r01 = (n0 - n1) / (n0 + n1);
  const r12 = (n1 - n2) / (n1 + n2);
  const r23 = (n2 - n3) / (n2 + n3);
  const phase1 = 4 * Math.PI * n1 * sioThickness / wavelength;
  const phase2 = 4 * Math.PI * n2 * nbThickness / wavelength;

  const c2 = Math.cos(phase2);
  const s2 = Math.sin(phase2);
  const numRe = r12 + r23 * c2;
  const numIm = r23 * s2;
  const denRe = 1 + r12 * r23 * c2;
  const denIm = r12 * r23 * s2;
  const denAbs2 = denRe * denRe + denIm * denIm;
  const stackRe = (numRe * denRe + numIm * denIm) / denAbs2;
  const stackIm = (numIm * denRe - numRe * denIm) / denAbs2;

  const c1 = Math.cos(phase1);
  const s1 = Math.sin(phase1);
  const propagatedRe = stackRe * c1 - stackIm * s1;
  const propagatedIm = stackRe * s1 + stackIm * c1;
  const totalNumRe = r01 + propagatedRe;
  const totalNumIm = propagatedIm;
  const totalDenRe = 1 + r01 * propagatedRe;
  const totalDenIm = r01 * propagatedIm;
  return (totalNumRe * totalNumRe + totalNumIm * totalNumIm)
    / (totalDenRe * totalDenRe + totalDenIm * totalDenIm);
}

function invertTmmMinimum(targetWavelength, nbThickness) {
  const nSio = interpolate(wavelengths, sioIndex, targetWavelength);
  const phaseGuess = 1.18 * targetWavelength / (4 * nSio);
  const slope = (sioThickness) => {
    const h = 0.02;
    if (targetWavelength <= wavelengths[0]) {
      return (tmmReflectance(targetWavelength + h, sioThickness, nbThickness)
        - tmmReflectance(targetWavelength, sioThickness, nbThickness)) / h;
    }
    if (targetWavelength >= wavelengths[wavelengths.length - 1]) {
      return (tmmReflectance(targetWavelength, sioThickness, nbThickness)
        - tmmReflectance(targetWavelength - h, sioThickness, nbThickness)) / h;
    }
    return (tmmReflectance(targetWavelength + h, sioThickness, nbThickness)
      - tmmReflectance(targetWavelength - h, sioThickness, nbThickness)) / (2 * h);
  };

  const scanStart = Math.max(20, phaseGuess - 18);
  const scanEnd = phaseGuess + 18;
  const scanStep = 0.25;
  const brackets = [];
  let previousX = scanStart;
  let previousY = slope(previousX);
  for (let x = scanStart + scanStep; x <= scanEnd + 1e-9; x += scanStep) {
    const y = slope(x);
    if (previousY === 0 || y === 0 || previousY * y < 0) brackets.push([previousX, x]);
    previousX = x;
    previousY = y;
  }
  if (brackets.length === 0) throw new Error(`TMM inverse root not found: lambda=${targetWavelength}, Nb=${nbThickness}`);

  const roots = brackets.map(([left, right]) => {
    let a = left;
    let b = right;
    let fa = slope(a);
    for (let i = 0; i < 45; i += 1) {
      const mid = (a + b) / 2;
      const fm = slope(mid);
      if (fa * fm <= 0) {
        b = mid;
      } else {
        a = mid;
        fa = fm;
      }
    }
    return (a + b) / 2;
  });
  return roots.reduce((best, value) => (
    Math.abs(value - phaseGuess) < Math.abs(best - phaseGuess) ? value : best
  ));
}

function expectedRow(rawRow, wavelengths, refractiveValues, settings, stackApproximation) {
  const values = rawRow.slice(7, 46).map(Number);
  const candidates = values
    .map((value, i) => ({ value, wavelength: wavelengths[i] }))
    .filter((p) => p.wavelength >= settings.start && p.wavelength <= settings.end);
  const extremeValue = settings.extreme === "최소"
    ? Math.min(...candidates.map((p) => p.value))
    : Math.max(...candidates.map((p) => p.value));
  const tied = candidates.filter((p) => p.value === extremeValue);
  const wavelength = tied.reduce((sum, p) => sum + p.wavelength, 0) / tied.length;
  const n = interpolate(wavelengths, refractiveValues, wavelength);
  const cosTheta = Math.cos(Math.asin(Math.sin(settings.angle * Math.PI / 180) / n));
  const p = settings.order;
  const thickness = stackApproximation
    ? (settings.extreme === "최소" ? (2 * p - 1) * wavelength / (4 * n * cosTheta) : p * wavelength / (2 * n * cosTheta))
    : (settings.extreme === "최소" ? p * wavelength / (2 * n * cosTheta) : (2 * p - 1) * wavelength / (4 * n * cosTheta));
  return { wavelength, n, cosTheta, thickness };
}

const measurementBook = await SpreadsheetFile.importXlsx(await FileBlob.load(measurementPath));
const measurementSheet = measurementBook.worksheets.getItem("Sheet1");
const stackRows = measurementSheet.getRange("A3:AT15").values;
const nbRows = measurementSheet.getRange("A19:AT31").values;

const refractiveBook = await SpreadsheetFile.importXlsx(await FileBlob.load(refractivePath));
const petRows = refractiveBook.worksheets.getItem("PET_TOYOBO").getRange("A2:B40").values;
const nbIndexRows = refractiveBook.worksheets.getItem("Nb2O5").getRange("A2:B40").values;
const sioIndexRows = refractiveBook.worksheets.getItem("SiO2").getRange("A2:B40").values;
const wavelengths = petRows.map((row) => Number(row[0]));
const petIndex = petRows.map((row) => Number(row[1]));
const nbIndex = nbIndexRows.map((row) => Number(row[1]));
const sioIndex = sioIndexRows.map((row) => Number(row[1]));

const workbook = Workbook.create();
const results = workbook.worksheets.add("결과");
const settings = workbook.worksheets.add("설정");
const nbInput = workbook.worksheets.add("Nb2O5 입력");
const stackInput = workbook.worksheets.add("2층 입력");
const refractive = workbook.worksheets.add("굴절률");
const tmmTable = workbook.worksheets.add("TMM 계산표");

results.showGridLines = false;
results.tabColor = colors.navy;
results.getRange("A2").values = [["반사율 극값 기반 두께 추정"]];
styleTitle(results, "O");
results.getRange("A3").values = [["Nb2O5 단층 두께와 Nb2O5/SiO2 적층의 SiO2 TMM 역산 두께를 13개 위치에서 계산합니다."]];
setFont(results.getRange("A3:O3"), 9, { italic: true, color: colors.muted });

results.getRange("A4:B4").values = [["Nb2O5 (1L)", "값"]];
results.getRange("D4:E4").values = [["SiO2 TMM 역산", "값"]];
styleSection(results.getRange("A4:B4"));
styleSection(results.getRange("D4:E4"));
results.getRange("A5:A9").values = [["평균 (nm)"], ["최소 (nm)"], ["최대 (nm)"], ["P-P (nm)"], ["± 균일도"]];
results.getRange("D5:D9").values = [["평균 (nm)"], ["최소 (nm)"], ["최대 (nm)"], ["P-P (nm)"], ["± 균일도"]];
results.getRange("B5:B9").formulas = [[
  `=IF(COUNT(G12:G24)=0,"",AVERAGE(G12:G24))`,
], [
  `=IF(COUNT(G12:G24)=0,"",MIN(G12:G24))`,
], [
  `=IF(COUNT(G12:G24)=0,"",MAX(G12:G24))`,
], [
  `=IF(COUNT(G12:G24)=0,"",MAX(G12:G24)-MIN(G12:G24))`,
], [
  `=IF(COUNT(G12:G24)=0,"",(MAX(G12:G24)-MIN(G12:G24))/(2*AVERAGE(G12:G24)))`,
]];
results.getRange("E5:E9").formulas = [[
  `=IF(COUNT(N12:N24)=0,"",AVERAGE(N12:N24))`,
], [
  `=IF(COUNT(N12:N24)=0,"",MIN(N12:N24))`,
], [
  `=IF(COUNT(N12:N24)=0,"",MAX(N12:N24))`,
], [
  `=IF(COUNT(N12:N24)=0,"",MAX(N12:N24)-MIN(N12:N24))`,
], [
  `=IF(COUNT(N12:N24)=0,"",(MAX(N12:N24)-MIN(N12:N24))/(2*AVERAGE(N12:N24)))`,
]];
styleBody(results.getRange("A5:B9"));
styleBody(results.getRange("D5:E9"));
results.getRange("B5:B8").format.numberFormat = "0.00";
results.getRange("E5:E8").format.numberFormat = "0.00";
results.getRange("B9").format.numberFormat = "0.00%";
results.getRange("E9").format.numberFormat = "0.00%";

const resultHeaders = [
  "시료", "위치(mm)", "Nb 최소 λ", "Nb 최대 λ", "Nb 사용 λ", "Nb n", "Nb 두께(nm)", "Nb 상태",
  "2L 최소 λ", "2L 최대 λ", "2L 사용 λ", "SiO2 n", "하부 Nb2O5(nm)", "SiO2 TMM 두께(nm)", "2L 상태",
];
results.getRange("A11:O11").values = [resultHeaders];
styleHeader(results.getRange("A11:O11"));
results.getRange("A12:O24").formulas = Array.from({ length: 13 }, (_, i) => {
  const inputRow = 6 + i;
  const positionRow = 16 + i;
  return [
    `='Nb2O5 입력'!A${inputRow}`, `='설정'!B${positionRow}`, `='Nb2O5 입력'!AV${inputRow}`, `='Nb2O5 입력'!AX${inputRow}`, `='Nb2O5 입력'!AZ${inputRow}`, `='Nb2O5 입력'!BA${inputRow}`, `='Nb2O5 입력'!BC${inputRow}`, `='Nb2O5 입력'!BD${inputRow}`,
    `='2층 입력'!AV${inputRow}`, `='2층 입력'!AX${inputRow}`, `='2층 입력'!AZ${inputRow}`, `='2층 입력'!BA${inputRow}`, `='2층 입력'!BB${inputRow}`, `='2층 입력'!BC${inputRow}`, `='2층 입력'!BD${inputRow}`,
  ];
});
styleBody(results.getRange("A12:O24"));
results.getRange("A12:B24").format.numberFormat = "0";
results.getRange("C12:E24").format.numberFormat = "0.0";
results.getRange("F12:F24").format.numberFormat = "0.00000";
results.getRange("G12:G24").format.numberFormat = "0.00";
results.getRange("I12:K24").format.numberFormat = "0.0";
results.getRange("L12:L24").format.numberFormat = "0.00000";
results.getRange("M12:N24").format.numberFormat = "0.00";
results.getRange("A12:G24").format.horizontalAlignment = "right";
results.getRange("H12:H24").format.horizontalAlignment = "left";
results.getRange("I12:N24").format.horizontalAlignment = "right";
results.getRange("O12:O24").format.horizontalAlignment = "left";
addStatusFormatting(results.getRange("H12:H24"));
addStatusFormatting(results.getRange("O12:O24"));

const chart = results.charts.add("line", [
  results.getRange("B11:B24"),
  results.getRange("G11:G24"),
  results.getRange("N11:N24"),
]);
chart.title = "위치별 추정 두께";
chart.titleTextStyle.fontSize = 12;
chart.titleTextStyle.typeface = fontName;
chart.legend = { position: "top", textStyle: { typeface: fontName } };
chart.xAxis = { axisType: "textAxis", textStyle: { typeface: fontName, fontSize: 10 } };
chart.yAxis = { numberFormatCode: "0.0", numberFormatSourceLinked: false, textStyle: { typeface: fontName }, title: { text: "두께 (nm)" } };
chart.xAxis.title.text = "위치 (mm)";
chart.yAxis.title.text = "두께 (nm)";
chart.setPosition("A27", "O46");
if (chart.series.items.length >= 2) {
  chart.series.items[0].fill = colors.navy;
  chart.series.items[1].fill = "#70AD47";
}

results.getRange("A1:A46").format.columnWidth = 8;
results.getRange("B1:B46").format.columnWidth = 11;
results.getRange("C1:G46").format.columnWidth = 12;
results.getRange("H1:H46").format.columnWidth = 18;
results.getRange("I1:M46").format.columnWidth = 13;
results.getRange("N1:N46").format.columnWidth = 18;
results.getRange("O1:O46").format.columnWidth = 18;

settings.showGridLines = false;
settings.tabColor = colors.blue;
settings.getRange("A2").values = [["계산 설정"]];
styleTitle(settings, "G");
settings.getRange("A3").values = [["노란색 셀을 조정하면 13개 위치의 결과가 즉시 다시 계산됩니다."]];
setFont(settings.getRange("A3:G3"), 9, { italic: true, color: colors.muted });
settings.getRange("A5:G5").values = [["계산 대상", "광학 모델", "극값", "차수 p", "검색 시작(nm)", "검색 끝(nm)", "입사각 θ0(°)"]];
styleHeader(settings.getRange("A5:G5"));
settings.getRange("A7:G8").values = [
  ["Nb2O5 (1L)", "air/Nb2O5/PET · 위상반전 1회", "최소", 1, 430, 600, 0],
  ["SiO2 (2L)", "air/SiO2/Nb2O5/PET · TMM 역산", "최소", 1, 430, 550, 0],
];
styleBody(settings.getRange("A7:G8"));
settings.getRange("C7:G7").format.fill = colors.input;
settings.getRange("E8:F8").format.fill = colors.input;
settings.getRange("C7:C7").dataValidation = { rule: { type: "list", values: ["최소", "최대"] } };
settings.getRange("D7:D7").dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: 10 } };
settings.getRange("E7:F8").dataValidation = { rule: { type: "list", formula1: "'굴절률'!$A$6:$A$44" } };
settings.getRange("G7:G7").dataValidation = { rule: { type: "whole", operator: "between", formula1: 0, formula2: 30 } };
settings.getRange("C7:G8").format.horizontalAlignment = "center";

settings.getRange("A10:G10").values = [["계산식", null, null, null, null, null, null]];
styleSection(settings.getRange("A10:G10"));
settings.getRange("A11:B13").values = [
  ["Nb2O5", "최소 d=pλ/(2ncosθf), 최대 d=(2p-1)λ/(4ncosθf)"],
  ["SiO2 (2L)", "하부 Nb2O5 두께와 PET를 포함한 TMM 반사율 최소 파장→SiO2 두께 역산(차수 p 미사용)"],
  ["굴절각", "θf=asin(sinθ0/n). 기본 θ0=0°이며 필요 시 설정에서 변경"],
];
styleBody(settings.getRange("A11:B13"));
settings.getRange("B11:B13").format.wrapText = true;

settings.getRange("A15:C15").values = [["시료", "위치(mm)", "하부 Nb2O5(nm)"]];
styleHeader(settings.getRange("A15:C15"));
settings.getRange("A16:C28").values = Array.from({ length: 13 }, (_, i) => [i + 1, 50 + i * 100, 10]);
styleBody(settings.getRange("A16:C28"));
settings.getRange("B16:C28").format.fill = colors.input;
settings.getRange("C16:C28").dataValidation = { rule: { type: "decimal", operator: "between", formula1: 8, formula2: 15 } };
settings.getRange("A16:B28").format.numberFormat = "0";
settings.getRange("C16:C28").format.numberFormat = "0.0";
settings.getRange("A16:C28").format.horizontalAlignment = "right";

settings.getRange("D15:G15").values = [["사용 방법", null, null, null]];
styleSection(settings.getRange("D15:G15"));
settings.getRange("D16:D21").values = [
  ["1. 각 입력 시트의 A:AT 노란색 영역에 CM-36dG 13개 행을 붙여넣습니다."],
  ["2. 설정 C16:C28에 2층 시료의 위치별 하부 Nb2O5 두께(8–15 nm)를 입력합니다."],
  ["3. 평탄부·경계값·잘못 선택된 극값은 AY열의 수동 파장으로 보정합니다."],
  ["4. 결과 시트에서 13개 위치의 파장, 굴절률, 두께와 균일도를 확인합니다."],
  ["5. 2층 SiO2는 TMM 계산이며 차수 p와 입사각 설정을 사용하지 않습니다."],
  ["노란색은 입력 셀입니다. 수식 셀은 직접 덮어쓰지 마세요."],
];
setFont(settings.getRange("D16:G21"), 10);
settings.getRange("D16:D21").format.wrapText = true;
settings.getRange("D16:D21").format.rowHeight = 32;

settings.getRange("A32:G32").values = [["가정 및 한계", null, null, null, null, null, null]];
styleSection(settings.getRange("A32:G32"));
settings.getRange("A33:B38").values = [
  ["단층식", "첨부 굴절률의 실수부 n만 사용한 광학 두께 기반 추정입니다."],
  ["2층 TMM", "Air/SiO2/Nb2O5/PET 구조의 정상입사 TMM으로 계산한 최저 반사율 파장과 측정 파장이 일치하는 SiO2 두께를 역산합니다."],
  ["간섭 차수", "한 개 극값만으로 차수 p를 고유하게 정할 수 없습니다. 알려진 목표 두께 또는 공정 조건과 함께 선택하십시오."],
  ["측정 각도", "기본값은 정상입사 0°입니다. CM-36dG d:8의 8°는 관찰각이므로 입사각으로 자동 적용하지 않았습니다."],
  ["극값 처리", "설정한 검색 구간의 전역 최소/최대를 사용하며, 동일한 극값이 여러 파장에 있으면 그 파장의 평균을 사용합니다. 원본 분해능은 10 nm입니다."],
  ["데이터 가정", "TMM은 굴절률 실수부 n, 정상입사, 투명·비흡수 막, 반무한 PET를 가정합니다. PET 후면 반사와 k는 포함하지 않습니다."],
];
styleBody(settings.getRange("A33:B38"));
settings.getRange("B33:B38").format.wrapText = true;
settings.getRange("A33:B38").format.rowHeight = 46;

settings.getRange("A1:A38").format.columnWidth = 18;
settings.getRange("B1:B38").format.columnWidth = 62;
settings.getRange("C1:C38").format.columnWidth = 12;
settings.getRange("D1:D38").format.columnWidth = 54;
settings.getRange("E1:G38").format.columnWidth = 15;

createInputSheet(nbInput, {
  title: "Nb2O5 (1L) 반사율 입력",
  sourceRows: nbRows,
  settingsRow: 7,
  nColumn: "C",
  stackApproximation: false,
});
createInputSheet(stackInput, {
  title: "Nb2O5 (1L) / SiO2 (2L) 반사율 입력",
  sourceRows: stackRows,
  settingsRow: 8,
  nColumn: "D",
  stackApproximation: true,
});

refractive.showGridLines = false;
refractive.tabColor = "#A5A5A5";
refractive.getRange("A2").values = [["파장별 굴절률"]];
styleTitle(refractive, "D");
refractive.getRange("A3").values = [["출처: 굴절률.xlsx (PET_TOYOBO, Nb2O5, SiO2 시트). 파장 단위는 문맥상 nm로 표기했습니다."]];
setFont(refractive.getRange("A3:D3"), 9, { italic: true, color: colors.muted });
refractive.getRange("A5:D5").values = [["파장(nm)", "PET_TOYOBO n", "Nb2O5 n", "SiO2 n"]];
styleHeader(refractive.getRange("A5:D5"));
refractive.getRange("A6:D44").values = wavelengths.map((w, i) => [w, petIndex[i], nbIndex[i], sioIndex[i]]);
styleBody(refractive.getRange("A6:D44"));
refractive.getRange("A6:A44").format.numberFormat = "0";
refractive.getRange("B6:D44").format.numberFormat = "0.00000";
refractive.getRange("A6:D44").format.horizontalAlignment = "right";
refractive.getRange("A46").values = [["10 nm 격자 사이 파장은 인접 두 값의 선형 보간을 사용하며 360–740 nm 밖은 계산하지 않습니다."]];
setFont(refractive.getRange("A46:D46"), 9, { italic: true, color: colors.muted });
refractive.getRange("A1:A46").format.columnWidth = 14;
refractive.getRange("B1:D46").format.columnWidth = 17;
refractive.freezePanes.freezeRows(5);

tmmTable.showGridLines = false;
tmmTable.tabColor = "#A5A5A5";
tmmTable.getRange("A2").values = [["2층 SiO2 TMM 역산 보정표"]];
styleTitle(tmmTable, "I");
tmmTable.getRange("A3").values = [["Air/SiO2/Nb2O5/PET, 정상입사, k=0. 각 파장이 반사율 최소가 되는 SiO2 두께(nm)입니다."]];
setFont(tmmTable.getRange("A3:I3"), 9, { italic: true, color: colors.muted });
const tmmNbHeaders = Array.from({ length: 8 }, (_, i) => 8 + i);
tmmTable.getRange("A5:I5").values = [["최저 파장(nm)", ...tmmNbHeaders]];
styleHeader(tmmTable.getRange("A5:I5"));
const tmmWavelengthGrid = Array.from({ length: 3801 }, (_, i) => 360 + i / 10);
const tmmCalibration = tmmWavelengthGrid.map((wavelength) => [
  wavelength,
  ...tmmNbHeaders.map((nbThickness) => invertTmmMinimum(wavelength, nbThickness)),
]);
tmmTable.getRange("A6:I3806").values = tmmCalibration;
styleBody(tmmTable.getRange("A6:I3806"));
tmmTable.getRange("A6:A3806").format.numberFormat = "0.0";
tmmTable.getRange("B6:I3806").format.numberFormat = "0.000";
tmmTable.getRange("A6:I3806").format.horizontalAlignment = "right";
tmmTable.getRange("A3808").values = [["Nb2O5 8–15 nm 사이의 입력은 인접한 TMM 곡선 사이를 선형 보간합니다. 굴절률 데이터를 변경하면 이 보정표를 재생성해야 합니다."]];
setFont(tmmTable.getRange("A3808:I3808"), 9, { italic: true, color: colors.muted });
tmmTable.getRange("A1:A3808").format.columnWidth = 18;
tmmTable.getRange("B1:I3808").format.columnWidth = 13;
tmmTable.freezePanes.freezeRows(5);

workbook.recalculate();

const inputCheck = await workbook.inspect({
  kind: "table",
  range: "Nb2O5 입력!AU5:BD18",
  include: "values,formulas",
  tableMaxRows: 14,
  tableMaxCols: 10,
  maxChars: 14000,
});
console.log("===== INPUT CHECK =====");
console.log(inputCheck.ndjson);

const resultCheck = await workbook.inspect({
  kind: "table",
  range: "결과!A2:O24",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 15,
  maxChars: 18000,
});
console.log("===== RESULT CHECK =====");
console.log(resultCheck.ndjson);

const errorCheck = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log("===== ERROR CHECK =====");
console.log(errorCheck.ndjson);

const representativeRows = [0, 6, 12];
for (const i of representativeRows) {
  const nbExpected = expectedRow(nbRows[i], wavelengths, nbIndex, { extreme: "최소", order: 1, start: 430, end: 600, angle: 0 }, false);
  const stackExpected = expectedRow(stackRows[i], wavelengths, sioIndex, { extreme: "최소", order: 1, start: 430, end: 550, angle: 0 }, true);
  const stackTmmExpected = invertTmmMinimum(stackExpected.wavelength, 10);
  const excelRow = 12 + i;
  const actual = results.getRange(`E${excelRow}:N${excelRow}`).values[0];
  const nbActual = Number(actual[2]);
  const stackActual = Number(actual[9]);
  if (!Number.isFinite(nbActual) || Math.abs(nbActual - nbExpected.thickness) > 1e-6) {
    throw new Error(`Nb2O5 verification failed at sample ${i + 1}: actual=${nbActual}, expected=${nbExpected.thickness}`);
  }
  if (!Number.isFinite(stackActual) || Math.abs(stackActual - stackTmmExpected) > 0.03) {
    throw new Error(`SiO2 TMM verification failed at sample ${i + 1}: actual=${stackActual}, expected=${stackTmmExpected}`);
  }
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`OUTPUT\t${outputPath}`);
