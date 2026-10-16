import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const outDir = 'C:/Users/lunarianmetal/Desktop/Test/outputs/01a11430-5b6d-7412-9d90-d6f0585219d5';
await fs.mkdir(outDir, { recursive: true });
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load('C:/Users/lunarianmetal/Desktop/Test/실습 데이터.xlsx'));
const sh = wb.worksheets.getItemAt(0);
console.log((await wb.inspect({kind:'workbook,sheet,table',maxChars:2500,tableMaxRows:2,tableMaxCols:13})).ndjson);
if (process.argv.includes('--inspect')) {
  const image = await wb.render({sheetName:'Sheet1',range:'A1:M3',scale:1.5,format:'png'});
  await fs.writeFile(`${outDir}/source-preview.png`,new Uint8Array(await image.arrayBuffer()));
  process.exit(0);
}
const headers = ['주문일자','생산법인','수출대상국','카테고리','제품명','판매채널','수량','단가(USD)','매출(USD)','이익(USD)','마케팅비(USD)','불량률(%)','납기준수율(%)'];
assert.deepEqual(sh.getRange('A1:M1').values[0],headers);
let seed=20261007;
const random=()=> { seed=(Math.imul(seed,1664525)+1013904223)>>>0; return seed/4294967296; };
const pick=a=>a[Math.floor(random()*a.length)];
const round=n=>Math.round(n*100)/100;
const products=[
  ['전자제품','무선 이어폰',49.9],['전자제품','블루투스 스피커',79.9],['전자제품','스마트워치',129.9],
  ['생활가전','공기청정기',199],['생활가전','로봇청소기',299],['생활가전','전기주전자',39.9],
  ['주방용품','스테인리스 냄비',54.9],['주방용품','텀블러',19.9],['주방용품','밀폐용기 세트',24.9],
  ['생활용품','수납함',14.9],['생활용품','LED 데스크 조명',34.9],['생활용품','욕실 매트',12.9],
  ['패션잡화','백팩',44.9],['패션잡화','여행용 캐리어',99.9],['패션잡화','지갑',29.9]
];
const rows=[];
for(let i=0;i<200;i++) {
  const p=pick(products);
  const channel=pick(['온라인몰','해외 대리점','도매','오프라인 매장','B2B']);
  const qty=channel==='도매'||channel==='B2B'?100+Math.floor(random()*900):10+Math.floor(random()*290);
  const price=round(p[2]*(0.85+random()*0.3));
  const sales=round(qty*price);
  const marketing=round(sales*(channel==='온라인몰'?0.05+random()*0.07:0.01+random()*0.04));
  const profit=round(sales*(0.18+random()*0.19)-marketing);
  const date=new Date(Date.UTC(2026,0,1)+Math.floor(random()*279)*86400000);
  rows.push([date,pick(['한국 생산법인','베트남 생산법인','중국 생산법인','태국 생산법인']),pick(['미국','일본','독일','프랑스','영국','캐나다','호주','싱가포르','베트남','인도']),p[0],p[1],channel,qty,price,sales,profit,marketing,Math.round((0.002+random()*0.035)*10000)/10000,Math.round((0.88+random()*0.12)*10000)/10000]);
}
rows.sort((a,b)=>a[0]-b[0]);
assert.equal(rows.length,200);
assert.equal(new Set(rows.map(r=>JSON.stringify(r))).size,200);
for(const r of rows) {
  assert.equal(r.length,13);
  assert.equal(r[8],round(r[6]*r[7]));
  assert.ok(r[9]>0&&r[9]<r[8]&&r[10]<r[8]);
  assert.ok(r[11]>=0&&r[11]<=1&&r[12]>=0&&r[12]<=1);
}
sh.getRange('A2:M201').values=rows;
sh.getRange('I2:I201').formulas=rows.map((_,i)=>[`=ROUND(G${i+2}*H${i+2},2)`]);
sh.getRange('A2:A201').setNumberFormat('yyyy-mm-dd');
sh.getRange('G2:G201').setNumberFormat('#,##0');
sh.getRange('H2:K201').setNumberFormat('#,##0.00');
sh.getRange('L2:M201').setNumberFormat('0.0%');
sh.getRange('A1:M201').format.autofitColumns();
sh.getRange('A1:M201').format.rowHeight=20;
sh.freezePanes.freezeRows(1);
wb.recalculate();
const values=sh.getRange('A2:M201').values;
assert.equal(values.length,200);
for(let i=0;i<200;i++) assert.ok(Math.abs(values[i][8]-rows[i][8])<0.001);
console.log((await wb.inspect({kind:'table',range:'Sheet1!A1:M4',include:'values,formulas',tableMaxRows:4,tableMaxCols:13,maxChars:3500})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
// Native rendering is unavailable on this host (Vulkan initialization failure).
const output=await SpreadsheetFile.exportXlsx(wb);
await output.save(`${outDir}/실습 데이터_더미200개.xlsx`);
console.log(JSON.stringify({records:200,fields:13,output:`${outDir}/실습 데이터_더미200개.xlsx`}));
