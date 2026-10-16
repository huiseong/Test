import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const source='C:/Users/lunarianmetal/Desktop/Test/outputs/01a11430-5b6d-7412-9d90-d6f0585219d5/실습 데이터_더미200개.xlsx';
const out='C:/Users/lunarianmetal/Desktop/Test/outputs/01a11430-5b6d-7412-9d90-d6f0585219d5/실습 데이터_경영자대시보드.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const data=wb.worksheets.getItem('Sheet1');
let dash=wb.worksheets.getItemOrNullObject?.('Dashboard');
if (!dash || dash.isNullObject) dash=wb.worksheets.add('Dashboard');

const cats=['전체','전자제품','생활가전','주방용품','생활용품','패션잡화'];
const months=['전체','2026-01','2026-02','2026-03','2026-04','2026-05','2026-06','2026-07','2026-08','2026-09','2026-10'];
dash.getRange('A1:N40').clear({applyTo:'all'});
dash.showGridLines=false;
dash.getRange('A1:N1').merge();
dash.getRange('A1').values=[['경영자 대시보드 | 글로벌 판매·수익성·운영 현황']];
dash.getRange('A2').values=[['필터를 변경하면 KPI와 분석표가 함께 갱신됩니다.']];
dash.getRange('A4').values=[['분석 조건']];
dash.getRange('A5:B5').values=[['카테고리','기간']];
dash.getRange('B6').values=[['전체']];
dash.getRange('C6').values=[['전체']];
dash.getRange('B6').dataValidation={rule:{type:'list',values:cats}};
dash.getRange('C6').dataValidation={rule:{type:'list',values:months}};
dash.getRange('A8:D8').values=[['핵심 KPI','값','경영자 해석','산출 기준']];
dash.getRange('A9:A14').values=[['총매출 (USD)'],['총이익 (USD)'],['이익률'],['총 판매수량'],['마케팅비 (USD)'],['납기준수율']];

const crit=(field, sel, all='<>')=>`IF(${sel}="전체","${all}",${sel})`;
const sum=(col, catCell='$B$6', monthCell='$C$6')=>`=IF(${monthCell}="전체",SUMIFS(Sheet1!$${col}$2:$${col}$201,Sheet1!$D$2:$D$201,${crit('D',catCell)}),SUMIFS(Sheet1!$${col}$2:$${col}$201,Sheet1!$D$2:$D$201,${crit('D',catCell)},Sheet1!$A$2:$A$201,">="&DATEVALUE(${monthCell}&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE(${monthCell}&"-01"),1)))`;
const count=()=>`=IF($C$6="전체",COUNTIFS(Sheet1!$D$2:$D$201,${crit('D','$B$6')}),COUNTIFS(Sheet1!$D$2:$D$201,${crit('D','$B$6')},Sheet1!$A$2:$A$201,">="&DATEVALUE($C$6&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE($C$6&"-01"),1)))`;
dash.getRange('B9:B10').formulas=[[sum('I')],[sum('J')]];
dash.getRange('B11').formulas=[['=IFERROR(B10/B9,0)']];
dash.getRange('B12').formulas=[[sum('G')]];
dash.getRange('B13').formulas=[[sum('K')]];
dash.getRange('B14').formulas=[[`=IF($C$6="전체",AVERAGE(Sheet1!$M$2:$M$201),AVERAGEIFS(Sheet1!$M$2:$M$201,Sheet1!$A$2:$A$201,">="&DATEVALUE($C$6&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE($C$6&"-01"),1)))`]];
dash.getRange('C9:C14').values=[['선택 조건의 매출 규모'],['매출을 실제로 남긴 금액'],['매출 대비 이익 비율'],['판매 활동 규모'],['매출 확보에 사용한 비용'],['고객 약속 이행 수준']];
dash.getRange('D9:D14').values=[['매출 합계'],['이익 합계'],['총이익 ÷ 총매출'],['수량 합계'],['마케팅비 합계'],['평균 납기준수율']];

dash.getRange('A17:D17').values=[['카테고리별 성과','매출 (USD)','이익 (USD)','이익률']];
dash.getRange('A18:A22').values=cats.slice(1).map(x=>[x]);
for(let r=18;r<=22;r++){
  dash.getRange(`B${r}`).formulas=[[sum('I','$A'+r)]];
  dash.getRange(`C${r}`).formulas=[[sum('J','$A'+r)]];
  dash.getRange(`D${r}`).formulas=[[`=IFERROR(C${r}/B${r},0)`]];
}
dash.getRange('F17:I17').values=[['판매채널별 성과','매출 (USD)','이익 (USD)','이익률']];
const channels=['온라인몰','해외 대리점','도매','오프라인 매장','B2B'];
dash.getRange('F18:F22').values=channels.map(x=>[x]);
for(let r=18;r<=22;r++){
  const ch=`$F${r}`;
  const catFilter=crit('D','$B$6');
  dash.getRange(`G${r}`).formulas=[[`=IF($C$6="전체",SUMIFS(Sheet1!$I$2:$I$201,Sheet1!$F$2:$F$201,${ch},Sheet1!$D$2:$D$201,${catFilter}),SUMIFS(Sheet1!$I$2:$I$201,Sheet1!$F$2:$F$201,${ch},Sheet1!$D$2:$D$201,${catFilter},Sheet1!$A$2:$A$201,">="&DATEVALUE($C$6&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE($C$6&"-01"),1)))`]];
  dash.getRange(`H${r}`).formulas=[[`=IF($C$6="전체",SUMIFS(Sheet1!$J$2:$J$201,Sheet1!$F$2:$F$201,${ch},Sheet1!$D$2:$D$201,${catFilter}),SUMIFS(Sheet1!$J$2:$J$201,Sheet1!$F$2:$F$201,${ch},Sheet1!$D$2:$D$201,${catFilter},Sheet1!$A$2:$A$201,">="&DATEVALUE($C$6&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE($C$6&"-01"),1)))`]];
  dash.getRange(`I${r}`).formulas=[[`=IFERROR(H${r}/G${r},0)`]];
}
dash.getRange('K17:N17').values=[['경영자 관점 체크','현재값','기준','해석']];
dash.getRange('K18:K21').values=[['이익률'],['납기준수율'],['불량률'],['마케팅비율']];
dash.getRange('L18:L21').formulas=[['=B11'],['=B14'],[`=IF($C$6="전체",AVERAGE(Sheet1!$L$2:$L$201),AVERAGEIFS(Sheet1!$L$2:$L$201,Sheet1!$A$2:$A$201,">="&DATEVALUE($C$6&"-01"),Sheet1!$A$2:$A$201,"<"&EDATE(DATEVALUE($C$6&"-01"),1)))`],['=IFERROR(B13/B9,0)']];
dash.getRange('M18:M21').values=[[0.2],[0.95],[0.02],[0.08]];
dash.getRange('N18:N21').formulas=[['=IF(L18>=M18,"양호","점검")'],['=IF(L19>=M19,"양호","점검")'],['=IF(L20<=M20,"양호","점검")'],['=IF(L21<=M21,"양호","점검")']];

dash.getRange('A1:N1').format={fill:'#1F4E78',font:{name:'Arial',size:16,bold:true,color:'#FFFFFF'}};
dash.getRange('A2:N2').format={font:{name:'Arial',size:10,color:'#666666',italic:true}};
for(const rng of ['A4:D4','A8:D8','A17:D17','F17:I17','K17:N17']) dash.getRange(rng).format={fill:'#5B9BD5',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center'};
dash.getRange('A5:C6').format={fill:'#FFF2CC',font:{name:'Arial',size:10,bold:true}};
dash.getRange('A9:D14').format.font={name:'Arial',size:10};
dash.getRange('A17:N22').format.font={name:'Arial',size:10};
dash.getRange('B9:B10').format.numberFormat='#,##0.00'; dash.getRange('B11').format.numberFormat='0.0%'; dash.getRange('B12').format.numberFormat='#,##0'; dash.getRange('B13').format.numberFormat='#,##0.00'; dash.getRange('B14').format.numberFormat='0.0%';
dash.getRange('B18:C22').format.numberFormat='#,##0.00'; dash.getRange('D18:D22').format.numberFormat='0.0%'; dash.getRange('G18:H22').format.numberFormat='#,##0.00'; dash.getRange('I18:I22').format.numberFormat='0.0%'; dash.getRange('L18:M21').format.numberFormat='0.0%';
dash.getRange('A8:N22').format.borders={preset:'all',style:'thin',color:'#D9E2F3'};
dash.getRange('A1:N40').format.verticalAlignment='center';
dash.getRange('A:N').format.columnWidth=14;
dash.getRange('A:A').format.columnWidth=18; dash.getRange('C:C').format.columnWidth=26; dash.getRange('D:D').format.columnWidth=18; dash.getRange('K:K').format.columnWidth=18; dash.getRange('N:N').format.columnWidth=14;
dash.freezePanes.freezeRows(6);
const chart1=dash.charts.add('bar',dash.getRange('A17:C22')); chart1.title='카테고리별 매출·이익'; chart1.setPosition('A25','G40'); chart1.hasLegend=true;
const chart2=dash.charts.add('bar',dash.getRange('F17:H22')); chart2.title='판매채널별 매출·이익'; chart2.setPosition('H25','N40'); chart2.hasLegend=true;
wb.recalculate();
const check=await wb.inspect({kind:'table',sheetId:'Dashboard',range:'A1:N22',include:'values,formulas',tableMaxRows:22,tableMaxCols:14,maxChars:8000});
await fs.writeFile(`${out}.inspect.ndjson`,check.ndjson);
const xlsx=await SpreadsheetFile.exportXlsx(wb); await xlsx.save(out);
console.log(JSON.stringify({output:out,check:check.ndjson.slice(0,1200)}));
