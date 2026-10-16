import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('C:/Users/lunarianmetal/Desktop/Test/outputs/01a11430-5b6d-7412-9d90-d6f0585219d5/실습 데이터_경영자대시보드.xlsx'));
const sh=wb.worksheets.getItem('Dashboard');
const before=sh.getRange('B9').values[0][0]; sh.getRange('B6').values=[['전자제품']]; wb.recalculate(); const after=sh.getRange('B9').values[0][0]; if(before===after) throw new Error('filter did not change KPI'); console.log(JSON.stringify({all:before,electronics:after}));
