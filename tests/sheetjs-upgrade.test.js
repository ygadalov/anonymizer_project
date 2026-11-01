// Regression coverage added for 6.3. Fixtures are independent BIFF files, not written by the reader under test.
'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const X=require('../vendor/xlsx.full.min.js');
const E=require('../vendor/exceljs.min.js');
const XL=require('../src/xls.js');
const A=require('../src/engine.js');
const load=name=>new Uint8Array(fs.readFileSync(path.join(__dirname,'fixtures',name+'.xls')));
for(const name of ['biff8','biff5-cp1251','biff5-cp866']) test(`${name}: кириллица, ведущие нули, даты, суммы, логические значения и XLSX`,async()=>{
 assert.equal(X.version,'0.20.3');
 const data=XL.extractBook(X,load(name),false),sh=data.sheets[0];
 const value=(r,c)=>sh.cells.find(x=>x[0]===r&&x[1]===c)[3];
 assert.equal(value(0,0),'Контрагент');
 assert.equal(value(1,0),name==='biff8'?'ООО «Вересков»':'ООО "Вересков"');
 assert.equal(value(1,1),'000154');assert.equal(value(2,1),'00/2210');
 assert.equal(value(1,2),1234.5);assert.equal(value(2,2),-2400);
 assert.equal(value(1,3),45000);assert.equal(value(1,4),0);
 assert.equal(value(1,5),false);assert.equal(value(2,5),true);
 assert.equal(data.sheets[1].hidden,1);
 const {wb}=XL.buildWorkbook(data,E);const reloaded=new E.Workbook();await reloaded.xlsx.load(await wb.xlsx.writeBuffer());
 const ws=reloaded.worksheets[0];assert.equal(ws.getCell('B2').value,'000154');assert.equal(ws.getCell('C2').value,1234.5);
 assert.equal(ws.getCell('D2').value.toISOString(),'2023-03-15T00:00:00.000Z');
 assert.equal(ws.getCell('F2').value,false);assert.equal(reloaded.worksheets[1].state,'hidden');
 if(name==='biff8') {assert.equal(ws.getCell('C4').master.address,'A4');assert.ok(ws.getColumn(1).width>25);}
});
for(const date1904 of [false,true])test(`XML Excel 2003: формулы с нулём, форматы и эпоха ${date1904?1904:1900}`,async()=>{
 const xml=`<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel"><ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">${date1904?'<Date1904/>':''}</ExcelWorkbook><Styles><Style ss:ID="date"><NumberFormat ss:Format="dd.mm.yyyy"/></Style></Styles><Worksheet ss:Name="Данные"><Table><Row><Cell><Data ss:Type="String">Контрагент</Data></Cell><Cell><Data ss:Type="String">Формула</Data></Cell><Cell><Data ss:Type="String">Дата</Data></Cell></Row><Row><Cell><Data ss:Type="String">ООО «Вересков»</Data></Cell><Cell ss:Formula="=1-1"><Data ss:Type="Number">0</Data></Cell><Cell ss:StyleID="date"><Data ss:Type="Number">45000</Data></Cell></Row></Table></Worksheet></Workbook>`;
 const data=XL.extractBook(X,new TextEncoder().encode(xml),false);assert.equal(data.date1904,date1904);
 const {wb,formulas}=XL.buildWorkbook(data,E);assert.equal(formulas,1);
 const ws=wb.worksheets[0];assert.equal(ws.getCell('B2').formula,'1-1');assert.equal(ws.getCell('B2').result,0);
 assert.equal(A.gridFromWorksheet(ws).rows[2][2].v,0);
 const next=new E.Workbook();await next.xlsx.load(await wb.xlsx.writeBuffer());
 assert.equal(next.properties.date1904,date1904);
 assert.equal(next.worksheets[0].getCell('B2').result,0);
 assert.equal(next.worksheets[0].getCell('C2').value.toISOString(),date1904?'2027-03-16T00:00:00.000Z':'2023-03-15T00:00:00.000Z');
});
