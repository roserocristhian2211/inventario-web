import ExcelJS from 'exceljs';
function value(cell){
  let v=cell.value;
  if(v&&typeof v==='object'&&'result' in v)v=v.result;
  if(v instanceof Date)return v.toISOString().slice(0,10);
  if(v?.richText)return v.richText.map(t=>t.text).join('');
  if(v&&typeof v==='object')return v.text||'';
  return v??'';
}
export async function readSheets(file){
  if(file.size>12*1024*1024)throw Error('Elige un libro de hasta 12 MB.');
  const book=new ExcelJS.Workbook();await book.xlsx.load(await file.arrayBuffer());
  return book.worksheets.filter(s=>s.state==='visible').map(s=>{
    const rows=[];let hidden=0;
    s.eachRow((row,n)=>{if(row.hidden){hidden++;return}const cells=[];row.eachCell((c,i)=>{cells[i-1]=value(c)});rows.push({n,cells})});
    return {name:s.name,rows,hidden};
  });
}
export async function workbookBlob(sheets){
  const book=new ExcelJS.Workbook();
  for(const sheet of sheets){
    const ws=book.addWorksheet(sheet.name);ws.addRows(sheet.rows);
    ws.views=[{state:'frozen',ySplit:1}];ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};
    ws.getRow(1).eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF17665E'}}});
    ws.columns.forEach((col,i)=>col.width=Math.max(13,Math.min(63,Math.max(...sheet.rows.slice(0,300).map(r=>String(r[i]??'').length))+2)));
    if(ws.columnCount)ws.autoFilter={from:{row:1,column:1},to:{row:1,column:ws.columnCount}};
  }
  return new Blob([await book.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
export async function exportSheets(name,sheets){
  const blob=await workbookBlob(sheets);
  const a=document.createElement('a'),url=URL.createObjectURL(blob);a.href=url;a.download=name+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
