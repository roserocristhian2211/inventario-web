export {exportSheets,readSheets,workbookBlob} from './excel.js';
import {dailyPDF} from './daily.js';
import {jsPDF} from 'jspdf';
import {autoTable} from 'jspdf-autotable';
import {exportSheets as downloadWorkbook} from './excel.js';

export function makePDF(report,asFile=false){
  const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  pdf.setProperties({title:report.title,author:report.author||'Inventario Web'});
  if(report.daily)return dailyPDF(pdf,report,asFile);
  let y=17;
  pdf.setFontSize(18);pdf.text(report.title,12,y);y+=8;
  pdf.setFontSize(10);
  for(const line of report.notes){pdf.text(line,12,y);y+=6}
  for(const section of report.sections){
    if(y>172){pdf.addPage();y=17}
    pdf.setFontSize(11);pdf.text(section.title,12,y);y+=3;
    autoTable(pdf,{head:[section.head],body:section.rows.map(row=>row.map(x=>typeof x==='number'?x.toLocaleString('es-EC',{maximumFractionDigits:3}):x)),startY:y,margin:{left:12,right:12,top:12,bottom:14},styles:{fontSize:8,cellPadding:1.6},headStyles:{fillColor:[15,105,101]},rowPageBreak:'avoid'});
    y=pdf.lastAutoTable.finalY+9;
  }
  for(let i=1;i<=pdf.getNumberOfPages();i++){pdf.setPage(i);pdf.setFontSize(8);pdf.text(i+' / '+pdf.getNumberOfPages(),281,203,{align:'right'})}
  if(asFile)return new File([pdf.output('blob')],report.filename+'.pdf',{type:'application/pdf'});
  pdf.save(report.filename+'.pdf');
}
export async function makeExcel(report){
  await downloadWorkbook(report.filename,report.sections.map((s,i)=>({name:i===0?'Consumo':'Totales',rows:[...report.notes.map(n=>[n]),s.head,...s.rows]})));
}
