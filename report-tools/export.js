import {dailyPDF} from './daily.js';
import {jsPDF} from 'jspdf';
import {autoTable} from 'jspdf-autotable';
import writeExcel from 'write-excel-file/universal';

export function makePDF(report){
  const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  pdf.setProperties({title:report.title,author:report.author||'Inventario Web'});
  if(report.daily)return dailyPDF(pdf,report);
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
  pdf.save(report.filename+'.pdf');
}
export async function makeExcel(report){
  const section=report.sections[0];
  const values=[...report.notes.map(n=>[n]),section.head,...section.rows];
  const blob=await writeExcel(values.map((row,i)=>row.map(value=>({value,type:typeof value==='number'?Number:String,fontWeight:i===report.notes.length?'bold':undefined}))),{sheet:'Consumo',columns:section.head.map(h=>({width:h==='Producto'?65:18}))}).toBlob();
  const a=document.createElement('a'),url=URL.createObjectURL(blob);a.href=url;a.download=report.filename+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
