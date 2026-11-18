import {autoTable} from 'jspdf-autotable';
import logos from './logos.json';

export function dailyPDF(pdf,report){
  const d=report.daily,s=report.sections[0],width=pdf.internal.pageSize.getWidth();
  const center=(text,x,y,size,rgb=[13,39,55])=>{pdf.setFont('helvetica','bold');pdf.setFontSize(size);pdf.setTextColor(...rgb);pdf.text(text,x,y,{align:'center'})};
  const heading=()=>{
    pdf.setDrawColor(22,74,94);pdf.setLineWidth(.3);
    for(const box of [[6,6,63,21],[72,6,149,21],[224,6,67,12],[224,20,67,7]])pdf.roundedRect(...box,2.7,2.7);
    pdf.addImage(logos[0],'PNG',10,8,26.8,12.4);pdf.addImage(logos[1],'PNG',38,9,27.3,11);
    center('GAMBIRASI SA',146.5,13,11);center('SISTEMA DE GESTIÓN DE TRAZABILIDAD',146.5,18.5,8);center('REPORTE DE CONSUMOS DIARIOS',146.5,24.5,10);
    center('CODIGO : SG/PRO/ F-07',257.5,11,7);center('VERSION : 02',257.5,16,7);center('APROBADA: JULIO 2022',257.5,24.6,7);
    center('No.',23,36,8);center(String(d.number),36,36,14,[208,20,23]);
    const date=new Intl.DateTimeFormat('es-EC',{dateStyle:'full',timeZone:'UTC'}).format(new Date(d.date+'T12:00:00Z'));
    pdf.setFontSize(8);pdf.setTextColor(15,45,58);pdf.text('Fecha: '+date,58,36);
    pdf.text('Especie Litopenaeus Vannamei',291,36,{align:'right'});
  };
  const format=x=>typeof x==='number'?x.toLocaleString('es-EC',{maximumFractionDigits:6,minimumFractionDigits:Number.isInteger(x)?0:2}):x;
  const rows=s.rows.map(row=>row.map(format));
  const groups=[{content:'',colSpan:3},{content:'PISCINA',colSpan:21},{content:'PRECRIADERO (PRE)',colSpan:11}];
  if(d.extra.length)groups.push({content:'OTROS',colSpan:d.extra.length});
  groups.push('');
  const mainCols=32+d.extra.length,small=(width-12-53-10-8-13)/mainCols,styles={0:{cellWidth:53,halign:'left'},1:{cellWidth:10},2:{cellWidth:8}};
  for(let i=0;i<mainCols;i++)styles[i+3]={cellWidth:small};
  styles[3+mainCols]={cellWidth:13};
  autoTable(pdf,{head:[groups,s.head],body:rows,foot:[d.totals.map(format)],showFoot:'lastPage',theme:'grid',startY:40,
    margin:{top:40,left:6,right:6,bottom:25},styles:{font:'helvetica',fontSize:5.4,lineWidth:.12,lineColor:[184,190,192],cellPadding:.55,minCellHeight:7,valign:'middle',halign:'center',textColor:20,overflow:'linebreak'},
    headStyles:{fillColor:255,textColor:[9,69,81],minCellHeight:4.4,fontSize:5.5},footStyles:{fillColor:255,textColor:15,fontStyle:'bold'},columnStyles:styles,rowPageBreak:'avoid',
    willDrawPage:heading,didParseCell:cell=>{if(cell.section==='body'&&cell.column.index===0)cell.cell.styles.textColor=[0,87,158];if(cell.section==='body'&&cell.column.index===2)cell.cell.styles.textColor=[145,24,129]}});
  const roles=['Bodega','Producción','Administración','Contabilidad'],labels=['Elaborado por','Revisado por','Aprobado por','Recibido por'];
  const y=194;roles.forEach((role,i)=>{const x=10+i*71.2;pdf.setDrawColor(30);pdf.line(x,y,x+63,y);center(labels[i]+': '+d.signers[i],x+31.5,y+3.8,d.signers[i].length>30?5:7);center(role,x+31.5,y+7.7,6.5)});
  pdf.save(report.filename+'.pdf');
}
