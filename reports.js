const seq=n=>Array.from({length:n},(_,i)=>String(i+1).padStart(2,'0'));
export const destinations=[...seq(9),'09B',...seq(20).slice(9),...seq(11).map(x=>'PRE'+x)];
export const familyOf=p=>p.family||(/\bBALANCEADO\b/i.test(p.name)?'balanceado':'insumo');
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const order=(a,b)=>String(a).localeCompare(String(b),'es',{numeric:true});
const sum=(rows,fn)=>rows.reduce((total,row)=>total+fn(row),0);
const stamp=date=>Date.parse(date+'T00:00:00Z')/86400000;
const delta=m=>m.type==='salida'?-m.q:m.type==='traspaso'?0:m.q;
const place=m=>m.data.dest?.trim()||'Sin destino';
const section=(title,head,rows)=>({title,head,rows});
const opt=(v,label=v)=>'<option value="'+esc(v)+'">'+esc(label)+'</option>';
function grouped(rows,products,byPlace){
  const totals=new Map();
  for(const m of rows){
    const p=products.get(m.data.pid),dest=place(m),key=JSON.stringify([p.id,dest,m.unit]);
    const line=totals.get(key)||[p.code,p.name,dest,0,m.unit];line[3]+=m.q;totals.set(key,line);
  }
  return [...totals.values()].sort((a,b)=>byPlace?order(a[2],b[2])||order(a[1],b[1]):order(a[1],b[1])||order(a[2],b[2]));
}
function pounds(p){
  if(p.lbPerSack)return p.lbPerSack;
  const match=p.name.match(/(\d+(?:[.,]\d+)?)\s*(LB[S]?|KG)\b/i);
  return match?Number(match[1].replace(',','.'))*(match[2].toUpperCase()==='KG'?2.20462262:1):null;
}
export function buildReport(kind,input,items,movements,anchor){
  const products=new Map(items.filter(p=>p.cat==='directo').map(p=>[p.id,p]));
  const moves=movements.filter(m=>products.has(m.data.pid)),out=moves.filter(m=>m.type==='salida');
  const {date,from,to,month,author,today}=input;
  const report={author,notes:[],sections:[]};
  if(kind==='saldo'){
    report.title='Saldo disponible';report.filename='Saldos-'+today;report.notes=['Fecha de consulta: '+today];
    report.sections.push(section('Productos con existencia',['Código','Descripción','Medida','Saldo'],[...products.values()].filter(p=>!p.deleted&&p.stock>0).map(p=>[p.code,p.name,p.unit,p.stock])));
  }else if(kind==='range'){
    if(!from||!to||from>to)throw Error('Selecciona un rango de fechas válido.');
    const rows=out.filter(m=>m.date>=from&&m.date<=to&&(!input.destination||place(m)===input.destination)&&(!input.family||familyOf(products.get(m.data.pid))===input.family));
    report.title='Consumo entre fechas';report.filename='Consumo-'+from+'_'+to;report.notes=[from+' al '+to+' · '+(input.destination||'Todos los destinos')+' · '+(input.family||'Todos los productos')];
    report.sections.push(section('Productos despachados',['Código','Producto','PSC / PRE','Cantidad','Medida'],grouped(rows,products,input.group==='dest')));
  }else if(kind==='daily'){
    if(!date)throw Error('Selecciona una fecha.');
    const day=out.filter(m=>m.date===date);
    const n=anchor?anchor.number+stamp(date)-stamp(anchor.date):null,serial=n>0?String(n).padStart(3,'0'):'';
    report.title='Consumo diario'+(serial?' · N.º '+serial:' · Borrador');
    report.filename='Consumo-diario-'+date+(serial?'-'+serial:'-borrador');
    report.notes=['Fecha: '+date+' · Elaborado por: '+author];
    const lines=grouped(day,products,false),byProd=new Map();
    for(const [code,name,d,q,unit] of lines){const key=JSON.stringify([code,name,unit]);const row=byProd.get(key)||{code,name,unit,qty:new Map()};row.qty.set(d,q);byProd.set(key,row)}
    const groups=new Map([...products.values()].map(p=>[p.code,familyOf(p)]));
    const sorted=[...byProd.values()].sort((a,b)=>order(groups.get(a.code),groups.get(b.code))||order(a.name,b.name));
    const extra=[...new Set(day.map(place).filter(d=>!destinations.includes(d)))].sort(order);
    for(const [label,places] of [['Piscinas',destinations.slice(0,21)],['Precrías',destinations.slice(21)],['Otros destinos',extra]]){
      for(let i=0;i<places.length;i+=11){
        const cols=places.slice(i,i+11);
        report.sections.push(section(label+' · '+cols[0]+' a '+cols.at(-1),['Producto','Medida',...cols,'Subtotal'],sorted.map(p=>[p.name,p.unit,...cols.map(d=>p.qty.get(d)||''),sum(cols,d=>p.qty.get(d)||0)])));
      }
    }
  }else{
    if(!/^\d{4}-\d{2}$/.test(month)||month>today.slice(0,7))throw Error('Selecciona un mes transcurrido o el actual.');
    const start=month+'-01',end=month===today.slice(0,7)?today:new Date(Date.UTC(+month.slice(0,4),+month.slice(5),0)).toISOString().slice(0,10);
    const days=stamp(end)-stamp(start)+1,period=moves.filter(m=>m.date>=start&&m.date<=end);
    const summary=[],rows=[];let empty=0,sacks=0,lbs=0,unknown=0;
    for(const p of products.values()){
      if(p.created> end || (p.removed&&p.removed<=end))continue;
      const all=moves.filter(m=>m.data.pid===p.id),during=period.filter(m=>m.data.pid===p.id);
      const closing=p.stock-sum(all.filter(m=>m.date>end),delta),opening=closing-sum(during,delta);
      const incoming=sum(during.filter(m=>m.type==='entrada'),m=>m.q),used=sum(during.filter(m=>m.type==='salida'),m=>m.q);
      if(closing<=0)empty++;
      if(familyOf(p)==='balanceado'&&used){
        if(p.unit==='SC'){sacks+=used;const factor=pounds(p);if(factor)lbs+=used*factor;else unknown+=used}
        if(p.unit==='KG')lbs+=used*2.20462262;
        if(/^LBS?$/.test(p.unit))lbs+=used;
      }
      const avg=used/days,mid=(opening+closing)/2;
      rows.push([p.code,p.name,p.unit,opening,incoming,used,closing,avg,avg?closing/avg:'—',mid>0?used/mid:'—']);
    }
    summary.push(['Balanceado consumido (SC)',sacks],['Libras con peso conocido',lbs],['Sacos pendientes de peso',unknown],['Promedio diario (SC/día)',sacks/days],['Productos agotados al cierre',empty]);
    report.title='Reporte mensual de KPI';report.filename='KPI-'+month;report.notes=[start+' al '+end+' · Elaborado por: '+author,'Cobertura en días = saldo / consumo diario. Rotación = consumo / inventario promedio de apertura y cierre.'];
    report.sections.push(section('Indicadores',['Indicador','Valor'],summary),section('Detalle por producto',['Código','Producto','Medida','Inicial','Entradas','Salidas','Saldo','Prom./día','Cobertura','Rotación'],rows),section('Consumo por destino',['Código','Producto','PSC / PRE','Cantidad','Medida'],grouped(period.filter(m=>m.type==='salida'),products,true)));
  }
  return report;
}
export function setupReportUI(ctx){
  const {$,today}=ctx,dialog=$('#report-dialog');
  $('#report-open').onclick=()=>{
    const date=today(),month=date.slice(0,7);
    dialog.innerHTML='<form method="dialog"><div class="modalhead"><h2>Reportes</h2><button aria-label="Cerrar">✕</button></div><label>Reporte<select id="report-kind">'+[['daily','Consumos diarios (correo)'],['range','Consumo entre fechas'],['saldo','Saldo disponible'],['kpi','Reporte mensual de KPI']].map(([k,l])=>opt(k,l)).join('')+'</select></label><div id="report-options"></div><div id="report-preview" aria-live="polite"></div><div class="actions"><button type="button" class="btn" id="report-pdf">Descargar PDF</button><button type="button" class="btn outline" id="report-xlsx">Descargar Excel</button></div></form>';
    function options(){
      const kind=$('#report-kind').value,author='<label>Elaborado por<input name="author" value="CRISTHIAN ROSERO" maxlength="80"></label>';
      const dates=(label,key,value,type='date')=>'<label>'+label+'<input type="'+type+'" name="'+key+'" value="'+esc(value)+'"></label>';
      let html='';
      if(kind==='daily'){
        html=dates('Fecha','date',date)+author;
        if(ctx.admin()){const a=ctx.anchor();html+='<details><summary>Numeración del reporte</summary><p>Usa un reporte conocido como referencia. El número avanza uno por día calendario.</p>'+dates('Fecha del reporte conocido','anchorDate',a?.date||date)+dates('Número de ese reporte','anchorNumber',a?.number||'','number')+'<button type="button" id="anchor-save" class="small">Guardar referencia</button></details>'}
      }else if(kind==='kpi')html=dates('Mes','month',month,'month')+author;
      else if(kind==='range'){
        const extra=[...new Set(ctx.moves().filter(m=>m.type==='salida').map(place))].filter(d=>!destinations.includes(d));
        html=dates('Desde','from',month+'-01')+dates('Hasta','to',date)+'<label>Agrupar por<select name="group">'+opt('product','Producto')+opt('dest','Piscina / precría')+'</select></label><label>Grupo<select name="family">'+opt('','Todos')+opt('balanceado','Balanceado')+opt('insumo','Insumos')+'</select></label><label>Destino<select name="destination">'+opt('','Todos')+[...destinations,...extra].map(d=>opt(d)).join('')+'</select></label>';
      }
      $('#report-options').innerHTML=html;$('#report-xlsx').hidden=kind!=='range';
      if($('#anchor-save'))$('#anchor-save').onclick=async()=>{try{const f=dialog.querySelector('form').elements,date=f.anchorDate.value,number=Number(f.anchorNumber.value);if(!date||!Number.isSafeInteger(number)||number<1)throw Error('Indica fecha y número del último reporte.');await ctx.saveAnchor({date,number});preview()}catch(e){ctx.fail(e)}};
      preview();
    }
    const get=()=>buildReport($('#report-kind').value,{...Object.fromEntries(new FormData(dialog.querySelector('form'))),today:date},ctx.inventory(),ctx.moves(),ctx.anchor());
    function preview(){
      try{const report=get(),s=report.sections[0];$('#report-preview').innerHTML='<h3>'+esc(report.title)+'</h3><div class="report-scroll"><table><thead><tr>'+s.head.map(v=>'<th>'+esc(v)+'</th>').join('')+'</tr></thead><tbody>'+s.rows.map(row=>'<tr>'+row.map(v=>'<td>'+esc(typeof v==='number'?v.toLocaleString('es-EC',{maximumFractionDigits:3}):v)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>';$('#report-pdf').disabled=$('#report-xlsx').disabled=!s.rows.length;if(!s.rows.length)$('#report-preview').insertAdjacentHTML('beforeend','<p>Sin registros para esta selección.</p>')}
      catch(e){$('#report-preview').textContent=e.message;$('#report-pdf').disabled=$('#report-xlsx').disabled=true}
    }
    $('#report-kind').onchange=options;$('#report-options').oninput=preview;
    for(const [id,fn] of [['#report-pdf','makePDF'],['#report-xlsx','makeExcel']])$(id).onclick=async()=>{try{const report=get();const writer=await import('./report-vendor.js?v=8');await writer[fn](report)}catch(e){ctx.fail(e)}};
    options();dialog.showModal();
  };
}
