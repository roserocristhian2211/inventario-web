const seq=n=>Array.from({length:n},(_,i)=>String(i+1).padStart(2,'0'));
export const referenceReport={date:'2026-11-17',number:2268};
export const destinations=[...seq(9),'09B',...seq(20).slice(9),...seq(11).map(x=>'PRE'+x)];
export const familyOf=p=>p.family||(/\bBALANCEADO\b/i.test(p.name)?'balanceado':'insumo');
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const order=(a,b)=>String(a).localeCompare(String(b),'es',{numeric:true});
export function compareProducts(a,b){
  return Number(familyOf(b)==='balanceado')-Number(familyOf(a)==='balanceado') ||
    order(a.name,b.name) || order(a.code,b.code);
}
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
    const line=totals.get(key)||{product:p,dest,quantity:0,unit:m.unit};
    line.quantity+=m.q;totals.set(key,line);
  }
  return [...totals.values()].sort((a,b)=>
    (byPlace?order(a.dest,b.dest):0) || compareProducts(a.product,b.product) || order(a.dest,b.dest)
  ).map(line=>[line.product.code,line.product.name,line.dest,line.quantity,line.unit]);
}
function pounds(p){
  if(p.lbPerSack)return p.lbPerSack;
  const match=p.name.match(/(\d+(?:[.,]\d+)?)\s*(LB[S]?|KG)\b/i);
  return match?Number(match[1].replace(',','.'))*(match[2].toUpperCase()==='KG'?2.20462262:1):familyOf(p)==='balanceado'?55:null;
}
export function buildReport(kind,input,items,movements,anchor){
  const products=new Map(items.filter(p=>p.cat==='directo').sort(compareProducts).map(p=>[p.id,p]));
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
    const subtotals=new Map();
    for(const line of report.sections[0].rows){
      const [code,name,,quantity,unit]=line,key=JSON.stringify([code,unit]);
      const r=subtotals.get(key)||[code,name,unit,0];r[3]+=quantity;subtotals.set(key,r);
    }
    report.sections.push(section('Totales por producto',['Código','Producto','Unidad','Total'],[...subtotals.values()].sort((a,b)=>{
      const p=items.find(x=>x.code===a[0]),q=items.find(x=>x.code===b[0]);return compareProducts(p,q)
    })));
  }else if(kind==='daily'){
    if(!date)throw Error('Selecciona una fecha.');
    const day=out.filter(m=>m.date===date);
    const base=anchor||referenceReport,n=base.number+stamp(date)-stamp(base.date);
    if(!Number.isSafeInteger(n)||n<1)throw Error('La fecha es anterior al inicio de la secuencia.');
    report.title='Reporte de consumos diarios · N.º '+n;
    report.filename='Reporte_consumos_diarios_'+date;
    report.notes=['Fecha: '+date+' · Elaborado por: '+author];
    const totals=new Map();
    for(const m of day){
      const p=products.get(m.data.pid),key=JSON.stringify([p.id,m.unit]);
      const line=totals.get(key)||{p,unit:m.unit,qty:new Map()};
      line.qty.set(place(m),(line.qty.get(place(m))||0)+m.q);totals.set(key,line);
    }
    const sorted=[...totals.values()].sort((a,b)=>compareProducts(a.p,b.p));
    const extra=[...new Set(day.map(place).filter(d=>!destinations.includes(d)))].sort(order);
    const cols=[...destinations,...extra];
    const head=['PRODUCTO','COD.','UND',...cols.map(d=>d.startsWith('PRE')?d.slice(3):d.replace(/^0/,'')),'TOTAL'];
    const rows=sorted.map(({p,unit,qty})=>[p.name,p.code,unit,...cols.map(d=>qty.get(d)||''),sum(cols,d=>qty.get(d)||0)]);
    report.daily={number:n,date,extra,signers:[author,input.reviewed||'',input.approved||'',input.received||''],totals:['TOTAL','','',...cols.map((_,i)=>sum(rows,r=>Number(r[i+3]))||''),sum(rows,r=>r.at(-1))]};
    report.sections.push(section('Piscinas y precriaderos',head,rows));
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
    const previous=new Date(Date.UTC(+month.slice(0,4),+month.slice(5)-2,1)).toISOString().slice(0,7);
    function metrics(list){
      const feed=list.filter(m=>m.type==='salida'&&familyOf(products.get(m.data.pid))==='balanceado');
      return [sum(feed.filter(m=>m.unit==='SC'),m=>m.q),sum(feed,m=>m.q*(m.unit==='SC'?pounds(products.get(m.data.pid))||0:m.unit==='KG'?2.20462262:/^LBS?$/.test(m.unit)?1:0)),new Set(list.filter(m=>m.type==='entrada').map(m=>m.data.group||[m.date,m.data.doc||m.id].join('|'))).size,list.filter(m=>m.type==='salida').length,new Set(list.filter(m=>['entrada','salida'].includes(m.type)).map(m=>m.data.actor)).size];
    }
    const thisMonth=metrics(period),lastMonth=metrics(moves.filter(m=>m.date.startsWith(previous)));
    const feedDays=new Map();
    for(const m of period.filter(m=>m.type==='salida'&&familyOf(products.get(m.data.pid))==='balanceado'&&m.unit==='SC'))feedDays.set(m.date,(feedDays.get(m.date)||0)+m.q);
    report.sections.push(section('Comparación con el mes anterior',['Indicador',previous,month],['Balanceado (SC)','Balanceado (lb)','Guías de ingreso','Salidas','Usuarios con movimientos'].map((label,i)=>[label,lastMonth[i],thisMonth[i]])),section('Consumo diario de balanceado',['Fecha','Sacos'],[...feedDays].sort((a,b)=>order(a[0],b[0]))),section('Cobertura inferior a siete días',['Código','Producto','Unidad','Saldo','Días'],rows.filter(r=>r[6]>0&&typeof r[8]==='number'&&r[8]<7).map(r=>[r[0],r[1],r[2],r[6],r[8]])));
  }
  return report;
}
export function setupReportUI(ctx){
  const {$,today}=ctx,dialog=$('#report-dialog');
  $('#report-open').onclick=()=>{
    const date=today(),month=date.slice(0,7);
    dialog.innerHTML='<form method="dialog"><div class="modalhead"><h2>Reportes</h2><button aria-label="Cerrar">✕</button></div><label>Reporte<select id="report-kind">'+[['daily','Consumos diarios (correo)'],['range','Consumo entre fechas'],['saldo','Saldo disponible'],['kpi','Reporte mensual de KPI']].map(([k,l])=>opt(k,l)).join('')+'</select></label><div id="report-options"></div><div id="report-preview" aria-live="polite"></div><div class="actions"><button type="button" class="btn" id="report-pdf">Descargar PDF</button><button type="button" class="btn outline" id="report-share">Compartir PDF</button><button type="button" class="btn outline" id="report-xlsx">Descargar Excel</button></div></form>';
    function options(){
      const kind=$('#report-kind').value,author='<label>Elaborado por<input name="author" value="CRISTHIAN ROSERO" maxlength="80"></label>';
      const dates=(label,key,value,type='date')=>'<label>'+label+'<input type="'+type+'" name="'+key+'" value="'+esc(value)+'"></label>';
      let html='';
      if(kind==='daily'){
        html=dates('Fecha','date',date)+author+'<details><summary>Responsables del reporte</summary>'+[['reviewed','Revisado por · Producción'],['approved','Aprobado por · Administración'],['received','Recibido por · Contabilidad']].map(([key,label])=>'<label>'+label+'<input name="'+key+'" maxlength="64" value="'+esc(localStorage.getItem('iw-report-'+key)||'')+'"></label>').join('')+'</details>';
        if(ctx.admin()){const a=ctx.anchor()||referenceReport;html+='<details><summary>Numeración del reporte</summary><p>Usa un reporte conocido como referencia. El número avanza uno por día calendario.</p>'+dates('Fecha del reporte conocido','anchorDate',a?.date||date)+dates('Número de ese reporte','anchorNumber',a?.number||'','number')+'<button type="button" id="anchor-save" class="small">Guardar referencia</button></details>'}
      }else if(kind==='kpi')html=dates('Mes','month',month,'month')+author;
      else if(kind==='range'){
        const extra=[...new Set(ctx.moves().filter(m=>m.type==='salida').map(place))].filter(d=>!destinations.includes(d));
        html=dates('Desde','from',month+'-01')+dates('Hasta','to',date)+'<label>Agrupar por<select name="group">'+opt('product','Producto')+opt('dest','Piscina / precría')+'</select></label><label>Grupo<select name="family">'+opt('','Todos')+opt('balanceado','Balanceado')+opt('insumo','Insumos')+'</select></label><label>Destino<select name="destination">'+opt('','Todos')+[...destinations,...extra].map(d=>opt(d)).join('')+'</select></label><label><input type="checkbox" name="allocated"> Aplicar traspasos de consumo</label>';
      }
      $('#report-options').innerHTML=html;$('#report-xlsx').hidden=kind!=='range';
      if($('#anchor-save'))$('#anchor-save').onclick=async()=>{try{const f=dialog.querySelector('form').elements,date=f.anchorDate.value,number=Number(f.anchorNumber.value);if(!date||!Number.isSafeInteger(number)||number<1)throw Error('Indica fecha y número del último reporte.');await ctx.saveAnchor({date,number});preview()}catch(e){ctx.fail(e)}};
      preview();
    }
    const get=()=>{
      const kind=$('#report-kind').value,input={...Object.fromEntries(new FormData(dialog.querySelector('form'))),today:date};
      let movements=ctx.moves();
      if(kind==='range'&&input.allocated){
        const byId=new Map(ctx.inventory().map(p=>[p.id,p]));
        movements=ctx.allocated(input.from,input.to).filter(r=>r.net>0.0000001).map(r=>({type:'salida',date:input.to,unit:r.unit,q:r.net,data:{pid:r.pid,dest:r.dest},prod:byId.get(r.pid)?.name}));
      }
      return buildReport(kind,input,ctx.inventory(),movements,ctx.anchor());
    };
    function preview(){
      try{const report=get(),s=report.sections[0];$('#report-preview').innerHTML='<h3>'+esc(report.title)+'</h3><div class="report-scroll"><table><thead><tr>'+s.head.map(v=>'<th>'+esc(v)+'</th>').join('')+'</tr></thead><tbody>'+s.rows.map(row=>'<tr>'+row.map(v=>'<td>'+esc(typeof v==='number'?v.toLocaleString('es-EC',{maximumFractionDigits:3}):v)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>';$('#report-share').disabled=$('#report-pdf').disabled=$('#report-xlsx').disabled=!s.rows.length;if(!s.rows.length)$('#report-preview').insertAdjacentHTML('beforeend','<p>Sin registros para esta selección.</p>')}
      catch(e){$('#report-preview').textContent=e.message;$('#report-share').disabled=$('#report-pdf').disabled=$('#report-xlsx').disabled=true}
    }
    $('#report-kind').onchange=options;$('#report-options').oninput=preview;
    for(const [id,fn] of [['#report-pdf','makePDF'],['#report-xlsx','makeExcel']])$(id).onclick=async()=>{try{const report=get();if(report.daily){for(const key of ['reviewed','approved','received'])localStorage.setItem('iw-report-'+key,dialog.querySelector('form').elements[key].value)}const writer=await import('./report-vendor.js?v=12');await writer[fn](report)}catch(e){ctx.fail(e)}};
    $('#report-share').hidden=!navigator.canShare;
    $('#report-share').onclick=async()=>{try{const writer=await import('./report-vendor.js?v=12');const files=[writer.makePDF(get(),true)];if(!navigator.canShare?.({files}))throw Error('Descarga el PDF para adjuntarlo.');await navigator.share({files})}catch(e){if(e.name!=='AbortError')ctx.fail(e)}};
    options();dialog.showModal();
  };
}
