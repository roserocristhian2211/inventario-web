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
  const dialog=ctx.$('#report-dialog');
  ctx.$('#report-open').onclick=()=>{
    const date=ctx.today(),month=date.slice(0,7);
    const field=(label,name,type,value)=>'<label>'+label+'<input name="'+name+'" type="'+type+'" value="'+esc(value)+'"></label>';
    const author=field('Elaborado por','author','text','Cristhian Rosero');
    const pair=(left,right)=>'<div class="report-pair">'+left+right+'</div>';
    const segments=(title,name,choices)=>'<fieldset class="filter-switch"><legend>'+title+'</legend><div class="switchbar">'+choices.map(([key,label],i)=>'<label><input type="radio" name="'+name+'" value="'+key+'" '+(i===0?'checked':'')+'><span>'+label+'</span></label>').join('')+'</div></fieldset>';
    const extra=[...new Set(ctx.moves().filter(m=>m.type==='salida').map(place))].filter(d=>!destinations.includes(d));
    const anchor=ctx.anchor()||referenceReport;
    const signatures=['reviewed','approved','received'];
    const advanced='<details><summary>Opciones del reporte diario</summary>'+signatures.map((key,i)=>field(['Revisado por · Producción','Aprobado por · Administración','Recibido por · Contabilidad'][i],key,'text',localStorage.getItem('iw-report-'+key)||'')).join('')+(ctx.admin()?pair(field('Fecha del reporte conocido','anchorDate','date',anchor.date),field('Número de ese reporte','anchorNumber','number',anchor.number))+'<button type="button" class="small" id="anchor-save">Guardar referencia</button>':'')+'</details>';
    const range=pair(field('Desde','from','date',month+'-01'),field('Hasta','to','date',date))+
      segments('Agrupar por','group',[['product','Producto'],['dest','Piscina / precría']])+
      segments('Grupo','family',[['','Todo'],['balanceado','Balanceado'],['insumo','Insumos']])+
      '<label>Piscina / precría<select name="destination">'+opt('','Todas')+[...destinations,...extra].map(d=>opt(d)).join('')+'</select></label>';
    const block=(kind,title,description,controls,pdf,more='')=>'<form class="report-block" data-report="'+kind+'" aria-label="'+title+'"><h3>'+title+'</h3>'+(description?'<p>'+description+'</p>':'')+controls+'<div class="report-result" aria-live="polite"></div><div class="report-downloads"><button type="button" class="btn" data-export="pdf">'+pdf+'</button>'+(kind==='range'?'<button type="button" class="btn outline" data-export="xlsx">Descargar Excel</button>':'')+'</div>'+more+(navigator.canShare?'<button class="report-share" type="button" data-export="share">Compartir PDF</button>':'')+'</form>';
    dialog.innerHTML='<button type="button" class="back-link" data-back><span aria-hidden="true">←</span> Volver</button><div class="modalhead"><h2>Reportes</h2><button type="button" aria-label="Cerrar">✕</button></div>'+
      block('saldo','Saldo disponible','','','Descargar PDF de saldos')+
      block('kpi','Reporte mensual de KPI','Balanceado consumido en sacos y libras, promedio diario, cobertura, rotación, entradas, salidas y consumo por piscina y producto.',pair(field('Mes','month','month',month),author),'Generar PDF de KPI')+
      block('daily','Consumos diarios (formato CORREO)','Piscinas y precrías en columnas. Solo aparecen los productos despachados ese día, primero los balanceados y después los insumos, en orden alfabético.',pair(field('Fecha','date','date',date),author),'Descargar PDF de consumos diarios',advanced)+
      block('range','Consumo entre fechas','',range,'Descargar PDF','<label class="allocation"><input type="checkbox" name="allocated"> Aplicar traspasos de consumo</label>');
    dialog.querySelector('[aria-label="Cerrar"]').onclick=()=>dialog.close();
    for(const form of dialog.querySelectorAll('.report-block')){
      const kind=form.dataset.report,result=form.querySelector('.report-result');
      function get(){
        const input={...Object.fromEntries(new FormData(form)),today:date};
        let movements=ctx.moves();
        if(kind==='range'&&input.allocated){
          movements=ctx.allocated(input.from,input.to).filter(r=>r.net>0.0000001).map(r=>({type:'salida',date:input.to,unit:r.unit,q:r.net,data:{pid:r.pid,dest:r.dest}}));
        }
        return buildReport(kind,input,ctx.inventory(),movements,ctx.anchor());
      }
      function preview(){
        let enabled=false;
        try{
          const report=get(),{head,rows}=report.sections[0];enabled=rows.length>0;
          if(kind==='saldo')result.textContent='Productos con saldo mayor a cero ('+rows.length+'), con código, descripción, unidad y saldo actual.';
          if(kind==='kpi')result.textContent='';
          if(kind==='daily'){
            const names=[...new Set(rows.map(row=>row[0]))];
            result.textContent=report.title+' · '+(names.length?names.length+' producto(s) con salidas: '+names.join(', '):'Sin salidas en la fecha seleccionada.');
          }
          if(kind==='range')result.innerHTML=enabled?'<div class="report-scroll"><table><thead><tr>'+head.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(x=>'<td>'+esc(typeof x==='number'?x.toLocaleString('es-EC',{minimumFractionDigits:2,maximumFractionDigits:6}):x)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>':'Sin despachos para esta selección.';
        }catch(e){result.textContent=e.message}
        for(const b of form.querySelectorAll('[data-export]'))b.disabled=!enabled;
      }
      form.oninput=preview;form.onsubmit=e=>e.preventDefault();
      const save=form.querySelector('#anchor-save');
      if(save)save.onclick=async()=>{try{
        const date=form.elements.anchorDate.value,number=Number(form.elements.anchorNumber.value);
        if(!date||!Number.isSafeInteger(number)||number<1)throw Error('Indica fecha y número del reporte conocido.');
        await ctx.saveAnchor({date,number});preview();
      }catch(e){ctx.fail(e)}};
      for(const b of form.querySelectorAll('[data-export]'))b.onclick=async()=>{try{
        const report=get(),writer=await import('./report-vendor.js?v=19');
        if(kind==='daily')for(const key of signatures)localStorage.setItem('iw-report-'+key,form.elements[key].value);
        if(b.dataset.export==='share'){
          const files=[writer.makePDF(report,true)];
          if(!navigator.canShare?.({files}))throw Error('Descarga el PDF para adjuntarlo.');
          await navigator.share({files});
        }else if(b.dataset.export==='xlsx')await writer.makeExcel(report);
        else await writer.makePDF(report);
      }catch(e){if(e.name!=='AbortError')ctx.fail(e)}};
      preview();
    }
    dialog.showModal();dialog.scrollTop=0;
  };
}
