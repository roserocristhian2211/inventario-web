import {destinations,familyOf,compareProducts} from './reports.js?v=12';

export function measurements(events){
  const current=new Map();
  for(const e of [...events].sort((a,b)=>(a.entryOrder||0)-(b.entryOrder||0)||a.ts.localeCompare(b.ts))){
    if(e.type==='reiniciarConteo')current.clear();
    if(e.type==='medicion')current.set(e.data.pid,e);
  }
  return current;
}
export function allocations(events){
  const removed=new Set(events.filter(e=>e.type==='anular').map(e=>e.data.target));
  return events.filter(e=>e.type==='reparto'&&!removed.has(e.id)).sort((a,b)=>a.date.localeCompare(b.date)||a.entryOrder-b.entryOrder);
}
export function pondRows(moves,events,from,to){
  if(!from||!to||from>to)throw Error('Revisa las fechas del período.');
  const portions=moves.filter(m=>m.type==='salida'&&m.date>=from&&m.date<=to).map(m=>({m,dest:m.data.dest||'Sin destino',q:m.q}));
  const totals=new Map();
  function row(m,dest){
    const key=JSON.stringify([dest,m.data.pid,m.unit]);
    if(!totals.has(key))totals.set(key,{pid:m.data.pid,unit:m.unit,dest,direct:0,received:0,sent:0});
    return totals.get(key);
  }
  for(const x of portions)row(x.m,x.dest).direct+=x.q;
  for(const {data:r,date} of allocations(events)){
    if(date>to)continue;
    const parts=new Map();
    // Merge each movement's portions at a destination before the next transfer.
    for(const x of portions.splice(0)){
      const add=(dest,q)=>{const key=JSON.stringify([x.m.id,dest]);const part=parts.get(key)||{m:x.m,dest,q:0};part.q+=q;parts.set(key,part)};
      if(x.dest!==r.origin||x.m.date<r.from||x.m.date>r.to){add(x.dest,x.q);continue}
      row(x.m,x.dest).sent+=x.q;
      let remaining=x.q;
      r.targets.forEach((t,i)=>{const q=i===r.targets.length-1?remaining:x.q*t.pct/100;remaining-=q;row(x.m,t.dest).received+=q;add(t.dest,q)});
    }
    for(const x of parts.values())portions.push(x);
  }
  return [...totals.values()].map(r=>({...r,net:r.direct+r.received-r.sent}));
}
export function weightFactor(p,unit){
  const u=unit.toUpperCase();if(u==='KG')return 2.20462262;if(/^LBS?$/.test(u))return 1;if(u!=='SC')return 0;
  if(p.lbPerSack)return p.lbPerSack;
  const lbs=p.name.match(/(\d+(?:[.,]\d+)?)\s*LBS?\b/i);if(lbs)return Number(lbs[1].replace(',','.'));
  const kg=p.name.match(/(\d+(?:[.,]\d+)?)\s*KG\b/i);
  return kg?Number(kg[1].replace(',','.'))*2.20462262:familyOf(p)==='balanceado'?55:0;
}
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const opt=(k,v=k)=>'<option value="'+escape(k)+'">'+escape(v)+'</option>';
const num=v=>typeof v==='number'?v.toLocaleString('es-EC',{maximumFractionDigits:4}):v;
const grid=(head,rows)=>'<div class="report-scroll"><table><thead><tr>'+head.map(h=>'<th>'+escape(h)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(v=>'<td>'+escape(num(v))+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>';
const button=(id,title)=>'<button class="btn outline" type="button" id="'+id+'">'+title+'</button>';
const label=(title,control)=>'<label>'+title+control+'</label>';
const input=(name,type,value='',extra='')=>'<input name="'+name+'" type="'+type+'" value="'+escape(value)+'" '+extra+'>';
export function installDirect(ctx){
  const {$,today,admin,products,events,moves,save,record,toast,fail}=ctx,dialog=$('#direct-dialog');
  const getProduct=id=>products().find(p=>p.id===id);
  const on=(id,fn)=>{const el=$(id);if(el)el.onclick=async()=>{if(el.disabled)return;el.disabled=true;try{await fn()}catch(e){fail(e)}finally{el.disabled=false}}};
  let parentView=null;
  function open(title,html,nested=false){
    parentView=nested&&dialog.open?{nodes:[...dialog.childNodes],scroll:dialog.scrollTop,focus:document.activeElement}:null;
    dialog.innerHTML='<button type="button" class="back-link" id="direct-back"><span aria-hidden="true">←</span> Volver</button><div class="modalhead"><h2>'+escape(title)+'</h2>'+button('direct-close','Cerrar')+'</div>'+html;
    on('#direct-back',()=>{
      if(!parentView){dialog.close();return}
      const {nodes,scroll,focus}=parentView;parentView=null;
      dialog.replaceChildren(...nodes);focus?.focus({preventScroll:true});dialog.scrollTop=scroll;
    });
    dialog.scrollTop=0;
    on('#direct-close',()=>dialog.close());if(!dialog.open)dialog.showModal();
  }
  const select=(name,options)=>'<select name="'+name+'">'+options+'</select>';
  const controls=()=>new FormData(dialog.querySelector('form'));
  const exportSheets=async(name,sheets)=>(await import('./report-vendor.js?v=12')).exportSheets(name,sheets);
  const direct=()=>products().filter(p=>p.cat==='directo');
  const current=()=>direct().filter(p=>!p.deleted);
  const directMoves=()=>{const ids=new Set(direct().map(p=>p.id));return moves().filter(m=>ids.has(m.data.pid))};
  function countData(){
    const counts=measurements(events());
    return current().map(p=>{
      const e=counts.get(p.id),diff=e?Math.round((e.data.physical-p.stock)*1e6)/1e6:null;
      return {p,e,diff,state:!e?'Pendiente':diff===0?'Correcto':diff<0?'Faltante':'Excedente'};
    });
  }
  const countSheet=()=>({name:'Conteo',rows:[['Código','Producto','Medida','Sistema','Físico','Diferencia','Estado','Contado por'],...countData().map(({p,e,diff,state})=>[p.code,p.name,p.unit,p.stock,e?.data.physical??'',diff??'',state,e?.data.actor||''])]});
  function countsView(){
    const rows=countData(),filter=$('#count-state').value,query=$('#count-search').value.toLowerCase();
    $('#count-summary').textContent=['Pendiente','Faltante','Excedente','Correcto'].map(s=>s+': '+rows.filter(r=>r.state===s).length).join(' · ');
    $('#count-results').innerHTML=rows.filter(r=>(!filter||r.state===filter)&&[r.p.name,r.p.code].some(v=>v.toLowerCase().includes(query))).map(({p,e,diff,state})=>'<div class="item"><div class="item-body"><h3>'+escape(p.name)+'</h3><p>'+escape(p.code)+' · Sistema '+num(p.stock)+' '+escape(p.unit)+(e?' · Físico '+num(e.data.physical)+' · '+escape(e.data.actor):'')+'</p><span>'+state+(diff?' '+num(Math.abs(diff)):'')+'</span></div><button class="small" data-measure="'+escape(p.id)+'">Contar</button></div>').join('');
    $('#count-reset').hidden=!admin();
  }
  $('#count-state').onchange=$('#count-search').oninput=countsView;
  on('#count-export',()=>exportSheets('Conteo-'+today(),[countSheet()]));
  on('#count-reset',async()=>{if(current().length&&confirm('¿Iniciar un nuevo conteo para todos los usuarios?')){await save([record('reiniciarConteo',{pid:current()[0].id})]);countsView()}});
  function productView(p){
    const history=directMoves().filter(m=>m.data.pid===p.id);
    const total=type=>history.filter(m=>m.type===type).reduce((n,m)=>n+m.q,0);
    open(p.name,'<p>'+escape(p.code)+' · '+escape(p.unit)+' · Mínimo '+num(p.min)+'</p>'+grid(['Inicio y ajustes','Entradas','Salidas','Saldo'],[[total('ajuste'),total('entrada'),total('salida'),p.stock]])+'<div class="actions">'+button('product-in','Entrada')+button('product-out','Salida')+'</div><h3>Historial reciente</h3>'+grid(['Fecha','Tipo','Cantidad','Documento / destino','Registrado por'],history.slice(0,30).map(m=>[m.date,m.type,m.q,m.data.doc||m.data.dest,m.data.actor])));
    for(const [id,type] of [['#product-in','entrada'],['#product-out','salida']])on(id,()=>{ctx.openMove(p.id,type)});
  }
  function movementView(m){
    const d=m.data,group=d.group?directMoves().filter(x=>x.data.group===d.group):[m];
    open(m.type==='entrada'?'Detalle de la entrada':'Detalle de la salida','<article class="movement-card"><h3>'+escape(m.prod)+'</h3><p>'+escape(m.code||'')+' · '+escape(m.date)+' · '+escape(d.dest||d.doc||'')+'</p><strong class="move-amount '+(m.type==='salida'?'dispatch':'')+'">'+(m.type==='salida'?'−':'+')+num(m.q)+' '+escape(m.unit)+'</strong><p>Registrado por '+escape(d.actor)+' · '+escape(new Date(m.ts).toLocaleString('es-EC'))+'</p></article>'+(d.editedBy?'<p>Modificado por '+escape(d.editedBy)+'</p>':'')+(group.length>1?grid(['Productos de la guía','Cantidad','Medida'],group.map(x=>[x.prod,x.q,x.unit])):'')+(admin()?'<div class="actions">'+button('movement-edit','Editar')+(isGuide(m)?button('guide-edit','Modificar guía completa'):'')+button('movement-remove','Eliminar')+'</div>':''));
    on('#movement-edit',()=>editMovement(m));on('#guide-edit',()=>guideForm(group));
    on('#movement-remove',async()=>{if(confirm('¿Eliminar este movimiento y recalcular el saldo?')){await save([record('anular',{pid:d.pid,target:m.id})]);dialog.close();toast('Movimiento anulado.')}});
  }
  const isGuide=m=>m.type==='entrada'&&m.data.group;
  function guideForm(group){
    const first=group[0],options=current().map(p=>opt(p.id,p.code+' · '+p.name)).join('');
    open('Modificar guía completa','<form>'+label('Guía o factura',input('doc','text',first.data.doc,'maxlength="75"'))+label('Fecha',input('date','date',first.date,'required'))+'<div id="guide-lines"></div>'+button('guide-add','Agregar producto')+button('guide-save','Guardar guía')+'</form>',true);
    function add(original){
      const parent=$('#guide-lines');if(parent.children.length>=100)return;
      const line=document.createElement('fieldset');line.dataset.original=original?.id||'';
      line.innerHTML=label('Producto',select('pid',options))+label('Cantidad',input('q','number',original?.q||'','min="0.000001" step="any" required'))+'<button type="button" class="small">Quitar producto</button>';
      if(original)line.querySelector('select').value=original.data.pid;
      line.querySelector('button').onclick=()=>line.remove();parent.append(line);
    }
    group.forEach(add);on('#guide-add',()=>add(null));let key,batch;
    on('#guide-save',async()=>{
      const form=dialog.querySelector('form'),lines=[...$('#guide-lines').children];
      if(!form.reportValidity()||!lines.length)throw Error('La guía debe incluir al menos un producto.');
      const rows=lines.map(line=>({original:line.dataset.original,pid:line.querySelector('select').value,q:Number(line.querySelector('input').value)}));
      const signature=JSON.stringify([form.elements.doc.value,form.elements.date.value,rows]);
      if(key!==signature){
        const replacements=rows.map(r=>{const old=group.find(m=>m.id===r.original),data={...(old||first).data,pid:r.pid,q:r.q,doc:form.elements.doc.value};delete data.replaces;if(old)data.replaces=old.id;
          const e=record('entrada',data);e.date=form.elements.date.value;return e});
        batch=[...replacements,...group.map(m=>record('anular',{pid:m.data.pid,target:m.id}))];key=signature;
      }
      await ctx.saveGuide(batch);dialog.close();toast('Guía actualizada.');
    });
  }
  function editMovement(m){
    const d=m.data,isIn=m.type==='entrada';
    open('Modificar '+m.type,'<form>'+label('Producto',select('pid',current().map(p=>opt(p.id,p.code+' · '+p.name)).join('')))+label('Cantidad',input('q','number',m.q,'min="0.000001" step="any" required'))+label('Fecha',input('date','date',m.date,'required'))+(isIn?label('Guía',input('doc','text',d.doc,'maxlength="75"')):label('Destino',select('dest',[...new Set([...destinations,d.dest].filter(Boolean))].map(v=>opt(v)).join(''))))+label('Observación',input('note','text',d.note,'maxlength="180"'))+button('edit-commit','Guardar corrección')+'</form>',true);
    dialog.querySelector('[name=pid]').value=d.pid;if(!isIn)dialog.querySelector('[name=dest]').value=d.dest;
    let batch,key;
    on('#edit-commit',async()=>{if(!dialog.querySelector('form').reportValidity())return;const f=Object.fromEntries(controls()),signature=JSON.stringify(f);
      if(key!==signature){const replacement=record(m.type,{...d,...f,q:Number(f.q),replaces:m.id});replacement.date=f.date;
        const cancel=record('anular',{pid:d.pid,target:m.id});batch=isIn?[replacement,cancel]:[cancel,replacement];key=signature}
      await save(batch);dialog.close();toast('Movimiento corregido.')});
  }
  function poolView(){
    open('Consumo por piscina y precría','<form>'+label('Desde',input('from','date',today().slice(0,7)+'-01'))+label('Hasta',input('to','date',today()))+label('Grupo',select('family',opt('','Todos')+opt('balanceado','Balanceado')+opt('insumo','Insumos')))+label('Destino',select('dest',opt('','Todos')+[...new Set([...destinations,...directMoves().map(m=>m.data.dest).filter(Boolean)])].map(d=>opt(d)).join('')))+'</form><div id="pool-results"></div><div class="actions">'+button('pool-excel','Descargar Excel')+(admin()?button('transfer-new','Nuevo traspaso'):'')+'</div><h3>Traspasos registrados</h3><div id="pool-history"></div>');
    const head=['Destino','Código','Producto','Unidad','Directo','Recibido','Transferido','Acumulado','Libras'];
    let rows=[];
    function paint(){try{const f=Object.fromEntries(controls());rows=pondRows(directMoves(),events(),f.from,f.to).filter(r=>(!f.dest||r.dest===f.dest)&&(!f.family||familyOf(getProduct(r.pid))===f.family)).sort((a,b)=>a.dest.localeCompare(b.dest,'es',{numeric:true})||compareProducts(getProduct(a.pid),getProduct(b.pid))).map(r=>{const p=getProduct(r.pid),factor=weightFactor(p,r.unit);return [r.dest,p.code,p.name,r.unit,r.direct,r.received,r.sent,r.net,factor?r.net*factor:'']});
      const summary=new Map();
      for(const row of rows){const d=row[0],u=row[3],v=summary.get(d)||{lb:0,rest:new Map()};if(row[8]!=='')v.lb+=row[8];else v.rest.set(u,(v.rest.get(u)||0)+row[7]);summary.set(d,v)}
      $('#pool-results').innerHTML=grid(['Destino','Consumo acumulado'],[...summary].map(([d,v])=>[d,[...(v.lb?[num(v.lb)+' lb']:[]),...[...v.rest].map(([u,q])=>num(q)+' '+u)].join(' · ')||'0']))+grid(head,rows);$('#pool-excel').disabled=!rows.length}catch(e){$('#pool-results').textContent=e.message;$('#pool-excel').disabled=true}}
    dialog.querySelector('form').oninput=paint;paint();on('#pool-excel',()=>exportSheets('Consumo-por-piscina',[{name:'Por piscina',rows:[head,...rows]}]));
    const list=allocations(events()).reverse();
    $('#pool-history').innerHTML=list.map(e=>'<p>'+escape(e.date)+' · '+escape(e.data.origin)+' → '+escape(e.data.targets.map(t=>t.dest+' '+num(t.pct)+'%').join(', '))+' · '+escape(e.data.from)+' / '+escape(e.data.to)+(admin()?' <button data-remove-transfer="'+escape(e.id)+'">Eliminar</button>':'')+'</p>').join('')||'Sin traspasos.';
    on('#transfer-new',transferForm);
  }
  function transferForm(){
    const choices=destinations.map(d=>opt(d)).join('');
    open('Traspaso de consumo','<form>'+label('Origen',select('origin',choices))+label('Consumo desde',input('from','date',today().slice(0,7)+'-01','required'))+label('Consumo hasta',input('to','date',today(),'required'))+label('Fecha de traspaso',input('when','date',today(),'required'))+'<fieldset><legend>Porcentaje por destino</legend><div class="ponds">'+destinations.map(d=>label(d,input(d,'number','','min="0" max="100" step="any"'))).join('')+'</div></fieldset><output id="transfer-total"></output>'+label('Nota',input('note','text','','maxlength="180"'))+button('transfer-save','Guardar traspaso')+'</form>',true);
    function values(){const f=Object.fromEntries(controls());const targets=destinations.map(dest=>({dest,pct:Number(f[dest])})).filter(t=>t.pct>0);return {f,targets,sum:targets.reduce((s,t)=>s+t.pct,0)}}
    const paint=()=>{$('#transfer-total').textContent='Total asignado: '+num(values().sum)+'%'};dialog.querySelector('form').oninput=paint;paint();
    let pending,key;
    on('#transfer-save',async()=>{const {f,targets,sum}=values();if(!current().length||!dialog.querySelector('form').reportValidity()||Math.abs(sum-100)>0.000001||targets.some(t=>t.dest===f.origin)||f.from>f.to||f.from>f.when)throw Error('Revisa el período, los destinos y la suma de 100%.');
      const data={pid:current()[0].id,origin:f.origin,targets,from:f.from,to:f.to,note:f.note};const k=JSON.stringify([data,f.when]);if(k!==key){pending=record('reparto',data);pending.date=f.when;key=k}await save([pending]);poolView()});
  }
  document.addEventListener('click',async e=>{const t=e.target.closest('[data-card],[data-detail],[data-measure],[data-remove-transfer]');if(!t)return;try{
    if(t.dataset.card)productView(getProduct(t.dataset.card));
    if(t.dataset.detail)movementView(moves().find(m=>m.id===t.dataset.detail));
    if(t.dataset.measure){const p=getProduct(t.dataset.measure);open('Conteo físico','<form><h3>'+escape(p.name)+'</h3><p>Sistema: '+num(p.stock)+' '+escape(p.unit)+'</p>'+label('Cantidad física',input('physical','number',measurements(events()).get(p.id)?.data.physical??'','min="0" step="any" required'))+button('count-save','Guardar conteo')+'</form>');
      on('#count-save',async()=>{if(!dialog.querySelector('form').reportValidity())return;const physical=Number(controls().get('physical'));await save([record('medicion',{pid:p.id,physical})]);countsView();dialog.close()})}
    if(t.dataset.removeTransfer&&admin()&&confirm('¿Eliminar el traspaso?')){const transfer=allocations(events()).find(x=>x.id===t.dataset.removeTransfer);await save([record('anular',{pid:transfer.data.pid,target:transfer.id})]);poolView()}
  }catch(error){fail(error)}});
  on('#pools',poolView);
  function workbook(){const stock={name:'Stock',rows:[['Código Producto','Descripción','Medida','Inicio y ajustes','Entradas','Salidas','Stock Actual'],...current().map(p=>{const hist=directMoves().filter(m=>m.data.pid===p.id),sum=t=>hist.filter(m=>m.type===t).reduce((n,m)=>n+m.q,0);return [p.code,p.name,p.unit,sum('ajuste'),sum('entrada'),sum('salida'),p.stock]})]};
    const sheets=['entrada','salida'].map(t=>({name:t==='entrada'?'Entradas':'Salidas',rows:[[t==='entrada'?'Nro. Documento':'PSC/PRE','Fecha','Codigo Producto','Descripción','Medida','Cantidad','IdMovimiento'],...directMoves().filter(m=>m.type===t).map(m=>[t==='entrada'?m.data.doc:m.data.dest,m.date,m.code,m.prod,m.unit,m.q,m.id])]}));
    return [stock,...sheets,countSheet()]}
  on('#book-export',()=>exportSheets('Inventario-directo-'+today(),workbook()));
  on('#book-drive',async()=>{const {workbookBlob}=await import('./report-vendor.js?v=12');await ctx.uploadBook(await workbookBlob(workbook()));toast('Libro guardado en Respaldos Inventario Web.')});
  $('#book-file').onchange=async e=>{
    const file=e.target.files[0];e.target.value='';if(!admin()||!file)return;
    try{const {readSheets}=await import('./report-vendor.js?v=12');const plan=ctx.planInventoryImport(await readSheets(file),current(),directMoves(),record);
      open('Importar inventario','<p>'+escape(file.name)+'</p><p>'+plan.added+' productos nuevos · '+plan.batch.filter(e=>['entrada','salida'].includes(e.type)).length+' movimientos nuevos · '+plan.skipped+' movimientos existentes.</p>'+button('book-confirm','Incorporar registros'));
      $('#book-confirm').disabled=!plan.batch.length;
      on('#book-confirm',async()=>{while(plan.batch.length){const part=plan.batch.slice(0,100);await save(part);plan.batch.splice(0,part.length)}dialog.close();toast('Inventario incorporado.')});
    }catch(error){fail(error)}
  };
  ctx.installImport({$,open,on,controls,grid,button,label,input,select,opt,current,directMoves,save,record,today,exportSheets,toast,fail,escape});
  return countsView;
}
