import {destinations} from './reports.js?v=12';
const norm=x=>String(x??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
const key=x=>[x.date,x.data.dest,x.data.pid,Math.round(x.data.q*1e6)].join('|');
function dest(x){
  const s=String(x??'').trim().toUpperCase().replace(/\s+/g,'');
  const m=s.match(/^(PSC|PISCINA|PRE|PRECRIA)?0*(\d{1,2})(B?)$/);
  return m?(m[1]?.startsWith('PRE')?'PRE':'')+m[2].padStart(2,'0')+m[3]:s;
}
function dateValue(x,fallback){
  if(x==null||x==='')return fallback;
  if(typeof x==='number'&&x>0&&x<110000)return new Date((x-25569)*86400000).toISOString().slice(0,10);
  const s=String(x).trim(),m=s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  return m?m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'):s;
}
export function parseConsumptions(sheet,items,fallback){
  const byCode=new Map(items.map(p=>[norm(p.code),p])),byName=new Map(items.map(p=>[norm(p.name),p]));
  const resolve=v=>byCode.get(norm(v))||byCode.get(norm(v).padStart(5,'0'))||byName.get(norm(v));
  let spec=null;const lines=[];
  for(const {n,cells} of sheet.rows){
    if(!spec){if(n>25)break;const head=cells.map(norm),col=re=>head.findIndex(h=>h&&re.test(h));
      const place=col(/^(psc|piscina|destino|precria)/),q=col(/^cantidad/),id=col(/^(codigo|cod\b|producto)/),name=col(/^(descripcion|nombre)/),date=col(/^fecha/);
      const matrix=cells.map((v,i)=>({p:resolve(v),i})).filter(t=>t?.p);
      if(place>=0&&((q>=0&&(id>=0||name>=0))||matrix.length))spec={place,q,id,name,date,matrix};continue}
    const {place,q,id,name,date,matrix}=spec;
    const add=(p,v)=>{
      if(v==null||v==='')return;
      const quantity=typeof v==='number'?v:Number(String(v).trim().replace(',','.')),destination=dest(cells[place]),day=dateValue(cells[date],fallback);
      const errors=[];
      if(!p)errors.push('Producto desconocido');if(!(quantity>0&&Number.isFinite(quantity)))errors.push('Cantidad inválida');
      if(!destinations.includes(destination))errors.push('Destino inválido');
      if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||!Number.isFinite(Date.parse(day))||new Date(day).toISOString().slice(0,10)!==day)errors.push('Fecha inválida');
      lines.push({row:n,date:day,p,quantity,destination,errors});
    };
    if(q>=0)add(resolve(cells[id])||resolve(cells[name]),cells[q]);
    else for(const {p,i} of matrix)if(Number(cells[i])!==0)add(p,cells[i]);
  }
  return spec?lines:null;
}
export function installImport(ui){
  const {$,open,on,grid,button,label,input,select,opt,current,directMoves,save,record,today,exportSheets,toast,fail,escape}=ui;
  on('#consumption-template',()=>exportSheets('Plantilla-consumo',[{name:'CONSUMO',rows:[['Fecha','Código','Nombre','Cantidad','Piscina']]},{name:'Catálogo',rows:[['Código','Descripción','Medida'],...current().map(p=>[p.code,p.name,p.unit])]}]));
  $('#consumption-file').onchange=async e=>{
    const file=e.target.files[0];e.target.value='';if(!file)return;
    try{
      const reader=await import('./report-vendor.js?v=12');
      const sheets=(await reader.readSheets(file)).filter(s=>!/(indirect|combustible|diesel|tanque|lubricante)/i.test(s.name)&&parseConsumptions(s,current(),today())?.length);
      if(!sheets.length)throw Error('No encontré filas de consumo. Revisa los encabezados de la plantilla.');
      open('Revisar consumo de Excel','<p>'+escape(file.name)+'</p><form>'+'<fieldset><legend>Hojas de consumo</legend>'+sheets.map((s,i)=>'<label><input type="checkbox" name="sheet" value="'+i+'" '+(i===0?'checked':'')+'> '+escape(s.name)+'</label>').join('')+'</fieldset>'+label('Fecha para filas sin fecha',input('day','date',today(),'required'))+label('Incluir posibles duplicados',input('duplicates','checkbox'))+'</form><div id="import-lines"></div>'+button('import-confirm','Registrar salidas'));
      const form=$('#direct-dialog form');let batch=[],ready=[],running=false;
      function preview(){
        const selected=new FormData(form).getAll('sheet').map(i=>sheets[+i]),rows=selected.flatMap(s=>parseConsumptions(s,current(),form.elements.day.value).map(r=>({...r,sheet:s.name}))),exists=new Set(directMoves().filter(m=>m.type==='salida').map(key)),stock=new Map(current().map(p=>[p.id,p.stock]));
        ready=[];batch=[];let errors=0,duplicates=0;
        const table=rows.map(r=>{const k=key({date:r.date,data:{dest:r.destination,pid:r.p?.id,q:r.quantity}}),dup=exists.has(k);exists.add(k);
          let reason=r.errors.join(' · ');if(reason)errors++;
          else if(dup&&!form.elements.duplicates.checked){duplicates++;reason='Posible duplicado'}
          else{const balance=stock.get(r.p.id)-r.quantity;if(balance < -0.0000001){reason='Stock insuficiente';errors++}else{ready.push(r);stock.set(r.p.id,balance)}}
          return [r.sheet+' · '+r.row,r.date,r.p?.name||'—',r.destination,r.quantity,reason||'Listo']});
        $('#import-lines').innerHTML='<p>'+ready.length+' listas · '+errors+' con error · '+duplicates+' duplicadas · '+selected.reduce((total,s)=>total+s.hidden,0)+' filas ocultas excluidas</p>'+grid(['Hoja / fila','Fecha','Producto','Destino','Cantidad','Estado'],table);
        $('#import-confirm').disabled=!ready.length;
      }
      form.oninput=preview;preview();
      on('#import-confirm',async()=>{
        if(running)return;running=true;form.inert=true;
        try{if(!batch.length)batch=ready.map(r=>{const e=record('salida',{pid:r.p.id,dest:r.destination,q:r.quantity});e.date=r.date;return e});
          while(batch.length){const part=batch.slice(0,100);await save(part);batch.splice(0,part.length)}
          $('#direct-dialog').close();toast(navigator.onLine?'Consumos registrados.':'Consumos pendientes de enviar al recuperar conexión.')}finally{running=false;form.inert=false}
      });
    }catch(e){fail(e)}
  };
}

export function planInventoryImport(sheets,products,moves,record){
  const named=n=>sheets.find(s=>norm(s.name)===n);
  const text=v=>String(v??'').trim(),canon=v=>typeof v==='number'?String(v).padStart(5,'0'):text(v);
  const stock=named('stock'),incoming=named('entradas'),out=named('salidas');
  if(!stock||!incoming||!out)throw Error('El libro debe incluir Stock, Entradas y Salidas.');
  function table(sheet){
    const i=sheet.rows.findIndex(r=>r.cells.some(c=>norm(c).startsWith('codigo')));
    if(i<0)throw Error('Faltan encabezados en '+sheet.name);
    const h=sheet.rows[i].cells.map(norm);
    const col=(regex,required=true)=>{const n=h.findIndex(c=>c&&regex.test(c));if(n<0&&required)throw Error('Encabezados incompatibles en '+sheet.name);return n};
    return {rows:sheet.rows.slice(i+1),col};
  }
  const lookup=new Map(products.map(p=>[norm(p.code),p])),batch=[],t=table(stock);
  const code=t.col(/^codigo/),name=t.col(/^descrip/),unit=t.col(/^(unidad|medida)/),opening=t.col(/^inicio/),category=t.col(/^categoria/,false);
  const added=new Set(),allowed=new Set();
  for(const {cells:r,n} of t.rows){
    if(!text(r[code])||!text(r[name])||(category>=0&&norm(r[category]).includes('indirect')))continue;
    const c=canon(r[code]),p=lookup.get(norm(c)),u=text(r[unit]);if(!u)throw Error('Falta la unidad en Stock, fila '+n);
    if(p&&(p.name!==text(r[name])||p.unit!==u))throw Error('El producto '+c+' difiere del catálogo actual.');
    if(allowed.has(norm(c)))throw Error('Código duplicado en Stock: '+c);
    allowed.add(norm(c));if(p)continue;
    const value=Number(r[opening]||0);if(!Number.isFinite(value)||value<0)throw Error('Apertura inválida en Stock, fila '+n);
    const pid=crypto.randomUUID(),data={pid,code:c,name:text(r[name]),unit:u,min:0,cat:'directo'};
    batch.push(record('create',data));if(value)batch.push(record('ajuste',{pid,q:value,expected:0,doc:'INVENTARIO INICIAL'}));lookup.set(norm(c),{id:pid,...data});added.add(pid);
  }
  const sig=e=>JSON.stringify([e.type,e.date,e.data.pid,e.data.q,text(e.type==='entrada'?e.data.doc:e.data.dest)]);
  const seen=new Map(),ids=new Map(moves.map(m=>[m.id,sig(m)]));
  for(const m of moves){const k=sig(m);seen.set(k,(seen.get(k)||0)+1)}
  const available=new Map([...lookup.values()].map(p=>[p.id,p.stock||0]));
  for(const e of batch)if(e.type==='ajuste')available.set(e.data.pid,e.data.q);
  let skipped=0;
  for(const [sheet,type] of [[incoming,'entrada'],[out,'salida']]){
    const {rows,col}=table(sheet),ic=col(/^codigo/),iq=col(/^cantidad/),it=col(/^fecha/),ir=col(type==='entrada'?/^(doc|nro)/:/^(psc|piscina|destino)/),ii=col(/^idmovimiento/,false);
    for(const {cells:r,n} of rows){
      const q=Number(r[iq]);if(!r[iq]||q===0)continue;const c=norm(canon(r[ic])),p=lookup.get(c);
      if(!p||!allowed.has(c))throw Error('Producto fuera del catálogo directo en '+sheet.name+', fila '+n);
      const day=dateValue(r[it],''),ref=text(r[ir]);
      if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||!Number.isFinite(q)||q<=0)throw Error('Fecha o cantidad inválida en '+sheet.name+', fila '+n);
      const e=record(type,{pid:p.id,q,[type==='entrada'?'doc':'dest']:ref});e.date=day;
      const k=sig(e),id=r[ii];
      if(id&&ids.has(id)&&ids.get(id)!==k)throw Error('Movimiento modificado. Revisa '+sheet.name+', fila '+n);
      if((id&&ids.has(id))||(seen.get(k)||0)>0){seen.set(k,Math.max(0,(seen.get(k)||0)-1));skipped++;continue}
      if(id&&/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(id)){e.id=id;ids.set(id,k)}
      const stock=available.get(p.id)+(type==='entrada'?q:-q);if(stock< -0.000001)throw Error('Saldo insuficiente para importar '+p.code+' en '+sheet.name+', fila '+n);
      available.set(p.id,stock);batch.push(e);
    }
  }
  return {batch,added:added.size,skipped};
}
