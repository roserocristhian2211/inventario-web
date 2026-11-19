import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {uuid} from './ledger.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const parse=text=>new DOMParser().parseFromString(text,'application/xml');
const xml=node=>new XMLSerializer().serializeToString(node);
const nodes=(parent,tag)=>Array.from(parent.getElementsByTagNameNS(NS,tag));
const newNode=(doc,name)=>doc.createElementNS(NS,name);
const first=(parent,name)=>parent?.getElementsByTagNameNS(NS,name)[0];
const iso=v=>Number.isInteger(+v)&&+v>0&&+v<110000?new Date((+v-25569)*86400000).toISOString().slice(0,10):'';
const serial=day=>Date.parse(day+'T00:00:00Z')/86400000+25569;
const signature=m=>JSON.stringify([m.type,m.date,m.pid,m.q,m.ref,...(m.note||m.origin?[m.note||'',m.origin||'']:[])]);
function check(ok,msg){if(!ok)throw Error(msg)}
function fields(e){return {type:e.type,date:e.date,pid:e.data.pid,q:e.data.q,note:e.data.note||'',origin:e.data.origin||'',ref:String(e.type==='entrada'?e.data.doc||'':e.data.dest||'')}}
function cellValues(row,strings){
  const values={};
  for(const c of nodes(row,'c')){
    const key=c.getAttribute('r').replace(/\d+$/,'');
    const v=first(c,'v')?.textContent||'';
    values[key]=c.getAttribute('t')==='s'?strings[+v]:c.getAttribute('t')==='inlineStr'?nodes(c,'t').map(x=>x.textContent).join(''):v;
  }
  return values;
}
function setCell(row,col,value,{numeric=false,formula=false,style}={}){
  const doc=row.ownerDocument,r=col+row.getAttribute('r');
  let c=nodes(row,'c').find(x=>x.getAttribute('r')===r);
  if(!c){c=newNode(doc,'c');c.setAttribute('r',r);const after=nodes(row,'c').find(cell=>cell.getAttribute('r').replace(/\d+$/,'')>col);row.insertBefore(c,after||null)}
  if(style)c.setAttribute('s',style);
  const f=formula&&first(c,'f');
  while(c.firstChild)c.removeChild(c.firstChild);
  if(f)c.appendChild(f);
  if(numeric||f){
    if(numeric)c.removeAttribute('t');else c.setAttribute('t','str');
    const v=newNode(doc,'v');v.textContent=String(value);c.appendChild(v);
  }else{
    c.setAttribute('t','inlineStr');const is=newNode(doc,'is'),t=newNode(doc,'t');
    t.textContent=String(value);is.appendChild(t);c.appendChild(is);
  }
}
function syncOperations({parts,read,save,strings,events,catalog,byCode,voided}){
  const relNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const book=read('xl/workbook.xml'),sheets=first(book,'sheets'),rels=read('xl/_rels/workbook.xml.rels');
  const name='Operaciones bodega',headers=['Tipo','Fecha','Código','Cantidad','Origen','Destino','Desde','Hasta','Distribución (%)','Nota','Estado','IdOperacion'];
  const list=Array.from(rels.documentElement.childNodes).filter(x=>x.nodeType===1);
  let sheet=nodes(book,'sheet').find(s=>s.getAttribute('name')===name),path,doc,dirty=false;
  if(!sheet){
    check(!Object.keys(parts).some(k=>k.startsWith('xl/worksheets/sheet')&&strFromU8(parts[k]).includes('IdOperacion')),'Restablece el nombre Operaciones bodega de la hoja vinculada');
    const n=1+Math.max(...nodes(book,'sheet').map(s=>+s.getAttribute('sheetId')));
    let file=n;while(parts['xl/worksheets/sheet'+file+'.xml'])file++;
    const rid='rId'+(1+Math.max(0,...list.map(s=>Number(s.getAttribute('Id').replace(/^rId/,''))||0)));
    sheet=newNode(book,'sheet');sheet.setAttribute('name',name);sheet.setAttribute('sheetId',n);sheet.setAttributeNS(relNS,'r:id',rid);sheets.appendChild(sheet);
    const rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');
    rel.setAttribute('Id',rid);rel.setAttribute('Type',relNS+'/worksheet');rel.setAttribute('Target','worksheets/sheet'+file+'.xml');
    rels.documentElement.appendChild(rel);path='xl/'+rel.getAttribute('Target');
    const types=read('[Content_Types].xml'),type=types.createElementNS(types.documentElement.namespaceURI,'Override');
    type.setAttribute('PartName','/'+path);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
    types.documentElement.appendChild(type);save('[Content_Types].xml',types);save('xl/_rels/workbook.xml.rels',rels);save('xl/workbook.xml',book);
    doc=parse('<worksheet xmlns="'+NS+'"><cols><col min="1" max="11" width="20" customWidth="1"/><col min="12" max="12" width="40" customWidth="1"/></cols><sheetData><row r="1"/></sheetData></worksheet>');
    headers.forEach((h,i)=>setCell(first(doc,'row'),String.fromCharCode(65+i),h));dirty=true;
  }else{
    const rel=list.find(r=>r.getAttribute('Id')===sheet.getAttributeNS(relNS,'id'));
    check(rel,'No se encontró la hoja '+name);const target=rel.getAttribute('Target');path=target.startsWith('/')?target.slice(1):'xl/'+target;doc=read(path);
  }
  const rows=nodes(doc,'row'),encodedHeader=cellValues(rows[0],strings);
  check(headers.every((h,i)=>h===encodedHeader[String.fromCharCode(65+i)]),'Revisa los encabezados de '+name);
  const ops=new Map(events.filter(e=>['traspaso','medicion','reparto','reiniciarConteo'].includes(e.type)).map(e=>[e.id,e]));
  const used=new Set(),edits=[];
  const encode=e=>{const d=e.data;return [e.type,e.date,catalog.get(d.pid)?.code||'',d.physical??d.q??'',d.origin||'',d.dest||'',d.from||'',d.to||'',(d.targets||[]).map(t=>t.dest+'='+t.pct).join('; '),d.note||'']};
  function write(row,values){values.forEach((v,i)=>setCell(row,String.fromCharCode(65+i),v,{numeric:i===3&&typeof v==='number'}));dirty=true}
  let end=1;
  for(const row of rows.slice(1)){
    const n=+row.getAttribute('r');end=Math.max(end,n);const v=cellValues(row,strings);
    if(!v.A&&!v.L&&!v.C)continue;
    const id=v.L||crypto.randomUUID(),e=ops.get(id);
    check(uuid(id)&&!used.has(id)&&(!v.L||!events.some(x=>x.id===id)||e),'IdOperacion inválido en fila '+n);
    used.add(id);
    if(e&&(voided.has(id)||e.type==='reiniciarConteo')){const state=voided.has(id)?'Anulado':'Registrado';if(v.K!==state){write(row,[...encode(e),state,id])}continue}
    const date=x=>iso(x)||x||'',product=byCode.get(v.C);
    check(product&&['traspaso','medicion','reparto'].includes(v.A)&&(!e||e.type===v.A),'Revisa tipo y producto en '+name+' fila '+n);
    const op={id,type:v.A,date:date(v.B),data:{pid:product.pid}};
    if(op.type==='medicion'){check(v.D!==undefined&&v.D!=='','Falta la cantidad contada');op.data.physical=+v.D}
    else if(op.type==='traspaso')Object.assign(op.data,{q:Number(v.D),origin:v.E||'',dest:v.F||'',note:v.J||''});
    else Object.assign(op.data,{origin:v.E||'',from:date(v.G),to:date(v.H),targets:String(v.I||'').split(';').map(pair=>{const [dest,pct]=pair.trim().split('=');return {dest,pct:Number(pct)}}),note:v.J||''});
    if(!e||JSON.stringify(encode(e))!==JSON.stringify(encode(op)))edits.push(op);
    if(!v.L||v.K!=='Registrado'||v.C!==product.code)write(row,[...encode(op),'Registrado',id]);
  }
  for(const e of ops.values())if(!used.has(e.id)){
    const row=newNode(doc,'row');row.setAttribute('r',++end);first(doc,'sheetData').appendChild(row);write(row,[...encode(e),voided.has(e.id)?'Anulado':'Registrado',e.id]);
  }
  if(dirty){first(doc,'dimension')?.setAttribute('ref','A1:L'+end);save(path,doc)}
  return {edits,dirty};
}

// The initial import used BASE's two existing tables. Bind only that layout.
export function reconcileWorkbook(bytes,events,linked,productLinks=new Map()){
  const parts=unzipSync(bytes),read=name=>{check(parts[name],'Falta una parte del libro');return parse(strFromU8(parts[name]))};
  const save=(name,doc)=>{parts[name]=strToU8(xml(doc))};
  const strings=(parts['xl/sharedStrings.xml']?nodes(read('xl/sharedStrings.xml'),'si'):[]).map(n=>nodes(n,'t').map(t=>t.textContent).join(''));
  const retired=new Set(),catalog=new Map(), byId=new Map(), byCode=new Map(), byName=new Map();
  const voided=new Set(events.filter(e=>e.type==="anular").map(e=>e.data.target));
  for(const e of events){
    if(e.type==='create'||e.type==='edit')catalog.set(e.data.pid,e.data);
    if(e.type==='delete')retired.add(e.data.pid);
    if(e.type==='entrada'||e.type==='salida')byId.set(e.id,e);
  }
  for(const p of catalog.values()){byCode.set(p.code,p);byName.set(p.name,p)}
  const stock=read('xl/worksheets/sheet1.xml'),stockRows=new Map();
  const stockTable=read('xl/tables/table1.xml'),table=stockTable.documentElement;
  const stockRange=table.getAttribute('ref').match(/^B3:[JKNO](\d+)$/);
  check(stockRange&&table.getAttribute('name')==='INVENTARIO','Revisa la tabla INVENTARIO de BASE');
  let stockEnd=+stockRange[1],stockDirty=false;
  const changes=[],productEvents=[],productsSeen=new Map(),seen=new Set(),legacy=new Map(),totals=new Map();let dirty=false;
  const day=new Date(Date.now()-5*3600000).toISOString().slice(0,10);
  const openings=new Map(),aliases=new Map(),original=new Map();
  for(const e of events){
    if(e.type==='create'){original.set(e.data.pid,e.data);aliases.set(e.data.code,e.data.pid)}
    if(e.type==='ajuste')openings.set(e.data.pid,(openings.get(e.data.pid)||0)+Math.round(e.data.q*1e6));
  }
  const data=first(stock,'sheetData'),rows=nodes(data,'row');
  const signatureOf=p=>[p.code,p.name,p.unit,(openings.get(p.pid)||0)/1e6,p.min||0,p.family||'',p.lbPerSack??''];
  const stockHeader=rows.find(r=>r.getAttribute('r')==='3');
  const extra=[['K','IdProducto'],['L','Mínimo'],['M','Grupo'],['N','Libras por saco'],['O','Estado']];
  const columns=first(table,'tableColumns');
  let columnId=Math.max(...nodes(columns,'tableColumn').map(col=>Number(col.getAttribute('id'))));
  const prior=nodes(columns,'tableColumn').length;
  for(const [i,[letter,label]] of extra.entries()){
    if(i+9<prior){check(nodes(columns,'tableColumn')[i+9].getAttribute('name')===label,'Revisa columna '+letter+' de Stock');continue}
    const col=newNode(stockTable,'tableColumn');col.setAttribute('id',++columnId);col.setAttribute('name',label);columns.appendChild(col);
    columns.setAttribute('count',10+i);setCell(stockHeader,letter,label);stockDirty=true;
  }
  function fillStock(row,values,id){
    for(const [i,col] of ['B','C','E','F','L','M','N'].entries())setCell(row,col,values[i],{numeric:[3,4,6].includes(i)&&values[i]!==''});
    setCell(row,'O',retired.has(id)?'Inactivo':'Activo');setCell(row,'D',values[0],{formula:true});setCell(row,'K',id);stockDirty=true;
  }
  for(const row of rows){
    const n=+row.getAttribute('r');if(n<4||n>stockEnd)continue;
    const v=cellValues(row,strings);if(!v.B&&!v.C&&!v.K)continue;
    const fromSheet=[String(v.B||'').trim(),String(v.C||'').trim(),String(v.E||'').trim(),Number(v.F||0),Number(v.L||0),String(v.M||'').trim().toLowerCase(),v.N?Number(v.N):''];
    check(fromSheet.slice(0,3).every(Boolean)&&Number.isFinite(fromSheet[3]),'Producto incompleto en Stock, fila '+n);
    let id=v.K||byCode.get(fromSheet[0])?.pid||aliases.get(fromSheet[0])||crypto.randomUUID();
    check(uuid(id)&&!productsSeen.has(id),'IdProducto repetido o inválido en Stock');
    let p=catalog.get(id);
    if(!p){
      check(!original.has(id),'Producto eliminado o identificador desconocido en Stock');
      p={pid:id,code:fromSheet[0],name:fromSheet[1],unit:fromSheet[2],cat:'directo',min:0};
      productEvents.push({id:crypto.randomUUID(),type:'create',date:day,data:p});catalog.set(id,p);
    }
    const app=signatureOf(p),baseline=productLinks.get(id)||[original.get(id)?.code??p.code,original.get(id)?.name??p.name,original.get(id)?.unit??p.unit,
      events.filter(e=>e.type==='ajuste'&&e.data.pid===id&&['INICIO','INVENTARIO INICIAL'].includes(e.data.doc)).reduce((sum,e)=>sum+e.data.q,0)];
    baseline[4]??=0;baseline[5]??='';baseline[6]??='';
    const resolved=app.map((value,i)=>{
      if(fromSheet[i]===value||fromSheet[i]===baseline[i])return value;
      check(value===baseline[i],'Cambio simultáneo en Stock, fila '+n+' ('+['código','descripción','medida','inicio','mínimo','grupo','libras por saco'][i]+')');return fromSheet[i];
    });
    if([0,1,2,4,5,6].some(i=>resolved[i]!==app[i])){
      p={...p,code:resolved[0],name:resolved[1],unit:resolved[2],min:resolved[4],family:resolved[5]||undefined,lbPerSack:resolved[6]||undefined};catalog.set(id,p);
      productEvents.push({id:crypto.randomUUID(),type:'edit',date:day,data:p});
    }
    if(resolved[3]!==app[3])productEvents.push({id:crypto.randomUUID(),type:'ajuste',date:day,data:{pid:id,q:resolved[3]-app[3],doc:'BASE',note:'Cambio de apertura en Excel'}});
    if(v.O!==(retired.has(id)?'Inactivo':'Activo')||!v.K||resolved.some((x,i)=>x!==fromSheet[i]))fillStock(row,resolved,id);
    productsSeen.set(id,resolved);stockRows.set(p.code,row);
  }
  for(const p of catalog.values()){
    if(productsSeen.has(p.pid))continue;
    check(!productLinks.has(p.pid)&&original.get(p.pid)?.actor!=='Historial BASE','Falta una fila vinculada en Stock: '+p.code);
    const last=nodes(data,'row').at(-1);check(+last.getAttribute('r')===stockEnd,'Hay contenido debajo de la tabla INVENTARIO');
    const row=last.cloneNode(true),n=++stockEnd;row.setAttribute('r',n);
    for(const cell of nodes(row,'c'))cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,n));
    data.appendChild(row);setCell(row,'A',n-3,{numeric:true});const values=signatureOf(p);
    fillStock(row,values,p.pid);productsSeen.set(p.pid,values);stockRows.set(p.code,row);
  }
  if(stockDirty){
    const ref='B3:O'+stockEnd;table.setAttribute('ref',ref);first(table,'autoFilter')?.setAttribute('ref',ref);first(stock,'dimension')?.setAttribute('ref','A1:O'+stockEnd);
  }
  for(const p of catalog.values()){byCode.set(p.code,p);byName.set(p.name,p)}
  const accumulate=(pid,kind,q)=>{
    const t=totals.get(pid)||{entrada:0,salida:0};t[kind]+=Math.round(q*1e6);totals.set(pid,t);
  };
  for(const e of byId.values())if(e.data.actor==='Historial BASE'&&!e.data.replaces&&!linked.has(e.id)){
    const k=signature(fields(e));if(!legacy.has(k))legacy.set(k,[]);legacy.get(k).push(e.id);
  }
  for(const [sheet,kind] of [[2,'entrada'],[3,'salida']]){
    const name=kind==='entrada'?'ENTRADAS':'SALIDAS',tabPath='xl/tables/table'+sheet+'.xml',sheetPath='xl/worksheets/sheet'+sheet+'.xml';
    const tab=read(tabPath),table=tab.documentElement,doc=read(sheetPath),rows=nodes(doc,'row'),range=table.getAttribute('ref').match(/^A3:([FGK])(\d+)$/);
    check(range&&table.getAttribute('name')===name,'Cambió la estructura de '+name);
    const headers=nodes(table,'tableColumn'),colNames=headers.map(c=>c.getAttribute('name'));
    check(colNames.slice(0,6).join('|')===[kind==='entrada'?'Nro. Documento':'PSC/PRE','Fecha','Codigo Producto','Descripción','Medida','Cantidad'].join('|'),'Revisa encabezados de '+name);
    const header=rows.find(r=>r.getAttribute('r')==='3'),end=+range[2];let sheetDirty=false;
    const extraColumns=[['G','IdMovimiento'],['H','Observación'],['I','Origen'],['J','Registrado por'],['K','Fecha de registro']];
    for(const [i,[letter,label]] of extraColumns.entries()){
      if(i+6<headers.length){check(colNames[i+6]===label,'Columna incompatible en '+name);continue}
      const columns=first(table,'tableColumns'),col=newNode(tab,'tableColumn');
      col.setAttribute('id',7+i);col.setAttribute('name',label);columns.appendChild(col);columns.setAttribute('count',7+i);
      setCell(header,letter,label);sheetDirty=true;
    }
    if(sheetDirty){const ref='A3:K'+end;table.setAttribute('ref',ref);first(table,'autoFilter')?.setAttribute('ref',ref)}
    const empty=[];let dateStyle,lastRow=end;
    for(const row of rows){
      const n=+row.getAttribute('r');if(n<=3||n>end)continue;
      const v=cellValues(row,strings);
      if(!v.F||+v.F===0){
        if(v.G&&voided.has(v.G))continue;
        check(!v.G,'Una fila vinculada fue vaciada: '+name+' '+n);
        if(!v.A&&!v.B&&!(kind==='entrada'?v.D:v.C))empty.push(row);
        continue;
      }
      const product=kind==='entrada'?(byName.get(v.D)||byCode.get(v.C)):(byCode.get(v.C)||byName.get(v.D));
      const movement={type:kind,date:iso(v.B),pid:product?.pid,q:Number(v.F),ref:String(v.A||''),note:v.H||'',origin:v.I||''};
      check(product&&movement.date&&Number.isFinite(movement.q)&&movement.q>0,'Revisa '+name+' fila '+n);
      dateStyle??=nodes(row,'c').find(c=>c.getAttribute('r')==='B'+n)?.getAttribute('s');
      let id=v.G;
      if(!id){const candidates=legacy.get(signature(movement));id=candidates?.pop()||crypto.randomUUID();setCell(row,'G',id);sheetDirty=true}
      check(uuid(id)&&!seen.has(id),'IdMovimiento repetido o inválido en '+name);
      const record=byId.get(id);
      if(record){
        if(headers.length<11){movement.note=record.data.note||'';movement.origin=record.data.origin||''}
        for(const [col,value] of [['H',movement.note],['I',movement.origin],['J',record.data.actor||''],['K',record.ts||'']]){
          if((v[col]||'')!==value){setCell(row,col,value);sheetDirty=true}
        }
      }
      if(voided.has(id)){
        check(byId.get(id)?.type===kind&&byId.has(id)&&signature(movement)===signature(fields(byId.get(id))),'Salida anulada con cambios en BASE');
        setCell(row,'F',0,{numeric:true});sheetDirty=true;continue;
      }
      const current=catalog.get(product.pid);
      for(const [col,value,formula] of [['C',current.code,kind==='entrada'],['D',current.name,kind==='salida'],['E',current.unit,true]]){
        if(v[col]!==value){setCell(row,col,value,{formula});sheetDirty=true}
      }
      accumulate(product.pid,kind,movement.q);seen.add(id);const previous=byId.get(id);
      if(!previous||signature(movement)!==signature(fields(previous))){
        check(!previous||linked.has(id),'No se pudo conciliar el historial inicial');
        check(!previous||(previous.type===kind&&linked.get(id)===signature(fields(previous))),'Movimiento modificado en ambos lados');
        changes.push({id,movement});
      }
    }
    for(const [id,e] of byId){
      if(voided.has(id)||e.type!==kind||seen.has(id))continue;
      check(!linked.has(id)&&(e.data.actor!=='Historial BASE'||e.data.replaces),'Falta una fila histórica o vinculada en '+name);
      let row=empty.shift();const product=catalog.get(e.data.pid);check(product,'Revisa catálogo de '+name);
      if(!row){
        check(+rows.at(-1).getAttribute('r')===end,'Hay contenido debajo de '+name);
        row=newNode(doc,'row');row.setAttribute('r',++lastRow);first(doc,'sheetData').appendChild(row);
        const ref='A3:K'+lastRow;table.setAttribute('ref',ref);first(table,'autoFilter')?.setAttribute('ref',ref);
      }
      const m=fields(e);accumulate(product.pid,kind,m.q);
      setCell(row,'A',m.ref);setCell(row,'B',serial(m.date),{numeric:true,style:dateStyle});
      setCell(row,'C',product.code,{formula:kind==='entrada'});setCell(row,'D',product.name,{formula:kind==='salida'});
      setCell(row,'E',product.unit,{formula:true});setCell(row,'F',m.q,{numeric:true});setCell(row,'G',id);
      for(const [col,value] of [['H',m.note],['I',m.origin],['J',e.data.actor],['K',e.ts]])if(value)setCell(row,col,value);
      seen.add(id);sheetDirty=true;
    }
    if(sheetDirty){
      for(const row of rows)if(row.hasAttribute('spans'))row.removeAttribute('spans');
      first(doc,'dimension')?.setAttribute('ref','A1:K'+lastRow);
      save(sheetPath,doc);save(tabPath,tab);dirty=true;
    }
  }
  // Keep cached stock visible to readers which do not run Excel's calculation engine.
  for(const p of catalog.values()){
    const t=totals.get(p.pid)||{entrada:0,salida:0},row=stockRows.get(p.code),old=cellValues(row,strings);
    const incoming=t.entrada/1e6,outgoing=t.salida/1e6,total=(Math.round(Number(old.F||0)*1e6)+t.entrada)/1e6;
    for(const [col,value] of [['G',incoming],['H',outgoing],['I',Math.round((total-outgoing)*1e6)/1e6],['J',total]]){
      if(Number(old[col])===value)continue;
      const cell=nodes(row,'c').find(c=>c.getAttribute('r')===col+row.getAttribute('r'));
      check(cell&&first(cell,'f'),'Revisa la fórmula de '+col+row.getAttribute('r')+' en Stock');
      setCell(row,col,value,{numeric:true,formula:true});stockDirty=true;
    }
  }
  if(stockDirty){save('xl/worksheets/sheet1.xml',stock);save('xl/tables/table1.xml',stockTable);dirty=true}
  const operations=syncOperations({parts,read,save,strings,events,catalog,byCode,voided});dirty||=operations.dirty;
  if(dirty){const book=read('xl/workbook.xml');first(book,'calcPr')?.setAttribute('fullCalcOnLoad','1');save('xl/workbook.xml',book)}
  return {bytes:dirty?zipSync(parts,{level:3}):null,changes,seen,productEvents,productsSeen,operations:operations.edits};
}
export const movementSignature=e=>signature(fields(e));
