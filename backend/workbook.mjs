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
const signature=m=>JSON.stringify([m.type,m.date,m.pid,m.q,m.ref]);
function check(ok,msg){if(!ok)throw Error(msg)}
function fields(e){return {type:e.type,date:e.date,pid:e.data.pid,q:e.data.q,ref:String(e.type==='entrada'?e.data.doc||'':e.data.dest||'')}}
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
  if(!c){c=newNode(doc,'c');c.setAttribute('r',r);row.appendChild(c)}
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
// The initial import used BASE's two existing tables. Bind only that layout.
export function reconcileWorkbook(bytes,events,linked){
  const parts=unzipSync(bytes),read=name=>{check(parts[name],'Falta una parte del libro');return parse(strFromU8(parts[name]))};
  const save=(name,doc)=>{parts[name]=strToU8(xml(doc))};
  const strings=nodes(read('xl/sharedStrings.xml'),'si').map(n=>nodes(n,'t').map(t=>t.textContent).join(''));
  const catalog=new Map(), byId=new Map(), byCode=new Map(), byName=new Map();
  const voided=new Set(events.filter(e=>e.type==="anular").map(e=>e.data.target));
  for(const e of events){
    if(e.type==='create'||e.type==='edit')catalog.set(e.data.pid,e.data);
    if(e.type==='delete')catalog.delete(e.data.pid);
    if(e.type==='entrada'||e.type==='salida')byId.set(e.id,e);
  }
  for(const p of catalog.values()){byCode.set(p.code,p);byName.set(p.name,p)}
  const stock=read('xl/worksheets/sheet1.xml'),stockByCode=new Map(),stockRows=new Map();
  const stockTable=read('xl/tables/table1.xml'),table=stockTable.documentElement;
  const stockRange=table.getAttribute('ref').match(/^B3:J(\d+)$/);
  check(stockRange&&table.getAttribute('name')==='INVENTARIO','Revisa la tabla INVENTARIO de BASE');
  let stockEnd=+stockRange[1],stockDirty=false;
  const changes=[],seen=new Set(),legacy=new Map(),totals=new Map();let dirty=false;
  for(const row of nodes(stock,'row')){
    if(+row.getAttribute('r')<4)continue;
    const v=cellValues(row,strings);if(v.B){stockByCode.set(v.B,v);stockRows.set(v.B,row)}
  }
  for(const p of catalog.values()){
    const v=stockByCode.get(p.code);
    if(v){check(v.C===p.name&&v.E===p.unit,'El catálogo de BASE y la app difiere para '+p.code);continue}
    // Only append products created in the app. A missing historical item is a conflict.
    const original=events.find(e=>e.type==='create'&&e.data.pid===p.pid);
    check(original&&original.data.actor!=='Historial BASE'&&original.data.code===p.code&&!stockByCode.has(original.data.code),'Revisa el producto '+p.code+' en BASE');
    const data=first(stock,'sheetData'),last=nodes(data,'row').at(-1);
    check(+last.getAttribute('r')===stockEnd,'Hay contenido debajo de la tabla INVENTARIO');
    const row=last.cloneNode(true),n=++stockEnd;row.setAttribute('r',n);
    for(const cell of nodes(row,'c'))cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,n));
    data.appendChild(row);setCell(row,'A',n-3,{numeric:true});
    for(const [col,value] of [['B',p.code],['C',p.name],['D',p.code],['E',p.unit]])setCell(row,col,value);
    const opening=events.filter(e=>e.type==='ajuste'&&e.data.pid===p.pid&&e.data.doc==='INVENTARIO INICIAL').reduce((q,e)=>q+e.data.q,0);
    setCell(row,'F',opening,{numeric:true});stockRows.set(p.code,row);
    stockByCode.set(p.code,{B:p.code,C:p.name,E:p.unit,F:opening});stockDirty=true;
    const ref='B3:J'+n;table.setAttribute('ref',ref);first(table,'autoFilter')?.setAttribute('ref',ref);
    first(stock,'dimension')?.setAttribute('ref','A1:J'+n);
  }
  const accumulate=(pid,kind,q)=>{
    const t=totals.get(pid)||{entrada:0,salida:0};t[kind]+=Math.round(q*1e6);totals.set(pid,t);
  };
  for(const e of byId.values())if(e.data.actor==='Historial BASE'&&!linked.has(e.id)){
    const k=signature(fields(e));if(!legacy.has(k))legacy.set(k,[]);legacy.get(k).push(e.id);
  }
  for(const [sheet,kind] of [[2,'entrada'],[3,'salida']]){
    const name=kind==='entrada'?'ENTRADAS':'SALIDAS',tabPath='xl/tables/table'+sheet+'.xml',sheetPath='xl/worksheets/sheet'+sheet+'.xml';
    const tab=read(tabPath),table=tab.documentElement,doc=read(sheetPath),rows=nodes(doc,'row'),range=table.getAttribute('ref').match(/^A3:([FG])(\d+)$/);
    check(range&&table.getAttribute('name')===name,'Cambió la estructura de '+name);
    const headers=nodes(table,'tableColumn'),colNames=headers.map(c=>c.getAttribute('name'));
    check(colNames.slice(0,6).join('|')===[kind==='entrada'?'Nro. Documento':'PSC/PRE','Fecha','Codigo Producto','Descripción','Medida','Cantidad'].join('|'),'Revisa encabezados de '+name);
    const header=rows.find(r=>r.getAttribute('r')==='3'),end=+range[2];let sheetDirty=false;
    if(headers.length===6){
      const columns=first(table,'tableColumns'),col=newNode(tab,'tableColumn');
      col.setAttribute('id','7');col.setAttribute('name','IdMovimiento');columns.appendChild(col);columns.setAttribute('count','7');
      const ref='A3:G'+end;table.setAttribute('ref',ref);first(table,'autoFilter')?.setAttribute('ref',ref);
      setCell(header,'G','IdMovimiento');sheetDirty=true;
    }else check(headers.length===7&&colNames[6]==='IdMovimiento','Columna de control incompatible');
    const empty=[];let dateStyle;
    for(const row of rows){
      const n=+row.getAttribute('r');if(n<=3||n>end)continue;
      const v=cellValues(row,strings);
      if(!v.F||+v.F===0){
        if(v.G&&voided.has(v.G)&&kind==='salida')continue;
        check(!v.G,'Una fila vinculada fue vaciada: '+name+' '+n);
        if(!v.A&&!v.B&&!(kind==='entrada'?v.D:v.C))empty.push(row);
        continue;
      }
      const product=kind==='entrada'?byName.get(v.D):byCode.get(v.C);
      const movement={type:kind,date:iso(v.B),pid:product?.pid,q:Number(v.F),ref:String(v.A||'')};
      check(product&&movement.date&&Number.isFinite(movement.q)&&movement.q>0,'Revisa '+name+' fila '+n);
      dateStyle??=nodes(row,'c').find(c=>c.getAttribute('r')==='B'+n)?.getAttribute('s');
      let id=v.G;
      if(!id){const candidates=legacy.get(signature(movement));id=candidates?.pop()||crypto.randomUUID();setCell(row,'G',id);sheetDirty=true}
      check(uuid(id)&&!seen.has(id),'IdMovimiento repetido o inválido en '+name);
      if(voided.has(id)){
        check(kind==='salida'&&byId.has(id)&&signature(movement)===signature(fields(byId.get(id))),'Salida anulada con cambios en BASE');
        setCell(row,'F',0,{numeric:true});sheetDirty=true;continue;
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
      check(!linked.has(id)&&e.data.actor!=='Historial BASE','Falta una fila histórica o vinculada en '+name);
      const row=empty.shift(),product=catalog.get(e.data.pid);check(row&&product,'Revisa catálogo o amplía la tabla '+name);
      const m=fields(e);accumulate(product.pid,kind,m.q);
      setCell(row,'A',m.ref);setCell(row,'B',serial(m.date),{numeric:true,style:dateStyle});
      setCell(row,'C',product.code,{formula:kind==='entrada'});setCell(row,'D',product.name,{formula:kind==='salida'});
      setCell(row,'E',product.unit,{formula:true});setCell(row,'F',m.q,{numeric:true});setCell(row,'G',id);
      seen.add(id);sheetDirty=true;
    }
    if(sheetDirty){
      for(const row of rows)if(row.hasAttribute('spans'))row.removeAttribute('spans');
      first(doc,'dimension')?.setAttribute('ref','A1:G'+end);
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
  if(dirty){const book=read('xl/workbook.xml');first(book,'calcPr')?.setAttribute('fullCalcOnLoad','1');save('xl/workbook.xml',book)}
  return {bytes:dirty?zipSync(parts,{level:3}):null,changes,seen};
}
export const movementSignature=e=>signature(fields(e));
