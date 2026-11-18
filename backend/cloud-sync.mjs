import {reconcileWorkbook,movementSignature} from './workbook.mjs';
import {authorizeSync} from './onedrive.mjs';
import {requireValue} from './ledger.mjs';

const period=300000;
const assertGraph=r=>requireValue(r.ok,'OneDrive respondió '+r.status+'. Se reintentará sin sobrescribir cambios.',r.status);
async function publish(url,headers,tag,bytes){
  const session=await fetch(url+'/createUploadSession',{method:'POST',headers:{...headers,'Content-Type':'application/json','If-Match':tag},body:JSON.stringify({deferCommit:true})});
  assertGraph(session);
  const {uploadUrl}=await session.json();
  const sent=await fetch(uploadUrl,{method:'PUT',headers:{'Content-Range':'bytes 0-'+(bytes.length-1)+'/'+bytes.length},body:bytes});
  assertGraph(sent);
  requireValue(sent.status===202,'OneDrive no confirmó la carga diferida.',409);
  const commit=await fetch(url,{method:'PUT',headers:{...headers,'If-Match':tag,'Content-Type':'application/json'},body:JSON.stringify({'@microsoft.graph.sourceUrl':uploadUrl,'@microsoft.graph.conflictBehavior':'replace'})});
  assertGraph(commit);
  return (await commit.json()).eTag;
}

function applyChange(db,{id,movement:m},admin){
  const old=db.first('SELECT payload FROM events WHERE id=?',id);
  requireValue(m.ref.length<=(m.type==='entrada'?75:80)&&Number.isSafeInteger(Math.round(m.q*1e6))&&Math.round(m.q*1e6)>0,'Cantidad o referencia inválida en BASE.',409);
  const data={pid:m.pid,q:m.q,[m.type==='entrada'?'doc':'dest']:m.ref};
  if(!old)return db.append({id,type:m.type,date:m.date,data},admin);
  const e=JSON.parse(old.payload);
  requireValue(e.type===m.type,'Cambió el tipo de movimiento.',409);
  const factor=e.type==='salida'?-1:1;
  const deltas=new Map([[e.data.pid,-factor*Math.round(e.data.q*1e6)]]);
  deltas.set(m.pid,(deltas.get(m.pid)||0)+factor*Math.round(m.q*1e6));
  for(const [pid,delta] of deltas){
    const row=db.first('SELECT stock FROM products WHERE id=? AND deleted=0',pid);
    requireValue(row&&Number.isSafeInteger(row.stock+delta)&&row.stock+delta>=0,'Saldo insuficiente para corregir la fila de BASE.',409);
    db.sql.exec('UPDATE products SET stock=stock+? WHERE id=?',delta,pid);
  }
  const saved={...e,date:m.date,data:{...e.data,...data}};
  db.sql.exec('UPDATE events SET seq=(SELECT MAX(seq)+1 FROM events), payload=? WHERE id=?',JSON.stringify(saved),id);
  return saved;
}
export async function syncCloud(ctx,env,db){
  const store=ctx.storage, connection=await store.get('onedrive-connection');
  if(!connection)return;
  await store.setAlarm(Date.now()+period);
  try{
    const token=await authorizeSync(store,env,connection),headers={Authorization:'Bearer '+token};
    const url='https://graph.microsoft.com/v1.0/drives/'+encodeURIComponent(connection.drive)+'/items/'+encodeURIComponent(connection.item);
    const meta=await fetch(url,{headers});assertGraph(meta);
    const file=await meta.json(),seq=db.first('SELECT MAX(seq) n FROM events').n||0;
    requireValue(file.name==='BASE.xlsm'&&file.size<=5*1024*1024&&file.eTag&&file['@microsoft.graph.downloadUrl'],'Revisa nombre, versión y tamaño de BASE.xlsm.',409);
    const last=await store.get('onedrive-checkpoint');
    if(last?.tag===file.eTag&&last.seq===seq){await store.put('onedrive-status','Al día');return}
    const download=await fetch(file['@microsoft.graph.downloadUrl']);assertGraph(download);
    const events=[...db.sql.exec('SELECT payload FROM events ORDER BY seq')].map(r=>JSON.parse(r.payload));
    db.sql.exec('CREATE TABLE IF NOT EXISTS workbook_links(id TEXT PRIMARY KEY,signature TEXT NOT NULL)');
    const linked=new Map([...db.sql.exec('SELECT id,signature FROM workbook_links')].map(r=>[r.id,r.signature]));
    const result=reconcileWorkbook(new Uint8Array(await download.arrayBuffer()),events,linked);
    let tag=file.eTag;
    if(result.bytes)tag=await publish(url,headers,tag,result.bytes);
    else{const latest=await fetch(url+'?$select=eTag',{headers});assertGraph(latest);requireValue((await latest.json()).eTag===tag,'BASE cambió durante la lectura.',409)}
    const admin={subject:env.ADMIN_SUB,email:env.ADMIN_EMAIL,role:'admin'};
    const committed=store.transactionSync(()=>{
      requireValue((db.first('SELECT MAX(seq) n FROM events').n||0)===seq,'La bodega cambió durante la lectura. Reintentando.',409);
      const byId=new Map(events.map(e=>[e.id,e]));
      for(const change of result.changes)byId.set(change.id,applyChange(db,change,admin));
      for(const id of result.seen){
        const signature=movementSignature(byId.get(id));
        if(linked.get(id)!==signature)db.sql.exec('INSERT OR REPLACE INTO workbook_links VALUES(?,?)',id,signature);
      }
      return db.first('SELECT MAX(seq) n FROM events').n||0;
    });
    await store.put('onedrive-checkpoint',{tag,seq:committed});
    await store.put('onedrive-status','Sincronizado · '+new Date().toISOString());
  }catch(err){await store.put('onedrive-status',err.message?.slice(0,180)||'Error de sincronización')}
}
