import {reconcileWorkbook,movementSignature} from './workbook.mjs';
import {authorizeSync} from './onedrive.mjs';
import {requireValue} from './ledger.mjs';

const period=60000;
async function assertGraph(response,step){
  if(response.ok)return;
  const error=(await response.json().catch(()=>({})))?.error;
  const code=String(error?.innerError?.code||error?.code||'').replace(/[^a-zA-Z0-9_.-]/g,'').slice(0,64);
  const messages={401:'Revisa la autorización de Microsoft.',403:'Revisa el permiso de edición.',404:'No se encontró BASE.xlsm.',412:'La versión cambió. Se volverá a leer BASE.'};
  const locked=response.status===423||/lock/i.test(code);
  const message=locked?'Guarda y cierra BASE en Excel. Se reintentará automáticamente.':messages[response.status]||'Se volverá a intentar automáticamente.';
  requireValue(false,'OneDrive '+response.status+' al '+step+(code?' ('+code+')':'')+'. '+message,response.status);
}
export async function publish(drive,file,headers,bytes){
  const url='https://graph.microsoft.com/v1.0/drives/'+encodeURIComponent(drive)+'/items/'+encodeURIComponent(file.id);
  const current=await fetch(url+'?$select=id,name,eTag',{headers});
  await assertGraph(current,'comprobar la versión');
  const version=await current.json();
  requireValue(version.id===file.id&&version.name==='BASE.xlsm'&&version.eTag===file.eTag,'BASE cambió antes de guardar. Se volverá a leer.',412);
  const saved=await fetch(url+'/content',{method:'PUT',headers:{...headers,'If-Match':file.eTag,'Content-Type':'application/octet-stream'},body:bytes});
  await assertGraph(saved,'guardar BASE');
  const item=await saved.json();
  requireValue(item.id===file.id&&item.eTag&&item.size===bytes.length,'OneDrive no confirmó el archivo completo. Se volverá a comprobar.',409);
  return item.eTag;
}

function applyChange(db,{id,movement:m},admin){
  const old=db.first('SELECT seq,payload FROM events WHERE id=?',id);
  requireValue(m.ref.length<=(m.type==='entrada'?75:80)&&Number.isSafeInteger(Math.round(m.q*1e6))&&Math.round(m.q*1e6)>0,'Cantidad o referencia inválida en BASE.',409);
  requireValue(m.note.length<=180&&m.origin.length<=80,'Observación u origen demasiado largo en BASE.',409);
  const data={pid:m.pid,q:m.q,note:m.note,origin:m.origin,[m.type==='entrada'?'doc':'dest']:m.ref};
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
  const saved={...e,entryOrder:e.entryOrder??old.seq,date:m.date,data:{...e.data,...data}};
  db.sql.exec('UPDATE events SET seq=(SELECT MAX(seq)+1 FROM events), payload=? WHERE id=?',JSON.stringify(saved),id);
  return saved;
}
export async function syncCloud(ctx,env,db,sessionToken,accessCache={}){
  const store=ctx.storage, connection=await store.get('onedrive-connection');
  if(connection)await store.setAlarm(Date.now()+period);
  if(!connection&&!sessionToken)return;

  try{
    if(!sessionToken&&!(accessCache.expires>Date.now()+90000)){
      Object.assign(accessCache,await authorizeSync(store,env,connection));
    }
    const headers={Authorization:'Bearer '+(sessionToken||accessCache.token)};
    const bound=connection||await store.get('onedrive-owner-file');
    const url=bound?'https://graph.microsoft.com/v1.0/drives/'+encodeURIComponent(bound.drive)+'/items/'+encodeURIComponent(bound.item):
      'https://graph.microsoft.com/v1.0/me/drive/root:/Inventario%20Web%20Independiente/BASE.xlsm';
    const meta=await fetch(url,{headers});await assertGraph(meta,'localizar BASE');
    const file=await meta.json(),seq=db.first('SELECT MAX(seq) n FROM events').n||0;
    requireValue(file.name==='BASE.xlsm'&&file.size<=5*1024*1024&&file.eTag&&file['@microsoft.graph.downloadUrl'],'Revisa nombre, versión y tamaño de BASE.xlsm.',409);
    const drive=file.parentReference?.driveId,item=file.id;
    requireValue(drive&&item,'OneDrive no identificó el libro.',409);
    if(!bound)await store.put('onedrive-owner-file',{drive,item});
    const itemUrl='https://graph.microsoft.com/v1.0/drives/'+encodeURIComponent(drive)+'/items/'+encodeURIComponent(item);
    const last=await store.get('onedrive-checkpoint');
    if(!await store.get('excel-deletion-review')&&last?.version===3&&last.drive===drive&&last.item===item&&last.tag===file.eTag&&last.seq===seq){await store.delete('excel-deletion-review');await store.put('onedrive-status','Al día');return 'Al día'}
    const download=await fetch(file['@microsoft.graph.downloadUrl']);await assertGraph(download,'leer BASE');
    const events=[...db.sql.exec('SELECT payload FROM events ORDER BY seq')].map(r=>JSON.parse(r.payload));
    db.sql.exec('CREATE TABLE IF NOT EXISTS workbook_links(id TEXT PRIMARY KEY,signature TEXT NOT NULL)');
    const linked=new Map([...db.sql.exec('SELECT id,signature FROM workbook_links')].map(r=>[r.id,r.signature]));
    db.sql.exec('CREATE TABLE IF NOT EXISTS workbook_products(id TEXT PRIMARY KEY,baseline TEXT NOT NULL)');
    const productLinks=new Map([...db.sql.exec('SELECT id,baseline FROM workbook_products')].map(r=>[r.id,JSON.parse(r.baseline)]));
    const result=reconcileWorkbook(new Uint8Array(await download.arrayBuffer()),events,linked,productLinks);
    let pendingReview=null;
    if(result.cancellations.length){
      requireValue(last?.drive===drive&&last.item===item,'Verifica el libro vinculado antes de importar eliminaciones.',409);
      const owner=await fetch('https://graph.microsoft.com/v1.0/me/drive?$select=id,owner',{headers});
      await assertGraph(owner,'verificar el propietario');
      const account=await owner.json();
      const ownerId=account.owner?.user?.id;
      const editorId=file.lastModifiedBy?.user?.id;
      const ownerEdited=account.id.toLowerCase()===drive.toLowerCase()&&ownerId&&editorId&&ownerId.toLowerCase()===editorId.toLowerCase();
      const review=await store.get('excel-deletion-review');
      if(!ownerEdited||review){
        const indexed=new Map(events.map(e=>[e.id,e]));
        const products=new Map(events.filter(e=>e.type==='create'||e.type==='edit').map(e=>[e.data.pid,e.data]));
        const rows=result.cancellations.map(({data:{target}})=>{
          const e=indexed.get(target),p=products.get(e.data.pid);
          return {id:target,type:e.type,date:e.date,product:p.name,q:e.data.q,unit:p.unit,ref:e.data.doc||e.data.dest||''};
        }).sort((a,b)=>a.id.localeCompare(b.id));
        const key=JSON.stringify([drive,item,file.eTag,rows]);
        if(review?.key!==key||!review.approved){
          pendingReview=review?.key===key?review:{key,rows,token:crypto.randomUUID(),approved:false};
          await store.put('excel-deletion-review',pendingReview);
        }
      }
    }else await store.delete('excel-deletion-review');
    const cancellations=pendingReview?[]:result.cancellations;
    const admin={subject:env.ADMIN_SUB,email:env.ADMIN_EMAIL,role:'admin'};
    function apply(){
      requireValue((db.first('SELECT MAX(seq) n FROM events').n||0)===seq,'La bodega cambió durante la lectura. Reintentando.',409);
      const byId=new Map(events.map(e=>[e.id,e]));
      for(const event of result.productEvents){
        if(event.type==='ajuste')event.data.expected=db.first('SELECT stock FROM products WHERE id=?',event.data.pid).stock/1e6;
        db.append(event,admin);
      }
      // Restore dispatched units before reversing receipts in the same Excel edit.
      for(const event of cancellations.filter(e=>e.targetType==='salida'))db.append(event,admin);
      for(const change of result.changes)byId.set(change.id,applyChange(db,change,admin));
      for(const event of cancellations.filter(e=>e.targetType==='entrada'))db.append(event,admin);
      for(const op of result.operations){
        const before=byId.get(op.id);
        if(!before){db.append(op,admin);continue}
        const revised=db.append({...op,id:crypto.randomUUID()},admin);
        db.sql.exec('DELETE FROM events WHERE id=?',revised.id);
        const data={...revised.data,actor:before.data.actor,editedBy:admin.email};
        if(before.type==='medicion')data.system=before.data.system;
        db.sql.exec('UPDATE events SET seq=?,payload=? WHERE id=?',revised.entryOrder,JSON.stringify({...before,date:op.date,data}),op.id);
      }
      return byId;
    }
    // Validate the merge against stock and uniqueness constraints before writing OneDrive.
    if(result.cancellations.length||result.productEvents.length||result.changes.length||result.operations.length){
      const rollback=new Error('Validación completada');
      try{store.transactionSync(()=>{apply();throw rollback})}catch(error){if(error!==rollback)throw error}
    }
    let tag=file.eTag;
    if(result.bytes)tag=await publish(drive,file,headers,result.bytes);
    else{const latest=await fetch(itemUrl+'?$select=eTag',{headers});await assertGraph(latest,'verificar BASE');requireValue((await latest.json()).eTag===tag,'BASE cambió durante la lectura.',409)}
    const committed=store.transactionSync(()=>{
      const byId=apply();
      for(const [id,fields] of result.productsSeen){
        const baseline=JSON.stringify(fields);
        if(JSON.stringify(productLinks.get(id))!==baseline)db.sql.exec('INSERT OR REPLACE INTO workbook_products VALUES(?,?)',id,baseline);
      }
      for(const id of result.seen){
        const signature=movementSignature(byId.get(id));
        if(linked.get(id)!==signature)db.sql.exec('INSERT OR REPLACE INTO workbook_links VALUES(?,?)',id,signature);
      }
      return db.first('SELECT MAX(seq) n FROM events').n||0;
    });
    await store.put('onedrive-checkpoint',{version:3,drive,item,tag,seq:committed});
    if(pendingReview){
      const message='Ingresos sincronizados. Quedan '+pendingReview.rows.length+' bajas por confirmar.';
      await store.put('onedrive-status',message);
      return message;
    }
    await store.delete('excel-deletion-review');
    const message='Sincronizado · '+new Date().toISOString();await store.put('onedrive-status',message);return message;
  }catch(err){if(err.status===401)accessCache.expires=0;const message=err.message?.slice(0,180)||'Error de sincronización';await store.put('onedrive-status',message);return message}
}
