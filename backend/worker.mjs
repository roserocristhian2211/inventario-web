import {createRemoteJWKSet,jwtVerify} from 'jose';
import {syncCloud} from './cloud-sync.mjs';
import {beginConnection,finishConnection} from './onedrive.mjs';
import {Ledger,requireValue,email,uuid} from './ledger.mjs';
const origin='https://roserocristhian2211.github.io';
const tenant='9188040d-6c67-4c5b-b112-36a304b66dad';
const issuer='https://login.microsoftonline.com/'+tenant+'/v2.0';
const keys=createRemoteJWKSet(new URL('https://login.microsoftonline.com/consumers/discovery/v2.0/keys'));
const cors={'Access-Control-Allow-Origin':origin,'Access-Control-Allow-Headers':'authorization,content-type','Access-Control-Allow-Methods':'POST,OPTIONS','Cache-Control':'no-store'};
const reply=(obj,status=200)=>Response.json(obj,{status,headers:cors});
async function requestBody(req){
  const limit=1800000;
  requireValue(Number(req.headers.get('Content-Length')||0)<=limit,'Solicitud demasiado grande.',413);
  if(!req.body)return {};
  const reader=req.body.getReader(),decoder=new TextDecoder();
  let size=0,text='';
  try{
    while(true){
      const {done,value}=await reader.read();if(done)break;
      size+=value.byteLength;
      if(size>limit){await reader.cancel();requireValue(false,'Solicitud demasiado grande.',413)}
      text+=decoder.decode(value,{stream:true});
    }
    text+=decoder.decode();
  }finally{reader.releaseLock()}
  const body=text?JSON.parse(text):{};
  requireValue(body&&typeof body==='object'&&!Array.isArray(body),'Solicitud inválida.');
  return body;
}
export default {async fetch(req,env){
  if(req.method==='OPTIONS')return new Response(null,{headers:cors});
  try{
    if(req.method==='GET' && new URL(req.url).pathname==='/onedrive/callback'){
      const stub=env.BODEGA.get(env.BODEGA.idFromName('inventario-central-v1'));
      return stub.fetch(req);
    }
    requireValue(req.method==='POST','Método inválido.',405);
    const path=new URL(req.url).pathname;
    let claims;
    try{({payload:claims}=await jwtVerify((req.headers.get('Authorization')||'').replace(/^Bearer /,''),keys,{issuer,audience:'0b22855f-0817-4124-8eb7-3e42296b5b17',algorithms:['RS256'],requiredClaims:['exp','sub','tid']}))}
    catch{throw Object.assign(new Error('Vuelve a iniciar sesión con Microsoft.'),{status:401})}
    requireValue(claims.tid===tenant,'Cuenta personal requerida.',403);
    const subject=claims.sub, label=String(claims.preferred_username||claims.email||'').trim().toLowerCase();
    if(path==='/identity')return reply({subject,email:label});
    requireValue(env.ADMIN_SUB&&env.ADMIN_EMAIL,'Administración pendiente de activar.',503);
    const body=await requestBody(req);
    const stub=env.BODEGA.get(env.BODEGA.idFromName('inventario-central-v1'));
    return await stub.fetch('https://internal'+path,{method:'POST',body:JSON.stringify({subject,label,body})});
  }catch(e){return reply({error:e.status?e.message:'Solicitud no procesada.'},e.status||400)}
}};
async function digest(value){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(value))),x=>x.toString(16).padStart(2,'0')).join('')}
export class Bodega {
  constructor(ctx,env){this.ctx=ctx;this.env=env;this.db=new Ledger(ctx.storage.sql);this.accessCache={}}
  synchronize(token){
    if(!this.syncing)this.syncing=syncCloud(this.ctx,this.env,this.db,token,this.accessCache).finally(()=>{this.syncing=null});
    return this.syncing;
  }
  async alarm(){if(!await this.ctx.storage.get('excel-paused'))await this.synchronize()}
  async wakeWorkbook(){
    const store=this.ctx.storage;
    if(await store.get('excel-paused')||!await store.get('onedrive-connection'))return;
    const next=Date.now()+1500,current=await store.getAlarm();
    if(!current||current>next)await store.setAlarm(next);
  }
  async fetch(req){
    try{
      if(req.method==='GET' && new URL(req.url).pathname==='/onedrive/callback'){
        await finishConnection(this.ctx.storage,this.env,new URL(req.url).searchParams,async token=>{
          const {payload}=await jwtVerify(token,keys,{issuer,audience:'0b22855f-0817-4124-8eb7-3e42296b5b17',algorithms:['RS256'],requiredClaims:['exp','sub','nonce']});
          return payload;
        });
        this.accessCache={};
        await this.ctx.storage.setAlarm(Date.now()+1000);
        return new Response('Cuenta de OneDrive conectada. Puedes volver a Inventario Web.',{headers:{'Content-Type':'text/plain;charset=utf-8','Cache-Control':'no-store','Referrer-Policy':'no-referrer'}});
      }
      const {subject,label,body}=await req.json(), path=new URL(req.url).pathname;
      if(path==='/onedrive/pause'){
        requireValue(subject===this.env.ADMIN_SUB&&typeof body.paused==='boolean','Acceso exclusivo del administrador.',403);
        await this.ctx.storage.put('excel-paused',body.paused);
        if(!body.paused)await this.wakeWorkbook();
        return reply({paused:body.paused});
      }
      if(path==='/onedrive/session'){
        requireValue(subject===this.env.ADMIN_SUB,'Acceso exclusivo del administrador.',403);
        requireValue(typeof body.token==='string'&&body.token.length>20,'Vuelve a iniciar sesión con Microsoft.');
        if(await this.ctx.storage.get('excel-paused'))return reply({message:'Sincronización de Excel en pausa.',paused:true});
        const cloud=!!await this.ctx.storage.get('onedrive-connection');
        const message=await this.synchronize(cloud?undefined:body.token);
        const review=await this.ctx.storage.get('excel-deletion-review');
        const latest=this.db.first("SELECT name FROM sqlite_master WHERE type='table' AND name='workbook_links'")&&this.db.first(`
          SELECT e.payload FROM events e JOIN workbook_links w ON w.id=e.id
          WHERE json_extract(e.payload,'$.type')='entrada' AND e.id NOT IN
            (SELECT json_extract(payload,'$.data.target') FROM events WHERE json_extract(payload,'$.type')='anular')
          ORDER BY e.seq DESC LIMIT 1`);
        return reply({message,paused:false,cloud,latestReceipt:latest?JSON.parse(latest.payload):null,review:review&&!review.approved?{token:review.token,rows:review.rows}:null});
      }
      if(path==='/onedrive/approve-deletions'){
        requireValue(subject===this.env.ADMIN_SUB,'Acceso exclusivo del administrador.',403);
        await this.ctx.storage.transaction(async tx=>{
          const review=await tx.get('excel-deletion-review');
          requireValue(review&&review.token===body.token,'La lista cambió. Vuelve a revisar BASE.',409);
          await tx.put('excel-deletion-review',{...review,approved:true});
        });
        await this.wakeWorkbook();
        return reply({approved:true});
      }
      if(path==='/onedrive/status'){
        requireValue(subject===this.env.ADMIN_SUB,'Acceso exclusivo del administrador.',403);
        return reply({message:await this.ctx.storage.get('onedrive-status')||'Pendiente de autorización'});
      }
      if(path==='/onedrive/start'){
        requireValue(subject===this.env.ADMIN_SUB,'Acceso exclusivo del administrador.',403);
        return reply(await beginConnection(this.ctx.storage,this.env));
      }
      const secret=path==='/invite'?crypto.randomUUID()+crypto.randomUUID():String(body.code||'');
      const hash=await digest(secret);
      const result=this.ctx.storage.transactionSync(()=>{
        const {db}=this, sql=db.sql, admin=subject===this.env.ADMIN_SUB;
        let member=db.first('SELECT email FROM members WHERE subject=?',subject);
        if(path==='/join'&&!member&&!admin){
          const invite=db.first('SELECT email FROM invites WHERE digest=? AND expires>?',hash,Date.now());
          requireValue(invite&&invite.email===label,'Invitación inválida para esta cuenta.',403);
          sql.exec('INSERT INTO members VALUES(?,?)',subject,label);
          sql.exec('DELETE FROM invites WHERE digest=?',hash);member={email:label};
        }
        requireValue(admin||member,'Cuenta sin acceso a esta bodega.',403);
        const user={subject,email:admin?this.env.ADMIN_EMAIL:member.email,role:admin?'admin':'operador'};
        if(path==='/me'||path==='/join'){
          const expired=[...sql.exec('SELECT id FROM receipt_expiry WHERE expires<=? LIMIT 100',Date.now())];
          for(const {id} of expired){sql.exec('DELETE FROM photos WHERE id=?',id);sql.exec('DELETE FROM receipt_expiry WHERE id=?',id)}
          return {...user,cargo:db.first('SELECT cargo FROM staff WHERE email=?',user.email)?.cargo||''};
        }
        if(path==='/changes'){
          const rows=[...sql.exec('SELECT seq,payload FROM events WHERE seq>? ORDER BY seq LIMIT 200',Number(body.after)||0)];
          return {anchor:db.first("SELECT value FROM config WHERE key='anchor'")?.value,rows:rows.map(r=>({seq:r.seq,event:{entryOrder:r.seq,...JSON.parse(r.payload)}})),name:db.first("SELECT value FROM config WHERE key='name'")?.value};
        }
        if(path==='/photo'){return db.first('SELECT p.data,x.expires FROM photos p LEFT JOIN receipt_expiry x ON x.id=p.id WHERE p.id=? AND (x.expires IS NULL OR x.expires>?)',String(body.id),Date.now())||{}}
        if(path==='/event')return db.append(body,user);
        if(path==='/guide-replace'){
          requireValue(admin,'Acceso exclusivo del administrador.',403);
          const batch=body.events;
          requireValue(Array.isArray(batch)&&batch.length>=2&&batch.length<=200,'Guía inválida.');
          const entries=batch.filter(e=>e.type==='entrada'),voids=batch.filter(e=>e.type==='anular');
          requireValue(entries.length>0&&entries.length<=100&&voids.length>0&&voids.length<=100&&entries.length+voids.length===batch.length,'Guía inválida.');
          requireValue(voids.every(e=>JSON.parse(db.first('SELECT payload FROM events WHERE id=?',String(e.data.target))?.payload||'{}').type==='entrada'),'Entrada inválida.');
          return {events:[...entries,...voids].map(e=>db.append(e,user))};
        }
        if(path==='/batch'){
          requireValue(Array.isArray(body.events)&&body.events.length>0&&body.events.length<=100,'Lote inválido.');
          return {events:body.events.map(e=>db.append(e,user))};
        }
        if(path==='/attachment'){
          requireValue(uuid(body.id)&&typeof body.data==='string'&&/^data:image\/jpeg;base64,[a-zA-Z0-9+/=]+$/.test(body.data),'Fotografía inválida.');
          const added=sql.exec('INSERT OR IGNORE INTO photos VALUES(?,?)',body.id,body.data).rowsWritten;
          if(added&&db.first("SELECT value FROM config WHERE key='expireNewPhotos'")?.value==='1')sql.exec('INSERT OR IGNORE INTO receipt_expiry VALUES(?,?)',body.id,Date.now()+30*86400000);
          return {expires:db.first('SELECT expires FROM receipt_expiry WHERE id=?',body.id)?.expires};
        }
        requireValue(admin,'Acceso exclusivo del administrador.',403);
        if(path==='/report-anchor'){
          requireValue(/^\d{4}-\d{2}-\d{2}$/.test(body.date)&&Number.isSafeInteger(body.number)&&body.number>0,'Referencia inválida.');
          sql.exec('INSERT OR REPLACE INTO config VALUES(?,?)','anchor',JSON.stringify({date:body.date,number:body.number}));return {};
        }
        if(path==='/name'){requireValue(typeof body.name==='string'&&body.name.length<=60,'Nombre inválido.');sql.exec('INSERT OR REPLACE INTO config VALUES(?,?)','name',body.name);return {}}
        if(path==='/photo-save'){requireValue(uuid(body.id)&&/^data:image\/jpeg;base64,[a-zA-Z0-9+/=]+$/.test(body.data),'Fotografía inválida.');sql.exec('INSERT OR IGNORE INTO photos VALUES(?,?)',body.id,body.data);return {}}
        if(path==='/import'){requireValue(Array.isArray(body.events)&&body.events.length<=100,'Lote inválido.');for(const e of body.events)db.append(e,user,true);return {}}
        if(path==='/retention'){sql.exec("INSERT OR REPLACE INTO config VALUES('expireNewPhotos',?)",body.enabled===true?'1':'0');return {}}
        if(path==='/members')return {expiry:db.first("SELECT value FROM config WHERE key='expireNewPhotos'")?.value==='1',active:[...sql.exec('SELECT m.email,s.cargo FROM members m LEFT JOIN staff s ON s.email=m.email ORDER BY m.email')],pending:[...sql.exec('SELECT i.email,s.cargo FROM invites i LEFT JOIN staff s ON s.email=i.email WHERE i.expires>? ORDER BY i.email',Date.now())]};
        const address=String(body.email||'').trim().toLowerCase();
        requireValue(email(address)&&address!==this.env.ADMIN_EMAIL.toLowerCase(),'Correo inválido o reservado.');
        if(path==='/cargo'){requireValue(typeof body.cargo==='string'&&body.cargo.length<=60,'Cargo inválido.');sql.exec('INSERT OR REPLACE INTO staff VALUES(?,?)',address,body.cargo.trim());return {}}
        if(path==='/invite'){requireValue(!db.first('SELECT email FROM members WHERE email=?',address),'Ya tiene acceso.',409);sql.exec('INSERT OR REPLACE INTO invites(email,digest,expires) VALUES(?,?,?)',address,hash,Date.now()+7*86400000);return {code:secret}}
        if(path==='/remove'){sql.exec('DELETE FROM invites WHERE email=?',address);sql.exec('DELETE FROM members WHERE email=?',address);return {}}
        throw Object.assign(new Error('Ruta inexistente.'),{status:404});
      });
      if(['/batch','/event','/import','/guide-replace'].includes(path))await this.wakeWorkbook();
      return reply(result);
    }catch(e){return reply({error:e.status?e.message:'No se pudo guardar la operación.'},e.status||400)}
  }
}
