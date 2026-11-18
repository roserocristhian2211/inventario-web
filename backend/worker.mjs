import {createRemoteJWKSet,jwtVerify} from 'jose';
import {Ledger,requireValue,email,uuid} from './ledger.mjs';
const origin='https://roserocristhian2211.github.io';
const tenant='9188040d-6c67-4c5b-b112-36a304b66dad';
const issuer='https://login.microsoftonline.com/'+tenant+'/v2.0';
const keys=createRemoteJWKSet(new URL('https://login.microsoftonline.com/consumers/discovery/v2.0/keys'));
const cors={'Access-Control-Allow-Origin':origin,'Access-Control-Allow-Headers':'authorization,content-type','Access-Control-Allow-Methods':'POST,OPTIONS','Cache-Control':'no-store'};
const reply=(obj,status=200)=>Response.json(obj,{status,headers:cors});
export default {async fetch(req,env){
  if(req.method==='OPTIONS')return new Response(null,{headers:cors});
  try{
    requireValue(req.method==='POST','Método inválido.',405);
    const path=new URL(req.url).pathname;
    let claims;
    try{({payload:claims}=await jwtVerify((req.headers.get('Authorization')||'').replace(/^Bearer /,''),keys,{issuer,audience:'0b22855f-0817-4124-8eb7-3e42296b5b17',algorithms:['RS256'],requiredClaims:['exp','sub','tid']}))}
    catch{throw Object.assign(new Error('Vuelve a iniciar sesión con Microsoft.'),{status:401})}
    requireValue(claims.tid===tenant,'Cuenta personal requerida.',403);
    const subject=claims.sub, label=String(claims.preferred_username||claims.email||'').trim().toLowerCase();
    if(path==='/identity')return reply({subject,email:label});
    requireValue(env.ADMIN_SUB&&env.ADMIN_EMAIL,'Administración pendiente de activar.',503);
    const raw=await req.text();requireValue(raw.length<1800000,'Solicitud demasiado grande.',413);
    const body=raw?JSON.parse(raw):{};
    const stub=env.BODEGA.get(env.BODEGA.idFromName('inventario-central-v1'));
    return await stub.fetch('https://internal'+path,{method:'POST',body:JSON.stringify({subject,label,body})});
  }catch(e){return reply({error:e.status?e.message:'Solicitud no procesada.'},e.status||400)}
}};
async function digest(value){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(value))),x=>x.toString(16).padStart(2,'0')).join('')}
export class Bodega {
  constructor(ctx,env){this.ctx=ctx;this.env=env;this.db=new Ledger(ctx.storage.sql)}
  async fetch(req){
    try{
      const {subject,label,body}=await req.json(), path=new URL(req.url).pathname;
      const secret=path==='/invite'?crypto.randomUUID()+crypto.randomUUID():String(body.code||'');
      const hash=await digest(secret);
      const result=this.ctx.storage.transactionSync(()=>{
        const {db}=this, sql=db.sql, admin=subject===this.env.ADMIN_SUB;
        let member=db.first('SELECT email FROM members WHERE subject=?',subject);
        if(path==='/join'&&!member&&!admin){
          const invite=db.first('SELECT email FROM invites WHERE digest=?',hash);
          requireValue(invite&&invite.email===label,'Invitación inválida para esta cuenta.',403);
          sql.exec('INSERT INTO members VALUES(?,?)',subject,label);
          sql.exec('DELETE FROM invites WHERE digest=?',hash);member={email:label};
        }
        requireValue(admin||member,'Cuenta sin acceso a esta bodega.',403);
        const user={subject,email:admin?this.env.ADMIN_EMAIL:member.email,role:admin?'admin':'operador'};
        if(path==='/me'||path==='/join')return user;
        if(path==='/changes'){
          const rows=[...sql.exec('SELECT seq,payload FROM events WHERE seq>? ORDER BY seq LIMIT 200',Number(body.after)||0)];
          return {rows:rows.map(r=>({seq:r.seq,event:JSON.parse(r.payload)})),name:db.first("SELECT value FROM config WHERE key='name'")?.value};
        }
        if(path==='/photo'){return db.first('SELECT data FROM photos WHERE id=?',String(body.id))||{}}
        if(path==='/event')return db.append(body,user);
        requireValue(admin,'Acceso exclusivo del administrador.',403);
        if(path==='/name'){requireValue(typeof body.name==='string'&&body.name.length<=60,'Nombre inválido.');sql.exec('INSERT OR REPLACE INTO config VALUES(?,?)','name',body.name);return {}}
        if(path==='/photo-save'){requireValue(uuid(body.id)&&/^data:image\/jpeg;base64,[a-zA-Z0-9+/=]+$/.test(body.data),'Fotografía inválida.');sql.exec('INSERT OR IGNORE INTO photos VALUES(?,?)',body.id,body.data);return {}}
        if(path==='/import'){requireValue(Array.isArray(body.events)&&body.events.length<=100,'Lote inválido.');for(const e of body.events)db.append(e,user,true);return {}}
        if(path==='/members')return {active:[...sql.exec('SELECT email FROM members ORDER BY email')],pending:[...sql.exec('SELECT email FROM invites ORDER BY email')]};
        const address=String(body.email||'').trim().toLowerCase();
        requireValue(email(address)&&address!==this.env.ADMIN_EMAIL.toLowerCase(),'Correo inválido o reservado.');
        if(path==='/invite'){requireValue(!db.first('SELECT email FROM members WHERE email=?',address),'Ya tiene acceso.',409);sql.exec('INSERT OR REPLACE INTO invites VALUES(?,?)',address,hash);return {code:secret}}
        if(path==='/remove'){sql.exec('DELETE FROM invites WHERE email=?',address);sql.exec('DELETE FROM members WHERE email=?',address);return {}}
        throw Object.assign(new Error('Ruta inexistente.'),{status:404});
      });return reply(result);
    }catch(e){return reply({error:e.status?e.message:'No se pudo guardar la operación.'},e.status||400)}
  }
}
