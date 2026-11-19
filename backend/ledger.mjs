export function requireValue(ok, message, status=400) {
  if (!ok) throw Object.assign(new Error(message),{status});
}
export const uuid = x => typeof x==='string' && /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(x);
export const email = x => typeof x==='string' && x.length<=254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x);
export class Ledger {
  constructor(sql){
    this.sql=sql;
    sql.exec('CREATE TABLE IF NOT EXISTS members(subject TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL)');
    sql.exec('CREATE TABLE IF NOT EXISTS invites(email TEXT PRIMARY KEY,digest TEXT UNIQUE NOT NULL)');
    sql.exec('CREATE TABLE IF NOT EXISTS events(seq INTEGER PRIMARY KEY,id TEXT UNIQUE NOT NULL,subject TEXT NOT NULL,payload TEXT NOT NULL)');
    sql.exec('CREATE TABLE IF NOT EXISTS products(id TEXT PRIMARY KEY,code TEXT NOT NULL COLLATE NOCASE,stock INTEGER NOT NULL,deleted INTEGER NOT NULL)');
    sql.exec('CREATE UNIQUE INDEX IF NOT EXISTS active_codes ON products(code) WHERE deleted=0');
    sql.exec('CREATE TABLE IF NOT EXISTS photos(id TEXT PRIMARY KEY,data TEXT NOT NULL)');
    sql.exec('CREATE TABLE IF NOT EXISTS receipt_expiry(id TEXT PRIMARY KEY,expires INTEGER NOT NULL)');
    sql.exec('CREATE INDEX IF NOT EXISTS expiry_time ON receipt_expiry(expires)');
    sql.exec('CREATE TABLE IF NOT EXISTS staff(email TEXT PRIMARY KEY,cargo TEXT NOT NULL)');
    sql.exec('CREATE TABLE IF NOT EXISTS config(key TEXT PRIMARY KEY,value TEXT NOT NULL)');
  }
  first(query,...args){return [...this.sql.exec(query,...args)][0]}
  append(e,user,historical=false){
    const admin=user.role==='admin', move=['entrada','salida'].includes(e?.type);
    requireValue(admin||move||e?.type==="medicion",'Solo el administrador puede realizar esta operación.',403);
    requireValue(uuid(e?.id)&&uuid(e.data?.pid)&&/^\d{4}-\d{2}-\d{2}$/.test(e.date),'Registro inválido.');
    const previous=this.first('SELECT subject,payload FROM events WHERE id=?',e.id);
    if(previous){requireValue(previous.subject===user.subject,'Identificador ocupado.',409);return JSON.parse(previous.payload)}
    const d=e.data, pid=d.pid, product=this.first('SELECT * FROM products WHERE id=?',pid);
    const text=(key,max)=>{const value=d[key]??'';requireValue(typeof value==='string'&&value.length<=max,'Campo inválido: '+key);return value};
    const data={pid,actor:historical&&typeof d.actor==='string'?d.actor.slice(0,254):user.email};
    if(['create','edit'].includes(e.type)){
      Object.assign(data,{code:text('code',35).trim(),name:text('name',160).trim(),unit:text('unit',32),cat:d.cat,min:d.min??0,photo:d.photo??null});
      if(d.family!==undefined){requireValue(['balanceado','insumo'].includes(d.family),'Grupo inválido.');data.family=d.family}
      if(d.lbPerSack!=null){requireValue(Number.isFinite(d.lbPerSack)&&d.lbPerSack>0,'Peso de saco inválido.');data.lbPerSack=d.lbPerSack}
      requireValue(data.code&&data.name&&data.unit&&['directo','indirecto'].includes(data.cat)&&Number.isFinite(data.min)&&data.min>=0&&(!data.photo||uuid(data.photo)),'Datos del producto inválidos.');
      requireValue(!this.first('SELECT id FROM products WHERE code=? AND id<>? AND deleted=0',data.code,pid),'El código ya existe.',409);
      if(e.type==='create'){requireValue(!product,'Producto existente.',409);this.sql.exec('INSERT INTO products VALUES(?,?,0,0)',pid,data.code)}
      else {requireValue(product&&!product.deleted,'Producto inactivo.',409);this.sql.exec('UPDATE products SET code=? WHERE id=?',data.code,pid)}
    }else if(e.type==='medicion'){
      requireValue(product&&!product.deleted&&Number.isFinite(d.physical)&&d.physical>=0&&Number.isSafeInteger(Math.round(d.physical*1e6)),'Conteo inválido.');
      Object.assign(data,{physical:d.physical,system:historical&&Number.isFinite(d.system)?d.system:product.stock/1e6});
    }else if(e.type==='reiniciarConteo'){
      requireValue(admin,'Solo administrador.',403);
    }else if(e.type==='reparto'){
      const pond=x=>typeof x==='string'&&/^(?:0[1-9]|09B|1[0-9]|20|PRE(?:0[1-9]|1[01]))$/.test(x);
      requireValue(pond(d.origin)&&Array.isArray(d.targets)&&d.targets.length>0&&d.targets.length<32,'Destinos inválidos.');
      requireValue([d.from,d.to].every(v=>typeof v==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(v))&&d.from<=d.to&&d.from<=e.date,'Período inválido.');
      requireValue(d.targets.every(t=>pond(t.dest)&&t.dest!==d.origin&&Number.isFinite(t.pct)&&t.pct>0)&&new Set(d.targets.map(t=>t.dest)).size===d.targets.length&&Math.abs(d.targets.reduce((s,t)=>s+t.pct,0)-100)<0.00001,'Los porcentajes deben sumar 100%.');
      Object.assign(data,{origin:d.origin,from:d.from,to:d.to,targets:d.targets.map(t=>({dest:t.dest,pct:t.pct})),note:text('note',180)});
    }else{
      requireValue(product&&(!product.deleted||e.type==='anular'),'Producto inexistente o dado de baja.',409);
      if(e.type==='anular'){
        requireValue(uuid(d.target),'Salida inválida.');
        const original=this.first('SELECT payload FROM events WHERE id=?',d.target);
        const salida=original&&JSON.parse(original.payload);
        requireValue(['salida','entrada','reparto'].includes(salida?.type)&&salida.data.pid===pid,'Salida inexistente.',409);
        requireValue(!this.first("SELECT id FROM events WHERE json_extract(payload,'$.type')='anular' AND json_extract(payload,'$.data.target')=?",d.target),'La salida ya fue anulada.',409);
        data.target=d.target;
        const restored=product.stock+(salida.type==='salida'?1:salida.type==='entrada'?-1:0)*Math.round((salida.data.q||0)*1e6);
        requireValue((historical||restored>=0)&&Number.isSafeInteger(restored),'Saldo fuera de rango.');
        this.sql.exec('UPDATE products SET stock=? WHERE id=?',restored,pid);
      }else if(e.type==='delete')this.sql.exec('UPDATE products SET deleted=1 WHERE id=?',pid);
      else{
        requireValue(move||['ajuste','traspaso'].includes(e.type),'Operación desconocida.');
        const units=Math.round(d.q*1e6);
        requireValue(Number.isFinite(d.q)&&Number.isSafeInteger(units)&&(e.type==='ajuste'||units>0),'Cantidad inválida.');
        Object.assign(data,{q:units/1e6,doc:text('doc',75),origin:text('origin',80),dest:text('dest',80),note:text('note',180)});
        if(e.type==='entrada'&&d.attachments){
          requireValue(Array.isArray(d.attachments)&&d.attachments.length<=10&&d.attachments.every(id=>uuid(id)&&(historical||this.first('SELECT id FROM photos WHERE id=?',id))),'Fotografía no disponible.');
          data.attachments=d.attachments;
        }
        if(d.group&&e.type==='entrada'){requireValue(uuid(d.group),'Guía inválida.');data.group=d.group}
        if(d.replaces){
          const original=uuid(d.replaces)&&this.first('SELECT payload FROM events WHERE id=?',d.replaces);
          requireValue(admin&&original,'Corrección no autorizada.',403);
          const old=JSON.parse(original.payload);requireValue(old.type===e.type,'Movimiento incompatible.');
          Object.assign(data,{replaces:d.replaces,actor:old.data.actor,editedBy:user.email});
        }
        const delta=e.type==='salida'?-units:e.type==='traspaso'?0:units;
        requireValue(Number.isSafeInteger(product.stock+delta),'Saldo fuera de rango.');
        requireValue(historical||product.stock+delta>=0,'Stock insuficiente. Actualiza el inventario.',409);
        if(e.type==='ajuste'&&!historical)requireValue(Math.round(d.expected*1e6)===product.stock,'El saldo cambió. Repite el conteo.',409);
        this.sql.exec('UPDATE products SET stock=stock+? WHERE id=?',delta,pid);
      }
    }
    const event={id:e.id,date:e.date,type:e.type,ts:historical?new Date(e.ts).toISOString():new Date().toISOString(),data,synced:true};
    event.entryOrder=(this.first('SELECT MAX(seq) n FROM events').n||0)+1;
    this.sql.exec('INSERT INTO events(seq,id,subject,payload) VALUES(?,?,?,?)',event.entryOrder,e.id,user.subject,JSON.stringify(event));return event;
  }
}
