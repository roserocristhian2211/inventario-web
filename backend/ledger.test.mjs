import test from 'node:test';
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {randomUUID} from 'node:crypto';
import {Ledger} from './ledger.mjs';
test('bodega común, permisos, retiros concurrentes y reintentos',async()=>{
 const db=new DatabaseSync(':memory:');
 const sql={exec(query,...args){const q=db.prepare(query);return q.columns().length?q.all(...args):(q.run(...args),[])}};
 const ledger=new Ledger(sql), pid=randomUUID();
 const admin={subject:'owner',email:'admin@example.test',role:'admin'};
 const operator={subject:'op1',email:'persona@example.test',role:'operador'};
 function apply(event,user){db.exec('BEGIN IMMEDIATE');try{const r=ledger.append(event,user);db.exec('COMMIT');return r}catch(e){db.exec('ROLLBACK');throw e}}
 const create=(type,data)=>({id:randomUUID(),date:'2026-11-18',type,data:{pid,...data}});
 const product=create('create',{name:'Balanceado',code:'SC1',unit:'SC',cat:'directo',min:1});
 assert.throws(()=>apply(product,operator),/administrador/);apply(product,admin);
 for(const type of ['delete','edit','ajuste','traspaso'])assert.throws(()=>apply(create(type,{}),operator),/administrador/);
 const incoming=create('entrada',{q:10,actor:'suplantado'});
 assert.equal(apply(incoming,operator).data.actor,operator.email);apply(incoming,operator);
 const results=await Promise.allSettled(Array.from({length:100},(_,i)=>Promise.resolve().then(()=>apply(create('salida',{q:8}),{...operator,subject:'operador-'+i}))));
 assert.equal(results.filter(x=>x.status==='fulfilled').length,1);
 assert.equal(ledger.first('SELECT stock FROM products').stock,2e6);
 assert.throws(()=>apply(create('ajuste',{expected:10,q:-1}),admin),/saldo cambió/);
 assert.equal(ledger.first('SELECT COUNT(*) AS n FROM events').n,3);
});
test('alta por invitación, revocación inmediata y administrador único',async()=>{
 const {Bodega}=await import('./worker.mjs');
 const db=new DatabaseSync(':memory:');
 const sql={exec(query,...args){const q=db.prepare(query);return q.columns().length?q.all(...args):(q.run(...args),[])}};
 const storage={sql,transactionSync(fn){db.exec('BEGIN');try{const r=fn();db.exec('COMMIT');return r}catch(e){db.exec('ROLLBACK');throw e}}};
 const owner='microsoft-owner', address='op@example.test';
 const b=new Bodega({storage},{ADMIN_SUB:owner,ADMIN_EMAIL:'owner@example.test'});
 async function call(path,subject,label,body={}){
   return b.fetch(new Request('https://internal/'+path,{method:'POST',body:JSON.stringify({subject,label,body})}))
 }
 assert.equal((await call('me','intruder','owner@example.test')).status,403);
 const {code}=await (await call('invite',owner,'owner@example.test',{email:address})).json();
 assert.ok(code);
 const expiring='caduca@example.test';
 const issued=await (await call('invite',owner,'owner@example.test',{email:expiring})).json();
 sql.exec('UPDATE invites SET expires=? WHERE email=?',Date.now()-1,expiring);
 assert.equal((await call('join','expired',expiring,{code:issued.code})).status,403);
 assert.equal((await (await call('members',owner,'owner@example.test')).json()).pending.some(x=>x.email===expiring),false);
 assert.equal((await call('join','other','other@example.test',{code})).status,403);
 assert.equal((await (await call('join','member',address,{code})).json()).role,'operador');
 assert.equal((await call('join','new',address,{code})).status,403);
 assert.equal((await call('invite','member',address,{email:'other@example.test'})).status,403);
 assert.equal((await call('members','member',address)).status,403);
 assert.equal((await call('remove',owner,'owner@example.test',{email:address})).status,200);
 assert.equal((await call('changes','member',address)).status,403);
 assert.equal((await call('me',owner,'owner@example.test')).status,200)
});

test('anular salida restituye el saldo una sola vez y requiere administrador',()=>{
 const db=new DatabaseSync(':memory:');
 const ledger=new Ledger({exec(query,...args){const q=db.prepare(query);return q.columns().length?q.all(...args):(q.run(...args),[])}});
 const pid=randomUUID(),admin={subject:'owner',email:'owner@example.test',role:'admin'};
 const event=(type,data)=>({type,id:randomUUID(),date:'2026-11-18',data:{pid,...data}});
 ledger.append(event('create',{name:'Zeolita',code:'Z1',unit:'SC',cat:'directo'}),admin);
 ledger.append(event('entrada',{q:20}),admin);
 const out=ledger.append(event('salida',{q:8,dest:'09B'}),admin);
 const cancel=event('anular',{target:out.id});
 assert.throws(()=>ledger.append(cancel,{...admin,role:'operador'}),/administrador/);
 ledger.append(cancel,admin);ledger.append(cancel,admin);
 assert.throws(()=>ledger.append({...cancel,id:randomUUID()},admin),/ya fue anulada/);
 assert.equal(ledger.first('SELECT stock FROM products').stock,20e6);
});

test('una salida digitada después conserva su posición aunque sea de una fecha anterior',()=>{
 const db=new DatabaseSync(':memory:');
 const ledger=new Ledger({exec(query,...args){const q=db.prepare(query);return q.columns().length?q.all(...args):(q.run(...args),[])}});
 const admin={subject:'owner',role:'admin',email:'owner@example.test'},pid=randomUUID();
 const event=(type,date,data)=>({type,date,id:randomUUID(),data:{pid,...data}});
 ledger.append(event('create','2026-11-01',{code:'A',name:'Producto',unit:'SC',cat:'directo'}),admin);
 ledger.append(event('entrada','2026-11-01',{q:4}),admin);
 const first=ledger.append(event('salida','2026-11-18',{q:1}),admin);
 const second=ledger.append(event('salida','2026-11-17',{q:1}),admin);
 assert.ok(second.entryOrder>first.entryOrder);
 assert.equal(ledger.append(first,admin).entryOrder,first.entryOrder);
});

test('conteo de operador, traspaso porcentual y corrección de entrada conservan los saldos',()=>{
 const db=new DatabaseSync(':memory:'),ledger=new Ledger({exec(s,...args){const q=db.prepare(s);return q.columns().length?q.all(...args):(q.run(...args),[])}});
 const pid=randomUUID(),owner={subject:'owner',email:'owner@example.test',role:'admin'},operator={subject:'member',email:'member@example.test',role:'operador'};
 const event=(type,data)=>({id:randomUUID(),date:'2026-11-18',type,data:{pid,...data}});
 ledger.append(event('create',{code:'TEST',name:'Balanceado',cat:'directo',unit:'SC'}),owner);
 const incoming=ledger.append(event('entrada',{q:10}),operator);
 assert.equal(ledger.append(event('medicion',{physical:8}),operator).data.system,10);
 const rule=event('reparto',{origin:'PRE01',from:'2026-11-01',to:'2026-11-18',targets:[{dest:'01',pct:100}]});
 assert.throws(()=>ledger.append(rule,operator),/administrador/);
 ledger.append(rule,owner);assert.equal(ledger.first('SELECT stock FROM products').stock,10e6);
 assert.throws(()=>ledger.append(event('anular',{target:incoming.id}),operator),/administrador/);
 const edited=ledger.append(event('entrada',{q:7,replaces:incoming.id}),owner);
 ledger.append(event('anular',{target:incoming.id}),owner);
 assert.equal(edited.data.actor,operator.email);assert.equal(edited.data.editedBy,owner.email);
 assert.equal(ledger.first('SELECT stock FROM products').stock,7e6);
});

test('BASE en OneDrive personal confirma el mismo archivo y respeta cambios concurrentes',async t=>{
  const {publish}=await import('./cloud-sync.mjs');
  const file={id:'base',eTag:'v1',parentReference:{id:'folder'}};
  let latest='v1',status=200,writes=0;
  t.mock.method(globalThis,'fetch',async(url,options)=>{
    if(url.endsWith('?$select=id,name,eTag'))return Response.json({id:'base',name:'BASE.xlsm',eTag:latest});
    assert.ok(url.endsWith('/items/base/content'));
    assert.equal(options.headers['If-Match'],'v1');
    assert.deepEqual([...options.body],[1,2]);
    writes++;
    return status===200?Response.json({id:'base',eTag:'v2',size:2}):Response.json({error:{code:'resourceLocked'}},{status});
  });
  const save=()=>publish('drive',file,{},new Uint8Array([1,2]));
  assert.equal(await save(),'v2');
  latest='v3';await assert.rejects(save(),/cambió antes de guardar/);
  assert.equal(writes,1);
  latest='v1';status=409;await assert.rejects(save(),/409 al guardar BASE .resourceLocked/);
  status=412;await assert.rejects(save(),/412/);
});
