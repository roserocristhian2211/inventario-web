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
 assert.equal((await call('join','other','other@example.test',{code})).status,403);
 assert.equal((await (await call('join','member',address,{code})).json()).role,'operador');
 assert.equal((await call('join','new',address,{code})).status,403);
 assert.equal((await call('invite','member',address,{email:'other@example.test'})).status,403);
 assert.equal((await call('members','member',address)).status,403);
 assert.equal((await call('remove',owner,'owner@example.test',{email:address})).status,200);
 assert.equal((await call('changes','member',address)).status,403);
 assert.equal((await call('me',owner,'owner@example.test')).status,200)
});
