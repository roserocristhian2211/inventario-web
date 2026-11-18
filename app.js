
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const DB='inventario-web-independiente-v1', GRAPH='https://graph.microsoft.com/v1.0',MARK='identidad-inventario-web.json';
const original='Inventario Gambirasi';
let db,events=[],photos=[],settings={},inventory=[],moves=[],alerts=[],page='inicio',editor=null,session=null,busy=false;
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','&quot;':'&quot;',"'":'&#39;'}[c]));
const uid=()=>crypto.randomUUID(), today=()=>{let d=new Date();return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-')};
const fmt=n=>new Intl.NumberFormat('es-EC',{maximumFractionDigits:3}).format(n||0);
const number=v=>{let n=Number(v);if(!Number.isFinite(n))throw Error('Cantidad inválida');return n};
function toast(t,bad=false){let el=$('#toast');el.textContent=t;el.className='show'+(bad?' error':'');clearTimeout(toast.id);toast.id=setTimeout(()=>el.className='',5000)}
function fail(e){console.error(e);toast(e.message||String(e),true)}
function req(r){return new Promise((resolve,reject)=>{r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
async function openDB(name=DB){return new Promise((resolve,reject)=>{const r=indexedDB.open(name,1);r.onupgradeneeded=()=>{let x=r.result;for(const n of ['events','photos','settings'])if(!x.objectStoreNames.contains(n))x.createObjectStore(n,{keyPath:'id'})};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)})}
async function store(n,mode,fn){const t=db.transaction(n,mode);let r=await fn(t.objectStore(n));await new Promise((yes,no)=>{t.oncomplete=yes;t.onerror=()=>no(t.error);t.onabort=()=>no(t.error)});return r}
const all=n=>store(n,'readonly',s=>req(s.getAll())),one=(n,id)=>store(n,'readonly',s=>req(s.get(id))),put=(n,x)=>store(n,'readwrite',s=>req(s.put(x)));
async function setting(k,v){settings[k]=v;await put('settings',{id:k,value:v})}
function dateLabel(s){return String(s||'').slice(0,10)}
function record(type,data){data={...data,actor:session.profile.label};return {id:uid(),type,date:today(),ts:new Date().toISOString(),data,synced:false}}
async function saveEvent(type,data){await put('events',record(type,data));await reload()}
function compute(){let p=new Map();moves=[];alerts=[];let sorted=[...events].sort((a,b)=>Number(b.type==='create')-Number(a.type==='create')||String(a.ts).localeCompare(String(b.ts))||String(a.id).localeCompare(String(b.id)));
for(const e of sorted){let d=e.data||{},x=p.get(d.pid);
if(e.type==='create'){if(p.has(d.pid)){alerts.push('Producto duplicado');continue}p.set(d.pid,{id:d.pid,code:d.code,name:d.name,unit:d.unit,cat:d.cat,min:d.min||0,photo:d.photo||null,stock:0,deleted:false});continue}
if(!x){alerts.push('Movimiento sin producto: '+e.id);continue}
if(e.type==='edit'){Object.assign(x,{code:d.code,name:d.name,unit:d.unit,cat:d.cat,min:d.min||0,photo:d.photo||null});continue}
if(e.type==='delete'){x.deleted=true;continue}
if(['entrada','salida','ajuste','traspaso'].includes(e.type)){let q=Number(d.q);if(!Number.isFinite(q)||(e.type!=='ajuste'&&q<=0)){alerts.push('Cantidad incorrecta: '+e.id);continue}if(e.type==='entrada')x.stock+=q;if(e.type==='salida')x.stock-=q;if(e.type==='ajuste')x.stock+=q;moves.push({...e,prod:x.name,code:x.code,unit:x.unit,q})}
}
inventory=[...p.values()].sort((a,b)=>a.name.localeCompare(b.name,'es'));moves.reverse();
let codes=new Set();for(const x of inventory){x.stock=Math.round(x.stock*100000)/100000;if(x.stock<0)alerts.push('Saldo negativo en '+x.code);if(!x.deleted&&x.cat==='directo'){if(codes.has(x.code.toLowerCase()))alerts.push('Código repetido '+x.code);codes.add(x.code.toLowerCase())}}}
async function reload(){events=await all('events');photos=await all('photos');compute();render()}
function active(){return inventory.filter(p=>!p.deleted&&p.cat==='directo')}
const pic=id=>photos.find(x=>x.id===id)?.data||'';
function productRow(p,actions){let state=p.stock<=0?'<span class="tag danger">Agotado</span>':p.min>0&&p.stock<p.min?'<span class="tag warning">Bajo mínimo</span>':'<span class="tag">Disponible</span>';let buttons=actions?'<div class="item-buttons"><button data-edit="'+esc(p.id)+'">Editar</button><button data-move="'+esc(p.id)+'">Movimiento</button><button class="danger" data-delete="'+esc(p.id)+'">Eliminar</button></div>':'';return '<div class="item">'+(p.photo?'<img class="thumb" src="'+esc(pic(p.photo))+'" alt="Foto del producto">':'')+'<div class="item-body"><h3>'+esc(p.name)+'</h3><p>'+esc(p.code)+' · '+esc(p.unit)+' · '+'Material directo'+'</p>'+buttons+'</div><div class="quantity"><strong>'+fmt(p.stock)+'</strong><small>'+esc(p.unit)+'</small>'+state+'</div></div>'}
function empty(t){return '<div class="empty">'+esc(t)+'</div>'}
function moveRow(m){let q=m.type==='salida'?'-':m.type==='entrada'?'+':m.type==='ajuste'&&m.q>=0?'+':'';return '<div class="item"><div class="item-body"><h3>'+esc(m.prod)+'</h3><p>'+esc(m.date)+(m.data.actor?' · '+esc(m.data.actor):'')+' · '+esc(m.data.doc||'Sin documento')+' · '+esc(m.data.origin||'')+(m.data.dest?' → '+esc(m.data.dest):'')+'</p><span class="tag">'+esc(m.type.toUpperCase())+'</span></div><div class="quantity"><strong>'+q+fmt(m.q)+'</strong><small>'+esc(m.unit)+'</small></div></div>'}
function render(){if(!session?.profile)return;let pp=active(),todayN=moves.filter(m=>m.date===today()&&inventory.find(x=>x.id===m.data.pid)?.cat==='directo').length;
$('#n-products').textContent=pp.length;$('#n-out').textContent=pp.filter(p=>p.stock<=0).length;$('#n-low').textContent=pp.filter(p=>p.min>0&&p.stock>0&&p.stock<p.min).length;$('#n-today').textContent=todayN;
$('#attention').innerHTML=pp.filter(p=>p.stock<=0||(p.min>0&&p.stock<p.min)).map(p=>productRow(p,false)).join('')||empty('No hay productos con alertas.');
let term=$('#find-product').value.toLowerCase();
$('#products').innerHTML=active().filter(p=>p.name.toLowerCase().includes(term)||p.code.toLowerCase().includes(term)).map(p=>productRow(p,true)).join('')||empty('No hay productos. Puedes agregar uno nuevo.');
let query=$('#find-move').value.toLowerCase(),type=$('#filter-type').value;
$('#moves').innerHTML=moves.filter(m=>inventory.find(p=>p.id===m.data.pid)?.cat==='directo'&&(type==='todos'||type===m.type)&&[m.prod,m.code,m.data.doc,m.data.origin,m.data.dest,m.date].some(v=>String(v||'').toLowerCase().includes(query))).slice(0,600).map(moveRow).join('')||empty('No hay movimientos registrados.');
$('#counts').innerHTML=active().map(p=>'<div class="item"><div class="item-body"><h3>'+esc(p.name)+'</h3><p>'+esc(p.code)+' · '+esc(p.unit)+'</p></div><div class="quantity"><strong>'+fmt(p.stock)+'</strong><button class="small" data-count="'+esc(p.id)+'">Contar</button></div></div>').join('')||empty('Registra productos para hacer un conteo.');
$('#appname').textContent=settings.name||'Inventario Web';
const pending=events.filter(e=>!e.synced).length;
$('#network').textContent=!navigator.onLine?'Sin conexión':busy?'Sincronizando…':pending?'Cambios pendientes':settings.lastSync?'Sincronizado':'OneDrive conectado';
$('#account-label').textContent=$('#cloud-account').textContent=session.profile.label;
$('#activity').innerHTML=moves.filter(m=>['entrada','salida'].includes(m.type)&&inventory.some(p=>p.id===m.data.pid&&p.cat==='directo')).slice(0,5).map(moveRow).join('')||empty('Aún no hay entradas ni salidas.');
$('#account-exit').disabled=busy;
$('#notice-status').textContent=settings.deviceNotices&&globalThis.Notification?.permission==='granted'?'Avisos del dispositivo activados.':'Los avisos aparecen dentro de la aplicación.';
$('#device-notices').textContent=settings.deviceNotices?'Desactivar notificaciones del dispositivo':'Activar notificaciones del dispositivo';
$('#sync-now').disabled=busy||!navigator.onLine||!session;$('#sync-top').disabled=busy||!navigator.onLine||!session;
$('#alert').classList.toggle('hidden',!alerts.length);$('#alert').textContent=alerts.slice(0,5).join(' · ');
$('#cloud-msg').textContent=events.filter(e=>!e.synced).length+' operaciones sin publicar en OneDrive. Última sincronización: '+(settings.lastSync||'nunca');
}
function show(id){page=id;$$('.page').forEach(el=>el.classList.toggle('hidden',el.id!=='p-'+id));$$('.navitem').forEach(el=>el.classList.toggle('on',el.dataset.page===id));render()}
const field=(lab,name,val='',type='text',extra='')=>'<label>'+lab+'<input name="'+name+'" type="'+type+'" value="'+esc(val)+'" '+extra+'></label>';
const select=(lab,name,opts,selected)=>'<label>'+lab+'<select name="'+name+'">'+opts.map(([v,t])=>'<option value="'+esc(v)+'" '+(v===selected?'selected':'')+'>'+esc(t)+'</option>').join('')+'</select></label>';
function openEditor(kind,pid=null,initial='entrada'){let p=inventory.find(x=>x.id===pid);editor={kind,pid};let h='';
if(kind==='producto'){$('#modal-title').textContent=p?'Editar producto':'Nuevo producto';h+=field('Código','code',p?.code||'','text','required maxlength="35"')+field('Descripción','name',p?.name||'','text','required maxlength="160"');h+=select('Unidad','unit',['SC','KG','LT','UND','CJ','PZ','GL','M'].map(x=>[x,x]),p?.unit||'UND');h+=field('Existencia mínima','min',p?.min??0,'number','required min="0" step="any"');h+='<label>Fotografía (opcional)<input type="file" name="picture" accept="image/jpeg,image/png,image/webp"></label>'}
if(kind==='movimiento'){if(!active().length){toast('Primero agrega un producto.',true);return}$('#modal-title').textContent='Registrar movimiento';h+=select('Tipo de movimiento','type',[['entrada','Entrada de mercadería'],['salida','Salida de mercadería'],['traspaso','Traspaso entre piscinas']],initial);h+=select('Producto','pid',active().map(x=>[x.id,x.code+' · '+x.name+' ('+fmt(x.stock)+' '+x.unit+')']),pid||active()[0].id);h+=field('Cantidad','q','','number','required min="0.000001" step="any"')+field('Fecha','date',today(),'date','required')+field('Documento/guía','doc','','text','maxlength="75"')+field('Origen','origin','','text','maxlength="80"')+field('Destino/piscina','dest','','text','maxlength="80"')+field('Observaciones','note','','text','maxlength="180"');h+='<p class="subtle">Un traspaso entre piscinas no resta mercadería de bodega. Para registrar consumo usa Salida.</p>'}
if(kind==='conteo'){if(!p)return;$('#modal-title').textContent='Conteo físico';h+='<div class="note">'+esc(p.name)+' · Existencia: '+fmt(p.stock)+' '+esc(p.unit)+'</div>'+field('Cantidad encontrada','q',p.stock,'number','required min="0" step="any"')+field('Observación','note','Conteo físico','text','maxlength="180"')}
$('#fields').innerHTML=h;$('#modal').showModal()}
async function imageData(file){if(file.size>10*1024*1024)throw Error('Imagen demasiado grande (máximo 10 MB).');if(!['image/png','image/jpeg','image/webp'].includes(file.type))throw Error('Formato de imagen inválido.');let bitmap=await createImageBitmap(file);let scale=Math.min(1,1200/Math.max(bitmap.width,bitmap.height));let can=document.createElement('canvas');can.width=Math.round(bitmap.width*scale);can.height=Math.round(bitmap.height*scale);can.getContext('2d').drawImage(bitmap,0,0,can.width,can.height);bitmap.close();return can.toDataURL('image/jpeg',.75)}
async function submitEditor(e){e.preventDefault();try{const f=new FormData(e.target),v=k=>String(f.get(k)||'').trim(),kind=editor.kind,pid=editor.pid;
if(kind==='producto'){let code=v('code'),name=v('name'),min=number(v('min'));if(!code||!name||min<0)throw Error('Completa el producto y el mínimo válido.');if(active().some(p=>p.id!==pid&&p.code.toLowerCase()===code.toLowerCase()))throw Error('Este código ya existe.');let prev=inventory.find(p=>p.id===pid),photo=prev?.photo||null,attached=f.get('picture');if(attached instanceof File&&attached.size){photo=uid();await put('photos',{id:photo,data:await imageData(attached),synced:false})}
await saveEvent(pid?'edit':'create',{pid:pid||uid(),code,name,min,cat:'directo',unit:v('unit'),photo})}
if(kind==='movimiento'){let x=active().find(p=>p.id===v('pid')),type=v('type'),q=number(v('q'));if(!x||q<=0)throw Error('Selecciona producto y cantidad válida.');if(type==='salida'&&q>x.stock)throw Error('Stock insuficiente para esta salida.');if(type==='traspaso'&&(!v('origin')||!v('dest')||v('origin')===v('dest')))throw Error('Indica piscinas de origen y destino distintas.');let ev=record(type,{pid:x.id,q,doc:v('doc'),origin:v('origin'),dest:v('dest'),note:v('note')});ev.date=v('date')||today();await put('events',ev);await reload();if(['entrada','salida'].includes(type))notifyMove(ev,x)}
if(kind==='conteo'){let p=active().find(x=>x.id===pid),q=number(v('q'));if(!p||q<0)throw Error('La cantidad no puede ser negativa.');await saveEvent('ajuste',{pid,q:q-p.stock,note:v('note'),doc:'CONTEO'})}
$('#modal').close();if(kind!=='movimiento'||v('type')==='traspaso')toast('Registro guardado en este dispositivo.');scheduleSync()}catch(err){fail(err)}}
async function exportJSON(){let json=JSON.stringify({format:'inventario-web-independent-v1',fecha:new Date().toISOString(),events,photos},null,2);download('respaldo-inventario-'+today()+'.json',json,'application/json')}
function download(name,body,type){const a=document.createElement('a'),url=URL.createObjectURL(new Blob([body],{type}));a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),3000)}
function csvEscape(x){let s=String(x??'');return /[;"\n\r]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}
function exportCSV(){download('productos-'+today()+'.csv','\uFEFF'+[['codigo','descripcion','unidad','categoria','minimo','stock'],...active().map(x=>[x.code,x.name,x.unit,x.cat,x.min,x.stock])].map(r=>r.map(csvEscape).join(';')).join('\r\n'),'text/csv')}
function csvParse(txt){let t=txt.replace(/^\uFEFF/,''),sep=t.split(/\r?\n/,1)[0].includes(';')?';':',';let rows=[],row=[],cell='',quote=false;for(let i=0;i<t.length;i++){let ch=t[i];if(ch==='"'){if(quote&&t[i+1]==='"'){cell+='"';i++}else quote=!quote}else if(ch===sep&&!quote){row.push(cell);cell=''}else if((ch==='\n'||ch==='\r')&&!quote){if(ch==='\r'&&t[i+1]==='\n')i++;row.push(cell);if(row.some(v=>v.trim()))rows.push(row);row=[];cell=''}else cell+=ch}if(quote)throw Error('CSV inválido: comillas sin cerrar');row.push(cell);if(row.some(v=>v.trim()))rows.push(row);return rows}
async function importCSV(file){let rows=csvParse(await file.text());let first=rows.shift()?.map(x=>x.toLowerCase().trim())||[];let get=(r,key)=>r[first.indexOf(key)];if(!first.includes('codigo')||!first.includes('descripcion'))throw Error('El CSV debe tener columnas codigo y descripcion.');let codes=new Set(active().map(x=>x.code.toLowerCase())),count=0;for(const r of rows){let code=(get(r,'codigo')||'').trim(),name=(get(r,'descripcion')||'').trim(),cat=(get(r,'categoria')||'').trim().toLowerCase();if((cat&&cat!=='directo'&&cat!=='material directo')||!code||!name||codes.has(code.toLowerCase()))continue;let min=Number(get(r,'minimo')||0);if(!Number.isFinite(min)||min<0)min=0;await put('events',record('create',{pid:uid(),code,name,unit:get(r,'unidad')||'UND',cat:'directo',min,photo:null}));codes.add(code.toLowerCase());count++}await reload();toast(count+' productos importados.')}
const SESSION_KEY='iw-oauth-session', CLIENT='0b22855f-0817-4124-8eb7-3e42296b5b17';
function loadSession(){
  try{return JSON.parse(sessionStorage.getItem(SESSION_KEY)||localStorage.getItem(SESSION_KEY)||'null')}
  catch{return null}
}
function setSession(value){
  session=value;
  if(value){(value.remember?localStorage:sessionStorage).setItem(SESSION_KEY,JSON.stringify(value))}
  else {sessionStorage.removeItem(SESSION_KEY);localStorage.removeItem(SESSION_KEY)}
}
function signOut(){
  if(busy)return;
  if(events.some(e=>!e.synced)&&!confirm('Hay cambios pendientes. Se conservarán para cuando vuelvas con esta cuenta. ¿Cerrar sesión?'))return;
  setSession(null);localStorage.setItem('iw-signed-out',uid());location.replace(redirect());
}
function notifyMove(event,product){
  const title=event.type==='entrada'?'Entrada registrada':'Salida registrada';
  const detail=fmt(event.data.q)+' '+product.unit+' · '+product.name;
  toast(title+' · '+detail);
  if(settings.deviceNotices&&globalThis.Notification?.permission==='granted'&&'serviceWorker' in navigator){
    navigator.serviceWorker.ready.then(reg=>reg.showNotification(title,{body:detail,tag:event.id,icon:'./icon.svg'})).catch(console.warn)
  }
}
let syncTimer;
function scheduleSync(){
  clearTimeout(syncTimer);
  if(session?.profile&&navigator.onLine)syncTimer=setTimeout(()=>{if(!busy)sync().catch(fail)},1200)
}
const redirect=()=>location.origin+location.pathname;
function random(){let b=crypto.getRandomValues(new Uint8Array(32));return btoa(String.fromCharCode(...b)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'')}
async function connect(){const client=CLIENT,tenant='consumers';let verifier=random(),state=random(),digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier));let challenge=btoa(String.fromCharCode(...new Uint8Array(digest))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');sessionStorage.setItem('iw-oauth-pending',JSON.stringify({client,tenant,verifier,state,remember:$('#remember').checked,redirect:redirect(),created:Date.now()}));let q=new URLSearchParams({client_id:client,response_type:'code',redirect_uri:redirect(),response_mode:'query',scope:'openid profile offline_access User.Read Files.ReadWrite',state,code_challenge:challenge,code_challenge_method:'S256',prompt:'select_account'});location.assign('https://login.microsoftonline.com/'+tenant+'/oauth2/v2.0/authorize?'+q)}
async function exchange(client,tenant,body){let r=await fetch('https://login.microsoftonline.com/'+tenant+'/oauth2/v2.0/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({client_id:client,scope:'openid profile offline_access User.Read Files.ReadWrite',...body})});let d=await r.json().catch(()=>({}));if(!r.ok){const error=Error(d.error_description||d.error||'Fallo de autenticación');error.code=d.error;throw error}return d}
async function callback(){let p=new URLSearchParams(location.search);if(!p.has('code')&&!p.has('error'))return;const pending=JSON.parse(sessionStorage.getItem('iw-oauth-pending')||'null');let err=p.get('error_description')||p.get('error'),code=p.get('code'),state=p.get('state');history.replaceState({},'',location.pathname);if(err)throw Error(err);if(!pending||pending.state!==state||pending.redirect!==redirect()||Date.now()-pending.created>600000)throw Error('Código OAuth caducado o inválido.');sessionStorage.removeItem('iw-oauth-pending');let x=await exchange(pending.client,pending.tenant,{grant_type:'authorization_code',code,redirect_uri:redirect(),code_verifier:pending.verifier});sessionStorage.removeItem(SESSION_KEY);localStorage.removeItem(SESSION_KEY);setSession({...x,remember:!!pending.remember,client:pending.client,tenant:pending.tenant,expires:Date.now()+x.expires_in*1000})}
async function token(){if(!session)throw Error('Conecta la cuenta OneDrive.');if(session.expires-Date.now()>90000)return session.access_token;try{let v=await exchange(session.client,session.tenant,{grant_type:'refresh_token',refresh_token:session.refresh_token});setSession({...session,...v,expires:Date.now()+v.expires_in*1000});return session.access_token}catch(e){if(e.code==='invalid_grant'){setSession(null);location.replace(redirect())}throw e}}
async function graph(path,opts={}){let auth=await token();let r=await fetch(GRAPH+path,{...opts,headers:{Authorization:'Bearer '+auth,...opts.headers}});if(!r.ok){let msg=await r.text().catch(()=>'');let e=Error('OneDrive HTTP '+r.status+' '+msg.slice(0,180));e.status=r.status;throw e}return r}
async function j(path,opts){return (await graph(path,opts)).json()}
const enc=encodeURIComponent;
async function children(id){let out=[],next='/me/drive/items/'+enc(id)+'/children?$top=200';while(next){let p=await j(next.startsWith('https://graph.microsoft.com/v1.0')?next.slice(GRAPH.length):next);out.push(...(p.value||[]));next=p['@odata.nextLink']||''}return out}
async function up(id,name,body,contentType){return j('/me/drive/items/'+enc(id)+':/'+enc(name)+':/content',{method:'PUT',headers:{'Content-Type':contentType},body})}
async function childFolder(id,name){let existing=(await children(id)).find(x=>x.name===name);if(existing){if(!existing.folder)throw Error('Existe un archivo con el nombre '+name);return existing}return j('/me/drive/items/'+enc(id)+'/children',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,folder:{},'@microsoft.graph.conflictBehavior':'fail'})})}
function folderName(){let x=$('#folder').value.trim();if(x.length<8||x.length>95||/[\\/:*?"<>|]/.test(x)||/gambirasi/i.test(x))throw Error('Usa un nombre de carpeta NUEVO, de 8 a 95 caracteres, diferente a Gambirasi.');return x}
async function sync(){if(busy)return;if(!session||!navigator.onLine)throw Error('Conecta OneDrive y verifica internet.');busy=true;render();try{let folder=folderName(),drive=await j('/me/drive?$select=id,driveType,webUrl'),binding=settings.binding;if(binding&&(binding.driveId!==drive.id||binding.folder!==folder))throw Error('Este inventario local ya está vinculado a otro OneDrive. No se modificarán los datos.');let root;try{root=await j('/me/drive/root:/'+enc(folder))}catch(e){if(e.status!==404)throw e}if(!root)root=await j('/me/drive/root/children',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:folder,folder:{},'@microsoft.graph.conflictBehavior':'fail'})});if(!root.folder)throw Error('La ruta no es una carpeta.');let items=await children(root.id),marker=items.find(x=>x.name===MARK);if(!marker){if(items.length)throw Error('Carpeta existente sin identidad de esta app: elige otra carpeta vacía.');await up(root.id,MARK,JSON.stringify({app:'inventario-web-independent-v1',created:new Date().toISOString()}),'application/json')}else{let d=await (await graph('/me/drive/items/'+enc(marker.id)+'/content')).json();if(d.app!=='inventario-web-independent-v1')throw Error('Esta carpeta pertenece a otra aplicación.')}
let evFolder=await childFolder(root.id,'eventos'),phFolder=await childFolder(root.id,'fotos');
await setting('binding',{driveId:drive.id,folder});
let remote=await children(evFolder.id),ids=new Set();let received=0,sent=0;const newMoves=[];for(const f of remote){if(!/^[0-9a-f-]{36}\.json$/i.test(f.name))continue;let id=f.name.slice(0,-5);ids.add(id);let x=await one('events',id);if(x){if(!x.synced)await put('events',{...x,synced:true});continue}let content=await (await graph('/me/drive/items/'+enc(f.id)+'/content')).json();if(!content.event||content.event.id!==id)throw Error('Movimiento remoto inválido');await put('events',{...content.event,synced:true});received++;if(['entrada','salida'].includes(content.event.type))newMoves.push(content.event)}
let local=await all('events'),remotePhotos=await children(phFolder.id),names=new Set(remotePhotos.map(x=>x.name));
for(const e of local){if(ids.has(e.id))continue;await up(evFolder.id,e.id+'.json',JSON.stringify({event:{id:e.id,type:e.type,ts:e.ts,date:e.date,data:e.data}}),'application/json');await put('events',{...e,synced:true});sent++}
const references=new Set(local.filter(x=>['create','edit'].includes(x.type)).map(x=>x.data.photo).filter(Boolean));let localPhotos=await all('photos');for(const p of localPhotos){if(!references.has(p.id))continue;if(names.has(p.id+'.jpg'))continue;let res=await fetch(p.data),blob=await res.blob();await up(phFolder.id,p.id+'.jpg',blob,'image/jpeg');await put('photos',{...p,synced:true})}
for(const p of remotePhotos){if(!/^[0-9a-f-]{36}\.jpg$/i.test(p.name))continue;let id=p.name.slice(0,-4);if(!references.has(id)||await one('photos',id))continue;let resp=await graph('/me/drive/items/'+enc(p.id)+'/content'),blob=await resp.blob(),data=await new Promise((yes,no)=>{let r=new FileReader();r.onload=()=>yes(r.result);r.onerror=()=>no(r.error);r.readAsDataURL(blob)});await put('photos',{id,data,synced:true})}
const previousSync=settings.lastSync;await setting('lastSync',new Date().toISOString());await reload();toast('Sincronización completa: '+received+' recibidos, '+sent+' publicados.');if(previousSync)for(const event of newMoves){const product=active().find(p=>p.id===event.data.pid);if(product)notifyMove(event,product)}}finally{busy=false;render()}}
function setup(){ $$('.navitem').forEach(x=>x.onclick=()=>show(x.dataset.page));$$('[data-action]').forEach(x=>x.onclick=()=>openEditor(x.dataset.action));$('#editor').addEventListener('submit',submitEditor);$('#close').onclick=$('#cancel').onclick=()=>$('#modal').close();
document.addEventListener('click',async e=>{let t=e.target.closest('[data-edit],[data-move],[data-delete],[data-count]');if(!t)return;try{if(t.dataset.edit)openEditor('producto',t.dataset.edit);if(t.dataset.move)openEditor('movimiento',t.dataset.move);if(t.dataset.count)openEditor('conteo',t.dataset.count);if(t.dataset.delete){let p=inventory.find(x=>x.id===t.dataset.delete);if(p&&confirm('¿Dar de baja el producto '+p.name+'? Sus movimientos se conservarán.'))await saveEvent('delete',{pid:p.id})}}catch(err){fail(err)}});
['#find-product','#find-move','#filter-type'].forEach(s=>$(s).addEventListener('input',render));
$('#save-name').onclick=async()=>{await setting('name',$('#name').value.trim()||'Inventario Web');render();toast('Nombre guardado')};
$('#export-json').onclick=exportJSON;$('#export-csv').onclick=exportCSV;
$('#import-csv').onchange=async e=>{try{if(e.target.files[0])await importCSV(e.target.files[0])}catch(err){fail(err)}e.target.value=''};
$('#account-exit').onclick=signOut;
$$('[data-kind]').forEach(button=>button.onclick=()=>openEditor('movimiento',null,button.dataset.kind));
$('#device-notices').onclick=async()=>{
  if(settings.deviceNotices){await setting('deviceNotices',false)}
  else if(globalThis.Notification&&'serviceWorker' in navigator&&await Notification.requestPermission()==='granted'){await setting('deviceNotices',true)}
  else {toast('Avisos disponibles dentro de la aplicación.')}
  render()
};
$('#sync-now').onclick=$('#sync-top').onclick=()=>sync().catch(fail);
window.addEventListener('online',()=>{render();scheduleSync()});window.addEventListener('offline',render);
}
async function selectDatabase(driveId){
  db=await openDB();
  let binding=(await one('settings','binding'))?.value;
  if(!binding && (await all('events')).length && confirm('Este dispositivo contiene un inventario anterior sin cuenta. ¿Vincularlo a '+session.profile.label+' y sincronizarlo con su OneDrive?')){
    binding=await store('settings','readwrite',async table=>{
      const current=await req(table.get('binding'));
      if(current)return current.value;
      const folder=(await req(table.get('folder')))?.value||'Inventario Web Independiente';
      const value={driveId,folder};
      await req(table.put({id:'binding',value}));return value
    })
  }
  if(binding?.driveId!==driveId){db.close();db=await openDB(DB+'-account-'+session.profile.id)}
}
async function openAccount(){
  if(!session.profile || navigator.onLine){
    const me=await j('/me?$select=id,displayName,mail,userPrincipalName');
    if(!me.id || (session.profile && me.id!==session.profile.id))throw Error('Vuelve a iniciar sesión para verificar tu cuenta.');
    const drive=await j('/me/drive?$select=id');
    if(!drive.id)throw Error('No se pudo identificar tu OneDrive.');
    setSession({...session,profile:{id:me.id,label:me.mail||me.userPrincipalName||me.displayName,driveId:drive.id}})
  }
  await selectDatabase(session.profile.driveId);
  for(const row of await all('settings'))settings[row.id]=row.value;
  $('#name').value=settings.name||'Inventario Web';
  $('#folder').value=settings.folder||'Inventario Web Independiente';
  setup();await reload();
  $('#access').hidden=true;$('#workspace').hidden=false;scheduleSync();
  setInterval(()=>{if(!document.hidden&&!busy)scheduleSync()},60000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)scheduleSync()})
}
async function main(){
  if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js').catch(console.error);
  window.addEventListener('storage',e=>{if(e.key==='iw-signed-out' || (e.key===SESSION_KEY&&session?.remember&&!e.newValue)){sessionStorage.removeItem(SESSION_KEY);location.replace(redirect())}});
  $('#sign-in').onclick=()=>connect().catch(fail);
  session=loadSession();
  try{await callback();if(session)await openAccount();else $('#access-msg').textContent=''}
  catch(error){$('#access-msg').textContent=navigator.onLine?'No se pudo abrir la cuenta. '+error.message:'Conéctate a internet para verificar tu cuenta por primera vez.'}
  finally{$('#sign-in').disabled=false}
}
main().catch(fail);
