import {setupReportUI,familyOf,compareProducts,destinations} from './reports.js?v=10';
import {TEAM_API} from './team-config.js?v=6';

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
let teamUser=null, pulling=null;
const canManage=()=>!TEAM_API||teamUser?.role==='admin';
async function teamCall(path,body={}){
  if(!navigator.onLine)throw Error('Conéctate para confirmar la operación en la bodega.');
  await token();
  const res=await fetch(TEAM_API+path,{method:'POST',headers:{Authorization:'Bearer '+session.id_token,'Content-Type':'application/json'},body:JSON.stringify(body)});
  const data=await res.json();
  if(!res.ok){
    if(res.status===403&&data.error==='Cuenta sin acceso a esta bodega.'&&teamUser){teamUser=null;location.replace(redirect())}
    throw Error(data.error||'No se pudo contactar a la bodega.')
  }
  return data
}
async function pullTeam(){
  if(pulling)return pulling;
  pulling=(async()=>{
    const previousSync=settings.lastSync, incoming=[];let rows;
    do{
      const result=await teamCall('/changes',{after:settings.cursor||0});
      rows=result.rows;
      for(const {event} of rows){
        for(const id of [event.data.photo,...(event.data.attachments||[])].filter(Boolean))if(!await one('photos',id)){
          const photo=await teamCall('/photo',{id});
          if(photo.data)await put('photos',{id,data:photo.data,synced:true})
        }
        if(previousSync&&['entrada','salida'].includes(event.type)&&!await one('events',event.id))incoming.push(event);
        await put('events',event)
      }
      if(rows.length)await setting('cursor',rows.at(-1).seq);
      if(result.anchor)await setting('reportAnchor',JSON.parse(result.anchor));
      if(result.name)await setting('name',result.name);
    }while(rows.length===200);
    await setting('lastSync',new Date().toISOString());await reload();
    for(const event of incoming){const product=active().find(p=>p.id===event.data.pid);if(product)notifyMove(event,product)}
  })().finally(()=>{pulling=null});
  return pulling
}
async function commitEvent(event){
  if(!TEAM_API){await put('events',event);return}
  if(event.data.photo){
    const photo=await one('photos',event.data.photo);
    if(photo&&!photo.synced){await teamCall('/photo-save',photo);await put('photos',{...photo,synced:true})}
  }
  const saved=await teamCall('/event',event);await put('events',saved)
}
async function listMembers(){
  const {active,pending}=await teamCall('/members');
  $('#member-list').innerHTML=[...active.map(x=>[x.email,'Activo']),...pending.map(x=>[x.email,'Invitación pendiente'])].map(([address,state])=>'<div class="item"><div class="item-body">'+esc(address)+'<p>'+state+'</p></div><button class="small" data-revoke="'+esc(address)+'">Retirar acceso</button></div>').join('')||empty('Todavía no has agregado usuarios.')
}
async function backupTeam(){
  await pullTeam();
  const name='respaldo-bodega-'+new Date().toISOString().replace(/[:.]/g,'-')+'.json';
  const folder=await childFolder((await j('/me/drive/root?$select=id')).id,'Respaldos Inventario Web');
  await up(folder.id,name,JSON.stringify({format:'inventario-web-independent-v1',events,photos}),'application/json');
  toast('Respaldo guardado en tu OneDrive.')
}
async function importTeam(file){
  const data=JSON.parse(await file.text());
  if(data.format!=='inventario-web-independent-v1'||!Array.isArray(data.events)||!Array.isArray(data.photos))throw Error('Selecciona un respaldo JSON de Inventario Web.');
  if(!confirm('¿Incorporar '+data.events.length+' registros de este respaldo a la bodega compartida?'))return;
  for(const photo of data.photos)await teamCall('/photo-save',photo);
  const ordered=[...data.events].sort((a,b)=>Number(b.type==='create')-Number(a.type==='create')||a.ts.localeCompare(b.ts)||a.id.localeCompare(b.id));
  for(let i=0;i<ordered.length;i+=100)await teamCall('/import',{events:ordered.slice(i,i+100)});
  await pullTeam();toast('Respaldo incorporado a la bodega.')
}
async function recoverPersonalInventory(){
  const profile=await j('/me?$select=id');
  const drive=await j('/me/drive?$select=id');
  let personal=await openDB();
  try{
    const binding=await req(personal.transaction('settings').objectStore('settings').get('binding'));
    if(!binding || binding.value.driveId!==drive.id){
      personal.close();
      personal=await openDB(DB+'-account-'+profile.id);
    }
    const transaction=personal.transaction(['events','photos']);
    const [events,photos]=await Promise.all(['events','photos'].map(name=>req(transaction.objectStore(name).getAll())));
    if(!events.length)throw Error('No se encontraron registros de esta cuenta en este navegador. Puedes incorporar el respaldo JSON.');
    const json=JSON.stringify({format:'inventario-web-independent-v1',events,photos});
    download('respaldo-personal-'+today()+'.json',json,'application/json');
    await importTeam(new Blob([json]));
  }finally{personal.close()}
}
function setupTeam(){
  $('#settings-link').addEventListener('click',async()=>{if(canManage())try{$('#onedrive-status').textContent=(await teamCall('/onedrive/status')).message}catch(error){fail(error)}});
  $('#onedrive-connect').onclick=async()=>{try{
    const result=await teamCall('/onedrive/start');
    location.assign(result.url);
  }catch(error){fail(error)}};
  $('#team-recover').onclick=()=>recoverPersonalInventory().catch(fail);
  $('#team-tools').hidden=!canManage();
  for(const el of $$('[data-action="producto"],#import-csv,#name,#save-name')){if(!canManage())(el.closest('label')||el).hidden=true}
  $('#settings-link').addEventListener('click',()=>{if(canManage())listMembers().catch(fail)});
  $('#invite-form').onsubmit=async e=>{
    e.preventDefault();
    try{const {code}=await teamCall('/invite',{email:$('#member-email').value});
      $('#invite-result').value=redirect()+'#invitar='+code;await listMembers()}catch(error){fail(error)}
  };
  $('#member-list').onclick=async e=>{
    const address=e.target.closest('[data-revoke]')?.dataset.revoke;
    if(address&&confirm('¿Retirar el acceso de '+address+'?'))try{await teamCall('/remove',{email:address});$('#invite-result').value='';await listMembers()}catch(error){fail(error)}
  };
  $('#team-backup').onclick=()=>backupTeam().catch(fail);
  $('#team-import').onchange=e=>{const file=e.target.files[0];if(file)importTeam(file).catch(fail);e.target.value=''};
}
function dateLabel(s){return String(s||'').slice(0,10)}
function record(type,data){data={...data,actor:session.profile.label};return {id:uid(),type,date:today(),ts:new Date().toISOString(),data,synced:false}}
async function saveEvent(type,data){await commitEvent(record(type,data));await reload()}
function compute(){let p=new Map();moves=[];alerts=[];let sorted=[...events].sort((a,b)=>Number(b.type==='create')-Number(a.type==='create')||String(a.ts).localeCompare(String(b.ts))||String(a.id).localeCompare(String(b.id)));
const voided=new Set(events.filter(e=>e.type==='anular').map(e=>e.data.target));
for(const e of sorted){if(voided.has(e.id))continue;let d=e.data||{},x=p.get(d.pid);
if(e.type==='create'){if(p.has(d.pid)){alerts.push('Producto duplicado');continue}p.set(d.pid,{created:e.date,id:d.pid,code:d.code,name:d.name,unit:d.unit,cat:d.cat,family:d.family,lbPerSack:d.lbPerSack,min:d.min||0,photo:d.photo||null,stock:0,deleted:false});continue}
if(!x){alerts.push('Movimiento sin producto: '+e.id);continue}
if(e.type==='edit'){Object.assign(x,{code:d.code,name:d.name,unit:d.unit,cat:d.cat,family:d.family,lbPerSack:d.lbPerSack,min:d.min||0,photo:d.photo||null});continue}
if(e.type==='delete'){x.deleted=true;x.removed=e.date;continue}
if(['entrada','salida','ajuste','traspaso'].includes(e.type)){let q=Number(d.q);if(!Number.isFinite(q)||(e.type!=='ajuste'&&q<=0)){alerts.push('Cantidad incorrecta: '+e.id);continue}if(e.type==='entrada')x.stock+=q;if(e.type==='salida')x.stock-=q;if(e.type==='ajuste')x.stock+=q;moves.push({...e,prod:x.name,code:x.code,unit:x.unit,q})}
}
inventory=[...p.values()].sort(compareProducts);moves.reverse();
let codes=new Set();for(const x of inventory){x.stock=Math.round(x.stock*100000)/100000;if(x.stock<0)alerts.push('Saldo negativo en '+x.code);if(!x.deleted&&x.cat==='directo'){if(codes.has(x.code.toLowerCase()))alerts.push('Código repetido '+x.code);codes.add(x.code.toLowerCase())}}}
async function reload(){events=await all('events');photos=await all('photos');compute();render()}
function active(){return inventory.filter(p=>!p.deleted&&p.cat==='directo')}
const pic=id=>photos.find(x=>x.id===id)?.data||'';
function productRow(p,actions){let state=p.stock<=0?'<span class="tag danger">Agotado</span>':p.min>0&&p.stock<p.min?'<span class="tag warning">Bajo mínimo</span>':'<span class="tag">Disponible</span>';if(actions&&!canManage())return productRow(p,false).replace('</h3>','</h3><button class="small" data-move="'+esc(p.id)+'">Movimiento</button>');let buttons=actions?'<div class="item-buttons"><button data-edit="'+esc(p.id)+'">Editar</button><button data-move="'+esc(p.id)+'">Movimiento</button><button class="danger" data-delete="'+esc(p.id)+'">Eliminar</button></div>':'';return '<div class="item">'+(p.photo?'<img class="thumb" src="'+esc(pic(p.photo))+'" alt="Foto del producto">':'')+'<div class="item-body"><h3>'+esc(p.name)+'</h3><p>'+esc(p.code)+' · '+esc(p.unit)+' · '+'Material directo'+'</p>'+buttons+'</div><div class="quantity"><strong>'+fmt(p.stock)+'</strong><small>'+esc(p.unit)+'</small>'+state+'</div></div>'}
function empty(t){return '<div class="empty">'+esc(t)+'</div>'}
function moveRow(m){const controls=(m.type==='salida'&&canManage()?'<button class="small danger" data-void="'+esc(m.id)+'">Eliminar salida</button>':'')+(m.data.attachments||[]).map((id,i)=>'<button class="small" data-picture="'+esc(id)+'">Foto '+(i+1)+'</button>').join('');let q=m.type==='salida'?'-':m.type==='entrada'?'+':m.type==='ajuste'&&m.q>=0?'+':'';return '<div class="item"><div class="item-body"><h3>'+esc(m.prod)+'</h3><p>'+esc(m.date)+(m.data.actor?' · '+esc(m.data.actor):'')+' · '+esc(m.data.doc||'Sin documento')+' · '+esc(m.data.origin||'')+(m.data.dest?' → '+esc(m.data.dest):'')+'</p><span class="tag">'+esc(m.type.toUpperCase())+'</span>'+controls+'</div><div class="quantity"><strong>'+q+fmt(m.q)+'</strong><small>'+esc(m.unit)+'</small></div></div>'}
function render(){if(!session?.profile)return;let pp=active(),todayN=moves.filter(m=>m.date===today()&&inventory.find(x=>x.id===m.data.pid)?.cat==='directo').length;
$('#n-products').textContent=pp.length;$('#n-out').textContent=pp.filter(p=>p.stock<=0).length;$('#n-low').textContent=pp.filter(p=>p.min>0&&p.stock>0&&p.stock<p.min).length;$('#n-today').textContent=todayN;
$('#attention').innerHTML=pp.filter(p=>p.stock<=0||(p.min>0&&p.stock<p.min)).map(p=>productRow(p,false)).join('')||empty('No hay productos con alertas.');
let term=$('#find-product').value.toLowerCase();
$('#products').innerHTML=active().filter(p=>($('#stock-filter').value!=='alert'||p.stock<=0||p.stock<p.min)&&(p.name.toLowerCase().includes(term)||p.code.toLowerCase().includes(term))).map(p=>productRow(p,true)).join('')||empty('No hay productos. Puedes agregar uno nuevo.');
let query=$('#find-move').value.toLowerCase(),type=$('#filter-type').value;
$('#moves').innerHTML=moves.filter(m=>inventory.find(p=>p.id===m.data.pid)?.cat==='directo'&&(!$('#moves-from').value||m.date>=$('#moves-from').value)&&(!$('#moves-to').value||m.date<=$('#moves-to').value)&&(type==='todos'||type===m.type)&&[m.prod,m.code,m.data.doc,m.data.origin,m.data.dest,m.date].some(v=>String(v||'').toLowerCase().includes(query))).slice(0,600).map(moveRow).join('')||empty('No hay movimientos registrados.');
$('#counts').innerHTML=active().map(p=>'<div class="item"><div class="item-body"><h3>'+esc(p.name)+'</h3><p>'+esc(p.code)+' · '+esc(p.unit)+'</p></div><div class="quantity"><strong>'+fmt(p.stock)+'</strong><button class="small" data-count="'+esc(p.id)+'">Contar</button></div></div>').join('')||empty('Registra productos para hacer un conteo.');
$('#appname').textContent=settings.name && settings.name!=='Inventario Web'?settings.name:'Inventario camaronera';
$('#display-date').textContent=new Intl.DateTimeFormat('es-EC',{dateStyle:'full'}).format(new Date());
const pending=events.filter(e=>!e.synced).length;
$('#network').textContent=!navigator.onLine?'Sin conexión':busy?'Sincronizando…':pending?'Cambios pendientes':settings.lastSync?'Sincronizado':'OneDrive conectado';
$('#account-label').textContent=$('#cloud-account').textContent=session.profile.label;
$('#activity').innerHTML=moves.filter(m=>['entrada','salida'].includes(m.type)&&inventory.some(p=>p.id===m.data.pid&&p.cat==='directo')).slice(0,5).map(moveRow).join('')||empty('Aún no hay entradas ni salidas.');
$('#account-exit').disabled=busy;
if(TEAM_API){
  for(const button of $$('[data-count]'))button.hidden=!canManage();
  $('#account-label').textContent=teamUser.email+' · '+(canManage()?'Administrador':'Operador');
}
$('#notice-status').textContent=settings.deviceNotices&&globalThis.Notification?.permission==='granted'?'Avisos del dispositivo activados.':'Los avisos aparecen dentro de la aplicación.';
$('#device-notices').textContent=settings.deviceNotices?'Desactivar notificaciones del dispositivo':'Activar notificaciones del dispositivo';
$('#sync-now').disabled=busy||!navigator.onLine||!session;
$('#alert').classList.toggle('hidden',!alerts.length);$('#alert').textContent=alerts.slice(0,5).join(' · ');
$('#cloud-msg').textContent=TEAM_API?'Bodega compartida · Última actualización: '+(settings.lastSync||'pendiente'):events.filter(e=>!e.synced).length+' operaciones sin publicar en OneDrive. Última sincronización: '+(settings.lastSync||'nunca');
}
function show(id){page=id;window.scrollTo(0,0);$$('.page').forEach(el=>el.classList.toggle('hidden',el.id!=='p-'+id));$$('.navitem').forEach(el=>el.classList.toggle('on',el.dataset.page===id));render()}
const field=(lab,name,val='',type='text',extra='')=>'<label>'+lab+'<input name="'+name+'" type="'+type+'" value="'+esc(val)+'" '+extra+'></label>';
const select=(lab,name,opts,selected)=>'<label>'+lab+'<select name="'+name+'">'+opts.map(([v,t])=>'<option value="'+esc(v)+'" '+(v===selected?'selected':'')+'>'+esc(t)+'</option>').join('')+'</select></label>';
function openEditor(kind,pid=null,initial='entrada'){let p=inventory.find(x=>x.id===pid);editor={kind,pid,expected:p?.stock};let h='';
if(kind==='producto'){$('#modal-title').textContent=p?'Editar producto':'Nuevo producto';h+=field('Código','code',p?.code||'','text','required maxlength="35"')+field('Descripción','name',p?.name||'','text','required maxlength="160"');h+=select('Medida','unit',[...new Set(['SC','KG','LT','UND','CJ','PZ','GL','M',...(p?[p.unit]:[])])].map(x=>[x,x]),p?.unit||'SC');
h+=select('Grupo','family',[['balanceado','Balanceado'],['insumo','Insumo']],p?familyOf(p):'insumo')+field('Libras por saco de balanceado (opcional)','lbs',p?.lbPerSack||'','number','min="0.000001" step="any"');
if(!p)h+=field('Inventario inicial','opening',0,'number','required min="0" step="any"');h+=field('Existencia mínima','min',p?.min??0,'number','required min="0" step="any"');h+='<label>Fotografía (opcional)<input type="file" name="picture" accept="image/jpeg,image/png,image/webp"></label>'}
if(kind==='movimiento'){if(!active().length){toast('Primero agrega un producto.',true);return}$('#modal-title').textContent='Registrar movimiento';h+=select('Tipo de movimiento','type',[['entrada','Entrada'],['salida','Salida'],...(canManage()?[['traspaso','Traspaso']]:[])],initial)+'<div id="movement-fields"></div>'}

if(kind==='conteo'){if(!p)return;$('#modal-title').textContent='Conteo físico';h+='<div class="note">'+esc(p.name)+' · Existencia: '+fmt(p.stock)+' '+esc(p.unit)+'</div>'+field('Cantidad encontrada','q',p.stock,'number','required min="0" step="any"')+field('Observación','note','Conteo físico','text','maxlength="180"')}
$('#fields').innerHTML=h;if(kind==='movimiento'){drawMovement(initial,pid);$('#fields select[name="type"]').onchange=e=>drawMovement(e.target.value,pid)}$('#modal').showModal()}
async function imageData(file){if(file.size>10*1024*1024)throw Error('Imagen demasiado grande (máximo 10 MB).');if(!['image/png','image/jpeg','image/webp'].includes(file.type))throw Error('Formato de imagen inválido.');let bitmap=await createImageBitmap(file);let scale=Math.min(1,1200/Math.max(bitmap.width,bitmap.height));let can=document.createElement('canvas');can.width=Math.round(bitmap.width*scale);can.height=Math.round(bitmap.height*scale);can.getContext('2d').drawImage(bitmap,0,0,can.width,can.height);bitmap.close();return can.toDataURL('image/jpeg',.75)}
async function submitEditor(e){e.preventDefault();const submit=e.target.querySelector('[type=submit]');if(submit.disabled)return;submit.disabled=true;try{const f=new FormData(e.target),v=k=>String(f.get(k)||'').trim(),kind=editor.kind,pid=editor.pid;
if(kind==='producto'){let code=v('code'),name=v('name'),min=number(v('min'));if(!code||!name||min<0)throw Error('Completa el producto y el mínimo válido.');if(active().some(p=>p.id!==pid&&p.code.toLowerCase()===code.toLowerCase()))throw Error('Este código ya existe.');let prev=inventory.find(p=>p.id===pid),photo=prev?.photo||null,attached=f.get('picture');if(attached instanceof File&&attached.size){photo=uid();await put('photos',{id:photo,data:await imageData(attached),synced:false})}
const productId=pid||(editor.productId||=uid()),data={pid:productId,code,name,min,cat:'directo',family:v('family'),lbPerSack:v('lbs')?number(v('lbs')):null,unit:v('unit'),photo};
if(photo&&TEAM_API){const obj=await one('photos',photo);if(obj&&!obj.synced){await teamCall('/photo-save',obj);await put('photos',{...obj,synced:true})}}
const batch=[record(pid?'edit':'create',data)],opening=number(v('opening'));
if(!pid&&opening>0)batch.push(record('ajuste',{pid:productId,q:opening,expected:0,doc:'INVENTARIO INICIAL'}));
await saveBatch(batch);await reload()}
if(kind==='movimiento'){await submitMovement(f);await reload()}
if(kind==='conteo'){let p=active().find(x=>x.id===pid),q=number(v('q'));if(!p||q<0)throw Error('La cantidad no puede ser negativa.');await saveEvent('ajuste',{pid,q:q-(TEAM_API?editor.expected:p.stock),expected:editor.expected,note:v('note'),doc:'CONTEO'})}
$('#modal').close();if(kind!=='movimiento'||v('type')==='traspaso')toast(TEAM_API?'Registro confirmado en la bodega.':'Registro guardado en este dispositivo.');scheduleSync()}catch(err){fail(err)}finally{submit.disabled=false}}
async function exportJSON(){let json=JSON.stringify({format:'inventario-web-independent-v1',fecha:new Date().toISOString(),events,photos},null,2);download('respaldo-inventario-'+today()+'.json',json,'application/json')}
function download(name,body,type){const a=document.createElement('a'),url=URL.createObjectURL(new Blob([body],{type}));a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),3000)}
function csvEscape(x){let s=String(x??'');return /[;"\n\r]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}
function exportCSV(){download('productos-'+today()+'.csv','\uFEFF'+[['codigo','descripcion','unidad','categoria','minimo','stock'],...active().map(x=>[x.code,x.name,x.unit,x.cat,x.min,x.stock])].map(r=>r.map(csvEscape).join(';')).join('\r\n'),'text/csv')}
function csvParse(txt){let t=txt.replace(/^\uFEFF/,''),sep=t.split(/\r?\n/,1)[0].includes(';')?';':',';let rows=[],row=[],cell='',quote=false;for(let i=0;i<t.length;i++){let ch=t[i];if(ch==='"'){if(quote&&t[i+1]==='"'){cell+='"';i++}else quote=!quote}else if(ch===sep&&!quote){row.push(cell);cell=''}else if((ch==='\n'||ch==='\r')&&!quote){if(ch==='\r'&&t[i+1]==='\n')i++;row.push(cell);if(row.some(v=>v.trim()))rows.push(row);row=[];cell=''}else cell+=ch}if(quote)throw Error('CSV inválido: comillas sin cerrar');row.push(cell);if(row.some(v=>v.trim()))rows.push(row);return rows}
async function importCSV(file){let rows=csvParse(await file.text());let first=rows.shift()?.map(x=>x.toLowerCase().trim())||[];let get=(r,key)=>r[first.indexOf(key)];if(!first.includes('codigo')||!first.includes('descripcion'))throw Error('El CSV debe tener columnas codigo y descripcion.');let codes=new Set(active().map(x=>x.code.toLowerCase())),count=0;for(const r of rows){let code=(get(r,'codigo')||'').trim(),name=(get(r,'descripcion')||'').trim(),cat=(get(r,'categoria')||'').trim().toLowerCase();if((cat&&cat!=='directo'&&cat!=='material directo')||!code||!name||codes.has(code.toLowerCase()))continue;let min=Number(get(r,'minimo')||0);if(!Number.isFinite(min)||min<0)min=0;await commitEvent(record('create',{pid:uid(),code,name,unit:get(r,'unidad')||'UND',cat:'directo',min,photo:null}));codes.add(code.toLowerCase());count++}await reload();toast(count+' productos importados.')}
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
async function sync(){if(busy)return;if(TEAM_API){busy=true;render();try{await pullTeam()}finally{busy=false;render()}return}if(!session||!navigator.onLine)throw Error('Conecta OneDrive y verifica internet.');busy=true;render();try{let folder=folderName(),drive=await j('/me/drive?$select=id,driveType,webUrl'),binding=settings.binding;if(binding&&(binding.driveId!==drive.id||binding.folder!==folder))throw Error('Este inventario local ya está vinculado a otro OneDrive. No se modificarán los datos.');let root;try{root=await j('/me/drive/root:/'+enc(folder))}catch(e){if(e.status!==404)throw e}if(!root)root=await j('/me/drive/root/children',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:folder,folder:{},'@microsoft.graph.conflictBehavior':'fail'})});if(!root.folder)throw Error('La ruta no es una carpeta.');let items=await children(root.id),marker=items.find(x=>x.name===MARK);if(!marker){if(items.length)throw Error('Carpeta existente sin identidad de esta app: elige otra carpeta vacía.');await up(root.id,MARK,JSON.stringify({app:'inventario-web-independent-v1',created:new Date().toISOString()}),'application/json')}else{let d=await (await graph('/me/drive/items/'+enc(marker.id)+'/content')).json();if(d.app!=='inventario-web-independent-v1')throw Error('Esta carpeta pertenece a otra aplicación.')}
let evFolder=await childFolder(root.id,'eventos'),phFolder=await childFolder(root.id,'fotos');
await setting('binding',{driveId:drive.id,folder});
let remote=await children(evFolder.id),ids=new Set();let received=0,sent=0;const newMoves=[];for(const f of remote){if(!/^[0-9a-f-]{36}\.json$/i.test(f.name))continue;let id=f.name.slice(0,-5);ids.add(id);let x=await one('events',id);if(x){if(!x.synced)await put('events',{...x,synced:true});continue}let content=await (await graph('/me/drive/items/'+enc(f.id)+'/content')).json();if(!content.event||content.event.id!==id)throw Error('Movimiento remoto inválido');await put('events',{...content.event,synced:true});received++;if(['entrada','salida'].includes(content.event.type))newMoves.push(content.event)}
let local=await all('events'),remotePhotos=await children(phFolder.id),names=new Set(remotePhotos.map(x=>x.name));
for(const e of local){if(ids.has(e.id))continue;await up(evFolder.id,e.id+'.json',JSON.stringify({event:{id:e.id,type:e.type,ts:e.ts,date:e.date,data:e.data}}),'application/json');await put('events',{...e,synced:true});sent++}
const references=new Set(local.flatMap(x=>[x.data.photo,...(x.data.attachments||[])]).filter(Boolean));let localPhotos=await all('photos');for(const p of localPhotos){if(!references.has(p.id))continue;if(names.has(p.id+'.jpg'))continue;let res=await fetch(p.data),blob=await res.blob();await up(phFolder.id,p.id+'.jpg',blob,'image/jpeg');await put('photos',{...p,synced:true})}
for(const p of remotePhotos){if(!/^[0-9a-f-]{36}\.jpg$/i.test(p.name))continue;let id=p.name.slice(0,-4);if(!references.has(id)||await one('photos',id))continue;let resp=await graph('/me/drive/items/'+enc(p.id)+'/content'),blob=await resp.blob(),data=await new Promise((yes,no)=>{let r=new FileReader();r.onload=()=>yes(r.result);r.onerror=()=>no(r.error);r.readAsDataURL(blob)});await put('photos',{id,data,synced:true})}
const previousSync=settings.lastSync;await setting('lastSync',new Date().toISOString());await reload();toast('Sincronización completa: '+received+' recibidos, '+sent+' publicados.');if(previousSync)for(const event of newMoves){const product=active().find(p=>p.id===event.data.pid);if(product)notifyMove(event,product)}}finally{busy=false;render()}}
function setup(){
  // This display is informational. The Worker independently verifies the signature and subject.
  try{
    const part=session.id_token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
    $('#identity-value').textContent=JSON.parse(atob(part)).sub;
  }catch{$('#identity-help').hidden=true}

  $('#cloud-link').onclick=()=>show('onedrive');
  $('#settings-link').onclick=()=>show('ajustes');
  $('#view-attention').onclick=()=>{$('#stock-filter').value='alert';$('#find-product').value='';show('productos')};
  $('#stock-filter').onchange=render;
  setupReportUI({$,today,inventory:()=>inventory,moves:()=>moves,anchor:()=>settings.reportAnchor,admin:canManage,fail,saveAnchor:async anchor=>{if(TEAM_API)await teamCall('/report-anchor',anchor);await setting('reportAnchor',anchor)}});
  document.addEventListener('click',async e=>{
    const id=e.target.closest('[data-void]')?.dataset.void, photo=e.target.closest('[data-picture]')?.dataset.picture;
    try{
      if(photo){const file=await one('photos',photo)||(TEAM_API&&await teamCall('/photo',{id:photo}));if(!file?.data)throw Error('Foto no disponible.');$('#evidence-img').src=file.data;$('#evidence').showModal()}
      if(id&&canManage()){const m=moves.find(x=>x.id===id);if(m&&confirm('¿Eliminar la salida de '+fmt(m.q)+' '+m.unit+' de '+m.prod+' del '+m.date+'? La cantidad volverá a bodega.')){await saveEvent('anular',{pid:m.data.pid,target:id});toast('Salida anulada y saldo restituido.')}}
    }catch(err){fail(err)}
  });
 $$('.navitem').forEach(x=>x.onclick=()=>show(x.dataset.page));$$('[data-action]').forEach(x=>x.onclick=()=>openEditor(x.dataset.action));$('#editor').addEventListener('submit',submitEditor);$('#close').onclick=$('#cancel').onclick=()=>$('#modal').close();
document.addEventListener('click',async e=>{let t=e.target.closest('[data-edit],[data-move],[data-delete],[data-count]');if(!t)return;try{if(t.dataset.edit)openEditor('producto',t.dataset.edit);if(t.dataset.move)openEditor('movimiento',t.dataset.move);if(t.dataset.count)openEditor('conteo',t.dataset.count);if(t.dataset.delete){let p=inventory.find(x=>x.id===t.dataset.delete);if(p&&confirm('¿Dar de baja el producto '+p.name+'? Sus movimientos se conservarán.'))await saveEvent('delete',{pid:p.id})}}catch(err){fail(err)}});
['#find-product','#find-move','#filter-type','#moves-from','#moves-to'].forEach(s=>$(s).addEventListener('input',render));
$('#save-name').onclick=async()=>{try{if(TEAM_API)await teamCall('/name',{name:$('#name').value.trim()||'Inventario Web'});await setting('name',$('#name').value.trim()||'Inventario Web');render();toast('Nombre guardado')}catch(error){fail(error)}};
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
$('#sync-now').onclick=()=>sync().catch(fail);
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
  if(TEAM_API){
    const code=sessionStorage.getItem('iw-team-invitation');
    teamUser=await teamCall(code?'/join':'/me',code?{code}:{});
    sessionStorage.removeItem('iw-team-invitation');
    setSession({...session,profile:{id:teamUser.subject,label:teamUser.email}});
    db=await openDB(DB+'-shared-'+encodeURIComponent(TEAM_API)+'-'+teamUser.subject);
    for(const row of await all('settings'))settings[row.id]=row.value;
    setup();setupTeam();await pullTeam();
    $('#name').value=settings.name||'Inventario Web';
    $('#folder').closest('label').hidden=true;
    $('#p-onedrive h2').textContent='Bodega compartida';
    $('#access').hidden=true;$('#workspace').hidden=false;
    setInterval(()=>{if(!document.hidden&&navigator.onLine)scheduleSync()},30000);
    return
  }
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
  const invite=new URLSearchParams(location.hash.slice(1)).get('invitar');
  if(TEAM_API&&invite){sessionStorage.setItem('iw-team-invitation',invite);history.replaceState({},'',location.pathname+location.search)}

  function setTheme(light){
    document.documentElement.classList.toggle('light',light);
    $('#theme-toggle').setAttribute('aria-label',light?'Activar tema oscuro':'Activar tema claro');
    document.querySelector('meta[name="theme-color"]').content=light?'#eff5f2':'#0b1514';
  }
  setTheme(localStorage.getItem('iw-theme')==='light');
  $('#theme-toggle').onclick=()=>{const light=!document.documentElement.classList.contains('light');localStorage.setItem('iw-theme',light?'light':'dark');setTheme(light)};

  if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js').catch(console.error);
  window.addEventListener('storage',e=>{if(e.key==='iw-signed-out' || (e.key===SESSION_KEY&&session?.remember&&!e.newValue)){sessionStorage.removeItem(SESSION_KEY);location.replace(redirect())}});
  $('#sign-in').onclick=()=>connect().catch(fail);
  session=loadSession();
  try{await callback();if(session)await openAccount();else $('#access-msg').textContent=''}
  catch(error){$('#access-msg').textContent=navigator.onLine?'No se pudo abrir la cuenta. '+error.message:'Conéctate a internet para verificar tu cuenta por primera vez.'}
  finally{$('#sign-in').disabled=false}
}
main().catch(fail);

async function saveBatch(batch){
  const key=JSON.stringify(batch.map(e=>[e.type,e.date,e.data]));
  if(editor.batch?.key===key)batch=editor.batch.events;
  else editor.batch={key,events:batch};
  const saved=TEAM_API?(await teamCall('/batch',{events:batch})).events:batch;
  await store('events','readwrite',s=>Promise.all(saved.map(e=>req(s.put(e)))));
}
function drawMovement(type,pid){
  const area=$('#movement-fields'),list=active(),selected=list.find(x=>x.id===pid);
  const options=list.map(p=>'<option value="'+esc(p.code+' · '+p.name)+'"></option>').join('');
  const product=(pre)=>'<label>Producto<input list="catalog-options" name="product" placeholder="Buscar por nombre o código" value="'+esc(pre?pre.code+' · '+pre.name:'')+'" required></label>';
  const quantity=()=>field('Cantidad','q','','number','min="0.000001" step="any" required');
  area.innerHTML='<datalist id="catalog-options">'+options+'</datalist>'+field('Fecha','date',today(),'date','required');
  if(type==='entrada'){
    area.insertAdjacentHTML('beforeend',field('Número de documento (guía o factura)','doc','','text','maxlength="75"')+'<div id="entry-lines"></div><button type="button" id="add-line" class="small">+ Agregar otro producto</button><label>Tomar foto<input type="file" accept="image/*" capture="environment" name="camera"></label><label>Galería (hasta 4 fotos en total)<input type="file" accept="image/jpeg,image/png,image/webp" multiple name="gallery"></label>');
    const add=()=>{const rows=$('#entry-lines');if(rows.children.length>=100)return;const row=document.createElement('fieldset');row.innerHTML=product(rows.children.length?null:selected)+quantity()+'<button type="button" class="small">Quitar</button>';rows.append(row);row.querySelector('button').onclick=()=>{if(rows.children.length>1)row.remove();$('#add-line').disabled=false};$('#add-line').disabled=rows.children.length>=100};
    $('#add-line').onclick=add;add();
  }else{
    area.insertAdjacentHTML('beforeend',product(selected)+quantity()+(type==='traspaso'?field('Origen','origin','','text','required maxlength="80"'):'')+'<fieldset><legend>Destino obligatorio</legend>'+['Piscinas','Precrías'].map((label,i)=>'<strong>'+label+'</strong><div class="ponds">'+destinations.filter(x=>x.startsWith('PRE')===!!i).map(d=>'<label><input type="radio" name="dest" value="'+d+'" required><span>'+d+'</span></label>').join('')+'</div>').join('')+'</fieldset><p id="balance-preview" class="note"></p>');
    const preview=()=>{const p=findProduct(area.querySelector('[name=product]').value),q=Number(area.querySelector('[name=q]').value);$('#balance-preview').textContent=p?'Stock actual: '+fmt(p.stock)+' '+p.unit+(type==='salida'&&q>0?' · Quedará en '+fmt(p.stock-q)+' '+p.unit:''):''};
    area.oninput=preview;preview();
  }
  if(type==='entrada')area.oninput=null;
  area.insertAdjacentHTML('beforeend',field('Observaciones','note','','text','maxlength="180"'));
}
function findProduct(value){return active().find(p=>[p.code,p.code+' · '+p.name].some(s=>s.toLowerCase()===value.trim().toLowerCase()))}
async function submitMovement(form){
  const val=k=>String(form.get(k)||'').trim(),type=val('type'),quantities=form.getAll('q');
  const batch=form.getAll('product').map((value,i)=>{
    const p=findProduct(value),q=Number(quantities[i]);if(!p||!Number.isFinite(q)||q<=0)throw Error('Selecciona un producto del catálogo y una cantidad positiva.');
    if(type==='salida'&&q>p.stock)throw Error('Stock insuficiente para '+p.name);
    if(type!=='entrada'&&!destinations.includes(val('dest')))throw Error('Selecciona una piscina o precría.');
    if(type==='traspaso'&&val('origin')===val('dest'))throw Error('Selecciona destinos distintos.');
    const e=record(type,{pid:p.id,q,doc:val('doc'),origin:val('origin'),dest:val('dest'),note:val('note')});e.date=val('date');return e;
  });
  const files=[...form.getAll('camera'),...form.getAll('gallery')].filter(f=>f?.size);
  if(files.length>4)throw Error('Selecciona hasta cuatro fotos.');
  const ids=[];editor.fileIds??=new Map();
  for(const file of files){
    const key=[file.name,file.size,file.lastModified].join('/');let id=editor.fileIds.get(key);
    if(!id){id=uid();const data=await imageData(file);if(TEAM_API)await teamCall('/attachment',{id,data});await put('photos',{id,data,synced:!!TEAM_API});editor.fileIds.set(key,id)}ids.push(id);
  }
  if(ids.length)for(const e of batch)e.data.attachments=ids;
  await saveBatch(batch);
  if(type!=='traspaso')notifyMove(batch[0],active().find(p=>p.id===batch[0].data.pid));
}
