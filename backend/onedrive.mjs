import { requireValue } from './ledger.mjs';

const appId = '0b22855f-0817-4124-8eb7-3e42296b5b17';
const callback = 'https://inventario-bodega.roserocristhian.workers.dev/onedrive/callback';
const authority = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/';
const scopes = 'openid offline_access Files.ReadWrite';
const enc = new TextEncoder();
const encoded = bytes => btoa(String.fromCharCode(...new Uint8Array(bytes))).replaceAll('+','-').replaceAll('/','_').replace(/=+$/,'');
const random = () => encoded(crypto.getRandomValues(new Uint8Array(32)));

export async function beginConnection(storage, env) {
  requireValue(env.ONEDRIVE_CLIENT_SECRET, 'Configura ONEDRIVE_CLIENT_SECRET en Cloudflare.', 503);
  const state = random(), verifier = random(), nonce = random();
  await storage.put('onedrive-pending', {state, verifier, nonce, expires: Date.now()+600000});
  const params = new URLSearchParams({client_id:appId,redirect_uri:callback,response_type:'code',
    scope:scopes,state,nonce,code_challenge_method:'S256',
    code_challenge:encoded(await crypto.subtle.digest('SHA-256',enc.encode(verifier))),
    login_hint:env.ADMIN_EMAIL});
  return {url:authority+'authorize?'+params};
}

export async function finishConnection(storage, env, query, verifyIdentity) {
  const pending = await storage.transaction(async tx => {
    const p = await tx.get('onedrive-pending');
    requireValue(p && p.state===query.get('state') && p.expires>Date.now(), 'Solicitud caducada. Vuelve a conectar desde Inventario Web.', 400);
    await tx.delete('onedrive-pending');
    return p;
  });
  requireValue(query.has('code'), 'Microsoft no concedió acceso.', 403);
  const response = await fetch(authority+'token',{method:'POST',body:new URLSearchParams({
    client_id:appId,client_secret:env.ONEDRIVE_CLIENT_SECRET,redirect_uri:callback,
    grant_type:'authorization_code',code_verifier:pending.verifier,code:query.get('code'),scope:scopes})});
  requireValue(response.ok,'Microsoft rechazó la autorización. Revisa la credencial del servidor.',403);
  const tokens = await response.json();
  const identity = await verifyIdentity(tokens.id_token);
  requireValue(identity.sub===env.ADMIN_SUB && identity.nonce===pending.nonce && tokens.refresh_token, 'Usa la cuenta administradora de la bodega.',403);
  const file = await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/Inventario%20Web%20Independiente/BASE.xlsm?$select=id,name,parentReference,file',
    {headers:{Authorization:'Bearer '+tokens.access_token}});
  requireValue(file.ok,'No se encontró BASE.xlsm en Inventario Web Independiente.',404);
  const item = await file.json();
  requireValue(item.name==='BASE.xlsm' && item.file && item.id && item.parentReference?.driveId,'El archivo no corresponde a BASE.xlsm.',409);
  const bytes = await crypto.subtle.digest('SHA-256',enc.encode('onedrive-storage-v1:'+env.ONEDRIVE_CLIENT_SECRET));
  const key = await crypto.subtle.importKey('raw',bytes,'AES-GCM',false,['encrypt']);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await crypto.subtle.encrypt({name:'AES-GCM',iv},key,enc.encode(tokens.refresh_token));
  await storage.put('onedrive-connection',{drive:item.parentReference.driveId,item:item.id,iv:encoded(iv),sealed:encoded(encrypted)});
}

const from64=s=>Uint8Array.from(atob(s.replaceAll('-','+').replaceAll('_','/')),c=>c.charCodeAt(0));
export async function authorizeSync(storage,env,connection){
  const key=await crypto.subtle.importKey('raw',await crypto.subtle.digest('SHA-256',enc.encode('onedrive-storage-v1:'+env.ONEDRIVE_CLIENT_SECRET)),'AES-GCM',false,['decrypt','encrypt']);
  const token=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:from64(connection.iv)},key,from64(connection.sealed)));
  const response=await fetch(authority+'token',{method:'POST',body:new URLSearchParams({
    client_id:appId,client_secret:env.ONEDRIVE_CLIENT_SECRET,grant_type:'refresh_token',refresh_token:token,scope:scopes})});
  requireValue(response.ok,'Vuelve a autorizar OneDrive.',403);
  const renewed=await response.json();
  if(renewed.refresh_token){
    const iv=crypto.getRandomValues(new Uint8Array(12));
    const sealed=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,enc.encode(renewed.refresh_token));
    await storage.put('onedrive-connection',{...connection,iv:encoded(iv),sealed:encoded(sealed)});
  }
  return {token:renewed.access_token,expires:Date.now()+Number(renewed.expires_in||0)*1000};
}
