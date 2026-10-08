import {timingSafeEqual} from 'node:crypto';
const encoder=new TextEncoder();
const digest=async value=>new Uint8Array(await crypto.subtle.digest('SHA-256',encoder.encode(value)));
const key=async code=>crypto.subtle.importKey('raw',await digest(code),{name:'HMAC',hash:'SHA-256'},false,['sign','verify']);
const cookieName='taisiam_session';
const hex=bytes=>Array.from(new Uint8Array(bytes),b=>b.toString(16).padStart(2,'0')).join('');
export function json(value,status=200,headers={}){return Response.json(value,{status,headers:{'Cache-Control':'no-store','X-Content-Type-Options':'nosniff',...headers}})}
export async function authenticated(request,env){
 if(!env.TESTER_ACCESS_CODE)return false;
 const token=(request.headers.get('Cookie')||'').split(';').map(s=>s.trim()).find(s=>s.startsWith(cookieName+'='))?.slice(cookieName.length+1)||'';
 const [expiry,nonce,signature]=token.split('.');
 if(!/^\d{13}$/.test(expiry||'')||!/^\w{32}$/.test(nonce||'')||! /^[a-f0-9]{64}$/.test(signature||'')||Number(expiry)<=Date.now()||Number(expiry)>Date.now()+28800000)return false;
 return crypto.subtle.verify('HMAC',await key(env.TESTER_ACCESS_CODE),Uint8Array.from(signature.match(/../g),s=>parseInt(s,16)),encoder.encode(expiry+'.'+nonce));
}
function cookie(request,value,age){return `${cookieName}=${value}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${age}${new URL(request.url).protocol==='https:'?'; Secure':''}`}
export async function authRoute(request,env,path,readJson){
 if(path==='/api/auth/session'&&request.method==='GET')return json({required:true,authenticated:await authenticated(request,env),storage:'cloudflare'});
 if(path==='/api/auth/logout'&&request.method==='POST')return json({ok:true},200,{'Set-Cookie':cookie(request,'',0)});
 if(path!=='/api/auth/login'||request.method!=='POST')return json({error:'ไม่พบคำขอ'},404);
 if(!env.TESTER_ACCESS_CODE||!env.DB)return json({error:'ผู้ดูแลต้องตั้งค่ารหัสผู้ทดลองและฐานข้อมูลก่อน'},503);
 const window=Math.floor(Date.now()/900000),ip=request.headers.get('CF-Connecting-IP')||'local';
 const rateKey=hex(await digest(ip+':'+window));
 const row=await env.DB.prepare('INSERT INTO login_attempts(key,expires,attempts) VALUES(?,?,1) ON CONFLICT(key) DO UPDATE SET attempts=attempts+1 RETURNING attempts').bind(rateKey,Date.now()+900000).first();
 if(row.attempts>5)return json({error:'ลองรหัสหลายครั้งเกินไป กรุณารอ 15 นาที'},429,{'Retry-After':'900'});
 const x=await readJson(request,4096);
 if(typeof x.code!=='string'||!timingSafeEqual(await digest(x.code),await digest(env.TESTER_ACCESS_CODE)))return json({error:'รหัสเข้าใช้ไม่ถูกต้อง'},401);
 await env.DB.prepare('DELETE FROM login_attempts WHERE expires < ?').bind(Date.now()).run();
 const value=String(Date.now()+28800000)+'.'+crypto.randomUUID().replaceAll('-','');
 const signature=hex(await crypto.subtle.sign('HMAC',await key(env.TESTER_ACCESS_CODE),encoder.encode(value)));
 return json({ok:true},200,{'Set-Cookie':cookie(request,value+'.'+signature,28800)});
}
export function loginPage(){return new Response(`<!doctype html><html lang="th"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>เข้าสู่ระบบ · ไตยสยาม</title><style>*{box-sizing:border-box}body{margin:0;background:#edf4f2;color:#163d36;font:16px system-ui;display:grid;place-items:center;min-height:100dvh;padding:24px}main{width:100%;max-width:440px;padding:36px;background:white;border:1px solid #dce8e4;border-radius:24px;box-shadow:0 20px 80px #153d3610}small{letter-spacing:2px;color:#658178}h1{font-size:28px;margin:20px 0 8px}p{color:#63756f;line-height:1.7}label{display:block;margin:28px 0 8px}input,button{font:inherit;width:100%;padding:14px;border-radius:12px}input{border:1px solid #b9cec6}button{background:#12594c;color:white;border:0;margin-top:16px;cursor:pointer}button:disabled{opacity:.5}#error{color:#ad2929;min-height:24px}</style><main><small>TAI SIAM / OPERATIONS</small><h1>ยินดีต้อนรับผู้ทดลอง</h1><p>ใส่รหัสที่ได้รับจากผู้ดูแล เพื่อเข้าใช้งานบัญชีของสาขา</p><form><label for="code">รหัสเข้าใช้</label><input id="code" type="password" autocomplete="current-password" required maxlength="256"><button>เข้าสู่ระบบ</button><p id="error" role="alert"></p></form><p>ข้อมูลบัญชีบันทึกส่วนกลาง ผู้ทดลองที่มีรหัสจะเห็นข้อมูลชุดเดียวกัน</p></main><script>document.querySelector('form').onsubmit=async e=>{e.preventDefault();const b=document.querySelector('button'),err=document.querySelector('#error');b.disabled=true;err.textContent='';try{const r=await fetch('/api/auth/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code:document.querySelector('#code').value})});const x=await r.json();if(!r.ok)throw Error(x.error);location.reload()}catch(e){err.textContent=e.message}finally{b.disabled=false}};</script></html>`,{headers:{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store','X-Frame-Options':'DENY','Referrer-Policy':'same-origin'}})}
