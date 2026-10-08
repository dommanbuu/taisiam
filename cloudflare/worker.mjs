import {authenticated,authRoute,json,loginPage} from './auth.mjs';
import {validRecord,commitData,fail} from './domain.mjs';
import accounting from '../prototype/excel/accounting.cjs';
import {parseWorkbook,exportWorkbook} from './excel.mjs';
const mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function readBody(request,limit=10000000){let length=0,chunks=[];if(!request.body)return new Uint8Array();const reader=request.body.getReader();try{while(true){const {value,done}=await reader.read();if(done)break;length+=value.byteLength;if(length>limit){await reader.cancel();fail('ข้อมูลใหญ่เกินขอบเขตที่รองรับ',413)}chunks.push(value)}}finally{reader.releaseLock()}let bytes=new Uint8Array(length),offset=0;for(const c of chunks){bytes.set(c,offset);offset+=c.length}return bytes}
async function readJson(request,limit=1000000){try{return JSON.parse(new TextDecoder().decode(await readBody(request,limit)))}catch(e){if(e.status)throw e;fail('รูปแบบข้อมูลไม่ถูกต้อง')}}
function bounded(value){let s=JSON.stringify(value);if(new TextEncoder().encode(s).length>1500000)fail('ข้อมูลเกินขอบเขตรุ่นทดลอง กรุณาให้ผู้ดูแลขยายฐานข้อมูลก่อน ข้อมูลเดิมยังอยู่',413);return s}
async function load(env){const row=await env.DB.prepare('SELECT data FROM ledger WHERE id=1').first();if(!row)fail('ยังไม่ได้เตรียมฐานข้อมูล',503);return JSON.parse(row.data)}
async function save(env,db,expected){const result=await env.DB.prepare('UPDATE ledger SET data=?, revision=? WHERE id=1 AND revision=?').bind(bounded(db),db.revision,expected).run();if(result.meta.changes!==1)fail('มีผู้อื่นบันทึกก่อนหน้า กรุณาโหลดข้อมูลล่าสุดแล้วลองใหม่',409)}
async function batch(env,id){if(!/^[a-f0-9]{64}$/.test(id||''))fail('รหัสไฟล์ไม่ถูกต้อง');const row=await env.DB.prepare('SELECT data FROM imports WHERE id=?').bind(id).first();if(!row)fail('ไม่พบไฟล์',404);return JSON.parse(row.data)}
function download(bytes,name,type=mime){return new Response(bytes,{headers:{'Content-Type':type,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(name),'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}})}
async function api(request,env,path){
 if(!env.DB)fail('ยังไม่ได้เชื่อมฐานข้อมูล',503);
 const route=path.slice('/api/excel'.length),method=request.method;
 if(route==='/state'&&method==='GET'){const db=await load(env),{results}=await env.DB.prepare('SELECT id,filename,imported_at AS importedAt,group_count AS groups FROM imports ORDER BY imported_at DESC').all();return json({...db,batches:results,storage:'cloudflare',access:'shared-tester-code'})}
 if(route.startsWith('/batch/')&&method==='GET')return json(await batch(env,route.split('/')[2]));
 if(route.startsWith('/source/')&&method==='GET'){const b=await batch(env,route.split('/')[2]);const {results}=await env.DB.prepare('SELECT data FROM source_chunks WHERE id=? ORDER BY part').bind(b.id).all();return download(new Blob(results.map(r=>new Uint8Array(r.data))),b.filename)}
 if(route==='/backup'&&method==='GET')return download(JSON.stringify(await load(env),null,2),'taisiam-ledger-backup.json','application/json; charset=utf-8');
 if(route==='/import'&&method==='POST'){
  const bytes=await readBody(request);if(bytes[0]!==80||bytes[1]!==75)fail('กรุณาเลือกไฟล์ .xlsx');
  const id=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),b=>b.toString(16).padStart(2,'0')).join('');
  const existing=await env.DB.prepare('SELECT data FROM imports WHERE id=?').bind(id).first();if(existing)return json(JSON.parse(existing.data));
  const b=await parseWorkbook(bytes);Object.assign(b,{id,sha256:id,filename:decodeURIComponent(request.headers.get('X-Filename')||'excel-report.xlsx').split(/[\\/]/).pop().slice(0,150),importedAt:new Date().toISOString()});
  const statements=[env.DB.prepare('INSERT OR IGNORE INTO imports VALUES(?,?,?,?,?)').bind(id,b.filename,b.importedAt,b.groups.length,bounded(b))];
  for(let i=0;i<bytes.length;i+=256000)statements.push(env.DB.prepare('INSERT OR IGNORE INTO source_chunks VALUES(?,?,?)').bind(id,i/256000,bytes.slice(i,i+256000).buffer));
  await env.DB.batch(statements);return json(b);
 }
 if(method==='POST'&&['/manual','/book-entry','/commit'].includes(route)){
  const x=await readJson(request),db=await load(env),revision=db.revision;if(x.revision!==revision)fail('ข้อมูลเปลี่ยน กรุณาโหลดข้อมูลล่าสุด',409);let result;
  if(route==='/commit')result=commitData(x,db,await batch(env,x.batchId));
  else if(route==='/book-entry'){result=accounting.validateEntry(x);db[result.kind].push(result);db.revision++;db.history.unshift({time:result.createdAt,action:result.kind==='expenses'?'บันทึกค่าใช้จ่าย':'บันทึกซื้อน้ำมัน',detail:result.branch+' / '+result.date})}
  else{let r=validRecord({branch:x.branch,date:x.date,lines:x.lines,note:typeof x.note==='string'?x.note.slice(0,2000):'',source:{filename:'กรอกผ่านเว็บ'},status:'รอตรวจ',id:crypto.randomUUID(),createdAt:new Date().toISOString()});for(const l of r.lines){if(l.opening==null||l.closing==null)fail('ต้องกรอกมิเตอร์เปิดและปิด');l.liters=l.closing-l.opening}validRecord(r);if(db.records.some(a=>a.branch===r.branch&&a.date===r.date))fail('วันและสาขานี้มีข้อมูลแล้ว ไม่บันทึกซ้ำ',409);db.records.push(r);db.revision++;db.history.unshift({time:r.createdAt,action:'บันทึกจากเว็บ',detail:r.branch+' / '+r.date});result=r}
  await save(env,db,revision);return json(result);
 }
 if(route==='/export'&&method==='POST'){
  const x=await readJson(request);let payload;
  if(x.batchId){const b=await batch(env,x.batchId);payload={mode:'review',branch:x.branch,records:b.groups.map(g=>({...g,branch:g.branch||x.branch||'TAI SIAM Tachileik',status:'รอตรวจ',source:{filename:b.filename,sheet:b.sheet,row:g.sourceRow},note:g.issues.join(' / ')}))}}
  else{const db=await load(env),filter=r=>(!x.branch||r.branch===x.branch)&&(!x.from||r.date>=x.from)&&(!x.to||r.date<=x.to);payload={mode:'ledger',branch:x.branch,from:x.from,to:x.to,records:db.records.filter(filter),expenses:db.expenses.filter(filter),purchases:db.purchases.filter(filter)}}
  return download(await exportWorkbook(payload),'taisiam-accounting.xlsx');
 }
 fail('ไม่พบคำขอ',404);
}
export default {
 async fetch(request,env){const url=new URL(request.url),path=url.pathname;try{
  if(path.startsWith('/api/')){
   if(request.headers.get('Origin')&&request.headers.get('Origin')!==url.origin||request.headers.get('Sec-Fetch-Site')==='cross-site')fail('ไม่อนุญาตคำขอข้ามเว็บไซต์',403);
   if(path.startsWith('/api/auth/'))return await authRoute(request,env,path,readJson);
   if(!await authenticated(request,env))return json({error:'กรุณาเข้าสู่ระบบด้วยรหัสผู้ทดลอง',code:'LOGIN_REQUIRED'},401);
   return await api(request,env,path);
  }
  if(!await authenticated(request,env))return loginPage();
  const response=await env.ASSETS.fetch(request);const headers=new Headers(response.headers);headers.set('Cache-Control','no-store');headers.set('X-Frame-Options','DENY');headers.set('X-Content-Type-Options','nosniff');return new Response(response.body,{status:response.status,headers});
 }catch(e){if(!e.status)console.error(JSON.stringify({event:'request_failed',path,name:e.name}));return json({error:e.status?e.message:'ดำเนินการไม่สำเร็จ กรุณาลองใหม่หรือติดต่อผู้ดูแล'},e.status||500)}}
};
