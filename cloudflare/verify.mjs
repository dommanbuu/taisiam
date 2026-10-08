import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import {Miniflare,convertV4MiniflareOptions} from 'miniflare';
import ExcelJS from 'exceljs';
import {parseWorkbook} from './excel.mjs';
const secret=crypto.randomBytes(24).toString('hex');
const mf=new Miniflare(convertV4MiniflareOptions({workers:[{name:"test",modules:true,scriptPath:'.wrangler/test-bundle/worker.js',compatibilityDate:'2026-10-08',compatibilityFlags:['nodejs_compat'],bindings:{TESTER_ACCESS_CODE:secret},d1Databases:['DB'],serviceBindings:{ASSETS:()=>new Response('<h1>assets</h1>')}}]}));
try{
 const db=await mf.getD1Database('DB');await db.exec((await fs.readFile('migrations/0001_pilot.sql','utf8')).replace(/^\uFEFF/,'').replaceAll('\n',' '));
 let cookie='';
 async function call(path,body,{auth=true,headers={},raw=false}={}){return mf.dispatchFetch('https://example.test'+path,{method:body===undefined?'GET':'POST',headers:{...(auth&&cookie?{Cookie:cookie}:{}),...(body===undefined?{}:{'Content-Type':raw?'application/octet-stream':'application/json',Origin:'https://example.test'}),...headers},body:body===undefined?undefined:raw?body:JSON.stringify(body)})}
 for(const path of ['/api/excel/state','/api/excel/backup','/api/excel/source/'+'a'.repeat(64)])assert.equal((await call(path,undefined,{auth:false})).status,401);
 assert.match(await (await call('/',undefined,{auth:false})).text(),/รหัสเข้าใช้/);
 assert.equal((await call('/api/auth/login',{code:'wrong'})).status,401);
 let response=await call('/api/auth/login',{code:secret});assert.equal(response.status,200);cookie=response.headers.get('set-cookie').split(';')[0];assert.match(response.headers.get('set-cookie'),/HttpOnly/);assert.match(response.headers.get('set-cookie'),/Secure/);
 assert.equal((await call('/api/excel/manual',{}, {headers:{Origin:'https://evil.test'}})).status,403);
 const state=async()=>{const r=await call('/api/excel/state');assert.equal(r.status,200,await r.clone().text());return r.json()};
 assert.equal((await state()).records.length,0);
 const sale={revision:0,branch:'TAI SIAM Tachileik',date:'2026-10-01',lines:[{pump:'A',fuel:'ดีเซล',opening:1000,closing:1500,liters:500,price:32,cost:29,legacyR:10000,legacyW:2}]};
 response=await call('/api/excel/manual',sale);assert.equal(response.status,200,await response.clone().text());const record=await response.json();assert.equal(record.sales,16000);
 assert.equal((await call('/api/excel/manual',sale)).status,409);
 let current=await state();assert.equal(current.revision,1);
 const results=await Promise.all([2,3].map(d=>call('/api/excel/manual',{...sale,revision:1,date:'2026-10-0'+d})));assert.deepEqual(results.map(r=>r.status).sort(),[200,409]);assert.equal((await state()).records.length,2);
 for(const kind of ['expenses','purchases']){current=await state();response=await call('/api/excel/book-entry',{revision:current.revision,kind,branch:sale.branch,date:sale.date,description:'test',reference:'test',amount:50,paid:null,fuel:'95',liters:100,price:30});assert.equal(response.status,200,await response.clone().text())}
 response=await call('/api/excel/export',{branch:sale.branch});assert.equal(response.status,200,await response.clone().text());const exported=new Uint8Array(await response.arrayBuffer());assert.ok(exported.length>5000);
 const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(exported);assert.equal(workbook.worksheets.length,8);assert.equal(workbook.getWorksheet('รายการเว็บ').getCell('A1').value,'TAISIAM_LEDGER_V1');assert.equal(workbook.getWorksheet('เงินรับจ่ายรายเดือน').getCell('H5').result??'','');
 workbook.getWorksheet('รายการเว็บ').getCell('M5').value=33;
 const edited=new Uint8Array(await workbook.xlsx.writeBuffer());
 response=await call('/api/excel/import',edited,{raw:true});assert.equal(response.status,200,await response.clone().text());const batch=await response.json();assert.equal(batch.groups[0].sales,16500);
 const original=await call('/api/excel/source/'+batch.id);assert.equal(original.status,200);assert.deepEqual(new Uint8Array(await original.arrayBuffer()),edited);
 current=await state();const commit={batchId:batch.id,branch:sale.branch,revision:current.revision,entries:[{sourceRow:batch.groups[0].sourceRow,date:sale.date,note:'verified test',confirmed:true,action:'create'}]};
 assert.equal((await call('/api/excel/commit',commit)).status,409);
 commit.entries[0].action='update';commit.entries[0].existingId=record.id;response=await call('/api/excel/commit',commit);assert.equal(response.status,200,await response.clone().text());current=await state();assert.equal(current.records.find(r=>r.id===record.id).sales,16500);assert.equal(current.records.find(r=>r.id===record.id).versions.length,1);
 if(process.argv[2]){const bytes=new Uint8Array(await fs.readFile(process.argv[2]));const legacy=await parseWorkbook(bytes);assert.equal(legacy.format,'legacy');assert.equal(legacy.groups.length,14);assert.equal(legacy.groups.reduce((s,g)=>s+g.lines.length,0),58);assert.equal(legacy.emptyPumpRows,72);assert.equal(legacy.groups.reduce((s,g)=>s+g.sales,0),legacy.sourceSummary.Q138);assert.ok(legacy.groups.some(g=>!g.date));response=await call('/api/excel/import',bytes,{raw:true});assert.equal(response.status,200,await response.clone().text());console.log('Legacy workbook matches original totals and split rows (isolated local test).')}
 assert.equal((await call('/api/auth/logout',{})).status,200);
 cookie='taisiam_session=bad';assert.equal((await call('/api/excel/state')).status,401);
 for(let i=0;i<6;i++)response=await call('/api/auth/login',{code:'bad'});assert.equal(response.status,429);
 console.log('PASS: access gate, secure session, CSRF, stale revision, concurrent writes, accounting entries, eight-sheet export, Excel review/update, source archive, invalid cookie, login throttling.');
}finally{await mf.dispose()}
