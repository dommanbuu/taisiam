const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),assert=require('node:assert/strict');
process.env.TAISIAM_DATA_DIR=fs.mkdtempSync(path.join(os.tmpdir(),'taisiam-books-'));
const bridge=require('./bridge.cjs');
const server=require('node:http').createServer((req,res)=>bridge.handle(req,res,new URL(req.url,'http://localhost')));
(async()=>{await new Promise(r=>server.listen(0,'127.0.0.1',r));const base='http://127.0.0.1:'+server.address().port+'/api/excel';
 const post=(p,x)=>fetch(base+p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(x)});
 try{
 const expense={revision:0,kind:'expenses',branch:'TAI SIAM Tachileik',date:'2026-10-02',description:'isolated test',amount:500,paid:null,reference:'TEST'};
 assert.equal((await post('/book-entry',{...expense,paid:600})).status,400);
 assert.equal((await post('/book-entry',{...expense,date:'2026-02-30'})).status,400);
 assert.equal((await post('/book-entry',expense)).status,200);
 assert.equal((await post('/book-entry',expense)).status,409);
 assert.equal((await post('/book-entry',{revision:1,kind:'purchases',branch:expense.branch,date:expense.date,description:'supplier test',fuel:'91',liters:1000,price:20,paid:10000})).status,200);
 const state=await (await fetch(base+'/state')).json();assert.equal(state.expenses[0].paid,null);assert.equal(state.purchases[0].amount,20000);assert.equal(state.records.length,0);
 for(const branch of [expense.branch,'MOK SIO']){const out=await post('/export',{branch});assert.equal(out.status,200);const bytes=Buffer.from(await out.arrayBuffer());assert.equal(bytes.slice(0,2).toString(),'PK');fs.writeFileSync(path.join(process.env.TAISIAM_DATA_DIR,branch==='MOK SIO'?'empty.xlsx':'populated.xlsx'),bytes)}
 console.log('PASS: valid/invalid inputs, revision guard, expense and purchase persistence, partial payments, combined XLSX with no sales, empty-branch export. '+process.env.TAISIAM_DATA_DIR);
 }finally{server.close()}
})().catch(e=>{console.error(e);process.exitCode=1});
