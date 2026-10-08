const fs=require('fs'),vm=require('vm'),assert=require('assert');
const nodes=new Map();const node=()=>({innerHTML:'',textContent:'',style:{},close(){},showModal(){}});
const ctx=vm.createContext({console,assert,document:{querySelector(s){if(!nodes.has(s))nodes.set(s,node());return nodes.get(s)},querySelectorAll(){return []}},window:{addEventListener(){},scrollTo(){}},location:{hash:''},setTimeout:fn=>fn(),FormData,Date});
for(const f of ['app.js','workflows.js','management.js','refinements.js'])vm.runInContext(fs.readFileSync('dist/'+f,'utf8'),ctx,{filename:f});
vm.runInContext(`
assert.equal(balance(days['b1|'+TODAY].lines[0],TODAY),12450);
assert.equal(total('sales'),270000);
setRole('Admin');assert(!allowed().includes('finance'));assert(allowed().includes('admin'));
setRole('พนักงาน');assert.equal(branch,'b1');assert(!scope('b2'));assert(!canApprove());
setRole('เจ้าของ');branch='all';reportPeriod='month';assert(reportData().every(r=>r.ending===null));
branch='b4';getDay();reportPeriod='day';assert.equal(reportData()[0].sales,null);
branch='b1';day=TODAY;let prior=getDay();let oldCount=prior.lines.length;
tanks.push({id:'test-tank',code:'TK-NEW',name:'Test',branch:'b1',fuel:'f1',opening:1000,start:TODAY,capacity:20000,warning:2000,active:true});
assert.equal(getDay().lines.length,oldCount);day='2026-09-28';assert.equal(getDay().lines.length,oldCount+1);
let advance=requests.find(r=>r.id==='REQ-002');advance.actual=1850;advance.returned=150;advance.evidence='sample';postExpense(advance);postExpense(advance);
assert.equal(finance.filter(f=>f.source===advance.id).length,1);assert.equal(finance.find(f=>f.source===advance.id).amount,1850);
let rendered=0,handlers=0;
for(const r of roles){setRole(r);for(const m of allowed()){route=m;render();let html=document.querySelector('#app').innerHTML;assert(html.length>100);for(const match of html.matchAll(/on(?:click|change|submit|input)="([^"]*)"/g)){new Function('event',match[1].replaceAll('&quot;','"'));handlers++}rendered++}}
for(const type of adminSections.map(x=>x[0])){setRole('Admin');adminForm(type);assert(document.querySelector('#modal').innerHTML.includes('ตรวจทาน'))}
console.log('PASS: oil balance, role scopes, monthly ending, missing data, tank snapshots, settlement idempotency; '+rendered+' role/screens, '+handlers+' inline handlers, 7 Admin forms');
`,ctx);
