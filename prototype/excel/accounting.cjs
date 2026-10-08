const crypto=require('node:crypto');
const branches=['TAI SIAM Tachileik','TAI SIAM 2','MOK SIO','TAI SIAM Murng Pan'];
function validateEntry(x){
 const fail=msg=>{let e=new Error(msg);e.status=400;throw e};
 if(!['expenses','purchases'].includes(x.kind))fail('หมวดไม่ถูกต้อง');
 if(!branches.includes(x.branch))fail('เลือกสาขาที่ถูกต้อง');
 if(!/^\d{4}-\d{2}-\d{2}$/.test(x.date||'')||!Number.isFinite(Date.parse(x.date))||new Date(x.date).toISOString().slice(0,10)!==x.date)fail('วันที่ไม่ถูกต้อง');
 const text=k=>typeof x[k]==='string'?x[k].trim().slice(0,500):'';
 const number=k=>{if(typeof x[k]!=='number'||!Number.isFinite(x[k])||x[k]<0)fail('ตรวจตัวเลข '+k);return x[k]};
 const r={id:crypto.randomUUID(),kind:x.kind,branch:x.branch,date:x.date,description:text('description'),reference:text('reference'),createdAt:new Date().toISOString()};
 if(!r.description)fail('กรอกรายละเอียด');
 if(x.kind==='purchases'){if(!['ดีเซล','95','91'].includes(x.fuel))fail('เลือกชนิดน้ำมัน');r.fuel=x.fuel;r.liters=number('liters');r.price=number('price');r.amount=Math.round(r.liters*r.price*100)/100}else r.amount=number('amount');
 r.paid=x.paid===null?null:number('paid');if(r.paid!==null&&r.paid>r.amount)fail('ยอดจ่ายมากกว่ายอดรายการ');
 return r;
}
module.exports={validateEntry};
