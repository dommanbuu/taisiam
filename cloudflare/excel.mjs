import ExcelJS from 'exceljs/excel.js';
import {fail,validDate} from './domain.mjs';
const number=v=>typeof v==='number'&&Number.isFinite(v)?v:null;
function value(cell){if(cell.isMerged&&cell.master.address!==cell.address)return null;const v=cell.value;if(v&&typeof v==='object'&&!(v instanceof Date)){if('formula'in v||'sharedFormula'in v)return v.result??null;if(v.richText)return v.richText.map(x=>x.text).join('');if(v.text)return v.text;return null}return v}
function raw(cell){return {value:cell.formula?'='+cell.formula:value(cell),cached:value(cell)}}
function day(v){if(v instanceof Date)return v.toISOString().slice(0,10);if(typeof v==='string'){if(validDate(v))return v;let m=v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(m){let y=+m[3];if(y>2400)y-=543;let d=`${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;if(validDate(d))return d}}return null}
function zipGuard(bytes){const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let total=0,count=0;for(let i=0;i+46<=v.byteLength;i++){if(v.getUint32(i,true)===0x02014b50){total+=v.getUint32(i+24,true);count++;i+=45+v.getUint16(i+28,true)+v.getUint16(i+30,true)+v.getUint16(i+32,true)}}if(!count||total>40000000||count>2000)fail('ไฟล์ Excel ใหญ่เกินขอบเขตที่รองรับ')}
const fuel={A:'ดีเซล',B:'ดีเซล',C:'95',D:'91'};
export async function parseWorkbook(bytes){
 zipGuard(bytes);const wb=new ExcelJS.Workbook();try{await wb.xlsx.load(bytes)}catch{fail('อ่าน Excel ไม่สำเร็จ กรุณาตรวจรูปแบบไฟล์')}
 const web=wb.getWorksheet('รายการเว็บ');let sheet,groups=[],empty=0;
 const isWeb=web&&value(web.getCell('A1'))==='TAISIAM_LEDGER_V1';
 if(isWeb)sheet=web;else{const matches=wb.worksheets.filter(s=>value(s.getCell('B1'))==='บัญชีขายรายวัน'&&value(s.getCell('C3'))==='เปิด'&&value(s.getCell('D3'))==='ปิด');if(matches.length!==1)fail('รองรับบัญชีขายรายวันรูปแบบเดิม หรือไฟล์ที่เว็บส่งออก');sheet=matches[0]}
 if(sheet.rowCount>10000||sheet.columnCount>100)fail('จำนวนแถวหรือคอลัมน์เกินกำหนด');
 const cell=(r,c)=>sheet.getRow(r).getCell(c),val=(r,c)=>value(cell(r,c)),n=(r,c)=>number(val(r,c));
 const rawRow=(r,end)=>Object.fromEntries(Array.from({length:end},(_,i)=>i+1).filter(c=>val(r,c)!=null).map(c=>[cell(r,c).address.replace(/\d/g,''),raw(cell(r,c))]));
 if(isWeb){const map=new Map();for(let r=5;r<=sheet.rowCount;r++){
  if(!val(r,6))continue;if([1,2,3,4,5,6,7,8,9,10,12,13,15,16,17,18,19,20,21].some(c=>cell(r,c).formula))fail(`แถว ${r}: ช่องกรอกต้องเป็นค่า ไม่ใช่สูตร`);
  const rid=String(val(r,3)||''),date=day(val(r,1)),branch=String(val(r,2)||'').trim(),key=JSON.stringify([rid,branch,date]);
  if(!map.has(key))map.set(key,{recordId:rid.startsWith('review:')?'':rid,branch,date,rawDate:String(val(r,1)||''),sourceRow:r,lines:[],issues:[]});
  const g=map.get(key),opening=n(r,8),closing=n(r,9),override=n(r,10),liters=override??(opening!==null&&closing!==null?closing-opening:null);
  const l={row:r,lineId:String(val(r,4)||r),pump:String(val(r,6)),fuel:String(val(r,7)),opening,closing,liters,split:val(r,5)==='แยกต้นทุน',raw:rawRow(r,24)};
  for(const [k,c]of Object.entries({cost:12,price:13,discount:15,credit:16,debtPaid:17,received:18,legacyR:19,legacyV:20,legacyW:21}))l[k]=n(r,c);
  g.lines.push(l);if(liters===null)g.issues.push(`แถว ${r}: ไม่มีมิเตอร์หรือยอดลิตรจัดสรร`);if(override!==null)g.issues.push(`แถว ${r}: ใช้ลิตรจัดสรรแทนผลต่างมิเตอร์ กรุณาตรวจ`);
 }groups=[...map.values()]}else{
 let g,last;for(let r=8;r<=sheet.rowCount;r++){
  if(val(r,1)!=null){g={sourceRow:r,rawDate:String(val(r,1)),date:day(val(r,1)),lines:[],issues:[],legacyDaily:n(r,30)};groups.push(g);last=null}if(!g)continue;
  let pump=val(r,2);if(pump!=null&&!fuel[pump])continue;
  const split=pump==null&&!!last&&[11,12,13].some(c=>!cell(r,c).formula&&n(r,c)!==null);
  const present=[3,4,5,18,20,21,22,23,24].some(c=>cell(r,c).value!=null&&!cell(r,c).formula);
  if(!present&&!split){if(pump)empty++;continue}if(pump)last=pump;pump=pump||last;if(!pump)continue;
  const l={row:r,pump,fuel:fuel[pump],split,raw:rawRow(r,30)};
  for(const [k,c]of Object.entries({opening:3,closing:4,liters:{A:11,B:11,C:12,D:13}[pump],received:5,purchasePrice:6,cost:14,price:16,legacySales:17,legacyR:18,legacyS:19,discount:20,credit:21,legacyV:22,legacyW:23,debtPaid:24}))l[k]=n(r,c);
  g.lines.push(l);if(l.liters===null)g.issues.push(`แถว ${r}: สูตรไม่มีค่าลิตรล่าสุด ให้เปิดและบันทึกใน Excel ก่อน`);
  if(!split&&[l.opening,l.closing,l.liters].every(v=>v!==null)&&Math.abs(l.closing-l.opening-l.liters)>.001)g.issues.push(`แถว ${r}: ลิตรตามมิเตอร์ต่างจากลิตรลงบัญชี อาจมีการแยกต้นทุน`);
  if([l.liters,l.price,l.legacySales].every(v=>v!==null)&&Math.abs(l.liters*l.price-l.legacySales)>.01)g.issues.push(`Q${r}: ยอดสูตรเดิมต่างจากลิตร × ราคาขาย`);
  if([l.legacyR,l.legacyV,l.legacyW].some(v=>v))g.issues.push(`แถว ${r}: วัดจริง/เกิน/หายเป็นลิตร ไม่หักจากยอดเงิน`);
 }groups=groups.filter(g=>g.lines.length)}
 const seen=new Set();let prev;for(const g of groups){if(!g.date)g.issues.push('วันที่ไม่ถูกต้อง ต้องระบุวันที่จริง');if(seen.has(g.date))g.issues.push('วันที่ซ้ำในไฟล์ ต้องตรวจรอบการขาย');if(g.date&&prev&&g.date<prev)g.issues.push('วันที่ย้อนกลับเมื่อเทียบกับบล็อกก่อนหน้า');if(g.date){seen.add(g.date);prev=g.date}g.sales=Math.round(g.lines.reduce((s,l)=>s+(l.liters||0)*(l.price||0),0)*100)/100;g.liters=g.lines.reduce((s,l)=>s+(l.liters||0),0)}
 return {format:isWeb?'web':'legacy',sheet:sheet.name,groups,emptyPumpRows:empty,sourceSummary:isWeb?{}:Object.fromEntries(Array.from({length:30},(_,i)=>i+1).filter(c=>val(138,c)!=null).map(c=>[cell(138,c).address,val(138,c)])),opening:isWeb?[]:[['ดีเซล','E4'],['95','E6'],['91','E7']].map(([fuel,c])=>({fuel,cell:c,liters:value(sheet.getCell(c))})),warnings:['ตรวจชื่อสาขา วันที่ และค่าเดิม–ค่าใหม่ก่อนยืนยัน','นำกลับเข้าเว็บเฉพาะชีต รายการเว็บ; ชีตหมวดอื่นยังไม่ซิงก์กลับ','ยอดวัดจริง/เกิน/หายเป็นลิตร ไม่ปนกับยอดเงิน','การลบแถวใน Excel ไม่ลบรายการบนเว็บอัตโนมัติ']};
}
const dateValue=d=>validDate(d||'')?new Date(d+'T00:00:00Z'):null;
const money=n=>Math.round(n*100)/100;
export async function exportWorkbook(payload){
 const wb=new ExcelJS.Workbook();wb.creator='TAI SIAM';wb.calcProperties.fullCalcOnLoad=true;
 function sheet(name,title,headers,header=4){let s=wb.addWorksheet(name,{views:[{state:'frozen',ySplit:header}]});s.getCell('A2').value=title;s.getCell('A2').font={size:16,bold:true};s.getRow(header).values=headers;headers.forEach((_,i)=>{s.getColumn(i+1).width=i===1?32:22});s.getRow(header).height=34;s.getRow(header).eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF244A81'}};c.font={bold:true,color:{argb:'FFFFFFFF'}};c.alignment={wrapText:true}});return s}
 function formula(s,address,f,result){s.getCell(address).value={formula:f,result:result??0}}
 const overview=sheet('ภาพรวม','ไตยสยาม · บัญชีรวม 5 หมวด',['หัวข้อ','ข้อมูล / วิธีใช้']);
 const summary=sheet('สรุปยอดขาย','ยอดขายก่อนปรับ ไม่ใช่เงินรับหรือกำไร',['วันที่','สาขา','ลิตรขาย','ยอดขายก่อนปรับ','ส่วนลด','ขายเชื่อ','รับชำระหนี้','รับเข้า (ลิตร)','สถานะ','วันที่ต้นทาง','หมายเหตุ'],5);
 const detail=sheet('รายการเว็บ','แก้ช่องกรอกแล้วนำกลับมาตรวจผลต่าง รักษารหัสรายการและรหัสแถว',['วันที่','สาขา','รหัสรายการ','รหัสแถว','ชนิดแถว','หัวจ่าย','ชนิดน้ำมัน','มิเตอร์เปิด','มิเตอร์ปิด','ลิตรจัดสรร (ถ้ามี)','ลิตรขาย','ต้นทุน/ลิตร','ราคาขาย/ลิตร','ยอดขายก่อนปรับ','ส่วนลด (บาท)','ขายเชื่อ (บาท)','รับชำระหนี้ (บาท)','รับเข้า (ลิตร)','วัดจริง (ลิตร)','เกิน (ลิตร)','หาย (ลิตร)','สถานะ','ต้นทาง','หมายเหตุ']);detail.getCell('A1').value='TAISIAM_LEDGER_V1';
 const source=sheet('ตรวจต้นทาง','ข้อมูลต้นทางสำหรับตรวจสอบ',['ไฟล์','ชีต','เซลล์','ค่า / สูตรเดิม (ข้อความ)','ค่าที่ Excel บันทึก','วันที่']);
 const expenses=sheet('ค่าใช้จ่ายรายวัน','แบบร่างค่าใช้จ่าย ไม่รวมซื้อน้ำมัน',['วันที่','สาขา','รายละเอียด','ยอดรายการ','จ่ายจริงวันนั้น','เอกสาร']);
 const purchases=sheet('ซื้อน้ำมันเข้า','แบบร่างซื้อน้ำมันเข้า',['วันที่','สาขา','ผู้ขาย / รายละเอียด','น้ำมัน','ลิตร','ราคา/ลิตร','ยอดรายการ','จ่ายจริงวันนั้น','เอกสาร']);
 const monthly=sheet('ยอดขายรายเดือน','ยอดขายและค่าใช้จ่าย ไม่ใช่กำไรสุทธิ',['เดือน','สาขา','ดีเซล (ลิตร)','95 (ลิตร)','91 (ลิตร)','ยอดขายก่อนปรับ','ค่าใช้จ่าย','ซื้อน้ำมัน']);
 const cash=sheet('เงินรับจ่ายรายเดือน','กรอกเงินรับจริง และยืนยันความครบถ้วนก่อนดูผลต่าง',['เดือน','สาขา','รับจากลูกค้าจริง','รับอื่น','จ่ายค่าใช้จ่าย','จ่ายซื้อน้ำมัน','จ่ายอื่น','เงินรับลบเงินจ่าย','ยืนยันครบ (พิมพ์ ครบ)']);
 let dr=5,sr=6,rr=5;const records=payload.records||[],es=payload.expenses||[],ps=payload.purchases||[];
 for(const r of records){let start=dr;for(const [i,l] of r.lines.entries()){
 const override=l.split||l.opening==null||l.closing==null||Math.abs(l.closing-l.opening-l.liters)>.00001?l.liters:null;
 detail.getRow(dr).values=[dateValue(r.date),r.branch,r.id||r.recordId||'review:'+r.sourceRow,l.lineId||String(l.row||i+1),l.split?'แยกต้นทุน':'หัวจ่าย',l.pump,l.fuel,l.opening,l.closing,override,null,l.cost,l.price,null,l.discount,l.credit,l.debtPaid,l.received,l.legacyR,l.legacyV,l.legacyW,r.status||'รอตรวจ',r.source?.filename||'',r.note||''];
 formula(detail,'K'+dr,`IF(J${dr}="",IF(OR(H${dr}="",I${dr}=""),"",I${dr}-H${dr}),J${dr})`,l.liters??'');formula(detail,'N'+dr,`IF(OR(K${dr}="",M${dr}=""),"",K${dr}*M${dr})`,l.liters==null||l.price==null?'':l.liters*l.price);
 for(const [col,raw]of Object.entries(l.raw||{}))source.getRow(rr++).values=[r.source?.filename||'',r.source?.sheet||'',col+(l.row||''),raw.value instanceof Object?JSON.stringify(raw.value):raw.value,raw.cached,dateValue(r.date)];dr++}
 summary.getRow(sr).values=[dateValue(r.date),r.branch,null,null,null,null,null,null,r.status||'รอตรวจ',r.rawDate||r.date||'',r.note||''];
 for(const [to,from,key]of [['C','K','liters'],['D','N','sales'],['E','O','discount'],['F','P','credit'],['G','Q','debtPaid'],['H','R','received']])formula(summary,to+sr,`SUM('รายการเว็บ'!${from}${start}:${from}${dr-1})`,r.lines.reduce((s,l)=>s+(key==='sales'?(l.liters||0)*(l.price||0):l[key]||0),0));sr++}
 summary.getCell('B'+sr).value='รวม';for(const c of ['C','D','E','F','G','H'])formula(summary,c+sr,sr>6?`SUM(${c}6:${c}${sr-1})`:'0',Array.from({length:sr-6},(_,i)=>summary.getCell(c+(i+6)).value.result).reduce((s,n)=>s+n,0));
 es.forEach((r,i)=>expenses.getRow(i+5).values=[dateValue(r.date),r.branch,r.description,r.amount,r.paid,r.reference]);
 ps.forEach((r,i)=>{purchases.getRow(i+5).values=[dateValue(r.date),r.branch,r.description,r.fuel,r.liters,r.price,null,r.paid,r.reference];formula(purchases,'G'+(i+5),`E${i+5}*F${i+5}`,r.amount)});
 const er=Math.max(es.length+4,104),pr=Math.max(ps.length+4,104),last=Math.max(dr-1,5);
 for(let i=ps.length+5;i<=pr;i++)formula(purchases,'G'+i,`IF(OR(E${i}="",F${i}=""),"",E${i}*F${i})`,'');
 const pairs=new Map();for(const r of [...records,...es,...ps])if(validDate(r.date||''))pairs.set(r.branch+'|'+r.date.slice(0,7),[r.branch,r.date.slice(0,7)]);
 if(!pairs.size)pairs.set('empty',[payload.branch||'TAI SIAM Tachileik',(payload.from||new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Bangkok'}).format(new Date())).slice(0,7)]);
 let mr=5;for(const [branch,month]of pairs.values()){
 monthly.getRow(mr).values=[dateValue(month+'-01'),branch];cash.getRow(mr).values=[dateValue(month+'-01'),branch];
 const ls=records.filter(r=>r.branch===branch&&r.date?.startsWith(month)).flatMap(r=>r.lines),e=es.filter(r=>r.branch===branch&&r.date.startsWith(month)),p=ps.filter(r=>r.branch===branch&&r.date.startsWith(month));
 const criteria=(s,end)=>`'${s}'!$B$5:$B$${end},B${mr},'${s}'!$A$5:$A$${end},">="&A${mr},'${s}'!$A$5:$A$${end},"<"&EDATE(A${mr},1)`;
 for(const [c,f]of [['C','ดีเซล'],['D','95'],['E','91']])formula(monthly,c+mr,`SUMIFS('รายการเว็บ'!$K$5:$K$${last},${criteria('รายการเว็บ',last)},'รายการเว็บ'!$G$5:$G$${last},"${f}")`,ls.filter(l=>l.fuel===f).reduce((s,l)=>s+l.liters,0));
 formula(monthly,'F'+mr,`SUMIFS('รายการเว็บ'!$N$5:$N$${last},${criteria('รายการเว็บ',last)})`,ls.reduce((s,l)=>s+l.liters*l.price,0));
 for(const [c,s,end,src,key]of [['G','ค่าใช้จ่ายรายวัน',er,e,'D'],['H','ซื้อน้ำมันเข้า',pr,p,'G']])formula(monthly,c+mr,`SUMIFS('${s}'!$${key}$5:$${key}$${end},${criteria(s,end)})`,src.reduce((s,r)=>s+r.amount,0));
 formula(cash,'E'+mr,`SUMIFS('ค่าใช้จ่ายรายวัน'!$E$5:$E$${er},${criteria('ค่าใช้จ่ายรายวัน',er)})`,e.reduce((s,r)=>s+(r.paid||0),0));formula(cash,'F'+mr,`SUMIFS('ซื้อน้ำมันเข้า'!$H$5:$H$${pr},${criteria('ซื้อน้ำมันเข้า',pr)})`,p.reduce((s,r)=>s+(r.paid||0),0));
 formula(cash,'H'+mr,`IF(OR(C${mr}="",D${mr}="",G${mr}="",I${mr}<>"ครบ",COUNTIFS(${criteria('ค่าใช้จ่ายรายวัน',er)},'ค่าใช้จ่ายรายวัน'!$E$5:$E$${er},"")>0,COUNTIFS(${criteria('ซื้อน้ำมันเข้า',pr)},'ซื้อน้ำมันเข้า'!$H$5:$H$${pr},"")>0),"",C${mr}+D${mr}-E${mr}-F${mr}-G${mr})`,'');mr++}
 [['ขายรายวัน','แก้ช่องกรอกใน รายการเว็บ แล้วนำกลับมาตรวจบนเว็บ'],['ยอดขายและค่าใช้จ่ายรายเดือน','ชีต ยอดขายรายเดือน รวมตามสาขาและเดือน'],['ค่าใช้จ่ายรายวัน','แบบร่างรอไฟล์จริง ไม่รวมซื้อน้ำมัน'],['ซื้อน้ำมันเข้า','แบบร่างรอไฟล์จริง'],['เงินรับ–เงินจ่ายรายเดือน','กรอกเงินรับจริง; ไม่ใช้ยอดขายแทนเงินสด'],['การนำเข้า','เฉพาะ รายการเว็บ; หมวดอื่นยังไม่ซิงก์กลับ'],['สถานะ',payload.mode==='review'?'รอตรวจเอกสาร ไม่ใช่ข้อมูลยืนยัน':'รายการที่บันทึกบนเว็บ'],['วันที่ไม่ถูกต้อง',records.filter(r=>!validDate(r.date||'')).length],['ช่องกรอกเพิ่ม',`ค่าใช้จ่ายถึงแถว ${er}; ซื้อถึงแถว ${pr}; สูตรเดือนอ้างอิงถึงแถวเหล่านี้`],['หน่วย','บาทตามไฟล์ต้นทาง; วัดจริง/เกิน/หายเป็นลิตร']].forEach((r,i)=>overview.getRow(i+5).values=r);
 overview.getColumn(1).width=34;overview.getColumn(2).width=90;
 for(const s of wb.worksheets){s.eachRow((r,i)=>{if(i>5){r.alignment={vertical:'top',wrapText:true};r.eachCell(c=>{if(c.value instanceof Date)c.numFmt='dd/mm/yyyy';else if(typeof c.value==='number'||c.formula)c.numFmt='#,##0.00'})}})}
 for(const s of [detail,expenses,purchases,source,monthly,cash])s.getColumn(1).numFmt=s===monthly||s===cash?'mmm yyyy':'dd/mm/yyyy';
 return new Uint8Array(await wb.xlsx.writeBuffer());
}


