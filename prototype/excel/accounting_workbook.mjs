export function addAccountingSheets(wb,payload,overview,detailEnd,summaryTotal){
 const branches=['TAI SIAM Tachileik','TAI SIAM 2','MOK SIO','TAI SIAM Murng Pan'];
 const date=d=>d?new Date(d+'T00:00:00Z'):null;
 const safe=v=>typeof v==='string'&&/^[=+@]/.test(v)?"'"+v:v??null;
 const expense=wb.worksheets.add('ค่าใช้จ่ายรายวัน'),purchase=wb.worksheets.add('ซื้อน้ำมันเข้า'),monthly=wb.worksheets.add('ยอดขายรายเดือน'),cash=wb.worksheets.add('เงินรับจ่ายรายเดือน');
 const specs=[];
 function setup(s,title,note,headers,rows){s.getRange('A2').values=[[title]];s.getRange('A3').values=[[note]];s.getRange(`A4:${String.fromCharCode(64+headers.length)}4`).values=[headers];if(rows.length)s.getRange(`A5:${String.fromCharCode(64+headers.length)}${rows.length+4}`).values=rows.map(r=>r.map(safe));let last=Math.max(29,rows.length+4);specs.push([s,last,headers.length]);return last}
 const expenses=payload.expenses||[];
 const incoming=(payload.records||[]).flatMap(r=>r.lines.filter(l=>l.received>0).map(l=>({date:r.date,branch:r.branch,description:'จาก Excel เดิม — รอตรวจเอกสารซื้อ',fuel:l.fuel,liters:l.received,price:l.purchasePrice??null,paid:null,reference:`${r.source?.filename||''} / E${l.row}`})));
 const purchases=payload.mode==='review'?incoming:(payload.purchases||[]);
 const er=setup(expense,'03 · ค่าใช้จ่ายรายวัน','แบบร่าง | ไม่รวมซื้อน้ำมัน | ยอดจ่ายหมายถึงจ่ายในวันที่รายการ | ช่องว่าง = ยังไม่ทราบ',['วันที่','สาขา','รายละเอียด','ยอดค่าใช้จ่าย','จ่ายจริงวันนั้น','เอกสาร / หมายเหตุ'],expenses.map(r=>[date(r.date),r.branch,r.description,r.amount,r.paid,r.reference]));
 const pr=setup(purchase,'04 · ซื้อน้ำมันเข้า','แบบร่าง | รายการจากไฟล์เดิมคือรับเข้า ยังไม่ยืนยันการซื้อหรือการจ่ายเงินจริง',['วันที่','สาขา','ผู้ขาย / รายละเอียด','น้ำมัน','ลิตรซื้อ','ราคาซื้อ / ลิตร','ยอดซื้อ','จ่ายจริงวันนั้น','เอกสาร / หมายเหตุ'],purchases.map(r=>[date(r.date),r.branch,r.description,({'95':'น้ำมัน 95','91':'น้ำมัน 91'}[r.fuel]||r.fuel),r.liters,r.price,null,r.paid,r.reference]));
 for(let r=5;r<=pr;r++)purchase.getRange(`G${r}`).formulas=[[`=IF(OR(E${r}="",F${r}=""),"",E${r}*F${r})`]];
 const periods=[...new Set([...(payload.records||[]).map(r=>r.date?.slice(0,7)),...expenses.map(r=>r.date.slice(0,7)),...purchases.map(r=>r.date?.slice(0,7)),payload.from?.slice(0,7)].filter(Boolean))].sort();if(!periods.length)periods.push(new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Bangkok'}).format(new Date()).slice(0,7));
 const selected=payload.branch?[payload.branch]:[...new Set([...branches,...(payload.records||[]).map(r=>r.branch)])];
 const monthRows=selected.flatMap(b=>periods.map(m=>[date(m+'-01'),b,null,null,null,null,null,null,null,'ยังไม่ยืนยันครบเดือน']));
 const mr=setup(monthly,'02 · ยอดขายและค่าใช้จ่ายรายเดือน','ยอดขายหลังส่วนลด ยังรวมขายเชื่อ | ซื้อน้ำมันแสดงแยก ไม่ใช่ต้นทุนขาย | วันที่ผิดไม่รวมรายเดือน',['เดือน','สาขา','ดีเซล (ลิตร)','95 (ลิตร)','91 (ลิตร)','ยอดขายก่อนลด','ส่วนลด','ยอดขายหลังลด','ค่าใช้จ่าย','ความครบถ้วน'],monthRows);
 const range=(sheet,col,end)=>`'${sheet}'!$${col}$5:$${col}$${end}`;
 function monthlySum(sheet,col,end,row){return `SUMIFS(${range(sheet,col,end)},${range(sheet,'B',end)},B${row},${range(sheet,'A',end)},">="&A${row},${range(sheet,'A',end)},"<"&EDATE(A${row},1))`}
 for(let r=5;r<monthRows.length+5;r++){
  for(const [col,fuel] of [['C','ดีเซล'],['D','95'],['E','91']])monthly.getRange(`${col}${r}`).formulas=[[`=SUMIFS(${range('รายการเว็บ','K',detailEnd)},${range('รายการเว็บ','B',detailEnd)},B${r},${range('รายการเว็บ','A',detailEnd)},">="&A${r},${range('รายการเว็บ','A',detailEnd)},"<"&EDATE(A${r},1),${range('รายการเว็บ','G',detailEnd)},"${fuel}")`]];
  monthly.getRange(`F${r}`).formulas=[['='+monthlySum('รายการเว็บ','N',detailEnd,r)]];
  monthly.getRange(`G${r}`).formulas=[['='+monthlySum('รายการเว็บ','O',detailEnd,r)]];
  monthly.getRange(`H${r}`).formulas=[[`=F${r}-G${r}`]];
  monthly.getRange(`I${r}`).formulas=[['='+monthlySum('ค่าใช้จ่ายรายวัน','D',er,r)]];
 }
 const cashRows=monthRows.map(r=>[r[0],r[1],null,null,null,null,null,null,'ยังไม่ครบ']);
 const cr=setup(cash,'05 · เงินรับ–เงินจ่ายรายเดือน','แบบร่าง | C รวมเงินรับจากลูกค้าจริงรวมชำระหนี้; D รายรับอื่น; G รายจ่ายอื่นที่ไม่ซ้ำ E/F | ใส่ I = ครบ จึงแสดงสุทธิ',['เดือน','สาขา','รับจากลูกค้าจริง','รายรับอื่นจริง','จ่ายค่าใช้จ่าย','จ่ายซื้อน้ำมัน','รายจ่ายอื่นจริง','เงินรับลบเงินจ่าย','ยืนยันข้อมูล'],cashRows);
 for(let r=5;r<cashRows.length+5;r++){
  cash.getRange(`E${r}`).formulas=[['='+monthlySum('ค่าใช้จ่ายรายวัน','E',er,r)]];
  cash.getRange(`F${r}`).formulas=[['='+monthlySum('ซื้อน้ำมันเข้า','H',pr,r)]];
  cash.getRange(`H${r}`).formulas=[[`=IF(OR(COUNTIFS('ค่าใช้จ่ายรายวัน'!$B$5:$B$${er},B${r},'ค่าใช้จ่ายรายวัน'!$A$5:$A$${er},">="&A${r},'ค่าใช้จ่ายรายวัน'!$A$5:$A$${er},"<"&EDATE(A${r},1),'ค่าใช้จ่ายรายวัน'!$E$5:$E$${er},"")>0,COUNTIFS('ซื้อน้ำมันเข้า'!$B$5:$B$${pr},B${r},'ซื้อน้ำมันเข้า'!$A$5:$A$${pr},">="&A${r},'ซื้อน้ำมันเข้า'!$A$5:$A$${pr},"<"&EDATE(A${r},1),'ซื้อน้ำมันเข้า'!$H$5:$H$${pr},"")>0,I${r}<>"ครบ",C${r}="",D${r}="",G${r}=""),"",C${r}+D${r}-E${r}-F${r}-G${r})`]];
 }
 setup(overview,'ไตยสยาม · บัญชีรวม 5 หมวด',payload.mode==='review'?'แบบร่างสำหรับทดสอบ — ข้อมูลขายเดิมย้อนหลัง 2015 รอตรวจวันที่ ไม่ใช่ยอดปัจจุบัน':'บัญชีรวมจากรายการที่บันทึก — หมวดใหม่เป็นแบบร่างและยังไม่ยืนยันความครบถ้วน',['หัวข้อ','ข้อมูล / วิธีใช้'],[
  ['01 ขายรายวัน','กรอก/ตรวจใน รายการเว็บ; สรุปยอดขาย แสดงรายวัน'],
  ['02 ยอดขายและค่าใช้จ่ายรายเดือน','ชีต ยอดขายรายเดือน ดึงตามสาขาและเดือน'],
  ['03 ค่าใช้จ่ายรายวัน','แบบร่างรอไฟล์จริง; ไม่รวมซื้อน้ำมัน'],
  ['04 ซื้อน้ำมันเข้า','แบบร่างรอไฟล์จริง; ปริมาณรับเข้าไม่ยืนยันการจ่าย'],
  ['05 เงินรับ–เงินจ่ายรายเดือน','กรอกเงินรับจริง ไม่ดึงยอดขายเชื่อมาเป็นเงินสด'],
  ['ยอดขายก่อนปรับทุกแถว',null],['ลิตรขายทุกแถว',null],
  ['ชุดรายการไม่มีวันที่ถูกต้อง',(payload.records||[]).filter(r=>!r.date).length],
  ['วันที่ซ้ำ / ข้อมูลย้อนหลัง','ต้องตรวจเอกสารก่อนยืนยันรายเดือน'],
  ['การนำกลับเข้าเว็บ','รองรับเฉพาะ รายการเว็บ; ชีตอื่นยังไม่ซิงก์กลับ'],
  ['การเติมแถว','แบบร่างค่าใช้จ่าย/ซื้อ: กรอกภายในแถวที่จัดไว้ สูตรรวมถึงแถว '+er+' / '+pr],
  ['สาขา',selected.join(' / ')],
  ['หน่วย','เงินบาทตามไฟล์ต้นทาง; น้ำมันวัดจริง/เกิน/หายเป็นลิตร'],
  ['เงินรับลบเงินจ่าย','ไม่ใช่กำไรสุทธิหรือเงินคงเหลือ; ยังไม่รวมยอดยกมา'],
  ['ข้อมูลจ่ายที่ยังว่าง','ตรวจช่องจ่ายจริงทุกใบก่อนยืนยันความครบถ้วน']
 ]);
 overview.getRange('B10').formulas=[[`='สรุปยอดขาย'!D${summaryTotal}`]];overview.getRange('B11').formulas=[[`='สรุปยอดขาย'!C${summaryTotal}`]];
 for(const [s,last,count] of specs){let end=String.fromCharCode(64+count);s.showGridLines=false;s.getRange(`A1:${end}${last}`).format={font:{name:'Arial',size:11,color:'#253858'},columnWidth:22,rowHeight:27};s.getRange(`A4:${end}4`).format={fill:'#244A81',font:{bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:44};s.getRange('A2').format.font={bold:true,size:16};s.getRange(`A3:${end}3`).merge();s.getRange('A3').format={wrapText:true,rowHeight:48};s.freezePanes.freezeRows(4);s.getRange(`B5:B${last}`).format.columnWidth=30;if(s!==overview){s.getRange(`A5:A${last}`).setNumberFormat(s===expense||s===purchase?'dd/mm/yyyy':'mmm yyyy');s.getRange(`C5:${end}${last}`).setNumberFormat('#,##0.00');s.getRange(`C5:${end}${last}`).format.font.color='#245BCD'}}
 overview.getRange('A1:A29').format.columnWidth=40;overview.getRange('B1:B29').format.columnWidth=100;overview.getRange('B5:B29').format.wrapText=true;overview.getRange('A5:B29').format.autofitRows();overview.getRange('B10:B11').setNumberFormat('#,##0.00');
 expense.getRange(`C5:C${er}`).format.columnWidth=42;purchase.getRange(`C5:C${pr}`).format.columnWidth=48;purchase.getRange(`G5:G${pr}`).format.fill='#EEF2F7';monthly.getRange(`C5:I${mr}`).format.fill='#EEF2F7';cash.getRange(`E5:F${cr}`).format.fill='#EEF2F7';cash.getRange(`H5:H${cr}`).format.fill='#EEF2F7';
 return ['ภาพรวม','ค่าใช้จ่ายรายวัน','ซื้อน้ำมันเข้า','ยอดขายรายเดือน','เงินรับจ่ายรายเดือน'];
}
