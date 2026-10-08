// Prototype-only utilities. No credentials, files, or business data leave the tab.
tanks.find(t=>t.id==='t3').warning=8000;
const baseModal=modal;modal=function(title,body){baseModal(title,body);document.querySelectorAll('#modal input[type=file]').forEach(input=>{let button=document.createElement('button');button.type='button';button.textContent='ใช้หลักฐานตัวอย่าง';button.onclick=()=>{let transfer=new DataTransfer();transfer.items.add(new File(['เอกสารสมมติสำหรับทดสอบหน้าจอเท่านั้น'],'หลักฐานตัวอย่าง.txt',{type:'text/plain'}));input.files=transfer.files;input.dispatchEvent(new Event('change',{bubbles:true}));toast('เลือกชื่อไฟล์ตัวอย่างแล้ว · ไม่มีการอัปโหลด')};input.insertAdjacentElement('afterend',button)})};
const baseDailyRows=dailyRows;dailyRows=function(){return baseDailyRows().filter(d=>!d.new)};
const baseReportData=reportData;reportData=function(){let result=baseReportData();for(let row of result){let b=branches.find(b=>b.name===row.branch);let saved=Object.values(days).filter(d=>d.branch===b.id&&!d.new&&(reportPeriod==='day'?d.date===TODAY:d.date.startsWith('2026-09')));if(!saved.length){row.sales=null;row.received=null;row.ending=null;row.status='รอข้อมูล';row.updated='ยังไม่มีข้อมูล'}}return result};
const baseTasks=tasks;tasks=function(){return baseTasks().replaceAll('ยังไม่บันทึก</small><button','ยังไม่ส่งข้อมูล</small><button')};
const baseOverview=overview;overview=function(){let html=baseOverview(),low=tanks.filter(t=>scope(t.branch)&&tankBalance(t).calc!==null&&tankBalance(t).calc<t.warning);return html.replace(`<button onclick="go('tanks')">ตรวจระดับถัง →</button>`,low.length?`<button onclick="tankDetail('${low[0].id}')">ดูถังต่ำ ${low.length} ถัง →</button>`:`<button onclick="go('tanks')">ตรวจระดับถัง →</button>`)};
// In-page registry validates enum explicitly even on providers that do not enforce schema.
// The main registration is optional; no network integration is required.

function setRole(value){if(roles.includes(value)){role=value;branch=['พนักงาน','ผู้จัดการสาขา'].includes(value)?'b1':'all';}}
const accessibleModal=modal;modal=function(title,body){accessibleModal(title,body);document.querySelectorAll('#modal input[type=file]').forEach((input,i)=>{input.id='modal-file-'+i;let label=input.parentElement.querySelector('label');if(label)label.htmlFor=input.id;})};
