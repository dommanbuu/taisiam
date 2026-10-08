let cloudAccess=false;
// Cloud API sessions expire after eight hours; reload to the server login screen.
const sessionFetch=window.fetch.bind(window);
window.fetch=async function(input,options){const response=await sessionFetch(input,options);const path=typeof input==='string'?input:'';if(path.startsWith('/api/excel')&&response.status===401){xlState=null;xlBatch=null;location.reload()}return response};
const cloudLayout=xlLayout;
xlLayout=function(title,sub,body){let html=cloudLayout(title,sub,body);if(xlState?.storage==='cloudflare')html=html.replace('ข้อมูล Excel จริงบันทึกอยู่ในเครื่องนี้ เปิดใหม่ข้อมูลยังอยู่','ข้อมูลบัญชีบันทึกส่วนกลางบน Cloudflare ผู้ทดลองใช้ข้อมูลชุดเดียวกัน').replace('อ่านเฉพาะในเครื่องนี้','อ่านไฟล์เพื่อตรวจบน Cloudflare');return html};
const cloudRender=render;
render=function(){cloudRender();if(cloudAccess||xlState?.storage==='cloudflare'){const banner=document.querySelector('.demo');if(banner){banner.replaceChildren();const label=document.createElement('span');label.textContent='กลุ่มผู้ทดลอง · บัญชีบันทึกบน Cloudflare · เมนูถังและภาพรวมยังเป็นข้อมูลสาธิต · ตัวเลือกบทบาทเป็นตัวอย่างหน้าจอ';const button=document.createElement('button');button.textContent='ออกจากระบบ';button.onclick=async()=>{await fetch('/api/auth/logout',{method:'POST'});location.reload()};banner.append(label,button)}}};
render();

fetch('/api/auth/session').then(r=>r.ok?r.json():null).then(s=>{cloudAccess=!!s?.authenticated;if(cloudAccess)render()}).catch(()=>{});
