'use strict';
const automatic=document.getElementById('automatic'),rules=document.getElementById('rules'),status=document.getElementById('status'),save=document.getElementById('save');
window.pagepilotSettings.read().then(v=>{automatic.checked=v.autoApprove;rules.value=v.headingRules;status.textContent='设置已载入';save.disabled=false;}).catch(e=>{status.textContent=e.message;});
save.addEventListener('click',async()=>{save.disabled=true;try{await window.pagepilotSettings.write({autoApprove:automatic.checked,headingRules:rules.value});status.textContent='已保存 · 当前 Word 会话生效';}catch(e){status.textContent=e.message;}finally{save.disabled=false;}});
document.getElementById('provider').addEventListener('click',()=>window.pagepilotSettings.provider().catch(e=>{status.textContent=e.message;}));
