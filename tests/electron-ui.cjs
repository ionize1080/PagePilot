// Test-only injection into a staging bootstrap. Uses Electron's own IPC/DOM,
// no Playwright/CDP/debugger sockets. Never included in the runtime bootstrap.
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const {BrowserWindow}=require('electron');
const folder=process.env.PAGEPILOT_QA_OUTPUT;fs.mkdirSync(folder,{recursive:true});
const log=m=>{fs.appendFileSync(path.join(folder,'ui-test.log'),m+'\n');console.log(m);};
const wait=ms=>new Promise(r=>setTimeout(r,ms));
const until=async(check,label)=>{for(let i=0;i<180;i++){if(await check())return;await wait(1000);}throw Error('Timed out: '+label);};
(async()=>{let w;await until(async()=>{w=BrowserWindow.getAllWindows()[0];return w&&!w.webContents.isLoading();},'window');
 const js=code=>w.webContents.executeJavaScript(code,true);
 await until(()=>js('!!document.querySelector("#pp-mode")'),'composer');
 const fill=async text=>{await js("(()=>{const t=document.querySelector('#message-input');t.focus();t.select();})()");await w.webContents.insertText(text);await wait(600);};
 const draft='保留中文草稿：侧栏与窗口切换。';await fill(draft);
 for(let i=0;i<6;i++){
  await js('document.querySelector(".pp-window-mode").click()');const mode=i%2===0?'sidebar':'window';
  await until(()=>js(`document.documentElement.dataset.ppMode==='${mode}'`),'mode');
  assert.equal(await js('document.querySelector("#message-input").value'),draft);
  if(mode==='sidebar'){
   assert.equal(await js('getComputedStyle(document.querySelector(".controls.self-start")).display'),'none');
   await js('document.querySelector("#message-input").focus()');w.webContents.sendInputEvent({type:'keyDown',keyCode:'End'});w.webContents.sendInputEvent({type:'keyUp',keyCode:'End'});await w.webContents.insertText('ABC123');
   assert.equal(await js('document.querySelector("#message-input").value'),draft+'ABC123');await fill(draft);
  }
  log('PASS dock/window transition '+(i+1)+' and editable retained draft');
 }
 fs.writeFileSync(path.join(folder,'assistant.png'),(await w.webContents.capturePage()).toPNG());
 await js('document.querySelector("#pp-approval").click()');log('WAIT confirmation: cancel');
 await until(()=>js('!document.querySelector("#pp-approval").disabled'),'cancel');
 assert.equal(await js('document.querySelector("#pp-approval").getAttribute("aria-checked")'),'false');log('PASS cancel preserves manual');
 await js('document.querySelector("#pp-approval").click()');log('WAIT confirmation: accept');
 await until(()=>js('document.querySelector("#pp-approval").getAttribute("aria-checked")==="true"'),'accept');
 assert.equal(await js('getComputedStyle(document.querySelector("#pp-approval")).color'),'rgb(220, 53, 69)');log('PASS automatic red state');
 fs.writeFileSync(path.join(folder,'automatic.png'),(await w.webContents.capturePage()).toPNG());
 await fill('读取当前 Word 测试文档，先查询能力和上下文。只检查数字逻辑：100 到 120、增长 10% 是否一致？不要修改文档，简要回答。');
 w.webContents.sendInputEvent({type:'keyDown',keyCode:'Enter'});w.webContents.sendInputEvent({type:'keyUp',keyCode:'Enter'});
 await until(()=>js('document.body.innerText.includes("读取当前 Word 测试文档，先查询能力和上下文。")&&document.querySelector("#message-input").value===""'),'exact sent prompt');await until(()=>js('document.body.innerText.includes("20%")'),'real Chatbox API result');log('PASS real Chatbox API numerical review');
 fs.writeFileSync(path.join(folder,'api-review.png'),(await w.webContents.capturePage()).toPNG());
 fs.writeFileSync(path.join(folder,'ui-ready'),'ready');
 while(!fs.existsSync(path.join(folder,'finish-ui')))await wait(1000);
 await js('document.querySelector("#pp-approval").click()');await until(()=>js('document.querySelector("#pp-approval").getAttribute("aria-checked")==="false"'),'manual');
 log('PASS manual restored');w.close();
})().catch(e=>{log('FAIL '+e.stack);fs.writeFileSync(path.join(folder,'ui-failed'),e.stack);});
