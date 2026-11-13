// Actual composer sends through the pinned Chatbox SDK, real API and stdio MCP.
// The test profile contains no copied chat history. Credentials stay in memory.
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const {app,BrowserWindow,ipcMain,dialog}=require('electron');
const {execFileSync}=require('node:child_process');
const root=path.resolve(__dirname,'..'),out=path.join(root,'.temp/workspace-send-live');
fs.mkdirSync(out,{recursive:true});
const log=value=>fs.appendFileSync(path.join(out,'progress.log'),value+'\n');
fs.writeFileSync(path.join(out,'progress.log'),'');
const configured=JSON.parse(fs.readFileSync(process.env.PAGEPILOT_QA_KEY_PATH,'utf8').replace(/^\uFEFF/,''));
const selected=configured.settings.defaultChatModel,provider=configured.settings.providers[selected.provider];
assert(provider?.apiKey&&selected.model,'A configured API provider and model are required');
const model={provider:selected.provider,model:selected.model},secret=provider.apiKey;
const strip=value=>JSON.parse(JSON.stringify(value).split(secret).join(''));
const profile=path.join(process.env.PAGEPILOT_PROFILE,'Chatbox-1.20.3');
fs.mkdirSync(profile,{recursive:true});
fs.writeFileSync(path.join(profile,'config.json'),JSON.stringify({lastShownAboutDialogVersion:'1.20.3',settings:{...strip(configured.settings),
 providers:{[selected.provider]:strip(provider)},defaultChatModel:selected,language:'zh-Hans',languageInited:true,
 autoGenerateTitle:false,autoCompaction:false,allowReportingAndTracking:false,spellCheck:false,autoUpdate:false,mcp:{servers:[],enabledBuiltinServers:[]},
}}));
// Inject only the selected credential into settings reads; strip it on all writes.
const register=ipcMain.handle.bind(ipcMain);
ipcMain.handle=(channel,fn)=>register(channel,async(event,...args)=>{
 if(channel==='setStoreValue'&&args[0]==='settings')args[1]=JSON.stringify(strip(JSON.parse(args[1])));
 if(channel==='setAllStoreValues')args[0]=JSON.stringify(strip(JSON.parse(args[0])));
 const value=await fn(event,...args);
 if(channel==='getSettings'||channel==='getStoreValue'&&args[0]==='settings')return {...value,providers:{...value?.providers,[selected.provider]:{...provider,apiKey:secret}}};
 return value;
});
const renderer=path.join(root,'dist/Chatbox/native/dist/renderer/js/index.TlLbGKvw.js');
const original=fs.readFileSync(renderer);
let restored=false;
const restore=()=>{if(!restored){restored=true;fs.writeFileSync(renderer,original);}};
app.on('will-quit',restore);app.on('quit',restore);process.on('exit',restore);
fs.appendFileSync(renderer,`\nwindow.pagepilotSendQA={session:()=>getSession(window.pagepilotWorkspaceSession()),tools:()=>Object.keys(mcpController.getAvailableTools()),navigate:path=>router.navigate({to:path})};`);
const nonce='WORKSPACE-'+crypto.randomBytes(8).toString('hex');
const reference=path.join(out,'send-reference.txt'),spreadsheet=path.join(out,'revenue.xlsx');
fs.copyFileSync(path.join(root,'.temp/workspace-tests/references/revenue.xlsx'),spreadsheet);
fs.writeFileSync(reference,'Synthetic reference. Evidence marker: '+nonce+'\nRevenue grew from 100 to 120, or 20%.\n');
const originalHash=crypto.createHash('sha256').update(fs.readFileSync(spreadsheet)).digest('hex');
dialog.showOpenDialog=async()=>({canceled:false,filePaths:[reference,spreadsheet]});
require('../dist/Chatbox/bootstrap.cjs');
const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const until=async(check,label,timeout=180000)=>{log('WAIT '+label);const end=Date.now()+timeout;while(Date.now()<end){if(await check()){log('READY '+label);return;}await wait(250);}throw Error('Timed out: '+label);};
const checks=[];const pass=label=>{checks.push(label);log('PASS '+label);};
let win;
const js=code=>Promise.race([win.webContents.executeJavaScript(code,true),wait(12000).then(()=>{throw Error('Renderer did not answer: '+code.slice(0,120));})]);
const shadow=`document.querySelector('pagepilot-workspace')?.shadowRoot`;
async function panel(){await js(`(()=>{const p=${shadow};if(!p.querySelector('#panel').matches(':popover-open'))p.querySelector('#trigger').click();})()`);}
async function send(prompt){
 await js(`(()=>{const p=${shadow};if(p?.querySelector('#panel').matches(':popover-open'))p.querySelector('#close').click();const t=document.querySelector('#message-input');t.focus();t.select();})()`);
 await win.webContents.insertText(prompt);await wait(500);
 assert.equal(await js(`document.querySelector('#message-input').value`),prompt);
 win.webContents.sendInputEvent({type:'keyDown',keyCode:'Enter'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Enter'});
 await until(()=>js(`document.querySelector('#message-input')?.value===''&&window.pagepilotWorkspaceSession()!=='new'`),'composer sends exact prompt');
}
async function completed(userCount){
 await until(()=>js(`(async()=>{const s=await window.pagepilotSendQA.session();const users=s?.messages.filter(m=>m.role==='user')||[];const last=s?.messages.findLastIndex(m=>m.role==='user')??-1;return users.length>=${userCount}&&s.messages.slice(last+1).some(m=>m.role==='assistant'&&m.contentParts?.some(p=>p.type==='text'&&p.text))&&!s.messages.some(m=>m.generating);})()`),'real model response completes',240000);
 return await js(`window.pagepilotSendQA.session()`);
}
(async()=>{
 await until(()=>{win=BrowserWindow.getAllWindows()[0];return win&&!win.webContents.isLoading()},'native Chatbox window');
 win.webContents.on('console-message',details=>{if(details.level==='error')log('RENDERER '+String(details.message).split(secret).join('[redacted]'));});
 log('INITIAL '+JSON.stringify(await js(`({path:location.href,body:document.body.innerText.slice(0,1200),ready:!!window.pagepilotSendQA})`)));
 await until(()=>js(`!!window.pagepilotSendQA`),'test introspection');
 await js(`window.pagepilotSendQA.navigate('/');void 0`);
 await until(()=>js(`!!document.querySelector('#message-input')&&!!${shadow}`),'actual composer and workspace');
 await until(()=>js(`window.pagepilotSendQA.tools().some(n=>n.endsWith('__workspace_excel_edit'))&&window.pagepilotSendQA.tools().some(n=>n.endsWith('__pagepilot_context'))`),'real MCP servers');
 await panel();await js(`${shadow}.querySelector('#add').click()`);
 await until(()=>js(`${shadow}.querySelectorAll('.row').length===2`),'choose references through workspace UI');
 const before=await js(`window.pagepilotFiles.context(window.pagepilotWorkspaceSession())`);
 await send('实际API回归测试：请列出当前工作区，并读取 send-reference.txt，回复其中的 Evidence marker 和收入增长率。必须实际调用工作区工具获取内容，不要修改 Word，也不要访问其他文件。');
 const sessionId=await js(`window.pagepilotWorkspaceSession()`);
 const after=await js(`window.pagepilotFiles.context(window.pagepilotWorkspaceSession())`);
 assert.equal(after.workspace_id,before.workspace_id);assert.equal(after.sourceCount,2);
 await panel();await until(()=>js(`${shadow}.querySelectorAll('.row').length===2`),'files visible immediately after first send');
 pass('first actual composer send preserves workspace ID and both visible files');
 const first=await completed(1),firstText=JSON.stringify(first);
 assert(firstText.includes(nonce),'Model must read the unpredictable evidence marker');
 assert(firstText.includes('workspace_read')&&firstText.includes('workspace_list'));
 pass('configured '+model.provider+'/'+model.model+' reads reference through actual MCP');
 await send('继续实际API测试：使用工作区中的 revenue.xlsx，先检查原生 Excel 能力和 Revenue 工作表，再用 workspace_excel_edit 创建一个 Excel 2010 兼容的分析副本：在 Revenue 新增 Profit 列，公式 Revenue-Cost；新建“原生透视”工作表，创建真正可刷新的透视表 ProfitPivot，行 Region、列 Year、值 Profit 求和；再新建“统计”工作表给出 Profit 的描述统计。最后调用 workspace_excel_read 读回透视表实际结果，报告 East/West 两个年份的利润和总利润。必须实际执行工具。不得修改源文件或 Word，不要用静态汇总代替原生透视表。');
 const second=await completed(2),secondText=JSON.stringify(second);
 fs.writeFileSync(path.join(out,'session.json'),JSON.stringify(strip(second),null,2));
 assert(secondText.includes('workspace_excel_edit')&&secondText.includes('workspace_excel_read'));
 const papers=path.join(process.env.PAGEPILOT_PROFILE,'Workspaces/results',before.workspace_id);
 const outputs=fs.existsSync(papers)?fs.readdirSync(papers).filter(n=>n.endsWith('.xlsx')):[];
 assert(outputs.length,'Real model must create an XLSX');
 const verification=JSON.parse(execFileSync(path.join(root,'dist/Python/python.exe'),['-c',
  `import json,sys,zipfile,openpyxl;from pathlib import Path;p=Path(sys.argv[1]);files=list(p.glob('*.xlsx'));out=[]\nfor f in files:\n b=openpyxl.load_workbook(f,data_only=True); pivots=[s for s in b if s._pivots]; values={s.title:list(s.values) for s in pivots};out.append(dict(name=f.name,pivots=values,sheets=b.sheetnames));b.close()\nprint(json.dumps(out,ensure_ascii=False))`,papers],{encoding:'utf8',windowsHide:true,env:{...process.env,PYTHONIOENCODING:'utf-8'}}));
 const native=verification.find(v=>Object.values(v.pivots).some(rows=>rows.some(r=>r[0]==='East'&&r.includes(40)&&r.includes(50))&&rows.some(r=>r[0]==='West'&&r.includes(30)&&r.includes(45))));
 assert(native,'Read actual saved native PivotTable cache and computed values');
 assert.equal(crypto.createHash('sha256').update(fs.readFileSync(spreadsheet)).digest('hex'),originalHash);
 pass('real API creates sheets, formulas, native PivotTable and statistics; saved values verified, source hash unchanged');
 await panel();await until(()=>js(`${shadow}.querySelector('#badge').textContent==='3'`),'generated papers appear automatically');
 await js(`window.pagepilotSendQA.navigate('/')`);await until(()=>js(`window.pagepilotWorkspaceSession()==='new'&&${shadow}?.querySelector('#badge').textContent===''`),'next draft isolated');
 await js(`window.pagepilotSendQA.navigate(${JSON.stringify('/session/'+sessionId)})`);
 await until(()=>js(`window.pagepilotWorkspaceSession()===${JSON.stringify(sessionId)}&&${shadow}?.querySelector('#badge').textContent==='3'`),'existing chat restores references');
 win.webContents.reload();
 await until(()=>js(`!!window.pagepilotSendQA&&!!${shadow}`).catch(()=>false),'renderer reopens');
 await js(`window.pagepilotSendQA.navigate(${JSON.stringify('/session/'+sessionId)})`);
 await until(()=>js(`${shadow}?.querySelector('#badge').textContent==='3'`),'persisted workspace after reopening');
 await panel();
 const names=await js(`Array.from(${shadow}.querySelectorAll('.name'),e=>e.textContent)`);
 assert(names.includes('send-reference.txt')&&names.includes('revenue.xlsx')&&names.includes('分析底稿'));
 pass('switching chats and reopening retain references and generated papers');
 fs.writeFileSync(path.join(out,'result.json'),JSON.stringify({passed:true,model,checks,workspace_id:before.workspace_id,sessionId,names,verification},null,2));
 restore();app.exit(0);
})().catch(async error=>{
 log('FAIL '+error.stack);
 try{fs.writeFileSync(path.join(out,'failed-state.json'),JSON.stringify(strip(await js(`(async()=>({body:document.body.innerText,session:await window.pagepilotSendQA?.session()}))()`)),null,2));}catch{}
 fs.writeFileSync(path.join(out,'result.json'),JSON.stringify({passed:false,error:error.message,model,checks},null,2));restore();app.exit(1);
});
setTimeout(()=>{log('FAIL global timeout');restore();app.exit(2);},450000).unref();
