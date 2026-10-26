const {app,ipcMain,BrowserWindow,screen}=require('electron');
const path=require('node:path'),net=require('node:net');
const root=process.env.PAGEPILOT_ROOT||path.resolve(__dirname,'..');
const pipeName=process.env.PAGEPILOT_PIPE,token=process.env.PAGEPILOT_TOKEN;
if(!/^PagePilot-[a-f0-9]{32}$/.test(pipeName||'')||!/^\w{64}$/.test(token||''))throw Error('请从 Word 的 PagePilot 打开助手。');
app.setPath('userData',path.join(app.getPath('appData'),'PagePilot','Chatbox-1.20.3'));
Object.defineProperty(app,'isPackaged',{get:()=>true});
Object.defineProperty(process,'resourcesPath',{value:path.join(__dirname,'resources')});
app.setAsDefaultProtocolClient=()=>false;
const host=require('./host.cjs');
let primary,displayMode='window';
function applyMode(mode){if(!primary)return;displayMode=mode;const docked=mode==='sidebar';primary.setHasShadow(!docked);primary.setSkipTaskbar(docked);primary.setMovable(!docked);primary.setMaximizable(!docked);primary.setMinimizable(!docked);primary.webContents.send('pagepilot:state',{mode});}
ipcMain.handle('pagepilot:command',async(event,action,args={})=>{
 if(!primary||event.sender!==primary.webContents||event.senderFrame!==primary.webContents.mainFrame||!['ai_mode','ai_approval','ping'].includes(action))throw Error('Invalid request');
 const r=await host(action,args);if(action==='ai_mode')applyMode(r.mode);return r;
});
function inject(value){value=value||{};const servers=value.mcp?.servers||[];return {...value,autoUpdate:false,mcp:{...value.mcp,
 enabledBuiltinServers:value.mcp?.enabledBuiltinServers||[],servers:[...servers.filter(s=>s.id!=='pagepilot-word'),{
 id:'pagepilot-word',name:'PagePilot Word',enabled:true,transport:{type:'stdio',command:path.join(root,'Python','python.exe'),args:[path.join(root,'bridge','server.py')],env:{PAGEPILOT_ROOT:root,PAGEPILOT_PIPE:pipeName,PAGEPILOT_TOKEN:token,PYTHONIOENCODING:'utf-8'}}}]}};}
ipcMain.handle('pagepilot:context',async event=>{
 if(!primary||event.sender!==primary.webContents||event.senderFrame!==primary.webContents.mainFrame||!event.senderFrame.url.startsWith('file:'))throw Error('Invalid sender');
 const c=await host('identity');return c.document;
});
const original=ipcMain.handle.bind(ipcMain);
ipcMain.handle=(channel,fn)=>{
 if(channel==='install-update')return original(channel,()=>{throw Error('请通过 PagePilot 发布包更新。');});
 if(channel==='getSettings')return original(channel,async(...args)=>inject(await fn(...args)));
 if(channel==='getStoreValue')return original(channel,async(event,key,...args)=>{const v=await fn(event,key,...args);return key==='settings'?inject(v):v;});
 return original(channel,fn);
};
app.on('browser-window-created',(_,window)=>{
 if(primary)return;primary=window;host("metric",{stage:"window_created"}).catch(()=>{});window.setMinimumSize(320,240);
 window.hookWindowMessage(0x8047,()=>{if(window.isMinimized())window.restore();window.show();window.focus();});
 window.webContents.on('will-navigate',(e,url)=>{if(!url.startsWith('file:'))e.preventDefault();});
 window.once('ready-to-show',async()=>{host('metric',{stage:'window_visible'}).catch(()=>{});try{const h=window.getNativeWindowHandle();const state=await host('ai_ready',{handle:(h.length===8?h.readBigUInt64LE():BigInt(h.readUInt32LE())).toString()});applyMode(state.mode);}catch(e){require('electron').dialog.showErrorBox('PagePilot',e.message);}});
 window.on('closed',()=>app.exit(0));
});
let probingInput=false;const inputDeadline=Date.now()+120000;const inputProbe=setInterval(async()=>{if(Date.now()>inputDeadline){clearInterval(inputProbe);return;}if(probingInput||!primary||primary.isDestroyed())return;probingInput=true;try{if(await primary.webContents.executeJavaScript("!!document.querySelector('#message-input')")){clearInterval(inputProbe);await host('metric',{stage:'input_ready'});}}catch{}finally{probingInput=false;}},100);inputProbe.unref();
require('./native/dist/main/main.js');
let polling=false,missing=0;
async function putPrompt(text){
 const focus="(()=>{const input=document.querySelector('#message-input');if(!input)return false;input.focus();input.select();return true;})()";
 if(!await primary.webContents.executeJavaScript(focus,true))primary.webContents.send('navigate-to','/');
 for(let i=0;i<50;i++){if(await primary.webContents.executeJavaScript(focus,true)){await primary.webContents.insertText(text);return;}await new Promise(r=>setTimeout(r,100));}
 throw Error('请打开一个 AI 对话后重试目录识别');
}
const health=setInterval(async()=>{if(polling)return;polling=true;try{const p=await host('ping',{role:'ai'});missing=0;if(!primary)return;if(p.activate){primary.show();primary.focus();}if(p.provider){primary.show();primary.focus();primary.webContents.send('navigate-to','/settings/provider');}if(p.prompt){primary.show();primary.focus();await putPrompt(p.prompt);}if(p.state.mode!==displayMode)applyMode(p.state.mode);primary.webContents.send('pagepilot:state',p.state);}catch(e){if(['ENOENT','ECONNREFUSED'].includes(e.code)&&++missing>=3)app.quit();}finally{polling=false;}},1500);health.unref();
