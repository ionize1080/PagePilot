const {app,ipcMain,BrowserWindow,screen}=require('electron');
const path=require('node:path'),net=require('node:net');
const root=process.env.PAGEPILOT_ROOT||path.resolve(__dirname,'..');
const pipeName=process.env.PAGEPILOT_PIPE,token=process.env.PAGEPILOT_TOKEN;
if(!/^PagePilot-[a-f0-9]{32}$/.test(pipeName||'')||!/^\w{64}$/.test(token||''))throw Error('请从 Word 的 PagePilot 打开助手。');
app.setPath('userData',path.join(app.getPath('appData'),'PagePilot','Chatbox-1.20.3'));
Object.defineProperty(app,'isPackaged',{get:()=>true});
Object.defineProperty(process,'resourcesPath',{value:path.join(__dirname,'resources')});
app.setAsDefaultProtocolClient=()=>false;
function host(action,args={}){return new Promise((resolve,reject)=>{
 const socket=net.createConnection('\\\\.\\pipe\\'+pipeName);let buffer='';
 socket.setTimeout(30000,()=>socket.destroy(Error('Word 响应超时')));
 socket.on('connect',()=>socket.write(JSON.stringify({...args,action,token})+'\n'));
 socket.on('error',reject);socket.on('data',data=>{buffer+=data;if(buffer.length>16*1024*1024){socket.destroy();reject(Error('响应过大'));return;}if(buffer.includes('\n')){socket.end();try{const m=JSON.parse(buffer);m.error?reject(Error(m.error)):resolve(m.result);}catch(e){reject(e);}}});
});}
let primary,settingsWindow;
function showSettings(){
 if(settingsWindow&&!settingsWindow.isDestroyed()){settingsWindow.show();settingsWindow.focus();return;}
 const area=screen.getDisplayNearestPoint(screen.getCursorScreenPoint()).workArea;
 const width=Math.min(760,area.width-40),height=Math.min(740,area.height-40);
 settingsWindow=new BrowserWindow({title:'PagePilot · 设置',x:Math.round(area.x+(area.width-width)/2),y:Math.round(area.y+(area.height-height)/2),width,height,minWidth:Math.min(580,width),minHeight:Math.min(520,height),autoHideMenuBar:true,backgroundColor:'#f5f7fb',webPreferences:{preload:path.join(__dirname,'settings-preload.cjs'),contextIsolation:true,nodeIntegration:false,sandbox:true}});
 settingsWindow.webContents.on('will-navigate',e=>e.preventDefault());
 settingsWindow.webContents.setWindowOpenHandler(()=>({action:'deny'}));
 settingsWindow.on('closed',()=>{settingsWindow=null;});
 settingsWindow.loadFile(path.join(__dirname,'settings.html'));
}
ipcMain.handle('pagepilot:settings',async(event,action,args)=>{
 if(!settingsWindow||event.sender!==settingsWindow.webContents||event.senderFrame!==settingsWindow.webContents.mainFrame||!event.senderFrame.url.startsWith('file:'))throw Error('Invalid sender');
 if(action==='read')return host('settings_read');
 if(action==='write'){
  if(typeof args?.autoApprove!=='boolean'||typeof args?.headingRules!=='string'||args.headingRules.length>64000)throw Error('Invalid settings');
  return host('settings_write',{autoApprove:args.autoApprove,headingRules:args.headingRules});
 }
 if(action==='provider'){primary?.show();primary?.focus();primary?.webContents.send('navigate-to','/settings/provider');return true;}
 throw Error('Unsupported settings action');
});
function inject(value){value=value||{};const servers=value.mcp?.servers||[];return {...value,autoUpdate:false,mcp:{...value.mcp,
 enabledBuiltinServers:value.mcp?.enabledBuiltinServers||[],servers:[...servers.filter(s=>s.id!=='pagepilot-word'),{
 id:'pagepilot-word',name:'PagePilot Word',enabled:true,transport:{type:'stdio',command:path.join(root,'Python','python.exe'),args:[path.join(root,'bridge','server.py')],env:{PAGEPILOT_ROOT:root,PAGEPILOT_PIPE:pipeName,PAGEPILOT_TOKEN:token,PYTHONIOENCODING:'utf-8'}}}]}};}
ipcMain.handle('pagepilot:context',async event=>{
 if(!primary||event.sender!==primary.webContents||event.senderFrame!==primary.webContents.mainFrame||!event.senderFrame.url.startsWith('file:'))throw Error('Invalid sender');
 const c=await host('context');return c.document?{path:c.document.path,name:c.document.name}:null;
});
const original=ipcMain.handle.bind(ipcMain);
ipcMain.handle=(channel,fn)=>{
 if(channel==='install-update')return original(channel,()=>{throw Error('请通过 PagePilot 发布包更新。');});
 if(channel==='getSettings')return original(channel,async(...args)=>inject(await fn(...args)));
 if(channel==='getStoreValue')return original(channel,async(event,key,...args)=>{const v=await fn(event,key,...args);return key==='settings'?inject(v):v;});
 return original(channel,fn);
};
app.on('browser-window-created',(_,window)=>{if(primary)return;primary=window;window.setTitle('PagePilot · Chatbox 1.20.3');window.setMinimumSize(640,480);window.webContents.on('will-navigate',(e,url)=>{if(!url.startsWith('file:'))e.preventDefault();});window.on('closed',()=>app.exit(0));});
require('./native/dist/main/main.js');
const health=setInterval(()=>host('ping').then(p=>{if(p.openSettings&&primary)showSettings();}).catch(()=>app.quit()),1500);health.unref();
