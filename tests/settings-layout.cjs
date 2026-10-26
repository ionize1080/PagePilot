// Renderer layout QA only. The real IPC/Word path is tested separately.
const {app,BrowserWindow,ipcMain}=require('electron'),path=require('node:path');
app.setPath('userData',path.resolve(__dirname,'../.temp/layout-profile-'+process.pid));
const fs=require('node:fs'),report=path.resolve(__dirname,'../.temp/settings-layout-result.json'),results=[];fs.mkdirSync(path.dirname(report),{recursive:true});
ipcMain.handle('pagepilot:settings',(_,action)=>action==='read'?{theme:0,navigationGap:36,aiMode:'window',headingRules:''}:{});
app.whenReady().then(async()=>{
 const win=new BrowserWindow({width:620,height:530,title:'PagePilot · 布局测试',webPreferences:{preload:path.resolve(__dirname,'../chatbox/settings-preload.cjs'),sandbox:true,contextIsolation:true,nodeIntegration:false}});
 await win.loadFile(path.resolve(__dirname,'../chatbox/settings.html'));win.setTitle('PagePilot · 布局测试');
 for(const width of [620,880]){win.setSize(width,530);await new Promise(r=>setTimeout(r,300));
  const data=await win.webContents.executeJavaScript(`(()=>{const boxes=[...document.querySelectorAll('.palette')].map(x=>x.getBoundingClientRect());return {innerWidth,ok:boxes.every(b=>b.left>=0&&b.right<=innerWidth)&&boxes.every((a,i)=>boxes.slice(i+1).every(b=>a.right<=b.left||b.right<=a.left||a.bottom<=b.top||b.bottom<=a.top))};})()`);results.push({width,content:win.getContentSize(),...data});fs.writeFileSync(report,JSON.stringify(results));if(!data.ok||Math.abs(data.innerWidth-win.getContentSize()[0])>2)throw Error('palette overlap or stale viewport '+width);console.log('PASS settings width '+width);
 }
 win.setSize(620,530);console.log('READY_LAYOUT_QA');setTimeout(()=>app.quit(),120000);
}).catch(e=>{console.error(e);app.exit(1)});
app.on('window-all-closed',()=>app.quit());
