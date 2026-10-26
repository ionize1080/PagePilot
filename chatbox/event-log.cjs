const fs=require('node:fs'),path=require('node:path');
const folder=process.env.PAGEPILOT_LOGS;
const file=folder&&path.join(folder,`electron-${process.pid}.jsonl`);
function log(event,detail='',level='info'){
 if(!file||(level==='debug'&&process.env.PAGEPILOT_LOG_LEVEL==='info'))return;
 try{
  fs.mkdirSync(folder,{recursive:true});
  if(fs.existsSync(file)&&fs.statSync(file).size>4*1024*1024){for(let i=3;i>=1;i--){let target=file+'.'+i,source=i===1?file:file+'.'+(i-1);if(fs.existsSync(target))fs.unlinkSync(target);if(fs.existsSync(source))fs.renameSync(source,target);}}
  const clean=String(detail).replace(/(\b(?:api[_-]?key|token|authorization)\b\s*[=:]\s*)([^\s,;]+)/gi,'$1[redacted]');
  fs.appendFileSync(file,JSON.stringify({utc:new Date().toISOString(),pid:process.pid,component:'electron',level,eventName:event,detail:clean.slice(0,12000)})+'\n','utf8');
 }catch{}
}
function install(app){
 log('process.start',`electron=${process.versions.electron} role=${process.argv[1]?.endsWith('settings-entry.cjs')?'settings':'assistant'}`);
 for(const name of ['ready','before-quit','will-quit','window-all-closed'])app.on(name,()=>log('app.'+name));
 app.on('browser-window-created',(_,w)=>{log('window.created',`id=${w.id}`);for(const e of ['show','hide','close','closed','unresponsive','responsive'])w.on(e,()=>log('window.'+e,`id=${w.id}`));w.webContents.on('render-process-gone',(_,d)=>log('renderer.exited',`reason=${d.reason} exitCode=${d.exitCode}`,'error'));w.webContents.on('did-fail-load',(_,code,description)=>log('renderer.load.failed',`code=${code} description=${description}`,'error'));});
 process.on('uncaughtExceptionMonitor',e=>log('process.exception',String(e?.name||'Error')+'\n'+String(e?.stack||'').split('\n').slice(1).join('\n'),'error'));
}
module.exports={log,install};
