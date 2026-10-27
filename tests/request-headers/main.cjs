// Loopback-only capture, no model credentials or external API calls.
const {app,BrowserWindow,net}=require('electron');
const fs=require('node:fs'),path=require('node:path'),http=require('node:http'),assert=require('node:assert/strict');
const output=path.resolve(process.argv[2]),patched=process.argv.includes('--patched');
app.setPath('userData',path.join(path.dirname(output),'headers-profile-'+(patched?'patched':'baseline')));
if(patched)require('../../chatbox/network-identity.cjs').install(app,'1.20.3');
app.on('window-all-closed',()=>{});
const deadline=setTimeout(()=>{fs.writeFileSync(output,JSON.stringify({error:'timeout'}));app.exit(1);},20000);
app.whenReady().then(async()=>{
 let server,window;
 try{
  const captures=[];
  server=http.createServer((req,res)=>{
   let body='';req.on('data',chunk=>body+=chunk);req.on('end',()=>{
    if(req.url==='/'){res.setHeader('Content-Type','text/html');res.end('<!doctype html><title>Loopback header check</title>');return;}
    captures.push({route:req.url,headers:req.headers,body});res.setHeader('Content-Type','application/json');res.end('{}');
   });
  });
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const origin='http://127.0.0.1:'+server.address().port;
  window=new BrowserWindow({show:false,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});
  await window.loadURL(origin);
  await window.webContents.executeJavaScript(`fetch('/model',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer local-test-only','x-sdk-marker':'preserve-me'},body:'{"model":"local-test","messages":[]}'}).then(r=>r.text())`);
  await window.webContents.executeJavaScript(`fetch('/sdk',{headers:{'User-Agent':'ai-sdk/test-provider','x-sdk-marker':'preserve-me'}}).then(r=>r.text())`);
  await net.fetch(origin+'/net');
  const renderer=captures.find(c=>c.route==='/model'),sdk=captures.find(c=>c.route==='/sdk'),main=captures.find(c=>c.route==='/net');
  assert.equal(renderer.headers.authorization,'Bearer local-test-only');assert.equal(renderer.headers['x-sdk-marker'],'preserve-me');
  assert.equal(renderer.body,'{"model":"local-test","messages":[]}');assert.equal(sdk.headers['x-sdk-marker'],'preserve-me');
  if(patched){
   assert.match(renderer.headers['user-agent'],/Chatbox\/1\.20\.3/);assert.doesNotMatch(renderer.headers['user-agent'],/Electron\//i);
   assert.doesNotMatch(main.headers['user-agent'],/Electron\//i);
   // Chromium may honor or ignore a script-supplied User-Agent. Both must avoid Electron.
   assert.doesNotMatch(sdk.headers['user-agent'],/Electron\//i);
  }else assert.match(renderer.headers['user-agent'],/Electron\//i);
  fs.writeFileSync(output,JSON.stringify({passed:true,patched,appVersion:app.getVersion(),fallback:app.userAgentFallback,captures},null,2));
  app.exit(0);
 }catch(e){fs.writeFileSync(output,JSON.stringify({error:e.stack}));app.exit(1);}
 finally{clearTimeout(deadline);if(window&&!window.isDestroyed())window.destroy();if(server)server.close();}
});
