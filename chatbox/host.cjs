const net=require('node:net');
const pipeName=process.env.PAGEPILOT_PIPE,token=process.env.PAGEPILOT_TOKEN;
if(!/^PagePilot-[a-f0-9]{32}$/.test(pipeName||''))throw Error('请从 Word 打开 PagePilot');
function host(action,args={}){return new Promise((resolve,reject)=>{
 const socket=net.createConnection('\\\\.\\pipe\\'+pipeName);let buffer='';
 socket.setTimeout(action==='ai_approval'?0:30000,()=>socket.destroy(Error('Word 响应超时')));
 socket.on('connect',()=>socket.write(JSON.stringify({...args,action,token})+'\n'));
 socket.on('error',reject);socket.on('data',data=>{buffer+=data;if(buffer.length>16*1024*1024){socket.destroy();reject(Error('响应过大'));return;}if(buffer.includes('\n')){socket.end();try{const m=JSON.parse(buffer);m.error?reject(Error(m.error)):resolve(m.result);}catch(e){reject(e);}}});
});}

module.exports=host;
