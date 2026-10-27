// Preserve Chatbox/provider headers. Only replace the browser's default client identity.
function install(app,version){
 if(!/^\d+\.\d+\.\d+(?:[-+][a-z0-9.-]+)?$/i.test(version))throw Error('Invalid Chatbox version');
 const base=app.userAgentFallback.replace(/\s(?:Electron|PagePilotChatbox|Chatbox|xyz\.chatboxapp\.app)\/[^\s]+/gi,'');
 const userAgent=`${base.trim()} Chatbox/${version}`;
 app.userAgentFallback=userAgent;
 return userAgent;
}
module.exports={install};
