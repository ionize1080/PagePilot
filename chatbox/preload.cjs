// Expose only context metadata, never host credentials or generic IPC.
require('electron').contextBridge.exposeInMainWorld('pagepilot',{
 context:()=>require('electron').ipcRenderer.invoke('pagepilot:context')
});
