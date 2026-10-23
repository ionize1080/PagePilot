const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('pagepilotSettings',{
 read:()=>ipcRenderer.invoke('pagepilot:settings','read'),
 write:values=>ipcRenderer.invoke('pagepilot:settings','write',values),
 provider:()=>ipcRenderer.invoke('pagepilot:settings','provider')
});
