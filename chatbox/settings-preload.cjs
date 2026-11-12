const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('pagepilotSettings',{
 workspace:value=>ipcRenderer.invoke('pagepilot:workspace-preferences',value),
 read:()=>ipcRenderer.invoke('pagepilot:settings','read'),
 ready:()=>ipcRenderer.invoke('pagepilot:settings','ready'),
 write:values=>ipcRenderer.invoke('pagepilot:settings','write',values),
 provider:()=>ipcRenderer.invoke('pagepilot:settings','provider')
});
