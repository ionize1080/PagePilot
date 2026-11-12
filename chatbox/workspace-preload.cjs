// Appended to the sandboxed preload. File objects cross the isolated bridge;
// paths are resolved by Electron, never supplied by document content.
(() => {
  const { contextBridge, ipcRenderer, webUtils } = require('electron');
  const invoke = (action, args) => ipcRenderer.invoke('pagepilot:workspace', action, args);
  contextBridge.exposeInMainWorld('pagepilotFiles', {
    summary: (session, directory_id = '', cursor = 0) => invoke('summary', { session, directory_id, cursor }),
    context: session => invoke('context', { session }),
    adopt: (from, session) => invoke('adopt', { from, session }),
    choose: (session, folder, parent = '') => invoke('choose', { session, folder, parent }),
    collection: (session, args) => invoke('collection', { session, ...args }),
    move: (session, ids, parent, before = '', after = '') => invoke('move', { session, ids, parent, before, after }),
    remove: (session, id) => invoke('remove', { session, id }),
    export: (session, id) => invoke('export', { session, id }),
    delete: (session, id) => invoke('delete', { session, id }),
    preferences: value => invoke('preferences', { value }),
    add: (session, files, parent = '') => invoke('add', { session, parent, paths: Array.from(files, file => {
      const resolved = webUtils.getPathForFile(file);
      if (!resolved) throw Error('此文件没有本机路径，请先保存到磁盘，再添加到工作区');
      return resolved;
    }) }),
  });
})();
