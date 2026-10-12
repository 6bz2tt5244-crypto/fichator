// Pont sécurisé entre l'interface et Windows
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('fichator', {
  getConfig: () => ipcRenderer.invoke('config:get'),
  markSeen: v => ipcRenderer.invoke('config:seen', v),
  chooseFolder: () => ipcRenderer.invoke('folder:choose'),
  save: (name, bytes) => ipcRenderer.invoke('file:save', { name, bytes }),
  saveAs: (name, bytes) => ipcRenderer.invoke('file:saveAs', { name, bytes }),
  open: p => ipcRenderer.invoke('file:open', p),
  reveal: p => ipcRenderer.invoke('file:reveal', p),
  installUpdate: () => ipcRenderer.invoke('update:install'),
  onUpdate: cb => ipcRenderer.on('update', (_e, data) => cb(data))
});
