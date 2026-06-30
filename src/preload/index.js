import { contextBridge, ipcRenderer } from 'electron'

contextBridge.exposeInMainWorld('api', {
  openFolder: () => ipcRenderer.invoke('dialog:openFolder'),
  readDir: (path) => ipcRenderer.invoke('fs:readDir', path),
  readFile: (path) => ipcRenderer.invoke('fs:readFile', path),
  writeFile: (path, arrayBuffer) => ipcRenderer.invoke('fs:writeFile', path, arrayBuffer),
  openPath: (path) => ipcRenderer.invoke('shell:openPath', path)
})
