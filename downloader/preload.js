const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("downloader", {
  getPlatform: () => ipcRenderer.invoke("get-platform"),
  downloadInstaller: (platformKey) => ipcRenderer.invoke("download-installer", platformKey),
  onProgress: (handler) => ipcRenderer.on("download-progress", (_event, progress) => handler(progress)),
});
