const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("qt", {
  getPaths: () => ipcRenderer.invoke("paths:get"),
  saveFile: (rel, bytes, folder = "data") =>
    ipcRenderer.invoke("file:save", { rel, bytes, folder }),
  loadCatalog: () => ipcRenderer.invoke("catalog:read")
});
