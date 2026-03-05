const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  openFileDialog: (options) => ipcRenderer.invoke("open-file-dialog", options || {}),
  saveFileDialog: () => ipcRenderer.invoke("save-file-dialog"),
  startProcessing: (args) => ipcRenderer.invoke("start-processing", args),
  resolveResponse: (response) => ipcRenderer.invoke("resolve-response", response),
  getDefaultZoteroDB: () => ipcRenderer.invoke("get-default-zotero-db"),
  onBackendMessage: (callback) => {
    ipcRenderer.on("backend-message", (event, msg) => callback(msg));
  },
});
