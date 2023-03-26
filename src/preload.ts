const { contextBridge, ipcRenderer  } = require('electron')
import $ from "jquery";

contextBridge.exposeInMainWorld('versions', {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron,
    ping: () => ipcRenderer.invoke('ping'),
    file_upload: () => ipcRenderer.invoke('file_upload'),
    get_columns: (sheetName: any) => ipcRenderer.invoke('get_columns', sheetName),
    parts_analyse: (parts_map: any) => ipcRenderer.invoke('parts_analyse', parts_map),
    save_results: () => ipcRenderer.invoke('save_results'),
    get_column_values_unique: (sheetName: any, colName: any) => ipcRenderer.invoke('get_column_values_unique', sheetName, colName),
    // we can also expose variables, not just functions
})

// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
window.addEventListener("DOMContentLoaded", () => {
  
  const replaceText = (selector: string, text: string) => {
    const element = document.getElementById(selector);
    if (element) {
      element.innerText = text;
    }
  };

  for (const type of ["chrome", "node", "electron"]) {
    replaceText(`${type}-version`, process.versions[type as keyof NodeJS.ProcessVersions]);
  }
});
