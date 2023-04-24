import { workbookType } from "./analyser/ExcelParser";

const { contextBridge, ipcRenderer  } = require('electron')
//import * as $ from "jquery";

contextBridge.exposeInMainWorld('versions', {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron,
    
    //flow functions
    xlsx_upload: (wbType: workbookType, filePath: any) => ipcRenderer.invoke('xlsx_upload', wbType, filePath),
    filter_from_table: (filter_map: any, main_table_col: any) => ipcRenderer.invoke('filter_from_table', filter_map, main_table_col),
    parts_analyse: (parts_map: any) => ipcRenderer.invoke('parts_analyse', parts_map),
    save_results: (table_name: string) => ipcRenderer.invoke('save_results', table_name),
    
    //get functions
    get_sheets_on_workbook: (wbType: workbookType) => ipcRenderer.invoke('get_sheets_on_workbook', wbType),
    get_columns_on_sheet: (wbType: workbookType, sheetName: any) => ipcRenderer.invoke('get_columns_on_sheet', wbType, sheetName),
    get_values_on_column: (wbType: workbookType, sheetName: any, colName: any) => ipcRenderer.invoke('get_values_on_column', wbType, sheetName, colName),

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
