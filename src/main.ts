import { app, BrowserWindow,ipcMain, shell } from "electron";
import * as path from "path";
import { AnalysesFlow } from "./flow/AnalysesFlow";

const flow: AnalysesFlow = new AnalysesFlow();

// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (require('electron-squirrel-startup')) {
  app.quit();
}

function createWindow() {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    fullscreenable: false, // disable fullscreen
    //resizable: false,
    maximizable: false,
    //minimizable: false,
    webPreferences: {
      //nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, "../index.html"));

  // Open the DevTools.
  //mainWindow.webContents.openDevTools();
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  createWindow();

  app.on("activate", function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

ipcMain.handle('xlsx_upload', async (event, wbType, filePath) => {
  if (await flow.WorkbookUpload(wbType, filePath))
  {
    return flow.GetSheetsOnWorkbook(wbType);
  }
  return null;
});

ipcMain.handle('filter_from_table', async (event, filter_map, main_table_col)=> {
  let res = flow.FilterFromTable(filter_map, main_table_col);
  return res;
});

ipcMain.handle('parts_analyse', async (event, parts_map) => {
  return flow.Analyse(parts_map);
});

ipcMain.handle('save_results', async (event, table_name) => {
  var saved_path = await flow.SaveAnalysisResult(table_name);
  saved_path = path.normalize(saved_path).replace(/\\/g, '/');
  if(saved_path.length > 1){
    shell.openPath(`file://${saved_path}`);
  }
  return saved_path;
});

ipcMain.handle('get_sheets_on_workbook', async (event, wbType) => {
  return flow.GetSheetsOnWorkbook(wbType);
});

ipcMain.handle('get_columns_on_sheet', async (event, wbtype, sheetName) => {
  return flow.GetSheetColumns(wbtype, sheetName);
});

ipcMain.handle('get_values_on_column', async (event, wbType, sheetName, colName) => {
  var filters = flow.GetColumnValues(wbType, sheetName, colName);
  return filters;
});

// In this file you can include the rest of your app"s specific main process
// code. You can also put them in separate files and require them here.
declare global {
  interface Window {
      versions: {
          node: () => string,
          chrome: () => string,
          electron: () => string,
          ping: () => Promise<any>,
          xlsx_upload: (wbType: any, filePath: any) => Promise<string[]>,
          filter_from_table: (filter_map: any, main_table_col: any) => Promise<string[]>,
          parts_analyse: (parts_map: any) => Promise<any>,
          save_results: (table_name: string) => Promise<any>,

          get_values_on_column: (wbType: any, sheetName: any, colName: any) => Promise<string[]>,
          get_columns_on_sheet: (wbType: any, sheetName: any) => Promise<string[]>,
          get_sheets_on_workbook: (wbType: any) => Promise<string[]>
      }
  }
};