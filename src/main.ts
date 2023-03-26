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
    width: 800,
    height: 600,
    fullscreenable: false, // disable fullscreen
    //resizable: false,
    maximizable: false,
    //minimizable: false,
    webPreferences: {
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, "../index.html"));

  // Open the DevTools.
  mainWindow.webContents.openDevTools();

  ipcMain.handle('file_upload', async (event) => {
    return flow.FileUpload();
  });

  ipcMain.handle('get_columns', async (event, sheetName) => {
    flow.SetWorksheet(sheetName);
    return flow.GetSheetColumns(sheetName);
  });

  ipcMain.handle('parts_analyse', async (event, parts_map) => {
    return flow.AnalysedMappedFields(parts_map);
  });

  ipcMain.handle('save_results', async (event) => {
    var saved_path = await flow.SaveAnalysisResult();
    saved_path = path.normalize(saved_path).replace(/\\/g, '/');
    if(saved_path.length > 1){
      shell.openPath(`file://${saved_path}`);
    }
    return saved_path;
  });

  ipcMain.handle('get_column_values_unique', async (event, sheetName, colName) => {
    var filters = flow.GetColumnValuesUnique(sheetName, colName);
    return filters;
  });
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

// In this file you can include the rest of your app"s specific main process
// code. You can also put them in separate files and require them here.
declare global {
  interface Window {
      versions: {
          node: () => string,
          chrome: () => string,
          electron: () => string,
          ping: () => Promise<any>,
          file_upload: () => Promise<string[]>,
          get_columns: (sheetName: any) => Promise<string[]>,
          parts_analyse: (parts_map: any) => Promise<any>,
          save_results: () => Promise<any>,
          get_column_values_unique: (sheetName: any, colName: any) => Promise<string[]>
      }
  }
};