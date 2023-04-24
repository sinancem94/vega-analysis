const path = require('path')
import { dialog } from 'electron';

//import axios from 'axios';

import { PartAnalyser } from "./../analyser/PartAnalyser";
import { ExcelParser, workbookType } from './../analyser/ExcelParser';
import { FieldMapper } from './../mapper/Mapper';
import { TableAnalyser } from '../analyser/TableAnalyser';
import { TableWriter } from '../analyser/TableWriter';

export class AnalysesFlow{

    filePath: string;
    parser: ExcelParser;
    mapper: FieldMapper;
    partAnalyser: PartAnalyser;

    constructor() {
      this.mapper  = new FieldMapper();
      this.parser = new ExcelParser();
    }

    WorkbookUpload = async function (wbtype: workbookType, filePath: string): Promise<boolean> {
      
      /*const result = await dialog.showOpenDialog({
        title: 'Select a file',
        properties: ['openFile'],
        filters: [
          { name: 'Spreadsheets', extensions: ['xlsx', 'csv'] },
          { name: 'All Files', extensions: ['*'] }
        ]
      }).then(async (result: Electron.OpenDialogReturnValue) => {

        if(result.filePaths.length == 0){
          console.log('empty file path');
          // Show an alert dialog box
          dialog.showMessageBoxSync({
            type: 'warning',
            message: 'if you gonna do some shit then do it',
            title: 'bitch please..',
            buttons: ['OK!', 'Sorry..']
          });
          return null;
        }
        var filePath = result.filePaths[0];
        const extension = path.extname(filePath);
        await this.parser.parse(filePath, extension);
      }).catch((err: Error) => {
        console.log('Error:', err);
      })*/
      
      this.filePath = filePath;
      const extension = path.extname(this.filePath);

      if(extension != '.xlsx' && extension != '.csv'){
        dialog.showMessageBoxSync({
          type: 'error',
          message: 'Only allowed .xlsx and .csv',
          title: 'Invalid ext',
          buttons: ['OK']
        });
        return false;
      }

      await this.parser.parse(wbtype, this.filePath, extension);
      return true;
    }

    GetSheetsOnWorkbook = async function (workbookIndex: workbookType): Promise<string[]> {
      return this.parser.getTableNames(workbookIndex);
    }

    GetSheetColumns = async function (workbookIndex: workbookType, sheetName: string): Promise<string[]> {
      return this.parser.getColumnsForTable(workbookIndex, sheetName);
    }

    GetColumnValues = async function (workbookIndex: workbookType, sheetName: string, colName: string): Promise<string[]> {
      let tempAnalyser = new TableAnalyser(this.parser.getExcelTable(workbookType.main, sheetName));
      let colNum = tempAnalyser.getColumnNumberOfField(colName);
      let values = tempAnalyser.getValuesOnColumn(colNum);
      
      return values;
    }

    FilterFromTable = async function(filter_map: any, mainSheetCol: string): Promise<string[]> {

      this.mapper.MapFilterTable(filter_map);
      let filterFields = this.mapper.filterTable;

      let filterAnalyser = new TableAnalyser(this.parser.getExcelTable(workbookType.filter, filterFields.sheetName));
      
      let allValues = this.partAnalyser.getUniqueValuesOnColumn(this.partAnalyser.getColumnNumberOfField(mainSheetCol));
      let filterTableVals = filterAnalyser.getUniqueValuesOnColumn(filterAnalyser.getColumnNumberOfField(filterFields.columnName));

      let unfilteredVals: string[] = [];

      allValues.forEach(function (val: string){
        if((filterTableVals.includes(val) && filterFields.isInclude) || (!filterTableVals.includes(val) && !filterFields.isInclude)){
          unfilteredVals.push(val);
        }
      });
      
      return unfilteredVals;
    }

    Analyse = async function name(parts_map: any): Promise<any> {
    
      this.mapper.mapFields(parts_map);
      let fields = this.mapper.joinedFields();

      this.partAnalyser = new PartAnalyser(this.parser.getExcelTable(workbookType.main, fields.sheetName));
      var analyseRes = this.partAnalyser.analyze(fields);
      return analyseRes;
    }

    SaveAnalysisResult = async function (table_name: string): Promise<string> {

      let writer = new TableWriter(this.parser.workbooks[workbookType.main]);
      
      if(table_name.length > 22){
        table_name = table_name.slice(0, 18) + "...";
      }

      let tableName = table_name + " Analysis";
      
      writer.createSheetOnWorkbook(tableName, writer.getResultColumnNames());

      for(let i = 0; i < this.partAnalyser.analysisResult.parts.length; i++){
        
        let colorGray = (i % 2 === 1) ? true : false;
        writer.addPartToAnalysisSheet(tableName, this.partAnalyser.analysisResult.parts[i], colorGray);
      }

      const saveRes = await writer.saveWorkbook(this.filePath);
      if(!saveRes){
        dialog.showMessageBoxSync({
          type: 'error',
          message: 'Kayıt başarısız, excel tablosu açık olabilir mi?',
          title: 'Save error',
          buttons: ['OK']
        });
        return "";
      }
      /*const options: Electron.SaveDialogOptions = {
        title: 'Save File',
        defaultPath: defaultPath, // Specify a default file name and extension
        filters: [{ name: 'Excel', extensions: ['xlsx'] }] // Specify the file types that should be shown in the dialog
      };
      
      const { filePath, canceled } = await dialog.showSaveDialog(options);
      if (!canceled && filePath) {
        // The user selected a file path, so you can save the file to that location
        const saveRes = await writer.saveWorkbook(filePath);
        if(!saveRes){
          dialog.showMessageBoxSync({
            type: 'error',
            message: 'Kayıt başarısız, excel tablosu açık olabilir mi?',
            title: 'Save error',
            buttons: ['OK']
          });
        }
      }*/

      return String(this.filePath);
    }
}