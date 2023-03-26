const path = require('path')
import { dialog } from 'electron';

//import axios from 'axios';

import { PartAnalyser } from "./../analyser/PartAnalyser";
import { ExcelParser } from './../analyser/ExcelParser';
import { FieldMapper } from './../mapper/Mapper';

export class AnalysesFlow{

    parser: ExcelParser;
    mapper: FieldMapper;
    partAnalyser: PartAnalyser;

    constructor() {
      this.mapper  = new FieldMapper();
      this.parser = new ExcelParser();

      this.partAnalyser = new PartAnalyser();
    }

    FileUpload = async function (): Promise<string[]> {
      const result = await dialog.showOpenDialog({
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
      })
      return this.parser.getSheetNames();
    }


    SetWorksheet = async function(sheetName: string): Promise<boolean> {
      this.partAnalyser.setWorksheet(this.parser.getExcelSheet(sheetName));
      return true;
    }

    GetSheetColumns = async function (sheetName: string): Promise<string[]> {
      return this.parser.getColumnsForSheet(sheetName);
    }

    GetColumnValuesUnique = async function (sheetName: string, colName: string): Promise<string[]> {
      
      let colNum = this.partAnalyser.getColumnNumberOfField(colName);
      let values = this.partAnalyser.getUniqueValuesOnColumn(colNum);
      
      return values;
    }

    AnalysedMappedFields = async function name(parts_map: any): Promise<any> {
    
      this.mapper.mapFields(parts_map);
      let fields = this.mapper.joinedFields();

      this.partAnalyser.reset();
      var analyseRes = this.partAnalyser.analyze(fields);
      return analyseRes;
    }

    SaveAnalysisResult = async function (): Promise<string> {

      this.parser.createSheetOnWorkbook("Part Analysis", this.partAnalyser.getResultColumnNames());
        var parts: any[] = [];
        for(let i = 0; i < this.partAnalyser.analysisResult.parts.length; i++){
          
          let colorGray = (i % 2 === 1) ? true : false;
          this.parser.addPartToAnalysisSheet("Part Analysis", this.partAnalyser.analysisResult.parts[i], colorGray);
        }

        const defaultPath = this.parser.filePath;
        const options: Electron.SaveDialogOptions = {
          title: 'Save File',
          defaultPath: defaultPath, // Specify a default file name and extension
          filters: [{ name: 'Excel', extensions: ['xlsx'] }] // Specify the file types that should be shown in the dialog
        };
        
        const { filePath, canceled } = await dialog.showSaveDialog(options);
        if (!canceled && filePath) {
          // The user selected a file path, so you can save the file to that location
          const saveRes = await this.parser.saveWorkbook(filePath);
          if(!saveRes){
            dialog.showMessageBoxSync({
              type: 'error',
              message: 'Kaydemedi, excel kapali mi bi bak bakam',
              title: 'sorry',
              buttons: ['<3', 'tabi canim benim']
            });
          }
        }

        return String(filePath);
    }

    /*SearchCommerce = async function(){
      const searchQuery = '2LA456342-04'; // Replace with your actual search query
      //const payload = {  your payload object  }; // Replace with your actual payload object

      const url = 'https://ecommerce.aircostcontrol.com/search'; // The search endpoint URL on the website

      axios.post(url, {
        params: {
          search_query: searchQuery
        }
      }).then(response => {
        console.log(response.data); // Replace with your desired code to handle the response data
      }).catch(error => {
        console.error(error); // Replace with your desired code to handle errors
      });
    }*/
}

//module.exports = AnalysesFlow;