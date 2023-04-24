import { Workbook } from "exceljs";
import * as ExcelJS from 'exceljs';

class sheetHolder {
    sheets: sheet[];
    constructor() {
        this.sheets = [];
    }
}

class sheet {
    name: string;
    columns: string[];
    constructor(name: string, columns: string[]) {
        this.name = name;
        this.columns = columns;
    }
}

export enum workbookType{
    main = 0,
    filter = 1
}

export class ExcelParser {
    static instance: ExcelParser;
    isParsing: boolean = false;
    relatedSheets: sheetHolder[] = [];// typeof WorkSheets;
    workbooks: Workbook[] = [];//new ExcelJS.Workbook();

    constructor() {
        if(!ExcelParser.instance) {
            ExcelParser.instance = this;
            this.isParsing = false;
            this.relatedSheets = [];
            this.workbooks = [] //new ExcelJS.Workbook();
        }
        return ExcelParser.instance;
    }

    async parse(wbType: workbookType,filePath: string, extension: string) {
        if(this.isParsing) {
            console.log("Parser is already working, please wait for it to complete.");
            return;
        }

        this.isParsing = true;
        var newSheetholder = new sheetHolder();
        var newWorkbook = new ExcelJS.Workbook();

        try {
            switch(extension){
                case ".csv":
                    const worksheet = await newWorkbook.csv.readFile(filePath);
                    break;
                case ".xlsx":
                    
                    await newWorkbook.xlsx.readFile(filePath)
                        .then(() => {
                            
                            // Iterate over each worksheet in the workbook
                            newWorkbook.eachSheet(worksheet => {
                                const columns: string[] = [];
                                const firstRow = worksheet.getRow(1); // get the first row of the worksheet
                                firstRow.eachCell(cell => {
                                    columns.push(String(cell.value));
                                });
                                var newSheet = new sheet(worksheet.name, columns);
                                newSheetholder.sheets.push(newSheet);
                            })
                        })
                        .catch(error => {
                            console.error(error)
                        })
                        .finally(() => {

                            if(wbType == workbookType.main){
                                if(this.workbooks.length > 0){
                                    this.workbooks[0] = newWorkbook;
                                    this.relatedSheets[0] = newSheetholder;
                                }else{
                                    this.workbooks.push(newWorkbook);
                                    this.relatedSheets.push(newSheetholder);
                                }
                            }   
                            else if(wbType == workbookType.filter){
                                if(this.workbooks.length > 1){
                                    this.workbooks[1] = newWorkbook;
                                    this.relatedSheets[1] = newSheetholder;
                                }else{
                                    this.workbooks.push(newWorkbook);
                                    this.relatedSheets.push(newSheetholder);
                                }
                            }
                        })
                    
                    break;
                default:
                    break;
            }
        } catch (error) {
            console.error(error);
        } finally {
            this.isParsing = false;
        }
    }

    getTableNames(workbookIndex: number) {
        return this.relatedSheets[workbookIndex].sheets.map((sheet: { name: string; }) => sheet.name);
    }

    getExcelTable(workbookIndex: number, sheetName: string): ExcelJS.Worksheet & any {
        try {
            return this.workbooks[workbookIndex].getWorksheet(sheetName);
        }
        catch (err) {
            console.error("Err: " + err);
            return null;
        }
    }

    getColumnsForTable(workbookIndex: number, sheetName: string) {
        const sheet = this.relatedSheets[workbookIndex].sheets.find((sheet: { name: string; }) => sheet.name === sheetName);
        if(sheet) {
            return sheet.columns;
        } else {
            console.log(`Sheet ${sheetName} not found. Please enter a valid sheet name.`);
            return null;
        }
    }
}