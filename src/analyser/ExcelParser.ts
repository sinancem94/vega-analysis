import { Workbook } from "exceljs";
import * as ExcelJS from 'exceljs';
import {PartAnalysis} from './Analyser'

const {WorkSheets, Sheet} = require("./../objects/WorkSheets")

export class ExcelParser {
    static instance: ExcelParser;
    isParsing: boolean = false;
    workSheets: typeof WorkSheets;
    workbook: Workbook = new ExcelJS.Workbook();

    constructor() {
        if(!ExcelParser.instance) {
            ExcelParser.instance = this;
            this.isParsing = false;
            this.workbook = new ExcelJS.Workbook();
        }
        return ExcelParser.instance;
    }

    async parse(filePath: string, extension: string) {
        if(this.isParsing) {
            console.log("Parser is already working, please wait for it to complete.");
            return;
        }

        this.isParsing = true;
        this.workSheets = new WorkSheets();
        this.workbook = new ExcelJS.Workbook();

        try {
            switch(extension){
                case ".csv":
                    const worksheet = await this.workbook.csv.readFile(filePath);
                    break;
                case ".xlsx":
                    
                    await this.workbook.xlsx.readFile(filePath)
                        .then(() => {

                            // Iterate over each worksheet in the workbook
                            this.workbook.eachSheet(worksheet => {
                                const columns: string[] = [];
                                const firstRow = worksheet.getRow(1); // get the first row of the worksheet
                                firstRow.eachCell(cell => {
                                    columns.push(String(cell.value));
                                });
                                var sheet = new Sheet(worksheet.name, columns);
                                this.workSheets.sheets.push(sheet);
                            })
                        })
                        .catch(error => {
                            console.error(error)
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

    getSheetNames() {
        if(this.workSheets) {
            return this.workSheets.sheets.map((sheet: { name: string; }) => sheet.name);
        } else {
            console.log("No worksheets found. Please parse an Excel file first.");
            return null;
        }
    }

    getExcelSheet(sheetName: string): ExcelJS.Worksheet & any {
        if(this.workbook.worksheets.length != 0){
            try {
                 return this.workbook.getWorksheet(sheetName);
            }
            catch (err) {
                console.error("Err: " + err);
                return null;
            }
        }
        else{
            console.log("No worksheets found. Please parse an Excel file first.");
            return null;
        }
    }

    getColumnsForSheet(sheetName: string) {
        if(this.workSheets) {
            const sheet = this.workSheets.sheets.find((sheet: { name: string; }) => sheet.name === sheetName);
            if(sheet) {
                return sheet.columns;
            } else {
                console.log(`Sheet ${sheetName} not found. Please enter a valid sheet name.`);
                return null;
            }
        } else {
            console.log("No worksheets found. Please parse an Excel file first.");
            return null;
        }
    }

    getValuesOnColumn(sheetName: string, columnName: string){
        const worksheet = this.getExcelSheet(sheetName);
    }

    createSheetOnWorkbook(newSheetName: string, columns: string[]) {

        const sheetExists = this.getExcelSheet(newSheetName);
        if(sheetExists){
            this.workbook.removeWorksheet(newSheetName);
        }

        const worksheet = this.workbook.addWorksheet(newSheetName);
        var sheetColumns: { header: string; key: string; width: number; }[] = [];
        for(const col in columns ){
            // Change the names of the columns
            const columnWidth = String(columns[col]).length + 2;
            sheetColumns.push({ header: columns[col], key: columns[col], width: columnWidth });
        }
        worksheet.columns = sheetColumns;
        // Freeze the first row
        worksheet.views = [
            {
            state: 'frozen',
            xSplit: 0,
            ySplit: 1,
            topLeftCell: 'B2', // Set the top left cell to the second row
            activeCell: 'B2', // Set the active cell to the second row
            },
        ];
        
        worksheet.getRow(1).eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '273B91' },
            };
            cell.font = {
                color: { argb: 'FFFFFFFF' },
            };
        });

        //TODO: test this text format
        worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
              cell.numFmt = '@';
            });
        });
    }

    addPartToAnalysisSheet(sheetName: string, analysisResult: PartAnalysis, colorGray: boolean = false){

        const analysisSheet = this.getExcelSheet(sheetName);
        const totalCellCount = analysisResult.quantities.length;/*Object.values(analysisResult.prices)
            .reduce((acc, curr) => acc + curr.length, 0);*/
        var rowToWrite: number = Number(analysisSheet.actualRowCount + 1);

        var rows = analysisSheet.getRows(rowToWrite, rowToWrite + totalCellCount);
        rows.forEach((cell: ExcelJS.Cell) => {
            cell.style.border= {
                top: {style:'thin', color: {argb:'FF000000'}},
                left: {style:'thin', color: {argb:'FF000000'}},
                bottom: {style:'thin', color: {argb:'FF000000'}},
                right: {style:'thin', color: {argb:'FF000000'}}
            };
            
            const fgColor = colorGray ? 'FFEFEFEF' : 'FFFFFFFF';
            const bgColor = colorGray ? '00000000' : 'FF000000';
            
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: fgColor },
                bgColor: { argb: bgColor }
              };
        });

        analysisSheet.getCell('A' + rowToWrite).value = analysisResult.partNo;
        analysisSheet.getCell('B' + rowToWrite).value = analysisResult.partDesc;

        let vendorCounter = 0;
        for (const vendorCode in analysisResult.vendors){
            analysisSheet.getCell('C' + (rowToWrite + vendorCounter)).value = vendorCode;
            analysisSheet.getCell('D' + (rowToWrite + vendorCounter)).value = analysisResult.vendors[vendorCode][0];
            
            var mergeCount = analysisResult.vendors[vendorCode][1] - 1;
            if(mergeCount > 0){
                try{
                    analysisSheet.mergeCells('C' + (rowToWrite + vendorCounter) + ':' + 'C' + (rowToWrite + vendorCounter + mergeCount)); 
                    analysisSheet.mergeCells('D' + (rowToWrite + vendorCounter) + ':' + 'D' + (rowToWrite + vendorCounter + mergeCount)); 
                }
                catch(err){
                    console.log(err);
                }
            }

            vendorCounter += mergeCount + 1; 
        }

        let totQty = 0;
        for(let q = 0; q < totalCellCount; q++)
        {
            let qty = analysisResult.quantities[q];
            analysisSheet.getCell('E' + (rowToWrite + q)).value = qty;
            totQty += qty;
            
        }

        analysisSheet.getCell('F' + rowToWrite).value = totQty;
        
        let currencyCounter = 0;
        for (const currency in analysisResult.prices){
            
            let totalPriceInCurr = 0;
            let priceCounter = 0;
            for(const p in analysisResult.prices[currency]){
                let price = analysisResult.prices[currency][p];
                analysisSheet.getCell('G' + (rowToWrite + currencyCounter + priceCounter)).value = price;
                totalPriceInCurr += price;
                priceCounter++;
            }

            var mergeCount = priceCounter - 1;
            analysisSheet.getCell('H' + (rowToWrite + currencyCounter)).value = totalPriceInCurr;
            if(mergeCount > 0){
                analysisSheet.mergeCells('H' + (rowToWrite + currencyCounter) + ':' + 'H' + (rowToWrite + currencyCounter + mergeCount)); 
            }

            analysisSheet.getCell('I' + (rowToWrite + currencyCounter)).value = currency;
            if(mergeCount > 0){
                analysisSheet.mergeCells('I' + (rowToWrite + currencyCounter) + ':' + 'I' + (rowToWrite + currencyCounter + mergeCount)); 
            }

            currencyCounter += priceCounter;
        }

        let orderTypeCounter = 0;
        for (const orderType in analysisResult.purchaseOrders){
            
            let orderCounter = 0;
            for(const po in analysisResult.purchaseOrders[orderType]){
                let order = analysisResult.purchaseOrders[orderType][po];
                analysisSheet.getCell('J' + (rowToWrite + orderTypeCounter + orderCounter)).value = order;
                orderCounter++;
            }

            var mergeCount = orderCounter - 1;
            analysisSheet.getCell('K' + (rowToWrite + orderTypeCounter)).value = orderType;
            if(mergeCount > 0){
                analysisSheet.mergeCells('K' + (rowToWrite + orderTypeCounter) + ':' + 'K' + (rowToWrite + orderTypeCounter + mergeCount)); 
            }

            orderTypeCounter += orderCounter;
        }

        if(totalCellCount > 1){
            analysisSheet.mergeCells('A' + rowToWrite + ':' + 'A' + (rowToWrite + totalCellCount - 1)); 
            analysisSheet.mergeCells('B' + rowToWrite + ':' + 'B' + (rowToWrite + totalCellCount - 1)); 
            
            analysisSheet.mergeCells('F' + rowToWrite + ':' + 'F' + (rowToWrite + totalCellCount - 1)); 
        }
    }

    async saveWorkbook(savePath: string): Promise<boolean> {
        try{
            await this.workbook.xlsx.writeFile(savePath);
            
            //const saveRes = await this.workbook.xlsx.writeFile(savePath);
            return true;
        }
        catch(err){
            return false;
        }
    }
}