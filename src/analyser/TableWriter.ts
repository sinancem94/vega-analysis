import * as ExcelJS from 'exceljs';
import {PartAnalysisResult} from './TableAnalyser'

interface resultColumns{
    partNumber: writeColumn;
    partDescription: writeColumn;
    vendorCode: writeColumn;
    vendorName: writeColumn;
    quantity: writeColumn;
    totalQuantity: writeColumn;
    price: writeColumn;
    totalPrice: writeColumn;
    currency: writeColumn;
    relatedPOs: writeColumn;
    orderType: writeColumn;
}

interface writeColumn{
    name: string;
    col: string;
    mandatory: boolean;
}

export class TableWriter {
    
    workbook: ExcelJS.Workbook;//new ExcelJS.Workbook();
    analysisTable: ExcelJS.Worksheet;

    constructor(workbook: ExcelJS.Workbook){
        this.workbook = workbook;
    }

    columns: resultColumns = {
        partNumber: { name: "Part Number", col: "A", mandatory: true },
        partDescription: { name: "Part Description", col: "B", mandatory: true },
        vendorCode: { name: "Vendor Code", col: "C", mandatory: false },
        vendorName: { name: "Vendor Name", col: "D", mandatory: true },
        quantity: { name: "Quantity", col: "E", mandatory: true },
        totalQuantity: { name: "Total Quantity", col: "F", mandatory: true },
        price: { name: "Price", col: "G", mandatory: false },
        totalPrice: { name: "Total Price", col: "H", mandatory: false },
        currency: { name: "Currency", col: "I", mandatory: false },
        relatedPOs: { name: "Related POs", col: "J", mandatory: false },
        orderType: { name: "Order Type", col: "K", mandatory: false },
      };

    getResultColumnNames(): string[] {
        return ["Part Number", "Part Description", "Vendor Name", "Vendor Code", "Quantity", "Total Quantity", "Price", "Total Price", "Currency", "Related POs", "Order Type"];
    }

    createSheetOnWorkbook(newSheetName: string, columns: string[]) {

        const sheetExists = this.workbook.getWorksheet(newSheetName);;
        if(sheetExists){
            this.workbook.removeWorksheet(newSheetName);
        }
    
        this.analysisTable = this.workbook.addWorksheet(newSheetName);
        var sheetColumns: { header: string; key: string; width: number; }[] = [];
        for(const col in columns ){
            // Change the names of the columns
            const columnWidth = String(columns[col]).length + 2;
            sheetColumns.push({ header: columns[col], key: columns[col], width: columnWidth });
        }
        this.analysisTable.columns = sheetColumns;
        // Freeze the first row
        this.analysisTable.views = [
            {
            state: 'frozen',
            xSplit: 0,
            ySplit: 1,
            topLeftCell: 'B2', // Set the top left cell to the second row
            activeCell: 'B2', // Set the active cell to the second row
            },
        ];
        
        this.analysisTable.getRow(1).eachCell((cell: ExcelJS.Cell) => {
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
        this.analysisTable.eachRow({ includeEmpty: true }, function(row: ExcelJS.Row, rowNumber: number) {
            row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
              cell.numFmt = '@';
            });
        });
    }
    
    addPartToAnalysisSheet(sheetName: string, analysisResult: PartAnalysisResult, colorGray: boolean = false){
    
        const totalCellCount = analysisResult.quantities.length;/*Object.values(analysisResult.prices)
            .reduce((acc, curr) => acc + curr.length, 0);*/
        var rowToWrite: number = Number(this.analysisTable.actualRowCount + 1);
    
        var rows = this.analysisTable.getRows(rowToWrite, totalCellCount);

        //let vendorKeys = Object.keys(analysisResult.vendors);
        rows.forEach((row: ExcelJS.Row, i: number) => {

            row.border= {
                top: {style:'thin', color: {argb:'FF000000'}},
                left: {style:'thin', color: {argb:'FF000000'}},
                bottom: {style:'thin', color: {argb:'FF000000'}},
                right: {style:'thin', color: {argb:'FF000000'}}
            };
            
            const fgColor = colorGray ? 'FFEFEFEF' : 'FFFFFFFF';
            const bgColor = colorGray ? '00000000' : 'FF000000';
            
            row.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: fgColor },
                bgColor: { argb: bgColor }
              };
              //let vendorIndex = Math.floor(i / vendorKeys.length);
              //let vendor = analysisResult.vendors[vendorKeys[vendorIndex]];
              //console.log(vendor);
              /*row.getCell('A').value = analysisResult.partNo;
              row.getCell('B').value = analysisResult.partDesc;
              
              row.getCell('C').value = analysisResult.vendors*/
        });
    
        this.analysisTable.getCell('A' + rowToWrite).value = analysisResult.partNo;
        this.analysisTable.getCell('B' + rowToWrite).value = analysisResult.partDesc;
    
        /*let vendorCounter = 0;
        for (const vendorCode in analysisResult.vendors){
            this.analysisTable.getCell('C' + (rowToWrite + vendorCounter)).value = vendorCode;
            this.analysisTable.getCell('D' + (rowToWrite + vendorCounter)).value = analysisResult.vendors[vendorCode][0];
            
            var mergeCount = analysisResult.vendors[vendorCode][1] - 1;
            if(mergeCount > 0){
                try{
                    this.analysisTable.mergeCells('C' + (rowToWrite + vendorCounter) + ':' + 'C' + (rowToWrite + vendorCounter + mergeCount)); 
                    this.analysisTable.mergeCells('D' + (rowToWrite + vendorCounter) + ':' + 'D' + (rowToWrite + vendorCounter + mergeCount)); 
                }
                catch(err){
                    console.log(err);
                }
            }
    
            vendorCounter += mergeCount + 1; 
        }*/

        let vendorCodeCounter = 0;
        for (const vendorCode in analysisResult.vendors){
            
            let vendorCounter = 0;
            for(const po in analysisResult.vendors[vendorCode]){
                let order = analysisResult.vendors[vendorCode][po];
                this.analysisTable.getCell('C' + (rowToWrite + vendorCodeCounter + vendorCounter)).value = order;
                vendorCounter++;
            }
    
            var mergeCount = vendorCounter - 1;
            this.analysisTable.getCell('D' + (rowToWrite + vendorCodeCounter)).value = vendorCode;
            if(mergeCount > 0){
                this.analysisTable.mergeCells('D' + (rowToWrite + vendorCodeCounter) + ':' + 'D' + (rowToWrite + vendorCodeCounter + mergeCount)); 
            }  
    
            vendorCodeCounter += vendorCounter;
        }
    
        let totQty = 0;
        for(let q = 0; q < totalCellCount; q++)
        {
            let qty = analysisResult.quantities[q];
            this.analysisTable.getCell('E' + (rowToWrite + q)).value = qty;
            totQty += qty;
        }
    
        this.analysisTable.getCell('F' + rowToWrite).value = totQty;
        
        let currencyCounter = 0;
        for (const currency in analysisResult.prices){
            
            let totalPriceInCurr = 0;
            let priceCounter = 0;
            for(const p in analysisResult.prices[currency]){
                let price = analysisResult.prices[currency][p];
                this.analysisTable.getCell('G' + (rowToWrite + currencyCounter + priceCounter)).value = price;
                let qty = analysisResult.quantities[currencyCounter + Number(p)];
                totalPriceInCurr += (price * qty);
                priceCounter++;
            }
    
            var mergeCount = priceCounter - 1;
            this.analysisTable.getCell('H' + (rowToWrite + currencyCounter)).value = totalPriceInCurr;
            if(mergeCount > 0){
                this.analysisTable.mergeCells('H' + (rowToWrite + currencyCounter) + ':' + 'H' + (rowToWrite + currencyCounter + mergeCount)); 
            }
    
            this.analysisTable.getCell('I' + (rowToWrite + currencyCounter)).value = currency;
            if(mergeCount > 0){
                this.analysisTable.mergeCells('I' + (rowToWrite + currencyCounter) + ':' + 'I' + (rowToWrite + currencyCounter + mergeCount)); 
            }
    
            currencyCounter += priceCounter;
        }
    
        let orderTypeCounter = 0;
        for (const orderType in analysisResult.purchaseOrders){
            
            analysisResult.purchaseOrders[orderType].sort(compareDates);

            let orderCounter = 0;
            for(const po in analysisResult.purchaseOrders[orderType]){
                let order = analysisResult.purchaseOrders[orderType][po];
                this.analysisTable.getCell('J' + (rowToWrite + orderTypeCounter + orderCounter)).value = order;
                orderCounter++;
            }
    
            var mergeCount = orderCounter - 1;
            this.analysisTable.getCell('K' + (rowToWrite + orderTypeCounter)).value = orderType;
            if(mergeCount > 0){
                this.analysisTable.mergeCells('K' + (rowToWrite + orderTypeCounter) + ':' + 'K' + (rowToWrite + orderTypeCounter + mergeCount)); 
            }  
    
            orderTypeCounter += orderCounter;
        }
    
        if(totalCellCount > 1){
            this.analysisTable.mergeCells('A' + rowToWrite + ':' + 'A' + (rowToWrite + totalCellCount - 1)); 
            this.analysisTable.mergeCells('B' + rowToWrite + ':' + 'B' + (rowToWrite + totalCellCount - 1)); 
            
            this.analysisTable.mergeCells('F' + rowToWrite + ':' + 'F' + (rowToWrite + totalCellCount - 1)); 
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

function convertToDate(dateString: string): Date {
    const parts = dateString.split(" ");
    const dateParts = parts[0].split(".");
    const timeParts = parts[1].split(":");
    
    const day = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1; // Subtract 1 from month since it's zero-based in JavaScript
    const year = parseInt(dateParts[2]);
    
    const hours = parseInt(timeParts[0]);
    const minutes = parseInt(timeParts[1]);
    const seconds = parseInt(timeParts[2]);
    
    return new Date(year, month, day, hours, minutes, seconds);
  }
  
  function compareDates(a: string, b: string): number {
    const dateA = convertToDate(a);
    const dateB = convertToDate(b);
    
    return dateB.getTime() - dateA.getTime();
  }