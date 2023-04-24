
import { Worksheet, CellValue, Workbook } from 'exceljs';
import { Fields } from "../mapper/Mapper"

export interface AnalysisResult{
  success: boolean;
  totalVolume: { [key: string]: number};
  uniquePartCount: number;
  uniqueVendorCount: number;
  parts: PartAnalysisResult[];
}

export interface PartAnalysisResult{
  partNo: string;
  partDesc: string;
  quantities: number[];
  purchaseOrders: { [orderType: string]: string[]}; 
  prices: { [currency: string]: number[]}; 
  vendors: { [vendorCode: string]: string[]};//{ [vendorCode: string]: [vendorName: string, index: number] };     
}

export interface PartTableIndexes{
  partNo: string;
  rowIndexes: number[];
  rowCount: number;
}

export class TableAnalyser {

  protected worksheet: Worksheet;
  analysisResult: any;

  constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;
  }

  reset(): void {
  
  }

  analyze(fields: any): any{
    return this.analysisResult;
  };


  getValuesOnColumn(column: number): string[] {
    const values: string[] = [];

    this.worksheet.eachRow((row: { getCell: (arg0: number) => any; }, rowIndex: number) => {
      if (rowIndex > 1) {
        const cell = row.getCell(column);

        if (cell.value !== null && cell.value !== undefined) {
          values.push(cell.value.toString());
        }
      }
    });

    return values;
  }

  getColumnNumberOfField(columnName: string): number | null {
    const row = this.worksheet.getRow(1);
    
    let column: number = null;
    for (let i = 1; i <= row.cellCount; i++) {
      const cell = row.getCell(i);

      if (cell.value && cell.value.toString() === columnName) {
        column = Number(cell.col);
        break;
      }
    }

    return column !== null ? column : null;
  }

  getUniqueValuesOnColumn(columnNumber: number): string[] {
    return columnNumber !== null ? Array.from(new Set(this.getValuesOnColumn(columnNumber))) : [];
  }

  getUniqueValuesWithRows(columnNumber: number): PartTableIndexes[] {// { [key: string]: number } {
    const uniqueParts: PartTableIndexes[] = [];// { [key: string]: number } = {};
    const column = this.worksheet.getColumn(columnNumber);
  
    column.eachCell((cell: { value: any; }, rowNumber: number) => {
      if (rowNumber === 1) return; // skip header row
  
      const cellValue = cell.value as CellValue;
      if (cellValue === undefined || cellValue === null) return; // skip empty cells
      
      const cellValueString = cellValue.toString();

      let parts = uniqueParts.filter((val) => val.partNo === cellValueString);
      if (parts.length > 0) {
        parts[0].rowCount += 1;
        parts[0].rowIndexes.push(rowNumber);
      } else {
        let part: PartTableIndexes = {partNo: cellValueString, rowCount: 1, rowIndexes: [rowNumber]}
        uniqueParts.push(part);
      }
    });
  
    return uniqueParts;
  }
}