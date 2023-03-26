
import { Worksheet, CellValue, Workbook } from 'exceljs';
import { AllMappedFields } from "../mapper/Mapper"

export interface PartsAnalysisResult{
  success: boolean;
  totalVolume: { [key: string]: number};
  uniquePartCount: number;
  uniqueVendorCount: number;
  parts: PartAnalysis[]
}

export interface PartAnalysis{
  partNo: string;
  partDesc: string;
  quantities: number[];
  purchaseOrders: { [key: string]: string[]}; 
  prices: { [key: string]: number[]}; 
  vendors: { [vendorCode: string]: [vendorName: string, count: number] };     
}

export abstract class Analyser {

  protected worksheet: Worksheet;
  analysisResult: any;
  fields: AllMappedFields;

  constructor() {
    this.fields = {partNumber: "", partDesc: "", partQuantity: "", unitPrices: "", unitCurrency: "", purchaseOrder: "", 
        vendorCode: "", vendorName: "", sheetName: "", orderType: "", typeFilter: [], currencyFilter: [], vendorFilter: []};
  }

  reset(): void {
    this.fields = {partNumber: "", partDesc: "", partQuantity: "", unitPrices: "", unitCurrency: "", purchaseOrder: "", 
    vendorCode: "", vendorName: "", sheetName: "", orderType: "", typeFilter: [], currencyFilter: [], vendorFilter: []};
  }

  abstract analyze(fields: AllMappedFields): any;
  abstract getResultColumnNames(): string[];

  private getColumnOfField(field: string): number | null {
    const row = this.worksheet.getRow(1);

    for (let i = 1; i <= row.cellCount; i++) {
      const cell = row.getCell(i);

      if (cell.value === field) {
        return Number(cell.col);
      }
    }

    return null;
  }

  private getColumnValues(column: number): CellValue[] {
    const values: CellValue[] = [];

    this.worksheet.eachRow((row: { getCell: (arg0: number) => any; }, rowIndex: number) => {
      if (rowIndex > 1) {
        const cell = row.getCell(column);

        if (cell.value !== null && cell.value !== undefined) {
          values.push(cell.value);
        }
      }
    });

    return values;
  }

  protected setFields(fields: AllMappedFields){
    this.fields = fields;
  }

  setWorksheet(worksheet: Worksheet){
    this.worksheet = worksheet;
  }

  getColumnNumberOfField(columnName: string): number | null {
    const column = this.getColumnOfField(columnName);
    return column !== null ? column : null;
  }

  getUniqueValuesOnColumn(columnNumber: number): CellValue[] {
    return columnNumber !== null ? Array.from(new Set(this.getColumnValues(columnNumber))) : [];
  }

  getUniqueValuesWithCount(columnNumber: number): { [key: string]: number } {
    const uniqueValues: { [key: string]: number } = {};
    const column = this.worksheet.getColumn(columnNumber);
  
    column.eachCell((cell: { value: any; }, rowNumber: number) => {
      if (rowNumber === 1) return; // skip header row
  
      const cellValue = cell.value as CellValue;
      if (cellValue === undefined || cellValue === null) return; // skip empty cells
  
      const cellValueString = cellValue.toString();
      if (uniqueValues[cellValueString]) {
        uniqueValues[cellValueString] += 1; // increment count for existing unique value
      } else {
        uniqueValues[cellValueString] = 1; // add new unique value to object
      }
    });
  
    return uniqueValues;
  }
}