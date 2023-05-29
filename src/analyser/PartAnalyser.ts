import * as ExcelJS from 'exceljs';
const {Part} = require("./../objects/Part");
import { TableAnalyser, PartAnalysisResult, AnalysisResult } from "./TableAnalyser";
import { FilterFields, Fields } from "./../mapper/Mapper"

export class PartAnalyser extends TableAnalyser{

    parts: typeof Part[] | null;
    declare analysisResult: AnalysisResult;

    constructor(worksheet: ExcelJS.Worksheet) {
        super(worksheet);
        this.parts = null;
        this.analysisResult = {
            success: false,
            totalVolume: {},
            uniquePartCount: 0,
            uniqueVendorCount: 0,
            parts: [],
          };
    }

    reset(): void{
        super.reset();
        this.parts = null;
        this.analysisResult = {
            success: false,
            totalVolume: {},
            uniquePartCount: 0,
            uniqueVendorCount: 0,
            parts: [],
          };
    }

    analyze(fields: Fields): AnalysisResult {
        
        var partNoColNum = this.getColumnNumberOfField(fields.part.partNumber);
        const uniquePNs = this.getUniqueValuesWithRows(Number(partNoColNum));
        
        uniquePNs.forEach((partTable) => {
            const rows: any[] = [];
            partTable.rowIndexes.forEach((index) => {
                const row = this.worksheet.getRow(index);
                rows.push(row);
            });
            
            let filters: FilterFields = fields.filter;
            var analysis = this.analyzeRows(fields, rows, filters);

            if (analysis.partNo != ""){

                this.analysisResult.parts.push(analysis);
            }

        });
        
        this.analysisResult.parts.sort((part1: PartAnalysisResult, part2: PartAnalysisResult) => 
                                        part2.quantities.reduce((a,b) => a+b, 0) - part1.quantities.reduce((a,b) => a+b, 0));

        this.analysisResult.uniqueVendorCount = Object.keys(this.analysisResult.parts.reduce((vendors, part) => ({ ...vendors, ...part.vendors }), {})).length;
        //this.getUniqueValuesOnColumn(Number(this.getColumnNumberOfField(fields.vendor.vendorCode))).length;
        this.analysisResult.uniquePartCount = this.analysisResult.parts.length;// Object.keys(uniquePNs).length;
        this.analysisResult.success = true;
        return this.analysisResult;
    }

    protected analyzeRows(fields: Fields, rows: ExcelJS.Row[], filters: FilterFields): PartAnalysisResult {

        var partNoColNum = this.getColumnNumberOfField(fields.part.partNumber);
        var partDescColNum = this.getColumnNumberOfField(fields.part.partDesc);
        var partQtyColNum = this.getColumnNumberOfField(fields.part.partQuantity);
        var unitPriceColNum = this.getColumnNumberOfField(fields.part.unitPrices);
        var unitCurrColNum = this.getColumnNumberOfField(fields.part.unitCurrency);
        var purchaseOrderColNum = this.getColumnNumberOfField(fields.part.purchaseOrder);
        var orderTypeColNum = this.getColumnNumberOfField(fields.part.orderType);
        var vendorCodeColNum = this.getColumnNumberOfField(fields.vendor.vendorCode);
        var vendorNameColNum = this.getColumnNumberOfField(fields.vendor.vendorName);

        var analysisRes: PartAnalysisResult = {
            partNo: "", partDesc: "", quantities: [],
            purchaseOrders: {}, prices: {}, vendors: {}       
        };

        let partNo = rows[0].getCell(Number(partNoColNum)).value as string;
        if ((filters.partFilter.length > 0 && filters.partFilter.every(pn => {return pn !== partNo;}))){
            return analysisRes;
        }

        for(let i = 0; i < rows.length; i++){
            var row = rows[i];

            //analysisRes.partNo = row.getCell(Number(partNoColNum)).value as string;
            var partDesc = row.getCell(Number(partDescColNum)).value as string;
            var vendorCode = row.getCell(Number(vendorCodeColNum)).value as string;

            var orderType = orderTypeColNum ? (row.getCell(Number(orderTypeColNum)).value as string) : "?";
            orderType = orderType ? orderType : "?";
            var currency = unitCurrColNum ? (row.getCell(Number(unitCurrColNum)).value as string) : "?";
            currency = currency ? currency : "?";
            var vendorName = vendorNameColNum ? (row.getCell(Number(vendorNameColNum)).value as string) : "?";

            //check filters
            if( (filters.typeFilter.length > 0 && filters.typeFilter.every(ot => {return ot !== orderType;})) || 
                (filters.currencyFilter.length > 0 && filters.currencyFilter.every(c => {return c !== currency;})) ||
                (filters.vendorFilter.length > 0 && filters.vendorFilter.every(vc => {return vc !== vendorCode;}))) {
                    
                    continue;
            }

            analysisRes.partNo = partNo;
            analysisRes.partDesc = partDesc;
            
            var qty = row.getCell(Number(partQtyColNum)).value as number;
            analysisRes.quantities.push(qty);
            
            if(unitPriceColNum){
                
                var price = row.getCell(Number(unitPriceColNum)).value as number;
                
                if(currency){
                    if (analysisRes.prices[currency]){
                        analysisRes.prices[currency].push(price);
                    }
                    else{
                        analysisRes.prices[currency] = [price];
                    }
                }
            }
            
            var purchaseOrder = purchaseOrderColNum ? (row.getCell(Number(purchaseOrderColNum)).value as string) : "?";

            if(orderType){
                if(analysisRes.purchaseOrders[orderType]){
                    analysisRes.purchaseOrders[orderType].push(purchaseOrder);
                }
                else{
                    analysisRes.purchaseOrders[orderType] = [purchaseOrder];
                }
            }
            
            if (analysisRes.vendors[vendorCode]){
                analysisRes.vendors[vendorCode].push(vendorName);
            }
            else{
                analysisRes.vendors[vendorCode] = [vendorName];
            }
        }

        return analysisRes;
    }
}