interface PartFields {
  partNumber: string,
  partDesc: string,
  partQuantity: string,
  unitPrices: string,
  unitCurrency: string,
  purchaseOrder: string,
  orderType: string,
}

interface VendorFields {
  vendorCode: string,
  vendorName: string,
}

export interface FilterTableFields {
  sheetName: string,
  columnName: string,
  isInclude: boolean,
}

export interface FilterFields {
  typeFilter: string[],
  currencyFilter: string[],
  vendorFilter: string[],
  partFilter: string[],
}

export interface Fields {
  sheetName: string;
  filter: FilterFields;
  part: PartFields;
  vendor: VendorFields;
}

export class FieldMapper {
    
  constructor() {
    this.vendor = { vendorCode: "", vendorName: "" };
    this.part = { partNumber: "", partDesc: "", partQuantity: "", unitPrices: "", unitCurrency: "", purchaseOrder: "", orderType: "" };
    this.filter = { typeFilter: [], currencyFilter: [], vendorFilter: [], partFilter: [] };
    this.filterTable = {sheetName: "", columnName: "", isInclude: true};
  }

  resetFields() {
    this.vendor = { vendorCode: "", vendorName: "" };
    this.part = { partNumber: "", partDesc: "", partQuantity: "", unitPrices: "", unitCurrency: "", purchaseOrder: "", orderType: "" };
    this.filter = { typeFilter: [], currencyFilter: [], vendorFilter: [], partFilter: [] };
    this.filterTable = {sheetName: "", columnName: "", isInclude: true};
  }

  sheetName: string;
  vendor: VendorFields;
  part: PartFields;
  filter: FilterFields;
  filterTable: FilterTableFields;

  joinedFields(): Fields {
    return {sheetName: this.sheetName, part: this.part, vendor: this.vendor, filter: this.filter};// { part: this.part, ...this.vendor, ...this.main, ...this.filter };
  }

  mapFields(form: any) {

    this.resetFields();

    this.sheetName = form.sheet_name[0] ?? this.getDefault();
    this.MapVendor(form);
    this.MapPart(form);
    this.MapFilter(form);
  }

  MapVendor(form: any){
    this.vendor.vendorCode = form.vendor_code[0] ?? this.getDefault();
    this.vendor.vendorName = form.vendor_name[0] ?? this.getDefault();
    return this.vendor;
  }

  MapPart(form: any){
    this.part.partNumber = form.part_number[0] ?? this.getDefault();
    this.part.partDesc = form.part_description[0] ?? this.getDefault();
    this.part.partQuantity = form.part_quantity[0] ?? this.getDefault();
    this.part.unitPrices = form.unit_prices[0] ?? this.getDefault();
    this.part.unitCurrency = form.unit_currency[0] ?? this.getDefault();
    this.part.purchaseOrder = form.purchase_order[0] ?? this.getDefault();
    this.part.orderType = form.order_type[0] ?? this.getDefault();
    return this.part;
  }

  MapFilter(form: any){

    for(let filterKey in form){
      //console.log(filterKey);

      if(this.filter.hasOwnProperty(filterKey)){
        console.log(form[filterKey]);
      }
    }

    if(form.ot_filter){
      for(let i = 0; i < form.ot_filter.length; i++){
        this.filter.typeFilter.push(form.ot_filter[i]);
      }
    }
    
    if(form.curr_filter){
      for(let i = 0; i < form.curr_filter.length; i++){
        this.filter.currencyFilter.push(form.curr_filter[i]);
      }
    }
    
    if(form.vc_filter){
      for(let i = 0; i < form.vc_filter.length; i++){
        this.filter.vendorFilter.push(form.vc_filter[i]);
      }
    }
    
    if(form.pn_filter){
      for(let i = 0; i < form.pn_filter.length; i++){
        this.filter.partFilter.push(form.pn_filter[i]);
      }
    }

    return this.filter;
  }

  MapFilterTable(form:any){
    this.filterTable.sheetName = form.sheet_name[0] ?? this.getDefault();
    this.filterTable.columnName = form.column_name[0] ?? this.getDefault();
    this.filterTable.isInclude = form.is_include[0] === "exclude" ? false : true;
  }

  protected getDefault(){
    return "";
  }
}