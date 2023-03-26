/*export {};

import { Mapper } from "./Mapper";

export class PartMapper extends Mapper<any, any> {
    constructor() {
        super({
          partNumber: '',
          partDesc: '',
          unitPrices: '',
          unitCurrency: '',
          vendorCode: '',
          vendorName: ''
        });
    }
  
    map(item: any): any {
      const vendorFields = this.getVendorFields(item);
      const partFields = this.getPartFields(item);
      return {
        partNumber: partFields.partNumber,
        partDesc: partFields.partDesc,
        unitPrices: partFields.unitPrices,
        unitCurrency: partFields.unitCurrency,
        vendorCode: vendorFields.vendorCode,
        vendorName: vendorFields.vendorName,
      }
    }
}

module.exports = PartMapper*/
  