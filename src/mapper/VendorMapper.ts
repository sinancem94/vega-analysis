/*export {};

import { Mapper } from "./Mapper";

class VendorMapper extends Mapper<any, any> {
    constructor() {
      super({
        vendorCode: '',
        vendorName: ''
      });
    }
  
    map(item: any): any {
        const vendorFields = this.getVendorFields(item);
        return {
            vendorCode: vendorFields.vendorCode,
            vendorName: vendorFields.vendorName,
        }
    }
}

module.exports = VendorMapper*/