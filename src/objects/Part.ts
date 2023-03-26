import {PO} from "./PO";

export class Part{
    PN: string;
    Desc: string;
    TotalVolume: bigint | null;
    TotalQty: number | null;
    RelatedPOs: typeof PO[] | null;

    constructor(PN: string, Desc: string) {
        this.PN = PN;
        this.Desc = Desc;
        this.TotalVolume = null;
        this.TotalQty = null;
        this.RelatedPOs = null;
    }
}