import {Part} from "./Part";
import {Vendor} from "./Vendor";

export class PO {
    vendor: Vendor;
    parts: typeof Part[];
}