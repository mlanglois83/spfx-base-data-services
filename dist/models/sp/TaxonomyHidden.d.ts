import { SPItem } from "../";
export declare class TaxonomyHidden extends SPItem {
    id: number;
    termId: string;
    constructor(item: any);
    convert(): Promise<any>;
}
