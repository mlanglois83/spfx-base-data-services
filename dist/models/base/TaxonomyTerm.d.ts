import { IBaseItem } from "../../interfaces/index";
export declare class TaxonomyTerm implements IBaseItem {
    wssids: Array<number>;
    id: string;
    title: string;
    path: string;
    customSortOrder?: string;
    customProperties: any;
    constructor(term: any);
    convert(): any;
}
