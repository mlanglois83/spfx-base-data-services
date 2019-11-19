import { UtilsService } from "../";
import { TaxonomyTerm } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { TaxonomyHiddenListService } from "../";
/**
 * Base service for sp termset operations
 */
export declare class BaseTermsetService<T extends TaxonomyTerm> extends BaseDataService<T> {
    protected taxonomyHiddenListService: TaxonomyHiddenListService;
    protected utilsService: UtilsService;
    protected itemType: (new (item?: any) => T);
    protected termsetnameorid: string;
    protected wssIds: any;
    /**
     * Associeted termset (pnpjs)
     */
    protected readonly termset: import("@pnp/sp-taxonomy").ITermSet;
    protected customSortOrder: string;
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    constructor(type: (new (item?: any) => T), termsetnameorid: string, tableName: string, cacheDuration?: number);
    getWssIds(termId: string): Promise<Array<number>>;
    /**
     * Retrieve all terms
     */
    protected getAll_Internal(): Promise<Array<T>>;
    getById_Internal(query: any): Promise<T>;
    protected get_Internal(query: any): Promise<Array<T>>;
    protected addOrUpdateItem_Internal(item: T): Promise<T>;
    protected deleteItem_Internal(item: T): Promise<void>;
    private getOrderedChildTerms;
    getAll(): Promise<Array<T>>;
}
