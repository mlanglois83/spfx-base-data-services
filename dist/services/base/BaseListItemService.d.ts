import { List } from "@pnp/sp";
import { IBaseItem } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { TaxonomyTerm, OfflineTransaction } from "../../models";
/**
 *
 * Base service for sp list items operations
 */
export declare class BaseListItemService<T extends IBaseItem> extends BaseDataService<T> {
    /***************************** Fields and properties **************************************/
    protected listRelativeUrl: string;
    protected initValues: any;
    protected tardiveLinks: any;
    readonly ItemFields: any;
    /**
     * Associeted list (pnpjs)
     */
    protected readonly list: List;
    /***************************** Constructor **************************************/
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param listRelativeUrl list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, tableName: string, cacheDuration?: number);
    /***************************** External sources init and access **************************************/
    private initialized;
    protected readonly isInitialized: boolean;
    private initPromise;
    protected init_internal(): Promise<void>;
    Init(): Promise<void>;
    private getServiceInitValues;
    /****************************** get item methods ***********************************/
    private getItemFromRest;
    private setFieldValue;
    /****************************** Send item methods ***********************************/
    private getSPRestItem;
    private setRestFieldValue;
    /********************** SP Fields conversion helpers *****************************/
    private convertTaxonomyFieldValue;
    private convertSingleUserFieldValue;
    /**
     *
     * @param wssid
     * @param terms
     */
    getTaxonomyTermByWssId<T extends TaxonomyTerm>(wssid: number, terms: Array<T>): T;
    /******************************************* Cache Management *************************************************/
    /**
     * Cache has to be reloaded ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    protected needRefreshCache(key?: string): Promise<boolean>;
    /***************** SP Calls associated to service standard operations ********************/
    /**
     * Get items by query
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected get_Internal(query: any): Promise<Array<T>>;
    /**
     * Get an item by id
     * @param id item id
     */
    protected getItemById_Internal(id: number): Promise<T>;
    /**
     * Get a list of items by id
     * @param id item id
     */
    protected getItemsById_Internal(ids: Array<number>): Promise<Array<T>>;
    /**
     * Retrieve all items
     *
     */
    protected getAll_Internal(): Promise<Array<T>>;
    /**
     * Add or update an item
     * @param item SPItem derived object to be converted
     */
    protected addOrUpdateItem_Internal(item: T): Promise<T>;
    /**
     * Delete an item
     * @param item SPItem derived class to be deleted
     */
    protected deleteItem_Internal(item: T): Promise<void>;
    /************************** Query filters ***************************/
    /**
     * Retrive all fields to include in odata setect parameter
     */
    private getOdataFieldNames;
    /**
     * convert full item to db format (with links only)
     * @param item full provisionned item
     */
    protected convertItemToDbFormat(item: T): T;
    /**
     * populate item from db storage
     * @param item db item with links in __internalLinks fields
     */
    mapItem(item: T): T;
    updateLinkedTransactions(oldId: number, newId: number, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>>;
    private updateLinksInDb;
    private updateWssIds;
}
