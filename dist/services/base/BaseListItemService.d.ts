import { List } from "@pnp/sp";
import { IBaseItem } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
/**
 *
 * Base service for sp list items operations
 */
export declare class BaseListItemService<T extends IBaseItem> extends BaseDataService<T> {
    protected itemType: (new (item?: any) => T);
    protected listRelativeUrl: string;
    protected readonly ItemFields: any;
    /**
     * If defined, will be called on each internal getter call to make associations (taxonomy, users...)
     */
    protected associate?: (...items: Array<T>) => Promise<Array<T>>;
    readonly listItemType: (new (item?: any) => T);
    /**
     * Associeted list (pnpjs)
     */
    protected readonly list: List;
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param listRelativeUrl list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, tableName: string, cacheDuration?: number);
    /**
     * Cache has to be reloaded ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    protected needRefreshCache(key?: string): Promise<boolean>;
    /**
     *
     * TODO avoid getting all fields
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected get_Internal(query: any): Promise<Array<T>>;
    /**
     *
     * @param id
     */
    protected getById_Internal(id: number): Promise<T>;
    /**
     * Retrieve all items
     *
     * TODO avoid getting all fields
     */
    protected getAll_Internal(): Promise<Array<T>>;
    protected addOrUpdateItem_Internal(item: T): Promise<T>;
    protected deleteItem_Internal(item: T): Promise<void>;
}
