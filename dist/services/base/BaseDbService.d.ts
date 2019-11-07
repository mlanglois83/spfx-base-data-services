import { DB, ObjectStore } from "idb";
import { IBaseItem } from "../../interfaces/IBaseItem";
import { IDataService } from "../../interfaces/IDataService";
import { BaseService } from "./BaseService";
import { IAddOrUpdateResult } from "../../interfaces";
/**
 * Base classe for indexedDB interraction using SP repository
 */
export declare class BaseDbService<T extends IBaseItem> extends BaseService implements IDataService<T> {
    protected tableName: string;
    protected db: DB;
    protected itemType: (new (item?: any) => T);
    /**
     *
     * @param tableName : Name of the db table the service interracts with
     */
    constructor(type: (new (item?: any) => T), tableName: string);
    protected getChunksRegexp(fileUrl: any): RegExp;
    protected getAllKeysInternal<TKey extends string | number>(store: ObjectStore<T, TKey>): Promise<Array<TKey>>;
    protected getNextAvailableKey(): Promise<number>;
    /**
     * Opens indexed db, update structure if needed
     */
    protected OpenDb(): Promise<void>;
    /**
     * Add or update an item in DB and returns updated item
     * @param item Item to add or update
     */
    addOrUpdateItem(item: T): Promise<IAddOrUpdateResult<T>>;
    deleteItem(item: T): Promise<void>;
    get(query?: string): Promise<Array<T>>;
    /**
     * add items in table (ids updated)
     * @param newItems
     */
    addOrUpdateItems(newItems: Array<T>, query?: string): Promise<Array<T>>;
    /**
     * Retrieve all items from db table
     */
    getAll(): Promise<Array<T>>;
    /**
     * Clear table and insert new items
     * @param newItems Items to insert in place of existing
     */
    replaceAll(newItems: Array<T>): Promise<void>;
    /**
     * Clear table
     */
    clear(): Promise<void>;
    getById(id: number | string): Promise<T>;
}
