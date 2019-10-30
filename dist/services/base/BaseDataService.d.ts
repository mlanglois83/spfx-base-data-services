import { IBaseItem, IAddOrUpdateResult, IDataService } from "../../interfaces";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp
 */
export declare abstract class BaseDataService<T extends IBaseItem> extends BaseService implements IDataService<T> {
    protected itemType: (new (item?: any) => T);
    protected transactionService: TransactionService;
    protected dbService: BaseDbService<T>;
    protected cacheDuration: number;
    /**
     * Stored promises to avoid multiple calls
     */
    protected static promises: {};
    updateLinkedItems?: (oldId: number | string, newId: number | string) => void;
    readonly serviceName: string;
    /**
     *
     * @param type type of items
     * @param context context of the current wp
     * @param tableName name of indexedDb table
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration?: number);
    protected getCacheKey(key?: string): string;
    protected getExistingPromise(key?: string): Promise<any>;
    protected storePromise(promise: Promise<any>, key?: string): void;
    protected removePromise(key?: string): void;
    /***
     *
     */
    protected getCachedData(key?: string): Promise<Date>;
    /**
     * Cache has to be relaod ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseDataService
     */
    protected needRefreshCache(key?: string): Promise<boolean>;
    protected UpdateCacheData(key?: string): void;
    protected abstract getAll_Internal(): Promise<Array<T>>;
    getAll(): Promise<Array<T>>;
    protected abstract get_Internal(query: any): Promise<Array<T>>;
    get(query: any): Promise<Array<T>>;
    protected abstract getById_Internal(id: number | string): Promise<T>;
    getById(id: number): Promise<T>;
    protected abstract addOrUpdateItem_Internal(item: T): Promise<T>;
    addOrUpdateItem(item: T): Promise<IAddOrUpdateResult<T>>;
    protected abstract deleteItem_Internal(item: T): Promise<void>;
    deleteItem(item: T): Promise<void>;
}
