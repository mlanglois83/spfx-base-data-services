import { BaseComponentContext } from "@microsoft/sp-component-base";
import { assign } from "@microsoft/sp-lodash-subset";
import { IBaseItem, IAddOrUpdateResult, IDataService } from "../../interfaces";
import { OfflineTransaction } from "../../models/index";
import { UtilsService } from "../index";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { Text } from "@microsoft/sp-core-library";
import { TransactionType, Constants } from "../../constants";


/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp 
 */
export abstract class BaseDataService<T extends IBaseItem> extends BaseService implements IDataService<T> {
    protected itemType: (new (item?: any) => T);
    protected transactionService: TransactionService;
    protected dbService: BaseDbService<T>;
    protected utilService: UtilsService;
    protected cacheDuration: number = -1;

    /**
     * Stored promises to avoid multiple calls
     */
    protected static promises = {};

    public updateLinkedItems?: (oldId: number | string, newId: number | string) => void;

    public get serviceName(): string {
        return this.constructor["name"];
    }

    /**
     * 
     * @param type type of items
     * @param context context of the current wp
     * @param tableName name of indexedDb table 
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration: number = -1) {
        super();
        this.itemType = type;
        this.cacheDuration = cacheDuration;
        this.dbService = new BaseDbService<T>(type, tableName);
        this.transactionService = new TransactionService();

        this.utilService = new UtilsService();
    }

    protected getCacheKey(key: string = "all"): string {
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, BaseService.Configuration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
    }

    protected getExistingPromise(key: string = "all"): Promise<any> {
        let pkey = this.serviceName + "-" + key;
        if (BaseDataService.promises[pkey]) {
            return BaseDataService.promises[pkey];
        }
        else return null;
    }

    protected storePromise(promise: Promise<any>, key: string = "all"): void {
        let pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = promise;
    }

    protected removePromise(key: string = "all"): void {
        let pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = undefined;
    }


    /***
     * 
     */
    protected async getCachedData(key: string = "all"): Promise<Date> {

        let cacheKey = this.getCacheKey(key);

        let lastDataLoadString = window.sessionStorage.getItem(cacheKey);
        let lastDataLoad: Date = null;

        if (lastDataLoadString) {
            lastDataLoad = new Date(JSON.parse(window.sessionStorage.getItem(cacheKey)));
        }

        return lastDataLoad;
    }


    /**
     * Cache has to be relaod ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseDataService
     */
    protected async needRefreshCache(key: string = "all"): Promise<boolean> {

        let result: boolean = this.cacheDuration == -1;
        //if cache defined
        if (!result) {

            let cachedDataDate = await this.getCachedData(key);
            if (cachedDataDate) {
                //add cache duration
                cachedDataDate.setMinutes(cachedDataDate.getMinutes() + this.cacheDuration);

                let now = new Date();

                //cache has expired
                result = cachedDataDate < now;
            } else {
                result = true;
            }

        }

        return result;
    }

    protected UpdateCacheData(key: string = "all") {
        let result: boolean = this.cacheDuration == -1;
        //if cache defined
        if (!result) {
            let cacheKey = this.getCacheKey(key);
            window.sessionStorage.setItem(cacheKey, JSON.stringify(new Date()));
        }

    }

    protected abstract getAll_Internal(): Promise<Array<T>>;

    /* 
     * Retrieve all elements from datasource depending on connection is enabled
     * If service is not configured as offline, an exception is thrown;
     */
    public async getAll(): Promise<Array<T>> {
        let promise = this.getExistingPromise();
        if (promise) {
            console.log(this.serviceName + " getAll : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result = new Array<T>();

                    //has to refresh cache
                    let reloadData = await this.needRefreshCache();
                    //if refresh is needed, test offline/online
                    if (reloadData) {
                        reloadData = await this.utilService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getAll_Internal();
                        await this.dbService.replaceAll(result);
                        this.UpdateCacheData();
                    }
                    else {
                        result = await this.dbService.getAll();
                    }


                    this.removePromise();
                    resolve(result);
                }
                catch (error) {
                    this.removePromise();
                    reject(error);
                }
            });
            this.storePromise(promise);
        }
        return promise;

    }

    protected abstract get_Internal(query: any): Promise<Array<T>>;


    public async get(query: any): Promise<Array<T>> {
        let keyCached = super.hashCode(query).toString();
        let promise = this.getExistingPromise(keyCached);
        if (promise) {
            console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result = new Array<T>();
                    //has to refresh cache
                    let reloadData = await this.needRefreshCache(keyCached);
                    //if refresh is needed, test offline/online
                    if (reloadData) {
                        reloadData = await this.utilService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.get_Internal(query);
                        await this.dbService.addOrUpdateItems(result, query);
                        this.UpdateCacheData(keyCached);
                    }
                    else {
                        result = await this.dbService.get(query);
                    }

                    this.removePromise(keyCached);
                    resolve(result);
                }
                catch (error) {
                    this.removePromise(keyCached);
                    reject(error);
                }
            });
            this.storePromise(promise, keyCached);
        }
        return promise;
    }

    protected abstract getById_Internal(id: number): Promise<T>;

    public async getById(id: number): Promise<T> {
        let keyCached = "getById_" + id.toString();
        let promise = this.getExistingPromise(keyCached);
        if (promise) {
            console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<T>(async (resolve, reject) => {
                try {
                    let result: T;

                    let reloadData = await this.needRefreshCache(keyCached);
                    //if refresh is needed, test offline/online
                    if (reloadData) {
                        reloadData = await this.utilService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getById_Internal(id);
                        await this.dbService.addOrUpdateItems([result], keyCached);
                        this.UpdateCacheData(super.hashCode(keyCached).toString());
                    }
                    else {
                        let temp = await this.dbService.get(keyCached);
                        if (temp && temp.length > 0) { result = temp[0]; }
                    }

                    this.removePromise(keyCached);
                    resolve(result);
                }
                catch (error) {
                    this.removePromise(keyCached);
                    reject(error);
                }
            });
            this.storePromise(promise, keyCached);
        }
        return promise;
    }

    protected abstract addOrUpdateItem_Internal(item: T): Promise<T>;


    public async addOrUpdateItem(item: T): Promise<IAddOrUpdateResult<T>> {
        let result: IAddOrUpdateResult<T> = null;
        let itemResult: T = null;
        let isconnected = await this.utilService.CheckOnline();
        if (isconnected) {
            try {
                itemResult = await this.addOrUpdateItem_Internal(item);
                await this.dbService.addOrUpdateItem(itemResult);
                result = {
                    item: itemResult
                };
            } catch (error) {
                if (error.name === Constants.Errors.ItemVersionConfict) {
                    itemResult = await this.getById_Internal(<number>item.id);
                    await this.dbService.addOrUpdateItems([itemResult]);
                    result = {
                        item: itemResult,
                        error: error
                    };
                }
                else {
                    result = {
                        item: item,
                        error: error
                    };
                }

            }
        }
        else {
            result = await this.dbService.addOrUpdateItem(item);
            // create a new transaction
            let ot: OfflineTransaction = new OfflineTransaction();
            ot.itemData = assign({}, result.item);
            ot.itemType = result.item.constructor["name"];
            ot.serviceName = this.serviceName;
            ot.title = TransactionType.AddOrUpdate;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return result;
    }

    protected abstract deleteItem_Internal(item: T): Promise<void>;

    public async deleteItem(item: T): Promise<void> {
        let isconnected = await this.utilService.CheckOnline();
        if (isconnected) {
            await this.deleteItem_Internal(item);
            await this.dbService.deleteItem(item);
        }
        else {
            await this.dbService.deleteItem(item);

            // create a new transaction
            let ot: OfflineTransaction = new OfflineTransaction();
            ot.itemData = assign({}, item);
            ot.itemType = item.constructor["name"];
            ot.serviceName = this.serviceName;
            ot.title = TransactionType.Delete;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return null;
    }

}