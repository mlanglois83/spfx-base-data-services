import { assign } from "@microsoft/sp-lodash-subset";
import { IBaseItem, IAddOrUpdateResult, IDataService } from "../../interfaces";
import { OfflineTransaction } from "../../models/index";
import { UtilsService } from "../index";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { Text } from "@microsoft/sp-core-library";
import { TransactionType, Constants } from "../../constants";
import { ServicesConfiguration } from "../..";


/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp 
 */
export abstract class BaseDataService<T extends IBaseItem> extends BaseService implements IDataService<T> {
    private itemModelType: (new (item?: any) => T);
    protected transactionService: TransactionService;
    protected dbService: BaseDbService<T>;
    protected cacheDuration: number = -1;

    public get ItemFields() {
        return {};
    }
    /**
     * Stored promises to avoid multiple calls
     */
    protected static promises = {};

    public get serviceName(): string {
        return this.constructor["name"];
    }
    
    public get itemType(): (new (item?: any) => T) {
        return this.itemModelType;
    }

    public async Init(): Promise<void> {
    }

    /**
     * 
     * @param type type of items
     * @param context context of the current wp
     * @param tableName name of indexedDb table 
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration: number = -1) {
        super();
        this.itemModelType = type;
        this.cacheDuration = cacheDuration;
        this.dbService = new BaseDbService<T>(type, tableName);
        this.transactionService = new TransactionService();
    }

    protected getCacheKey(key: string = "all"): string {
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
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
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getAll_Internal();
                        let convresult = await Promise.all(result.map((res) => {
                            return this.convertItemToDbFormat(res);
                        }));
                        await this.dbService.replaceAll(convresult);
                        this.UpdateCacheData();
                    }
                    else {
                        let tmp = await this.dbService.getAll();
                        result = await Promise.all(tmp.map((res) => {
                            return this.mapItem(res);
                        }));
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
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.get_Internal(query);
                        let convresult = await Promise.all(result.map((res) => {
                            return this.convertItemToDbFormat(res);
                        }));
                        await this.dbService.addOrUpdateItems(convresult, query);
                        this.UpdateCacheData(keyCached);
                    }
                    else {
                        let tmp = await this.dbService.get(query);
                        result = await Promise.all(tmp.map((res) => {
                            return this.mapItem(res);
                        }));
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

    protected abstract getItemById_Internal(id: number | string): Promise<T>;

    public async getItemById(id: number): Promise<T> {
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
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getItemById_Internal(id);
                        let converted = await this.convertItemToDbFormat(result);
                        await this.dbService.addOrUpdateItem(converted);
                        this.UpdateCacheData(super.hashCode(keyCached).toString());
                    }
                    else {
                        let temp = await this.dbService.getItemById(id);
                        if (temp) { 
                            result = await this.mapItem(temp); 
                        }
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

    protected abstract getItemsById_Internal(ids: Array<number | string>): Promise<Array<T>>;

    public async getItemsById(ids: Array<number | string>): Promise<Array<T>> {
        let keyCached = "getByIds_" + ids.join();
        let promise = this.getExistingPromise(keyCached);
        if (promise) {
            console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let results: Array<T>;

                    let reloadData = await this.needRefreshCache(keyCached);
                    //if refresh is needed, test offline/online
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        results = await this.getItemsById_Internal(ids);
                        let convresults = await Promise.all(results.map(async (res) => {
                            return this.convertItemToDbFormat(res)
                        }));
                        await this.dbService.addOrUpdateItems(convresults);
                        this.UpdateCacheData(super.hashCode(keyCached).toString());
                    }
                    else {
                        let tmp = await this.dbService.getItemsById(ids);
                        results = await Promise.all(tmp.map((res) => {
                            return this.mapItem(res);
                        }));
                    }
                    this.removePromise(keyCached);
                    resolve(results);
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

        let isconnected = true
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            try {
                itemResult = await this.addOrUpdateItem_Internal(item);
                let converted = await this.convertItemToDbFormat(itemResult);
                await this.dbService.addOrUpdateItem(converted);
                result = {
                    item: itemResult
                };
            } catch (error) {
                console.error(error);
                if (error.name === Constants.Errors.ItemVersionConfict) {
                    itemResult = await this.getItemById_Internal(item.id);
                    let converted = await this.convertItemToDbFormat(itemResult);
                    await this.dbService.addOrUpdateItem(converted);
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
            let dbItem = await this.convertItemToDbFormat(item);
            let resultitem = await this.dbService.addOrUpdateItem(dbItem);
            result = {
                item: item,
                error: resultitem.error
            };
            // update id (only field modified in db)
            result.item.id = resultitem.item.id;
            // create a new transaction
            let ot: OfflineTransaction = new OfflineTransaction();
            ot.itemData = assign({}, dbItem);
            ot.itemType = result.item.constructor["name"];
            ot.title = TransactionType.AddOrUpdate;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return result;
    }

    protected abstract deleteItem_Internal(item: T): Promise<void>;

    public async deleteItem(item: T): Promise<void> {
        let isconnected = true
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            await this.deleteItem_Internal(item);
            await this.dbService.deleteItem(item);
        }
        else {
            await this.dbService.deleteItem(item);

            // create a new transaction
            let ot: OfflineTransaction = new OfflineTransaction();
            let converted = await this.convertItemToDbFormat(item);
            ot.itemData = assign({}, converted);
            ot.itemType = item.constructor["name"];
            ot.title = TransactionType.Delete;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return null;
    }


    protected async convertItemToDbFormat(item: T): Promise<T> {        
        delete item.__internalLinks;
        return item;
    }

    public mapItem(item: T): Promise<T> {
        return Promise.resolve(item);
    }
    
    public async updateLinkedTransactions(oldId: number | string, newId: number | string, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        return nextTransactions;
    }

    public __getFromCache(id: string): Promise<T> {
        return this.dbService.getItemById(id);
    }

    public __getAllFromCache(): Promise<Array<T>> {
        return this.dbService.getAll();
    }

    public __updateCache(...items: Array<T>): Promise<Array<T>> {
        return this.dbService.addOrUpdateItems(items);
    }
}