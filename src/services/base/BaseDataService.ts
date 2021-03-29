import { assign, cloneDeep, findIndex } from "@microsoft/sp-lodash-subset";
import { IDataService, IQuery, ILogicalSequence, IPredicate, IFieldDescriptor } from "../../interfaces";
import { BaseItem, OfflineTransaction, TaxonomyTerm } from "../../models";
import { UtilsService } from "../UtilsService";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { Text } from "@microsoft/sp-core-library";
import { TransactionType, Constants, LogicalOperator, TestOperator, QueryToken, FieldType } from "../../constants";
import { ServicesConfiguration } from "../../configuration";
import { stringIsNullOrEmpty } from "@pnp/common";
import { ServiceFactory } from "../ServiceFactory";
import { Decorators } from "../../decorators";
const trace = Decorators.trace;

/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp 
 */
export abstract class BaseDataService<T extends BaseItem> extends BaseService implements IDataService<T> {
    private itemModelType: (new (item?: any) => T);

    protected transactionService: TransactionService;
    protected dbService: BaseDbService<T>;
    protected cacheDuration = -1;      

    protected get debug(): boolean {
        return ServicesConfiguration.configuration.debug === true;
    }

    public get serviceName(): string {
        return this.constructor["name"];
    }

    public get itemType(): (new (item?: any) => T) {
        return this.itemModelType;
    }

    public cast<Tdest extends BaseDataService<T>> (): Tdest {
        return this as unknown as Tdest;
    }

    /**
     * 
     * @param type - type of items
     * @param context - context of the current wp
     */
    constructor(type: (new (item?: any) => T), cacheDuration = -1) {
        super();
        if(ServiceFactory.isServiceManaged(type["name"]) && !ServiceFactory.isServiceInitializing(type["name"])) {
            console.warn(`Service constructor called out of Service factory. Please use ServiceFactory.getService(${type["name"]}) or ServiceFactory.getServiceByModelName("${type["name"]}")`);
        }
        this.itemModelType = type;
        this.cacheDuration = cacheDuration;
        this.dbService = new BaseDbService<T>(type, type["name"]);
        this.transactionService = new TransactionService();
    }

    /***************************** External sources init and access **************************************/
    protected initValues: {[modelName: string]: BaseItem[]} = {};


    protected initialized = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }
    private initPromise: Promise<void> = null;

    protected async init_internal(): Promise<void> {
        return;
    }
    
    @trace()
    private async initLinkedFields(): Promise<void> {
        const fields = this.ItemFields;
        const models: string[] = [];
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDescription = fields[key];
                if (fieldDescription.modelName &&
                    models.indexOf(fieldDescription.modelName) === -1 &&
                    fieldDescription.fieldType !== FieldType.Lookup &&
                    fieldDescription.fieldType !== FieldType.LookupMulti &&
                    fieldDescription.fieldType !== FieldType.Json) {
                    models.push(fieldDescription.modelName);
                }
            }
        }
        await Promise.all(models.map(async (modelName) => {
            if (!this.initValues[modelName]) {
                const service = ServiceFactory.getServiceByModelName(modelName);
                const values = await service.getAll();
                this.initValues[modelName] = values;
            }
        }));
    }

    public async Init(): Promise<void> {
        if (!this.initPromise) {
            this.initPromise = new Promise<void>(async (resolve, reject) => {
                if (this.initialized) {
                    resolve();
                }
                else {
                    this.initValues = {};
                    try {
                        if (this.init_internal) {
                            await this.init_internal();
                        }
                        await this.initLinkedFields();
                        this.initialized = true;
                        this.initPromise = null;
                        resolve();
                    }
                    catch (error) {
                        this.initPromise = null;
                        reject(error);
                    }
                }
            });
        }
        return this.initPromise;

    }

    protected getServiceInitValues<Tvalue extends BaseItem>(model: new (data?: any) => Tvalue): Tvalue[] {
        return this.getServiceInitValuesByName<Tvalue>(model["name"]);
    }

    protected getServiceInitValuesByName<Tvalue extends BaseItem>(modelName: string): Tvalue[] {
        return this.initValues[modelName] as Tvalue[];
    }

    protected updateInitValues(modelName: string, ...items: BaseItem[]): void {
        this.initValues[modelName] = this.initValues[modelName] || [];
        items.forEach(i => {
            const idx = findIndex(this.initValues[modelName], iv => iv.id === i.id);
            if(idx !== -1) {
                this.initValues[modelName][idx] = i;
            }
            else {
                this.initValues[modelName].push(i);
            }
        });
    }
    /*************************************************************************************************************************/

    /********************************************* Fields Management *********************************************************/
    
    public get ItemFields(): {[propertyName: string]: IFieldDescriptor} {        
        return ServiceFactory.getModelFields(this.itemType["name"]);
    }
    public get Identifier(): Array<string> {

        const fields = this.ItemFields;
        const fieldNames = new Array<string>();

        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key];
                if (fieldDesc.identifier) {
                    fieldNames.push(key);
                    break;
                }
            }
        }
        return fieldNames;
    }


    /*****************************************************************************************************************************************************************/



    /**************************************************************** Promise Concurency ******************************************************************************/

    /**
     * Stored promises to avoid multiple calls
     */
     protected static promises = {};

    protected getExistingPromise(key = "all"): Promise<any> {
        const pkey = this.serviceName + "-" + key;
        if (BaseDataService.promises[pkey]) {
            return BaseDataService.promises[pkey];
        }
        else return null;
    }

    ///add semaphore for store and remove
    protected storePromise(promise: Promise<any>, key = "all"): void {
        const pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = promise;
    }

    ///add semaphore
    protected removePromise(key = "all"): void {
        const pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = undefined;
    }
    /*****************************************************************************************************************************************************************/

    /************************************************************************* Cache expiration ************************************************************************************/
    
    protected getCacheKey(key = "all"): string {
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
    }
    /***
     * 
     */
    protected getCachedData(key = "all"): Date {

        const cacheKey = this.getCacheKey(key);

        const lastDataLoadString = window.sessionStorage.getItem(cacheKey);
        let lastDataLoad: Date = null;

        if (lastDataLoadString) {
            lastDataLoad = new Date(JSON.parse(window.sessionStorage.getItem(cacheKey)));
        }

        return lastDataLoad;
    }
    protected getIdLastLoad(id: number | string): Date {
        const cacheKey = this.getCacheKey("ids");
        const idTableString = window.sessionStorage.getItem(cacheKey);
        let lastDataLoad: Date = null;
        if (!stringIsNullOrEmpty(idTableString)) {
            const converted = JSON.parse(idTableString);
            if (converted[id]) {
                lastDataLoad = new Date(converted[id]);
            }
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
    protected async needRefreshCache(key = "all"): Promise<boolean> {

        let result: boolean = this.cacheDuration === -1;
        //if cache defined
        if (!result) {

            const cachedDataDate = this.getCachedData(key);
            if (cachedDataDate) {
                //add cache duration
                cachedDataDate.setMinutes(cachedDataDate.getMinutes() + this.cacheDuration);

                const now = new Date();

                //cache has expired
                result = cachedDataDate < now;
            } else {
                result = true;
            }

        }

        return result;
    }

    protected async getExpiredIds(...ids: Array<number | string>): Promise<Array<number | string>> {
        const expired = ids.filter((id) => {
            let result = true;
            const lastLoad = this.getIdLastLoad(id);
            if (lastLoad) {
                lastLoad.setMinutes(lastLoad.getMinutes() + this.cacheDuration);
                const now = new Date();
                //cache has expired
                result = lastLoad < now;
            }
            return result;
        });
        return expired;
    }

    protected UpdateIdsLastLoad(...ids: Array<number | string>): void {
        if (this.cacheDuration !== -1) {
            const cacheKey = this.getCacheKey("ids");
            const initTableString = sessionStorage.getItem(cacheKey);
            let idTable;
            if (!stringIsNullOrEmpty(initTableString)) {
                idTable = JSON.parse(initTableString) || {};
            }
            else {
                idTable = {};
            }
            const now = new Date();
            ids.forEach((id) => {
                idTable[id] = now;
            });
            window.sessionStorage.setItem(cacheKey, JSON.stringify(idTable));
        }
    }
    protected UpdateCacheData(key = "all"): void {
        const result: boolean = this.cacheDuration === -1;
        //if cache defined
        if (!result) {
            const cacheKey = this.getCacheKey(key);
            window.sessionStorage.setItem(cacheKey, JSON.stringify(new Date()));
        }

    }

    /*********************************************************************************************************************************************************/

    /*********************************************************** Data operations ***************************************************************************/
    protected abstract getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>>;

    /* 
     * Retrieve all elements from datasource depending on connection is enabled
     * If service is not configured as offline, an exception is thrown;
     */
    
    @trace()
    public async getAll(linkedFields?: Array<string>): Promise<Array<T>> {
        let promise = this.getExistingPromise();
        if (promise) {
            if(this.debug)
                console.log(this.serviceName + " getAll : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result: Array<T>;

                    //has to refresh cache
                    let reloadData = await this.needRefreshCache();
                    //if refresh is needed, test offline/online
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getAll_Internal(linkedFields);
                        const convresult = await Promise.all(result.map((res) => {
                            return this.convertItemToDbFormat(res);
                        }));
                        await this.dbService.replaceAll(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData();
                    }
                    else {
                        const tmp = await this.dbService.getAll();
                        result = await this.mapItems(tmp, linkedFields);
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

    
    protected abstract get_Internal(query: IQuery, linkedFields?: Array<string>): Promise<Array<T>>;

    
    @trace()
    public async get(query: IQuery, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(query).toString() + super.hashCode(linkedFields).toString();
        let promise = this.getExistingPromise(keyCached);
        if (promise) {
            if(this.debug)
                console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result: Array<T>;
                    //has to refresh cache
                    let reloadData = await this.needRefreshCache(keyCached);
                    //if refresh is needed, test offline/online
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.get_Internal(query, linkedFields);
                        //check if data exist for this query in database
                        let tmp = await this.dbService.get(query);
                        tmp = this.filterItems(query, tmp);

                        //if data exists trash them 
                        if (tmp && tmp.length > 0) {
                            await this.dbService.deleteItems(tmp);
                        }

                        const convresult = await Promise.all(result.map((res) => {
                            return this.convertItemToDbFormat(res);
                        }));
                        await this.dbService.addOrUpdateItems(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData(keyCached);
                    }
                    else {
                        const tmp = await this.dbService.get(query);
                        result = await this.mapItems(tmp, linkedFields);
                        // filter
                        result = this.filterItems(query, result);
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

    protected abstract getItemById_Internal(id: number | string, linkedFields?: Array<string>): Promise<T>;

    
    @trace()
    public async getItemById(id: number, linkedFields?: Array<string>): Promise<T> {
        const promiseKey = "getById_" + id.toString();
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
            if(this.debug)
                console.log(this.serviceName + " " + promiseKey + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<T>(async (resolve, reject) => {
                try {
                    let result: T;
                    const deprecatedIds = await this.getExpiredIds(id);
                    let reloadData = deprecatedIds.length > 0;
                    //if refresh is needed, test offline/online
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }

                    if (reloadData) {
                        result = await this.getItemById_Internal(id, linkedFields);
                        const converted = await this.convertItemToDbFormat(result);
                        await this.dbService.addOrUpdateItem(converted);
                        this.UpdateIdsLastLoad(id);
                    }
                    else {
                        const temp = await this.dbService.getItemById(id);
                        if (temp) {
                            const res = await this.mapItems([temp], linkedFields);
                            result = res.shift();
                        }
                    }

                    this.removePromise(promiseKey);
                    resolve(result);
                }
                catch (error) {
                    this.removePromise(promiseKey);
                    reject(error);
                }
            });
            this.storePromise(promise, promiseKey);
        }
        return promise;
    }

    protected abstract getItemsById_Internal(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>>;

    
    public async getItemsFromCacheById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const tmp = await this.dbService.getItemsById(ids);
        return this.mapItems(tmp, linkedFields);
    }
    
    @trace()
    public async getItemsById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const promiseKey = "getByIds_" + ids.join();
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
            if(this.debug)
                console.log(this.serviceName + " " + promiseKey + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                if (ids.length > 0) {
                    try {
                        let results: Array<T>;
                        const deprecatedIds = await this.getExpiredIds(...ids);
                        let reloadData = deprecatedIds.length > 0;
                        //if refresh is needed, test offline/online
                        if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                            reloadData = await UtilsService.CheckOnline();
                        }

                        if (reloadData) {
                            const expired = await this.getItemsById_Internal(deprecatedIds, linkedFields);
                            const tmpcached = await this.dbService.getItemsById(ids.filter((i) => { return deprecatedIds.indexOf(i) === -1; }));
                            const cached = await this.mapItems(tmpcached, linkedFields);
                            results = expired.concat(cached);
                            const convresults = await Promise.all(results.map(async (res) => {
                                return this.convertItemToDbFormat(res);
                            }));
                            await this.dbService.addOrUpdateItems(convresults);
                            this.UpdateIdsLastLoad(...ids);
                        }
                        else {
                            const tmp = await this.dbService.getItemsById(ids);
                            results = await this.mapItems(tmp, linkedFields);
                        }
                        this.removePromise(promiseKey);
                        resolve(results);
                    }
                    catch (error) {
                        this.removePromise(promiseKey);
                        reject(error);
                    }
                }
                else {
                    this.removePromise(promiseKey);
                    resolve([]);
                }
            });
            this.storePromise(promise, promiseKey);
        }
        return promise;
    }

    protected abstract addOrUpdateItem_Internal(item: T): Promise<T>;

    
    @trace()
    public async addOrUpdateItem(item: T): Promise<T> {
        let result: T = null;
        let itemResult: T = null;

        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            try {

                const oldId = item.id;
                itemResult = await this.addOrUpdateItem_Internal(item);
                if (oldId < -1) { // created item allready stored in db
                    this.dbService.deleteItem(item);
                }
                const converted = await this.convertItemToDbFormat(itemResult);
                await this.dbService.addOrUpdateItem(converted);
                this.UpdateIdsLastLoad(converted.id);
                result = itemResult;


            } catch (error) {
                console.error(error);
                if (error.name === Constants.Errors.ItemVersionConfict) {
                    itemResult = await this.getItemById_Internal(item.id);
                    const converted = await this.convertItemToDbFormat(itemResult);
                    await this.dbService.addOrUpdateItem(converted);
                    itemResult.error = error;
                    result = itemResult;
                    this.UpdateIdsLastLoad(converted.id);
                }
                else {
                    item.error = error;
                    result = item;
                }

            }
        }
        else {
            const dbItem = await this.convertItemToDbFormat(item);
            const resultitem = await this.dbService.addOrUpdateItem(dbItem);
            item.error = resultitem.error;
            result = item;
            // update id (only field modified in db)
            result.id = resultitem.id;
            // create a new transaction
            const ot: OfflineTransaction = new OfflineTransaction();
            ot.itemData = assign({}, dbItem);
            ot.itemType = result.constructor["name"];
            ot.title = TransactionType.AddOrUpdate;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return result;
    }

    protected abstract addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>>;

    
    @trace()
    public async addOrUpdateItems(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        let results: Array<T> = [];

        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            results = await this.addOrUpdateItems_Internal(items, onItemUpdated);
            const versionErrors = results.filter((res) => {
                return res.error && res.error.name === Constants.Errors.ItemVersionConfict;
            });
            // find back items with version error
            if (versionErrors.length > 0) {
                const spitems = await this.getItemsById_Internal(versionErrors.map(ve => ve.id));
                spitems.forEach((retrieved) => {
                    const idx = findIndex(versionErrors, { id: retrieved.id });
                    if (idx > -1) {
                        versionErrors[idx] = retrieved;
                    }
                });
            }
            for (const item of results) {
                const converted = await this.convertItemToDbFormat(item);
                await this.dbService.addOrUpdateItem(converted);
                this.UpdateIdsLastLoad(converted.id);
            }


        }
        else {
            for (const item of items) {
                const copy = cloneDeep(item);
                const dbItem = await this.convertItemToDbFormat(item);
                const resultitem = await this.dbService.addOrUpdateItem(dbItem);
                copy.error = resultitem.error;
                // update id (only field modified in db)
                copy.id = resultitem.id;
                results.push(copy);
                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                ot.itemData = assign({}, dbItem);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.AddOrUpdate;
                await this.transactionService.addOrUpdateItem(ot);
                if (onItemUpdated) {
                    onItemUpdated(copy, item);
                }
            }

        }

        return results;
    }

    protected abstract deleteItem_Internal(item: T): Promise<T>;

    
    @trace()
    public async deleteItem(item: T): Promise<T> {
        if(typeof(item.id) === "number" && item.id === -1) {
            item.deleted = true;
        }
        else {
            let isconnected = true;
            if (ServicesConfiguration.configuration.checkOnline) {
                isconnected = await UtilsService.CheckOnline();
            }
            if (isconnected) {
                if(typeof(item.id) !== "number" || item.id > -1) {
                    item = await this.deleteItem_Internal(item);
                }
                if(item.deleted || item.id < -1) {
                    item = await this.dbService.deleteItem(item);
                }
            }
            else {
                item = await this.dbService.deleteItem(item);

                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                const converted = await this.convertItemToDbFormat(item);
                ot.itemData = assign({}, converted);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.Delete;
                await this.transactionService.addOrUpdateItem(ot);
            }
        }
        return item;
    }

    protected abstract deleteItems_Internal(items: Array<T>): Promise<Array<T>>;
    
    @trace()
    public async deleteItems(items: Array<T>): Promise<Array<T>> {
        items.filter(i => (typeof(i.id) === "number" && i.id === -1)).forEach(i => {
            i.deleted = true;
        });
        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            await this.deleteItems_Internal(items.filter(i => (typeof(i.id) !== "number" || i.id > -1)));
            await this.dbService.deleteItems(items.filter(i=>i.deleted || i.id < -1));
        }
        else { 
            await this.dbService.deleteItems(items.filter(i => i.id > -1));
            const transactions: Array<OfflineTransaction> = [];
            for (const item of items) {
                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                const converted = await this.convertItemToDbFormat(item);
                ot.itemData = assign({}, converted);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.Delete;
                transactions.push(ot);
            }   
            await this.transactionService.addOrUpdateItems(transactions);
        }

        return items;
    }

    
    @trace()
    public async persistItemData(data: any, linkedFields?: Array<string>, lookupLoaded?: boolean): Promise<T> {
        const result = await this.persistItemData_internal(data, linkedFields, lookupLoaded);
        const convresult = await this.convertItemToDbFormat(result);
        await this.dbService.addOrUpdateItem(convresult);
        this.UpdateIdsLastLoad(convresult.id);  
        return result;
    }

    protected abstract persistItemData_internal(data: any, linkedFields?: Array<string>, lookupLoaded?: boolean): Promise<T>;

    
    @trace()
    public async persistItemsData(data: any[], linkedFields?: Array<string>, lookupLoaded?: boolean): Promise<T[]> {
        const result = await this.persistItemsData_internal(data, linkedFields, lookupLoaded);
        const convresult = await Promise.all(result.map(r => this.convertItemToDbFormat(r)));
        await this.dbService.addOrUpdateItems(convresult);
        this.UpdateIdsLastLoad(...convresult.map(cr => cr.id));  
        return result;
    }

    protected async persistItemsData_internal(data: any[], linkedFields?: Array<string>, lookupLoaded?: boolean): Promise<T[]> {
        let result = null;
        if (data) {
            result = await Promise.all(data.map(d => this.persistItemData_internal(d, linkedFields, lookupLoaded)));
        }
        return result;
    }

    /*****************************************************************************************************************************************************************/

    /********************************************************************** Cached data management ******************************************************************************/

    protected async convertItemToDbFormat(item: T): Promise<T> {
        const result: T = cloneDeep(item);
        result.cleanBeforeStorage();
        return result;
    }

    @trace()
    public async mapItems(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> { // eslint-disable-line @typescript-eslint/no-unused-vars
        items.forEach(i => i.__clearEmptyInternalLinks());
        return items;
    }

    @trace()
    public async updateLinkedTransactions(oldId: number | string, newId: number | string, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        return nextTransactions;
    }

    public __getFromCache(id: number | string): Promise<T> {
        return this.dbService.getItemById(id);
    }

    public __getAllFromCache(): Promise<Array<T>> {
        return this.dbService.getAll();
    }

    public __updateCache(...items: Array<T>): Promise<Array<T>> {
        return this.dbService.addOrUpdateItems(items);
    }

    /**
     * Refresh cached data
     */
     public async refreshData(): Promise<void>  {
        // Invalidate cache
        const cacheKey = this.getCacheKey(); // Default key is "ALL"
        window.sessionStorage.removeItem(cacheKey);    
        // remove local cache
        this.initValues = {};  
        // Reload all data
        await this.getAll();
    }

    /*****************************************************************************************************************************************************************/

    /********************************************************************* Queries ************************************************************************************/
    private filterItems(query: IQuery, items: Array<T>): Array<T> {
        // filter items by test
        let results = query.test ? items.filter((i) => { return this.getTestResult(query.test, i); }) : cloneDeep(items);
        // order by
        if (query.orderBy) {
            results.sort(function (a, b) {
                for (const order of query.orderBy) {
                    const aKey = a[order.propertyName];
                    const bKey = b[order.propertyName];
                    if (typeof (aKey) === "string" || typeof (bKey) === "string") {
                        if ((aKey || "").localeCompare(bKey || "") < 0) {
                            return order.ascending ? -1 : 1;
                        }
                        if ((aKey || "").localeCompare(bKey || "") > 0) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                    else if(aKey instanceof Date || bKey instanceof Date) {
                        const aval = aKey && aKey.getTime ? aKey.getTime() : 0;
                        const bval = bKey && bKey.getTime ? bKey.getTime() : 0;
                        if (aval < bval) {
                            return order.ascending ? -1 : 1;
                        }
                        if (aval > bval) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                    else if (aKey.id && bKey.id) {
                        if ((aKey.title || "").localeCompare(bKey.title || "") < 0) {
                            return order.ascending ? -1 : 1;
                        }
                        if ((aKey.title || "").localeCompare(bKey.title || "") > 0) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                    else {
                        if (aKey < bKey) {
                            return order.ascending ? -1 : 1;
                        }
                        if (aKey > bKey) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                }
                return 0;
            });
        }
        // Paged query
        if (query.lastId) {
            const idx = findIndex(results, (r) => { return r.id === query.lastId; });
            if (idx > -1) {
                results = results.slice(idx + 1);
            }
            else {
                results = [];
            }
        }
        // Paged query
        if(query.lastId) {
            const idx = findIndex(results, (r) => {return r.id === query.lastId;});
            if(idx > -1) {
                results = results.slice(idx + 1);
            }
            else {
                results= [];
            }
        }
        // limit
        if (query.limit) {
            results.splice(query.limit);
        }
        return results;
    }
    private getTestResult(testElement: IPredicate | ILogicalSequence, item: T): boolean {
        return (
            testElement.type === "predicate" ?
                this.getPredicateResult(testElement, item) :
                this.getSequenceResult(testElement, item)
        );
    }
    private getPredicateResult(predicate: IPredicate, item: T): boolean {
        let result = false;
        let value = item[predicate.propertyName];
        let refVal = predicate.value;
        // Dates
        if (refVal === QueryToken.Now) {
            refVal = new Date();
        }
        else if (refVal === QueryToken.Today) {
            const now = new Date();
            refVal = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        }
        if (refVal && refVal instanceof Date && !predicate.includeTimeValue) {
            refVal = new Date(refVal.getFullYear(), refVal.getMonth(), refVal.getDate());
        }
        if (value && value instanceof Date && !predicate.includeTimeValue) {
            value = new Date(value.getFullYear(), value.getMonth(), value.getDate());
        }
        // Lookups
        if (refVal === QueryToken.UserID) {
            refVal = ServicesConfiguration.configuration.currentUserId;
        }
        if (predicate.lookupId) {
            if (value && value.id && typeof (value.id) === "number") {
                value = value.id;
            }
        }
        else if (value && value.id && typeof (value.id) === "number") {
            value = value.title;
        }

        switch (predicate.operator) {
            case TestOperator.BeginsWith:
                result = (
                    value && typeof (value) === "string" &&
                        refVal && typeof (refVal) === "string" ?
                        value.indexOf(refVal) === 0 :
                        false
                );
                break;
            case TestOperator.Contains:
                result = (
                    value && typeof (value) === "string" &&
                        refVal && typeof (refVal) === "string" ?
                        value.indexOf(refVal) !== -1 :
                        false
                );
                break;
            case TestOperator.Eq:
                if (value instanceof TaxonomyTerm && predicate.lookupId) {
                    if(typeof(refVal) === "number") {
                        result = value.wssids.indexOf(refVal) !== -1;
                    }
                    else {
                        // not in sp list --> test on id
                        result = value.id === refVal;
                    }
                }
                else {
                    result = value === refVal;
                }
                break;
            case TestOperator.Geq:
                result = (typeof (value) === "string" && typeof (refVal) === "string") ?
                    value.localeCompare(refVal) >= 0 :
                    value >= refVal;
                break;
            case TestOperator.Gt:
                result = (typeof (value) === "string" && typeof (refVal) === "string") ?
                    value.localeCompare(refVal) > 0 :
                    value > refVal;
                break;
            case TestOperator.In:
                if (value instanceof TaxonomyTerm && predicate.lookupId) {
                    if(Array.isArray(refVal) && refVal.length > 0 && typeof(refVal[0]) === "number") {
                        result = refVal.some(v => value.wssids.indexOf(v) !== -1);
                    }
                    else {
                        // not in sp list --> test on id
                        result = refVal.some(v => value.id === v);
                    }                    
                }
                else {
                    result = Array.isArray(refVal) && refVal.some(v => v === value);
                }
                break;
            case TestOperator.Includes:
                if (Array.isArray(value)) {
                    result = value.some((lookup) => {
                        let test = false;
                        if (predicate.lookupId) {
                            if (lookup && lookup.id) {
                                if (typeof (lookup.id) === "number") {
                                    test = lookup === refVal;
                                }
                                else if (lookup instanceof TaxonomyTerm) {
                                    if(typeof(refVal) === "number"){
                                        test = lookup.wssids.indexOf(refVal) !== -1;
                                    }
                                    else
                                    {
                                        test = lookup.id === refVal;
                                    }
                                }

                            }
                        }
                        else if (lookup && lookup.id) {
                            test = lookup.title === refVal;
                        }
                        return test;
                    });
                }
                break;
            case TestOperator.IsNotNull:
                result = value !== null && value !== undefined && value !== "";
                break;
            case TestOperator.IsNull:
                result = value === null || value === undefined || value === "";
                break;
            case TestOperator.Leq:
                result = (typeof (value) === "string" && typeof (refVal) === "string") ?
                    value.localeCompare(refVal) <= 0 :
                    value <= refVal;
                break;
            case TestOperator.Lt:
                result = (typeof (value) === "string" && typeof (refVal) === "string") ?
                    value.localeCompare(refVal) < 0 :
                    value < refVal;
                break;
            case TestOperator.Neq:
                if (value instanceof TaxonomyTerm && predicate.lookupId) {
                    if(typeof(refVal) === "number"){
                        result = value.wssids.indexOf(refVal) === -1;
                    }
                    else
                    {
                        result = value.id !== refVal;
                    }                    
                }
                else {
                    result = value !== refVal;
                }
                break;
            case TestOperator.NotIncludes:
                if (Array.isArray(value)) {
                    result = !value.some((lookup) => {
                        let test = false;
                        if (predicate.lookupId) {
                            if (lookup && lookup.id) {
                                if (typeof (lookup.id) === "number") {
                                    test = lookup === refVal;
                                }
                                else if (lookup instanceof TaxonomyTerm) {
                                    if(typeof(refVal) === "number"){
                                        test = lookup.wssids.indexOf(refVal) !== -1;
                                    }
                                    else
                                    {
                                        test = lookup.id === refVal;
                                    }    
                                }
                            }
                        }
                        else if (lookup && lookup.id) {
                            test = lookup.title === refVal;
                        }
                        return test;
                    });
                }
                break;
            default:
                break;
        }
        return result;
    }

    private getSequenceResult(sequence: ILogicalSequence, item: T): boolean {
        // and : find first false, or : find first true
        let result = sequence.operator === LogicalOperator.And;
        for (const subTest of sequence.children) {
            const tmp = subTest.type === "predicate" ? this.getPredicateResult(subTest, item) : this.getSequenceResult(subTest, item);
            if (!tmp && sequence.operator === LogicalOperator.And) {
                result = false;
                break;
            }
            else if (tmp && sequence.operator === LogicalOperator.Or) {
                result = true;
                break;
            }
        }
        return result;
    }
}