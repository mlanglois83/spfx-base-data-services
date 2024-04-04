import { assign, cloneDeep, find, findIndex } from "lodash";
import { IDataService, IQuery, ILogicalSequence, IPredicate, IFieldDescriptor, IBaseDataServiceOptions, IBaseSPServiceOptions } from "../../interfaces";
import { BaseItem, OfflineTransaction, TaxonomyTerm } from "../../models";
import { UtilsService } from "../UtilsService";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./cache/BaseDbService";
import { BaseService } from "./BaseService";
import { TransactionType, Constants, LogicalOperator, TestOperator, QueryToken, FieldType, TraceLevel } from "../../constants";
import { ServicesConfiguration } from "../../configuration";
import { isArray, stringIsNullOrEmpty } from "@pnp/core";
import { ServiceFactory } from "../ServiceFactory";
import { Decorators } from "../../decorators";
import { BaseCacheService } from "./cache/BaseCacheService";
import { BaseLocalStorageService } from "./cache/BaseLocalStorageService";
const trace = Decorators.trace;

/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp 
 */
export abstract class BaseDataService<T extends BaseItem<string | number>> extends BaseService implements IDataService<T> {


    protected transactionService: TransactionService;
    protected cacheService: BaseCacheService<T>;
    protected serviceOptions: IBaseDataServiceOptions;

    protected _itemType: (new (item?: any) => T);
    public get itemType(): (new (item?: any) => T) {
        return this._itemType;
    }

    public cast<Tdest extends BaseDataService<T>>(): Tdest {
        return this as unknown as Tdest;
    }

    protected get singleLinkedTypes(): FieldType[] {
        return [
            FieldType.Lookup,
            FieldType.Taxonomy,
            FieldType.User
        ];
    }
    protected get multipleLinkedTypes(): FieldType[] {
        return [
            FieldType.LookupMulti,
            FieldType.TaxonomyMulti,
            FieldType.UserMulti
        ];
    }
    protected get allLinkedTypes(): FieldType[] {
        return this.singleLinkedTypes.concat(...this.multipleLinkedTypes);
    }

    /**
     * 
     * @param type - type of items
     * @param context - context of the current wp
     */
    constructor(itemType: (new (item?: any) => T), options?: IBaseDataServiceOptions, ...args: any []) {
        super(options, ...args);
        if (ServiceFactory.isServiceManaged(itemType["name"]) && !ServiceFactory.isServiceInitializing(itemType["name"])) {
            console.warn(`Service constructor called out of Service factory. Please use ServiceFactory.getService(${itemType["name"]}) or ServiceFactory.getServiceByModelName("${itemType["name"]}")`);
        }
        this._itemType = itemType;
        this.serviceOptions = options || {};
        if(ServicesConfiguration.configuration.useLocalStorage) {
            this.cacheService = new BaseLocalStorageService<T>(itemType, itemType["name"]);
        }
        else {
            this.cacheService = new BaseDbService<T>(itemType, itemType["name"]);
        }
        this.transactionService = new TransactionService();
    }

    /***************************** External sources init and access **************************************/
    protected initialized = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }

    protected async init_internal(): Promise<void> {
        return;
    }


    public async Init(): Promise<void> {
        
        if (!this.initialized) {
            return this.callAsyncWithPromiseManagement(async (): Promise<void> => {
                if (this.init_internal) {
                    await this.init_internal();
                }
                this.initialized = true;
            }, "init");            
        }
    }


    /*************************************************************************************************************************/

    /********************************************* Fields Management *********************************************************/

    public get ItemFields(): { [propertyName: string]: IFieldDescriptor } {
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

    /************************************************************************* Cache expiration ************************************************************************************/

    protected get cacheKeyUrl(): string {
        return ServicesConfiguration.serverRelativeUrl;
    }

    protected getCacheKey(key = "all"): string {
        return UtilsService.formatText(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName, this.hashCode(this.__thisArgs), key);
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

    public get hasCache(): boolean {
        return this.serviceOptions?.cacheDuration > 0 || ServicesConfiguration.configuration.checkOnline;
    }
    /**
     * Cache has to be relaod ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseDataService
     */
    protected needRefreshCache(key = "all"): boolean {

        let result = !(this.serviceOptions?.cacheDuration > 0);
        //if cache defined
        if (!result) {

            const cachedDataDate = this.getCachedData(key);
            if (cachedDataDate) {
                //add cache duration
                cachedDataDate.setMinutes(cachedDataDate.getMinutes() + this.serviceOptions.cacheDuration);

                const now = new Date();

                //cache has expired
                result = cachedDataDate < now;
            } else {
                result = true;
            }

        }

        return result;
    }

    protected getExpiredIds(...ids: Array<number | string>): Array<number | string> {
        const expired = ids.filter((id) => {
            let result = true;
            const lastLoad = this.getIdLastLoad(id);
            if (lastLoad) {
                lastLoad.setMinutes(lastLoad.getMinutes() + (this.serviceOptions?.cacheDuration || 0));
                const now = new Date();
                //cache has expired
                result = lastLoad < now;
            }
            return result;
        });
        return expired;
    }

    protected UpdateIdsLastLoad(...ids: Array<number | string>): void {
        if (this.serviceOptions?.cacheDuration > 0) {
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

    protected resetIdsLastLoad(): void {
        const cacheKey = this.getCacheKey("ids");           
        window.sessionStorage.removeItem(cacheKey);
    }

    protected UpdateCacheData(key = "all"): void {
        const result = !(this.serviceOptions?.cacheDuration > 0);
        //if cache defined
        if (!result) {
            const cacheKey = this.getCacheKey(key);
            window.sessionStorage.setItem(cacheKey, JSON.stringify(new Date()));
        }

    }
    protected resetCacheData(key = "all"): void {
        const cacheKey = this.getCacheKey(key);
        window.sessionStorage.removeItem(cacheKey);
    }

    /*********************************************************************************************************************************************************/

    /*********************************************************** Data operations ***************************************************************************/
    protected populateItem(data: any): T {
        const item = new this.itemType(data);
        const allProperties = Object.keys(this.ItemFields);
        // set field values
        for (const propertyName of allProperties) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescription = this.ItemFields[propertyName];
                this.populateFieldValue(data, item, propertyName, fieldDescription);
            }
        }
        return item;
    }
    protected populateFieldValue(data: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
        const defaultValue = cloneDeep(fieldDescriptor.defaultValue);
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
            case FieldType.Boolean:
            case FieldType.Number:
                destItem[propertyName] = data[fieldDescriptor.fieldName] !== null && data[fieldDescriptor.fieldName] !== undefined ? data[fieldDescriptor.fieldName] : defaultValue;
                break;
            case FieldType.Url:
                destItem[propertyName] = data[fieldDescriptor.fieldName] !== null && data[fieldDescriptor.fieldName] !== undefined ? {
                    url: data[fieldDescriptor.fieldName].Url,
                    description: data[fieldDescriptor.fieldName].Description
                 } : defaultValue;
                break;
            case FieldType.Date:
                destItem[propertyName] = data[fieldDescriptor.fieldName] ? new Date(data[fieldDescriptor.fieldName]) : defaultValue;
                break;
            case FieldType.Json:
                if (data[fieldDescriptor.fieldName]) {
                    try {
                        if (fieldDescriptor.containsFullObject) {
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const itemType = ServiceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                                destItem[propertyName] = assign(new itemType(), data[fieldDescriptor.fieldName]);
                            }
                            else {
                                destItem[propertyName] = data[fieldDescriptor.fieldName];
                            }
                        }
                        else {
                            const jsonObj = JSON.parse(data[fieldDescriptor.fieldName]);
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const itemType = ServiceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                                destItem[propertyName] = assign(new itemType(), jsonObj);
                            }
                            else {
                                destItem[propertyName] = jsonObj;
                            }
                        }
                    }
                    catch (error) {
                        console.error(error);
                        destItem[propertyName] = defaultValue;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            default:
                destItem[propertyName] = defaultValue;
                break;
        }
    }
    protected get ignoredFields(): string[] {
        return [];
    }
    protected isFieldIgnored(item: T, propertyName: string, fieldDescriptor: IFieldDescriptor): boolean {
        return (this.ignoredFields && this.ignoredFields.indexOf(fieldDescriptor.fieldName) !== -1)
            ||
            (propertyName === "id" && item.isLocal);
    }

    protected async convertItem(item: T): Promise<any> {
        const result = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
             await this.convertFieldValue(item, result, propertyName, fieldDescription);
        }));
        return result;
    }
    protected async convertFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor: IFieldDescriptor): Promise<void> {
        const itemValue = item[propertyName];
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;

        if (!this.isFieldIgnored(item, propertyName, fieldDescriptor)) {
            switch (fieldDescriptor.fieldType) {
                case FieldType.Simple:
                case FieldType.Date:
                case FieldType.Boolean:
                case FieldType.Number:
                    destItem[fieldDescriptor.fieldName] = itemValue;
                    break;
                case FieldType.Url:
                    destItem[fieldDescriptor.fieldName] = itemValue ? {Url: itemValue.url, Description: itemValue.description}: itemValue;
                    break;
                case FieldType.Json:
                    if (fieldDescriptor.containsFullObject) {
                        destItem[fieldDescriptor.fieldName] = itemValue;
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                    }
                    break;
                default: break;
            }
        }
    }

    protected abstract getAll_Query(linkedFields?: Array<string>): Promise<Array<any>>;

    @trace(TraceLevel.Internal)
    protected async getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>> {
        let results: Array<T> = [];
        await this.Init();
        const items = await this.getAll_Query(linkedFields);

        if (items && items.length > 0) {
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLinkedFields(linkedFields)) {
                await this.populateLinkedFields(results, linkedFields, preloaded);
            }
        }

        return results;
    }

    /* 
     * Retrieve all elements from datasource depending on connection is enabled
     * If service is not configured as offline, an exception is thrown;
     */

    @trace(TraceLevel.Service)
    public async getAll(linkedFields?: Array<string>): Promise<Array<T>> {

        return this.callAsyncWithPromiseManagement(async () => {
            let result: Array<T>;

            //has to refresh cache
            let reloadData = this.needRefreshCache();            

            //if refresh is needed, test offline/online
            if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                reloadData = navigator.onLine;
            }

            if(!reloadData) {
                try {
                    const tmp = await this.cacheService.getAll();
                    if (this.isMapItemsAsync(linkedFields)) {
                        result = await this.mapItemsAsync(tmp, linkedFields);
                    }
                    else {
                        result = this.mapItemsSync(tmp);
                    }
                } catch (error) {                    
                    reloadData = !ServicesConfiguration.configuration.checkOnline || navigator.onLine;
                }
            }

            if (reloadData) {
                result = await this.getAll_Internal(linkedFields);
                const convresult = result.map(res => this.convertItemToDbFormat(res));
                if(this.hasCache) {
                    try {
                        await this.cacheService.replaceAll(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData();
                    }
                    catch(error) {
                        this.resetIdsLastLoad();
                        this.resetCacheData();
                        console.error(error);
                    }
                }
            }
            return result;
        });            

    }

    protected abstract get_Query(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<any>>;

    /**
     * Get items by query
     * @protected
     * @param {IQuery} query - query used to retrieve items
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    @trace(TraceLevel.Internal)
    protected async get_Internal(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        let results = new Array<T>();

        await this.Init();

        const items = await this.get_Query(query, linkedFields);

        if (items && items.length > 0) {
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLinkedFields(linkedFields)) {
                await this.populateLinkedFields(results, linkedFields, preloaded);
            }
        }
        return results;
    }


    @trace(TraceLevel.Service)
    public async get(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(query).toString() + super.hashCode(linkedFields).toString();
        return this.callAsyncWithPromiseManagement(async () => {
            let result: Array<T>;
            //has to refresh cache
            let reloadData = this.needRefreshCache(keyCached);
            //if refresh is needed, test offline/online
            if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                reloadData = navigator.onLine;
            }

            if(!reloadData) {
                try {
                    const tmp = await this.cacheService.get(query);
                    if (this.isMapItemsAsync(linkedFields)) {
                        result = await this.mapItemsAsync(tmp, linkedFields);
                    }
                    else {
                        result = this.mapItemsSync(tmp);
                    }
                    // filter
                    result = this.filterItems(query, result);
                } catch (error) {                    
                    reloadData = !ServicesConfiguration.configuration.checkOnline || navigator.onLine;
                }
            }

            if (reloadData) {

                result = await this.get_Internal(query, linkedFields);
                if(this.hasCache) {
                    try {
                        //check if data exist for this query in database
                        let tmp = await this.cacheService.get(query);
                        tmp = this.filterItems(query, tmp);
                        //if data exists trash them 
                        if (tmp && tmp.length > 0) {
                            await this.cacheService.deleteItems(tmp);
                        }
                        const convresult = result.map(res => this.convertItemToDbFormat(res));
                        await this.cacheService.addOrUpdateItems(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData(keyCached);
                    }
                    catch(error) {
                        this.resetIdsLastLoad();
                        this.resetCacheData(keyCached);
                        console.error(error);
                    }
                }

            }
            return result;
        }, keyCached);
        
    }

    protected abstract getItemById_Query(id: number | string, linkedFields?: Array<string>): Promise<any>;

    @trace(TraceLevel.Internal)
    protected async getItemById_Internal(id: number | string, linkedFields?: Array<string>): Promise<T> {
        let result = null;
        await this.Init();
        const temp = await this.getItemById_Query(id, linkedFields);
        if (temp) {
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner([temp], linkedFields);
            }
            result = this.populateItem(temp);
            if (this.hasLinkedFields(linkedFields)) {
                await this.populateLinkedFields([result], linkedFields, preloaded);
            }
        }
        return result;
    }


    @trace(TraceLevel.Service)
    public async getItemById(id: number | string, linkedFields?: Array<string>): Promise<T> {
        const promiseKey = "getById_" + id.toString();
        return this.callAsyncWithPromiseManagement(async () => {
            let result: T;
            const deprecatedIds = await this.getExpiredIds(id);
            let reloadData = deprecatedIds.length > 0;
            //if refresh is needed, test offline/online
            if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                reloadData = navigator.onLine;
            }

            if(!reloadData) {
                try {
                    const temp = await this.cacheService.getItemById(id);
                    if (temp) {
                        let res: Array<T>;
                        if (this.isMapItemsAsync(linkedFields)) {
                            res = await this.mapItemsAsync([temp], linkedFields);
                        }
                        else {
                            res = this.mapItemsSync([temp]);
                        }
                        result = res.shift();
                    }
                } catch (error) {                    
                    reloadData = !ServicesConfiguration.configuration.checkOnline || navigator.onLine;
                }
            }

            if (reloadData) {
                result = await this.getItemById_Internal(id, linkedFields);
                const converted = this.convertItemToDbFormat(result);
                if(this.hasCache) {
                    try {
                        await this.cacheService.addOrUpdateItem(converted);
                        this.UpdateIdsLastLoad(id);
                    }
                    catch(error) {
                        this.resetIdsLastLoad();
                        console.error(error);
                    }
                }
            }
            return result;
        }, promiseKey);
        
    }



    protected abstract getItemsById_Query(id: Array<number | string>, linkedFields?: Array<string>): Promise<any>;
    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    @trace(TraceLevel.Internal)
    protected async getItemsById_Internal(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {

        let results = new Array<T>();
        await this.Init();
        const items = await this.getItemsById_Query(ids, linkedFields);
        if (items && items.length > 0) {
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLinkedFields(linkedFields)) {
                await this.populateLinkedFields(results, linkedFields, preloaded);
            }
        }
        return results;
    }

    public async getItemsFromCacheById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const tmp = await this.cacheService.getItemsById(ids);
        if (this.isMapItemsAsync(linkedFields)) {
            return this.mapItemsAsync(tmp, linkedFields);
        }
        else {
            return this.mapItemsSync(tmp);
        }
    }

    @trace(TraceLevel.Service)
    public async getItemsById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const promiseKey = "getByIds_" + ids.join();
        return this.callAsyncWithPromiseManagement(async () => {
            if (ids.length > 0) {
                let results: Array<T>;
                const deprecatedIds = this.hasCache ? await this.getExpiredIds(...ids): ids;
                let reloadData = deprecatedIds.length > 0;
                //if refresh is needed, test offline/online
                if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                    reloadData = navigator.onLine;
                }

                if(!reloadData) {
                    try {
                        const tmp = await this.cacheService.getItemsById(ids);
                        if (this.isMapItemsAsync(linkedFields)) {
                            results = await this.mapItemsAsync(tmp, linkedFields);
                        }
                        else {
                            results = this.mapItemsSync(tmp);
                        }
                    } catch (error) {                    
                        reloadData = !ServicesConfiguration.configuration.checkOnline || navigator.onLine;
                    }
                }

                if (reloadData) {
                    results = await this.getItemsById_Internal(deprecatedIds, linkedFields);
                    if(this.hasCache) {
                        const tmpcached = await this.cacheService.getItemsById(ids.filter((i) => { return deprecatedIds.indexOf(i) === -1; }));
                        let cached: Array<T>;
                        if (this.isMapItemsAsync(linkedFields)) {
                            cached = await this.mapItemsAsync(tmpcached, linkedFields);
                        }
                        else {
                            cached = this.mapItemsSync(tmpcached);
                        }
                        results = results.concat(cached);
                        try {
                            const convresults = results.map(res => this.convertItemToDbFormat(res));
                            await this.cacheService.addOrUpdateItems(convresults);
                            this.UpdateIdsLastLoad(...ids);
                        }
                        catch(error) {
                            this.resetIdsLastLoad();
                            console.error(error);
                        }
                    }
                }
                return results;
            }
            else {
                return [];
            }
        }, promiseKey);        
    }

    protected abstract addOrUpdateItem_Internal(item: T): Promise<T>;


    @trace(TraceLevel.Service)
    public async addOrUpdateItem(item: T): Promise<T> {
        item.error = undefined;
        this.updateInternalLinks(item);
        let result: T = null;

        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = navigator.onLine;
        }
        if (isconnected) {
            try {
                result = await this.addOrUpdateItem_Internal(item);
                if(this.hasCache) {
                    try {
                        if (item.isCreatedOffline) { // created item allready stored in db
                            this.cacheService.deleteItem(item);
                        }
                        const converted = this.convertItemToDbFormat(result);
                        await this.cacheService.addOrUpdateItem(converted);
                        this.UpdateIdsLastLoad(converted.id);
                    }
                    catch(error) {
                        console.error(error);
                    }
                }


            } catch (error) {
                console.error(error);
                if (error.name === Constants.Errors.ItemVersionConfict) {
                    result = await this.getItemById_Internal(item.id);
                    result.error = error;
                    if(this.hasCache) {
                        try {
                            const converted = this.convertItemToDbFormat(result);
                            await this.cacheService.addOrUpdateItem(converted);
                            this.UpdateIdsLastLoad(converted.id);
                        }
                        catch(error) {
                            console.error(error);
                        }
                    }
                }
                else {
                    item.error = error;
                    result = item;
                }

            }
        }
        else {
            const dbItem = this.convertItemToDbFormat(item);
            const resultitem = await this.cacheService.addOrUpdateItem(dbItem);
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

    protected abstract addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void, onRefreshItems?: (index: number, length: number) => void): Promise<Array<T>>;


    @trace(TraceLevel.Service)
    public async addOrUpdateItems(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void, onRefreshItems?: (index: number, length: number) => void): Promise<Array<T>> {
        items.forEach(item => {
            item.error = undefined;
            this.updateInternalLinks(item);
        });

        let results: Array<T> = [];

        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = navigator.onLine;
        }
        if (isconnected) {
            results = await this.addOrUpdateItems_Internal(items, onItemUpdated, onRefreshItems);
            const versionErrors = results.filter((res) => {
                return res.error && res.error.name === Constants.Errors.ItemVersionConfict;
            });
            // find back items with version error
            if (versionErrors.length > 0) {
                const spitems = await this.getItemsById_Internal(versionErrors.map(ve => ve.id));
                spitems.forEach((retrieved) => {
                    const idx = findIndex(results, (r) => r.id === retrieved.id);
                    if (idx > -1) {
                        retrieved.error = results[idx].error;
                        results[idx] = retrieved;
                    }
                });
            }
            // TODO: promise.All (concurrency on idslastload ?)
            if(this.hasCache) {
                try {
                    for (const item of results) {
                        const converted = this.convertItemToDbFormat(item);
                        await this.cacheService.addOrUpdateItem(converted);
                        this.UpdateIdsLastLoad(converted.id);
                    }
                }
                catch(error) {
                    console.error(error);
                }
            }


        }
        else {
            // TODO: promise.All
            for (const item of items) {
                const copy = cloneDeep(item);
                const dbItem = this.convertItemToDbFormat(item);
                const resultitem = await this.cacheService.addOrUpdateItem(dbItem);
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

    @trace(TraceLevel.Service)
    public async deleteItem(item: T): Promise<T> {
        item.error = undefined;
        if (item.id === item.defaultKey) {
            item.deleted = true;
        }
        else {
            let isconnected = true;
            if (ServicesConfiguration.configuration.checkOnline) {
                isconnected = navigator.onLine;
            }
            if (isconnected) {
                if (!item.isLocal) {
                    item = await this.deleteItem_Internal(item);
                }
                if ((item.deleted || item.isCreatedOffline) && this.hasCache) {
                    try {
                        item = await this.cacheService.deleteItem(item);
                    }
                    catch(error) {
                        console.error(error);
                    }
                }
            }
            else {
                item = await this.cacheService.deleteItem(item);

                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                const converted = this.convertItemToDbFormat(item);
                ot.itemData = assign({}, converted);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.Delete;
                await this.transactionService.addOrUpdateItem(ot);
            }
        }
        return item;
    }

    protected abstract recycleItems_Internal(items: Array<T>): Promise<Array<T>>;

    @trace(TraceLevel.Service)
    public async recycleItems(items: Array<T>): Promise<Array<T>> {
        items.filter(i => (i.id === i.defaultKey)).forEach(i => {
            i.error = undefined;
            i.deleted = true;
        });
        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = navigator.onLine;
        }
        if (isconnected) {
            await this.deleteItems_Internal(items.filter(i => !i.isLocal));
            if(this.hasCache) {
                try {
                    await this.cacheService.deleteItems(items.filter(i => i.deleted || i.isCreatedOffline));
                }
                catch(error){
                    console.error(error);
                }
            }
        }
        else {
            await this.cacheService.deleteItems(items.filter(i => !i.isLocal));
            const transactions: Array<OfflineTransaction> = [];
            // TODO: promise.All
            for (const item of items) {
                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                const converted = this.convertItemToDbFormat(item);
                ot.itemData = assign({}, converted);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.Delete;
                transactions.push(ot);
            }
            await this.transactionService.addOrUpdateItems(transactions);
        }

        return items;
    }

    protected abstract deleteItems_Internal(items: Array<T>): Promise<Array<T>>;

    @trace(TraceLevel.Service)
    public async deleteItems(items: Array<T>): Promise<Array<T>> {
        items.filter(i => (i.id === i.defaultKey)).forEach(i => {
            i.error = undefined;
            i.deleted = true;
        });
        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = navigator.onLine;
        }
        if (isconnected) {
            await this.deleteItems_Internal(items.filter(i => !i.isLocal));
            if(this.hasCache) {
                try {
                    await this.cacheService.deleteItems(items.filter(i => i.deleted || i.isCreatedOffline));
                }
                catch(error){
                    console.error(error);
                }
            }
        }
        else {
            await this.cacheService.deleteItems(items.filter(i => !i.isLocal));
            const transactions: Array<OfflineTransaction> = [];
            // TODO: promise.All
            for (const item of items) {
                // create a new transaction
                const ot: OfflineTransaction = new OfflineTransaction();
                const converted = this.convertItemToDbFormat(item);
                ot.itemData = assign({}, converted);
                ot.itemType = item.constructor["name"];
                ot.title = TransactionType.Delete;
                transactions.push(ot);
            }
            await this.transactionService.addOrUpdateItems(transactions);
        }

        return items;
    }


    @trace(TraceLevel.Service)
    public async persistItemData(data: any, linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem<string | number>[] }): Promise<T> {
        let results: Array<T>;
        if (this.isPersistItemsDataAsync(linkedFields, preloaded)) {
            results = await this.persistItemsDataAsync_internal([data], linkedFields, preloaded);
        }
        else {
            results = this.persistItemsDataSync_internal([data]);
        }
        const result = results.shift();
        const convresult = this.convertItemToDbFormat(result);
        await this.cacheService.addOrUpdateItem(convresult);
        this.UpdateIdsLastLoad(convresult.id);
        return result;
    }

    @trace(TraceLevel.Service)
    public async persistItemsData(data: any[], linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem<string | number>[] }): Promise<T[]> {
        let result: Array<T>;
        if (this.isPersistItemsDataAsync(linkedFields, preloaded)) {
            result = await this.persistItemsDataAsync_internal(data, linkedFields, preloaded);
        }
        else {
            result = this.persistItemsDataSync_internal(data);
        }
        if(this.hasCache) {
            try {
                const convresult = result.map(r => this.convertItemToDbFormat(r));
                result = await this.cacheService.addOrUpdateItems(convresult);
                this.UpdateIdsLastLoad(...convresult.map(cr => cr.id));
            }
            catch(error) {
                console.error(error);
            }
        }
        return result;
    }

    protected isPersistItemsDataAsync(linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem<string | number>[] }): boolean {
        return !this.initialized || (!preloaded && this.needsPersistInner(linkedFields)) || this.hasLinkedFields(linkedFields);
    }

    @trace(TraceLevel.Internal)
    protected async persistItemsDataAsync_internal(data: any[], linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem<string | number>[] }): Promise<T[]> {
        let result = null;
        await this.Init();
        if (data) {
            if (!preloaded && this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(data, linkedFields);
            }
            result = data.map(d => this.populateItem(d));
            if (this.hasLinkedFields(linkedFields)) {
                await this.populateLinkedFields(result, linkedFields, preloaded);
            }
        }
        return result;
    }
    @trace(TraceLevel.Internal)
    protected persistItemsDataSync_internal(data: any[]): T[] {
        let result = null;
        if (data) {
            result = data.map(d => this.populateItem(d));
        }
        return result;
    }

    public needsPersistInner(linkedFields?: Array<string>): boolean {
        const fields = this.ItemFields;
        const keys = Object.keys(fields).filter(k => fields.hasOwnProperty(k) &&
            (!linkedFields || (linkedFields.length === 1 && linkedFields[0] === 'loadAll') || linkedFields.indexOf(fields[k].fieldName) !== -1) &&
            (this.allLinkedTypes.some(lt => lt === fields[k].fieldType)) &&
            fields[k].containsFullObject &&
            !stringIsNullOrEmpty(fields[k].modelName)
        );
        return keys.length > 0;
    }

    protected async persistInner(objects: any[], linkedFields?: Array<string>): Promise<{ [modelName: string]: BaseItem<string | number>[] }> {
        let level = 0;
        let result: { [modelName: string]: BaseItem<string | number>[] } = undefined;
        let sortedByLevel = undefined;
        // get inner objects sorted with level in tree
        let innerItems = this.getInnerValuesForLevel({ [this.itemType["name"]]: objects }, linkedFields);
        while (innerItems !== undefined) {
            level++;
            sortedByLevel = sortedByLevel || {};
            for (const key in innerItems) {
                if (innerItems.hasOwnProperty(key)) {
                    // init
                    sortedByLevel[key] = sortedByLevel[key] || {};
                    // set max level
                    sortedByLevel[key].maxLevel = level;
                    // add objects
                    sortedByLevel[key].objects = sortedByLevel[key].objects || [];
                    sortedByLevel[key].objects.push(...innerItems[key]);


                }
            }
            innerItems = this.getInnerValuesForLevel(innerItems);
        }

        // persist by level desc
        for (let index = level; index > 0; index--) {
            result = result || {};
            // get models for level
            const keys = Object.keys(sortedByLevel).filter(k => sortedByLevel.hasOwnProperty(k) &&
                sortedByLevel[k].maxLevel === index);
            // persist by model
            await Promise.all(keys.map(async k => {
                result[k] = result[k] || [];
                // get service
                const service = ServiceFactory.getServiceByModelName(k);
                const persisted = await service.persistItemsData(sortedByLevel[k].objects, undefined, result);
                result[k].push(...persisted);
            }));


        }
        return result;
    }
    protected getInnerValuesForLevel(objects: { [modelName: string]: any[] }, linkedFields?: Array<string>): { [modelName: string]: any[] } {
        let result: { [modelName: string]: any[] } = undefined;
        if (objects) {
            // get inner lookups by model name
            const inner = [];
            for (const key in objects) {
                if (objects.hasOwnProperty(key) && objects[key] && objects[key].length > 0) {
                    const innerResult = this.getInnerValuesForSingleType(key, objects[key], linkedFields);
                    if (innerResult) {
                        inner.push(innerResult);
                    }
                }
            }
            // merge results
            inner.forEach(i => {
                for (const key in i) {
                    if (i.hasOwnProperty(key) && i[key] && i[key].length > 0) {
                        result = result || {};
                        result[key] = result[key] || [];
                        result[key].push(...i[key]);
                    }
                }
            });

        }
        return result;
    }



    protected getInnerValuesForSingleType(modelName: string, objects: any[], linkedFields?: Array<string>): { [modelName: string]: any[] } {
        let result: { [modelName: string]: any[] } = undefined;
        if (objects && objects.length > 0) {
            // get service to find fields
            const service = ServiceFactory.getServiceByModelName(modelName);
            const fields = service.ItemFields;
            const keys = Object.keys(fields).filter(k => fields.hasOwnProperty(k) &&
                (!linkedFields || (linkedFields.length === 1 && linkedFields[0] === 'loadAll') || linkedFields.indexOf(fields[k].fieldName) !== -1) &&
                (fields[k].fieldType === FieldType.Lookup || fields[k].fieldType === FieldType.LookupMulti) &&
                fields[k].containsFullObject &&
                !stringIsNullOrEmpty(fields[k].modelName)
            );
            for (const key of keys) {
                const descriptor = fields[key];
                const destModelName = descriptor.modelName;
                objects.forEach(o => {
                    if (o[descriptor.fieldName]) {
                        if (this.singleLinkedTypes.some(lt => lt === descriptor.fieldType)) {
                            result = result || {};
                            result[destModelName] = result[destModelName] || [];
                            result[destModelName].push(o[descriptor.fieldName]);
                        }
                        else if (Array.isArray(o[descriptor.fieldName]) && o[descriptor.fieldName].length > 0) {
                            result = result || {};
                            result[destModelName] = result[destModelName] || [];
                            result[destModelName].push(...o[descriptor.fieldName]);
                        }
                    }
                });
            }
        }
        return result;
    }

    /*****************************************************************************************************************************************************************/


    /********************************** Link to lookups  *************************************/
    protected linkedFields(loadLookups?: Array<string>): Array<IFieldDescriptor> {
        const result: Array<IFieldDescriptor> = [];
        const fields = this.ItemFields;
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key];
                if (this.allLinkedTypes.some(lt => lt === fieldDesc.fieldType) && !stringIsNullOrEmpty(fieldDesc.modelName)) {
                    if (!loadLookups || (loadLookups.length === 1 && loadLookups[0] === 'loadAll') || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                        result[key] = fieldDesc;
                    }
                }
            }
        }

        return result;
    }

    public hasLinkedFields(linkedFields?: Array<string>): boolean {
        const serviceLinkedFields = this.linkedFields(linkedFields);
        return Object.keys(serviceLinkedFields).filter(k => serviceLinkedFields.hasOwnProperty(k)).length > 0;
    }

    @trace(TraceLevel.ServiceUtilities)
    protected async populateLinkedFields(items: Array<T>, loadLinked?: Array<string>, innerItems?: { [modelName: string]: BaseItem<string | number>[] }): Promise<void> {
        await this.Init();        
        // get linked fields
        const linkedFields = this.linkedFields(loadLinked);
        // init values and retrieve all ids by model
        const allIds = {};
        const innerResult = {};
        for (const key in linkedFields) {
            if (linkedFields.hasOwnProperty(key)) {
                const fieldDesc = linkedFields[key];
                allIds[fieldDesc.modelName] = allIds[fieldDesc.modelName] || [];
                const ids = allIds[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const links = item.__getInternalLinks(key);
                    //init value 
                    if (this.allLinkedTypes.some(lt => lt === fieldDesc.fieldType)) {
                        item[key] = fieldDesc.defaultValue;
                    }
                    if (this.singleLinkedTypes.some(lt => lt === fieldDesc.fieldType) &&
                        // lookup has value
                        links &&
                        links !== -1) {
                        // check in preloaded
                        let inner = undefined;
                        if (innerItems && innerItems[fieldDesc.modelName]) {
                            inner = find(innerItems[fieldDesc.modelName], ii => ii.id === links);
                        }
                        // inner found
                        if (inner) {
                            innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                            innerResult[fieldDesc.modelName].push(inner);
                        }
                        else {
                            ids.push(links);
                        }
                    }
                    else if (this.multipleLinkedTypes.some(lt => lt === fieldDesc.fieldType) &&
                        links &&
                        links.length > 0) {
                        links.forEach((id) => {
                            let inner = undefined;
                            if (innerItems && innerItems[fieldDesc.modelName]) {
                                inner = find(innerItems[fieldDesc.modelName], ii => ii.id === id);
                            }
                            // inner found
                            if (inner) {
                                innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                                innerResult[fieldDesc.modelName].push(inner);
                            }
                            else {
                                ids.push(id);
                            }
                        });
                    }
                });

            }
        }

        const resultItems: { [modelName: string]: BaseItem<string | number>[] } = innerResult;
        
        // Init queries       
        const promises: Array<() => Promise<BaseItem<string | number>[]>> = [];
        for (const modelName in allIds) {
            if (allIds.hasOwnProperty(modelName)) {
                const ids = allIds[modelName];
                if (ids && ids.length > 0) {
                    const options: IBaseSPServiceOptions = {};
                    // for sp services
                    if(this.serviceOptions.hasOwnProperty('baseUrl')) {
                        options.baseUrl = (this.serviceOptions as IBaseSPServiceOptions).baseUrl;
                    }
                    const service = ServiceFactory.getServiceByModelName(modelName, options);
                    promises.push(() => service.getItemsById(ids));
                }
            }
        }
        // execute and store
        const results = await UtilsService.executePromisesInStacks(promises, 3);
        results.forEach(itemsTab => {
            if (itemsTab.length > 0) {
                resultItems[itemsTab[0].constructor["name"]] = resultItems[itemsTab[0].constructor["name"]] || [];
                resultItems[itemsTab[0].constructor["name"]].push(...itemsTab);
            }
        });

        // Associate to items
        for (const propertyName in linkedFields) {
            if (linkedFields.hasOwnProperty(propertyName)) {
                const fieldDesc = linkedFields[propertyName];
                const refCol = resultItems[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const links = item.__getInternalLinks(propertyName);
                    if (this.singleLinkedTypes.some(lt => lt === fieldDesc.fieldType) &&
                        links &&
                        links !== -1) {
                        const litem = find(refCol, { id: links });
                        if (litem) {
                            item[propertyName] = litem;
                        }

                    }
                    else if (this.multipleLinkedTypes.some(lt => lt === fieldDesc.fieldType) &&
                        links &&
                        links.length > 0) {
                        item[propertyName] = [];
                        links.forEach((id) => {
                            const litem = find(refCol, { id: id });
                            if (litem) {
                                item[propertyName].push(litem);
                            }
                        });
                    }
                });
            }
        }
    }

    protected updateInternalLinks(item: T, loadLinkedFields?: Array<string>): void {
        const linkedFields = this.linkedFields(loadLinkedFields);
        for (const propertyName in linkedFields) {
            if (linkedFields.hasOwnProperty(propertyName)) {
                const fieldDesc = linkedFields[propertyName];
                if (!loadLinkedFields || loadLinkedFields.indexOf(fieldDesc.fieldName) !== -1) {                    
                    const obj = ServiceFactory.getItemByName(fieldDesc.modelName);
                    if (this.singleLinkedTypes.some(lt => lt === fieldDesc.fieldType)) {                        
                        item.__deleteInternalLinks(propertyName);
                        if (item[propertyName] && item[propertyName].id !== obj.defaultKey) {
                            item.__setInternalLinks(propertyName, item[propertyName].id);
                        }
                    }
                    else if (this.multipleLinkedTypes.some(lt => lt === fieldDesc.fieldType)) {
                        item.__deleteInternalLinks(propertyName);
                        if (item[propertyName] && item[propertyName].length > 0) {
                            item.__setInternalLinks(propertyName, item[propertyName].filter(l => l.id !== obj.defaultKey).map(l => l.id));
                        }
                    }
                }
            }
        }
    }


    /********************************************************************** Cached data management ******************************************************************************/



    /**
     * convert full item to db format (with links only)
     * @param item - full provisionned item
     */
    protected convertItemToDbFormat(item: T): T {
        const result: T = cloneDeep(item);
        result.cleanBeforeStorage();
        result.__clearInternalLinks();
        for (const propertyName in result) {
            if (result.hasOwnProperty(propertyName)) {
                if (this.ItemFields.hasOwnProperty(propertyName)) {
                    const fieldDescriptor = this.ItemFields[propertyName];
                    switch (fieldDescriptor.fieldType) {
                        case FieldType.User:
                        case FieldType.Taxonomy:
                        case FieldType.Lookup:
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                //link defered
                                if (item[propertyName]) {
                                    result.__setInternalLinks(propertyName, (item[propertyName] as unknown as BaseItem<string | number>).id);
                                }
                                delete result[propertyName];
                            }
                            break;
                        case FieldType.UserMulti:
                        case FieldType.TaxonomyMulti:
                        case FieldType.LookupMulti:
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const ids = [];
                                if (item[propertyName]) {
                                    (item[propertyName] as unknown as BaseItem<string | number>[]).forEach(element => {
                                        if (element?.id) {
                                            ids.push(element.id);
                                        }
                                    });
                                }
                                if (ids.length > 0) {
                                    result.__setInternalLinks(propertyName, ids.length > 0 ? ids : []);
                                }
                                delete result[propertyName];
                            }
                            break;
                        default:
                            break;
                    }
                } else if (typeof (result[propertyName]) === "function") {
                    delete result[propertyName];
                }
            }
        }
        return result;
    }

    public isMapItemsAsync(linkedFields?: Array<string>): boolean {
        return !this.initialized || this.hasLinkedFields(linkedFields);
    }

    /**
     * populate item from db storage
     * @param item - db item with links in internalLinks fields
     */
    @trace(TraceLevel.ServiceUtilities)
    public async mapItemsAsync(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        let results: Array<T> = [];
        await this.Init();
        if (items && items.length > 0) {
            results = this.mapItems_internal(items);
        }
        if (this.hasLinkedFields(linkedFields)) {
            await this.populateLinkedFields(results, linkedFields);
        }
        return results;
    }

    @trace(TraceLevel.ServiceUtilities)
    public mapItemsSync(items: Array<T>): Array<T> {
        let results: Array<T> = [];
        if (items && items.length > 0) {
            results = this.mapItems_internal(items);
        }
        return results;
    }

    protected mapItems_internal(items: Array<T>): Array<T> {
        const results: Array<T> = [];
        for (const item of items) {
            const result: T = cloneDeep(item);
            if (item) {
                for (const propertyName in this.ItemFields) {
                    if (this.ItemFields.hasOwnProperty(propertyName)) {
                        const fieldDescriptor = this.ItemFields[propertyName];                        
                        if (fieldDescriptor.fieldType === FieldType.Json && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            const itemType = ServiceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                            result[propertyName] = assign(new itemType(), item[propertyName]);
                        }
                        else {
                            result[propertyName] = item[propertyName];
                        }
                    }
                }
            }
            result.__clearEmptyInternalLinks();
            results.push(result);
        }
        return results;
    }

    public async updateLinkedTransactions(oldId: number | string, newId: number | string, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        const updatedTransactions = [];
        // Update items pointing to this in transactions
        nextTransactions.forEach(transaction => {
            let currentObject = null;
            let needUpdate = false;
            const service = ServiceFactory.getServiceByModelName(transaction.itemType);
            const fields = service.ItemFields;
            // search for lookup fields
            for (const propertyName in fields) {
                if (fields.hasOwnProperty(propertyName)) {
                    const fieldDescription: IFieldDescriptor = fields[propertyName];
                    if (fieldDescription.refItemName === this.itemType["name"] || fieldDescription.modelName === this.itemType["name"]) {
                        // get object if not done yet
                        if (!currentObject) {
                            currentObject = ServiceFactory.getItemByName(transaction.itemType);
                            assign(currentObject, transaction.itemData);
                        }
                        if (fieldDescription.fieldType === FieldType.Lookup) {
                            if (fieldDescription.modelName) {
                                // search in internalLinks
                                const link = currentObject.__getInternalLinks(propertyName);
                                if (link && link === oldId) {
                                    currentObject.__setInternalLinks(propertyName, newId);
                                    needUpdate = true;
                                }
                            }
                            else if (currentObject[propertyName] === oldId) {
                                // change field
                                currentObject[propertyName] = newId;
                                needUpdate = true;
                            }
                        }
                        else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                            if (fieldDescription.modelName) {
                                // serch in internalLinks
                                const links = currentObject.__getInternalLinks(propertyName);
                                if (links && isArray(links)) {
                                    // find item
                                    const lookupidx = findIndex(links, (id) => { return id === oldId; });
                                    // change id
                                    if (lookupidx > -1) {
                                        currentObject.__setReplaceInternalLinks(propertyName, oldId, newId);
                                        needUpdate = true;
                                    }
                                }
                            }
                            else if (currentObject[propertyName] && isArray(currentObject[propertyName])) {
                                // find index
                                const lookupidx = findIndex(currentObject[propertyName], (id) => { return id === oldId; });
                                // change field
                                // change id
                                if (lookupidx > -1) {
                                    currentObject[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                        }

                    }

                }
            }
            if (needUpdate) {
                transaction.itemData = assign({}, currentObject);
                updatedTransactions.push(transaction);
            }
        });
        if (updatedTransactions.length > 0) {
            const updateResult = await this.transactionService.addOrUpdateItems(updatedTransactions);
            updateResult.forEach(r => {
                const idx = findIndex(nextTransactions, { id: r.id });
                if (idx > -1) {
                    nextTransactions[idx] = r;
                }
            });
        }
        return nextTransactions;
    }

    @trace(TraceLevel.ServiceUtilities)
    protected async updateLinksInDb(oldId: number, newId: number): Promise<void> {
        const allFields = assign({}, this.itemType["Fields"]);
        let parentType = this.itemType;
        do {
            delete allFields[parentType["name"]];
            parentType = Object.getPrototypeOf(parentType);
        } while (parentType["name"] !== BaseItem["name"]);
        for (const modelName in allFields) {
            if (allFields.hasOwnProperty(modelName)) {
                const modelFields = allFields[modelName];
                const lookupProperties = Object.keys(modelFields).filter((prop) => {
                    return (modelFields[prop].refItemName &&
                        modelFields[prop].refItemName === this.itemType["name"] || modelFields[prop].modelName === this.itemType["name"]);
                });
                if (lookupProperties.length > 0) {
                    let service: BaseDataService<BaseItem<string | number>>;
                    try {
                        service = ServiceFactory.getServiceByModelName(modelName);
                    } catch {
                        console.warn("No service found for '" + modelName + "'");
                    }
                    if (service && service.hasCache) {
                        const allitems = await service.__getAllFromCache();
                        const updated = [];
                        allitems.forEach(element => {
                            let needUpdate = false;
                            lookupProperties.forEach(propertyName => {
                                const fieldDescription = modelFields[propertyName];
                                if (fieldDescription.fieldType === FieldType.Lookup) {
                                    if (fieldDescription.modelName) {
                                        // search in internalLinks
                                        const link = element.__getInternalLinks(propertyName);
                                        if (link && link === oldId) {
                                            element.__setInternalLinks(propertyName, newId);
                                            needUpdate = true;
                                        }
                                    }
                                    else if (element[propertyName] === oldId) {
                                        // change field
                                        element[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                                else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                    if (fieldDescription.modelName) {
                                        // search in internalLinks
                                        const links = element.__getInternalLinks(propertyName);
                                        if (links && isArray(links)) {
                                            // find item
                                            const lookupidx = findIndex(links, (id) => { return id === oldId; });
                                            // change id
                                            if (lookupidx > -1) {
                                                element.__setInternalLinks(propertyName, newId);
                                                needUpdate = true;
                                            }
                                        }
                                    }
                                    else if (element[propertyName] && isArray(element[propertyName])) {
                                        // find index
                                        const lookupidx = findIndex(element[propertyName], (id) => { return id === oldId; });
                                        // change field
                                        // change id
                                        if (lookupidx > -1) {
                                            element[propertyName] = newId;
                                            needUpdate = true;
                                        }
                                    }
                                }
                            });
                            if (needUpdate) {
                                updated.push(element);
                            }
                        });

                        if (updated.length > 0) {
                            try {
                                await service.__updateCache(...updated);
                            }
                            catch(error) {
                                console.error(error);
                            }
                        }
                    }
                }
            }
        }
    }

    public __getFromCache(id: number | string): Promise<T> {
        return this.cacheService.getItemById(id);
    }

    public __getAllFromCache(): Promise<Array<T>> {
        return this.cacheService.getAll();
    }

    public __updateCache(...items: Array<T>): Promise<Array<T>> {
        return this.cacheService.addOrUpdateItems(items);
    }

    /**
     * Refresh cached data
     */
    public async refreshData(): Promise<void> {
        // Invalidate cache
        const cacheKey = this.getCacheKey(); // Default key is "ALL"
        window.sessionStorage.removeItem(cacheKey);
        // remove local cache
        this.initialized = false;
        // Reload all data
        await this.getAll();
    }

    /*****************************************************************************************************************************************************************/

    /********************************************************************* Queries ************************************************************************************/
    private filterItems(query: IQuery<T>, items: Array<T>): Array<T> {
        // filter items by test
        let results = query.test ? items.filter((i) => { return this.getTestResult(query.test, i); }) : cloneDeep(items);
        // order by
        if (query.orderBy) {
            results.sort(function (a, b) {
                for (const order of query.orderBy) {
                    const aKey = a[order.propertyName.toString()];
                    const bKey = b[order.propertyName.toString()];
                    if (typeof (aKey) === "string" || typeof (bKey) === "string") {
                        if ((aKey || "").localeCompare(bKey || "") < 0) {
                            return order.ascending ? -1 : 1;
                        }
                        if ((aKey || "").localeCompare(bKey || "") > 0) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                    else if (aKey instanceof Date || bKey instanceof Date) {
                        const aval = aKey && aKey.getTime ? aKey.getTime() : 0;
                        const bval = bKey && bKey.getTime ? bKey.getTime() : 0;
                        if (aval < bval) {
                            return order.ascending ? -1 : 1;
                        }
                        if (aval > bval) {
                            return order.ascending ? 1 : -1;
                        }
                    }
                    else if (aKey && bKey && aKey.id && bKey.id) {
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
        // limit
        if (query.limit) {
            results.splice(query.limit);
        }
        return results;
    }
    private getTestResult(testElement: IPredicate<T, keyof T> | ILogicalSequence<T>, item: T): boolean {
        return (
            testElement.type === "predicate" ?
                this.getPredicateResult(testElement, item) :
                this.getSequenceResult(testElement, item)
        );
    }
    private getPredicateResult(predicate: IPredicate<T, keyof T>, item: T): boolean {
        let result = false;
        let value = item[predicate.propertyName.toString()];
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
        // url
        if(value && value.hasOwnProperty("url")) {
            value = value.url;
        }
        // Lookups
        if (refVal === QueryToken.UserID) {
            refVal = ServicesConfiguration.configuration.currentUserId;
        }        
        if (value && value.id && typeof (value.id) === "number") {
            value = predicate.lookupId ? value.id : value.title;
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
                    if (typeof (refVal) === "number") {
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
                    if (Array.isArray(refVal) && refVal.length > 0 && typeof (refVal[0]) === "number") {
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
                                    test = lookup.id === refVal;
                                }
                                else if (lookup instanceof TaxonomyTerm) {
                                    if (typeof (refVal) === "number") {
                                        test = lookup.wssids.indexOf(refVal) !== -1;
                                    }
                                    else {
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
                    if (typeof (refVal) === "number") {
                        result = value.wssids.indexOf(refVal) === -1;
                    }
                    else {
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
                                    if (typeof (refVal) === "number") {
                                        test = lookup.wssids.indexOf(refVal) !== -1;
                                    }
                                    else {
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

    private getSequenceResult(sequence: ILogicalSequence<T>, item: T): boolean {
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