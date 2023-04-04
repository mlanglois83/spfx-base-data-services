import { assign, cloneDeep, find, findIndex } from "lodash";
import { IDataService, IQuery, ILogicalSequence, IPredicate, IFieldDescriptor } from "../../interfaces";
import { BaseItem, OfflineTransaction, TaxonomyTerm } from "../../models";
import { UtilsService } from "../UtilsService";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { TransactionType, Constants, LogicalOperator, TestOperator, QueryToken, FieldType, TraceLevel } from "../../constants";
import { ServicesConfiguration } from "../../configuration";
import { isArray, stringIsNullOrEmpty } from "@pnp/core";
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


    public get itemType(): (new (item?: any) => T) {
        return this.itemModelType;
    }

    public cast<Tdest extends BaseDataService<T>>(): Tdest {
        return this as unknown as Tdest;
    }

    /**
     * 
     * @param type - type of items
     * @param context - context of the current wp
     */
    constructor(type: (new (item?: any) => T), cacheDuration = -1) {
        super();
        if (ServiceFactory.isServiceManaged(type["name"]) && !ServiceFactory.isServiceInitializing(type["name"])) {
            console.warn(`Service constructor called out of Service factory. Please use ServiceFactory.getService(${type["name"]}) or ServiceFactory.getServiceByModelName("${type["name"]}")`);
        }
        this.itemModelType = type;
        this.cacheDuration = cacheDuration;
        this.dbService = new BaseDbService<T>(type, type["name"]);
        this.transactionService = new TransactionService();
    }

    /***************************** External sources init and access **************************************/
    protected initValues: { [modelName: string]: BaseItem[] } = {};
    protected cachedLookups: { [modelName: string]: Array<string | number> } = {};

    protected getServiceCachedLookupIds<Tvalue extends BaseItem>(model: new (data?: any) => Tvalue): Array<number | string> {
        return this.getServiceCachedLookupIdsByName(model["name"]);
    }

    protected getServiceCachedLookupIdsByName(modelName: string): Array<string | number> {
        return this.cachedLookups[modelName] as Array<string | number>;
    }
    protected updateServiceCachedLookupIds(modelName: string, ...items: BaseItem[]): void {
        this.cachedLookups[modelName] = this.cachedLookups[modelName] || [];
        items.forEach(i => {
            const idx = findIndex(this.cachedLookups[modelName], iv => iv === i.id);
            if (idx !== -1) {
                this.cachedLookups[modelName][idx] = i.id;
            }
            else {
                this.cachedLookups[modelName].push(i.id);
            }
        });
    }
    protected isServiceLookupIdCached(modelName: string, id: number | string): boolean {
        const values = this.getServiceCachedLookupIdsByName(modelName);
        return values && values.indexOf(id) !== -1;
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
            if (idx !== -1) {
                this.initValues[modelName][idx] = i;
            }
            else {
                this.initValues[modelName].push(i);
            }
        });
    }




    protected initialized = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }

    protected async init_internal(): Promise<void> {
        return;
    }

    @trace(TraceLevel.ServiceUtilities)
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
        if (!this.initialized) {
            let initPromise = this.getExistingPromise("init");
            if (!initPromise) {
                initPromise = new Promise<void>(async (resolve, reject) => {
                    this.initValues = {};
                    try {
                        if (this.init_internal) {
                            await this.init_internal();
                        }
                        await this.initLinkedFields();
                        this.initialized = true;
                        resolve();
                    }
                    catch (error) {
                        reject(error);
                    }
                });
            }
            this.storePromise(initPromise, "init");
            return initPromise;
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

    protected getCacheKey(key = "all"): string {
        return UtilsService.formatText(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.serverRelativeUrl, this.serviceName, key);
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
    protected needRefreshCache(key = "all"): boolean {

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

    protected getExpiredIds(...ids: Array<number | string>): Array<number | string> {
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
            (propertyName === "id" && typeof (item.id) === "number" && item.id <= 0);
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
                    destItem[fieldDescriptor.fieldName] = itemValue;
                    break;
                case FieldType.Url: // TODO
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


        const items = await this.getAll_Query(linkedFields);


        if (items && items.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLookup(linkedFields)) {
                await this.populateLookups(results, linkedFields, preloaded);
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

        let promise = this.getExistingPromise();
        if (promise) {
            if (this.debug)
                console.log(this.serviceName + " getAll : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result: Array<T>;

                    //has to refresh cache

                    let reloadData = this.needRefreshCache();

                    //if refresh is needed, test offline/online
                    if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                        reloadData = await UtilsService.CheckOnline();
                    }


                    if (reloadData) {
                        result = await this.getAll_Internal(linkedFields);
                        const convresult = result.map(res => this.convertItemToDbFormat(res));
                        await this.dbService.replaceAll(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData();

                    }
                    else {
                        const tmp = await this.dbService.getAll();
                        if (this.isMapItemsAsync(linkedFields)) {
                            result = await this.mapItemsAsync(tmp, linkedFields);
                        }
                        else {
                            result = this.mapItemsSync(tmp);
                        }
                    }

                    resolve(result);
                }
                catch (error) {
                    reject(error);
                }
            });
            this.storePromise(promise);
        }
        return promise;

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


        const items = await this.get_Query(query, linkedFields);

        if (items && items.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLookup(linkedFields)) {
                await this.populateLookups(results, linkedFields, preloaded);
            }
        }
        return results;
    }


    @trace(TraceLevel.Service)
    public async get(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(query).toString() + super.hashCode(linkedFields).toString();
        let promise = this.getExistingPromise(keyCached);
        if (promise) {
            if (this.debug)
                console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
        }
        else {
            promise = new Promise<Array<T>>(async (resolve, reject) => {
                try {
                    let result: Array<T>;
                    //has to refresh cache
                    let reloadData = this.needRefreshCache(keyCached);
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

                        const convresult = result.map(res => this.convertItemToDbFormat(res));
                        await this.dbService.addOrUpdateItems(convresult);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData(keyCached);

                    }
                    else {

                        const tmp = await this.dbService.get(query);
                        if (this.isMapItemsAsync(linkedFields)) {
                            result = await this.mapItemsAsync(tmp, linkedFields);
                        }
                        else {
                            result = this.mapItemsSync(tmp);
                        }
                        // filter
                        result = this.filterItems(query, result);
                    }
                    resolve(result);
                }
                catch (error) {
                    reject(error);
                }
            });
            this.storePromise(promise, keyCached);
        }
        return promise;
    }

    protected abstract getItemById_Query(id: number | string, linkedFields?: Array<string>): Promise<any>;

    @trace(TraceLevel.Internal)
    protected async getItemById_Internal(id: number | string, linkedFields?: Array<string>): Promise<T> {
        let result = null;
        const temp = await this.getItemById_Query(id, linkedFields);
        if (temp) {
            if (!this.initialized) {
                await this.Init();
            }
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner([temp], linkedFields);
            }
            result = this.populateItem(temp);
            if (this.hasLookup(linkedFields)) {
                await this.populateLookups([result], linkedFields, preloaded);
            }
        }
        return result;
    }


    @trace(TraceLevel.Service)
    public async getItemById(id: number | string, linkedFields?: Array<string>): Promise<T> {
        const promiseKey = "getById_" + id.toString();
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
            if (this.debug)
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
                        const converted = this.convertItemToDbFormat(result);
                        await this.dbService.addOrUpdateItem(converted);
                        this.UpdateIdsLastLoad(id);
                    }
                    else {
                        const temp = await this.dbService.getItemById(id);
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
                    }
                    resolve(result);
                }
                catch (error) {
                    reject(error);
                }
            });
            this.storePromise(promise, promiseKey);
        }
        return promise;
    }



    protected abstract getItemsById_Query(id: Array<number | string>, linkedFields?: Array<string>): Promise<any>;
    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    @trace(TraceLevel.Internal)
    protected async getItemsById_Internal(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {

        let results = new Array<T>();
        const items = await this.getItemsById_Query(ids, linkedFields);
        if (items && items.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
            let preloaded = undefined;
            if (this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(items, linkedFields);
            }
            results = items.map(r => this.populateItem(r));
            if (this.hasLookup(linkedFields)) {
                await this.populateLookups(results, linkedFields, preloaded);
            }
        }
        return results;
    }

    public async getItemsFromCacheById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const tmp = await this.dbService.getItemsById(ids);
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
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
            if (this.debug)
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
                            let cached: Array<T>;
                            if (this.isMapItemsAsync(linkedFields)) {
                                cached = await this.mapItemsAsync(tmpcached, linkedFields);
                            }
                            else {
                                cached = this.mapItemsSync(tmpcached);
                            }
                            results = expired.concat(cached);
                            const convresults = results.map(res => this.convertItemToDbFormat(res));
                            await this.dbService.addOrUpdateItems(convresults);
                            this.UpdateIdsLastLoad(...ids);
                        }
                        else {
                            const tmp = await this.dbService.getItemsById(ids);
                            if (this.isMapItemsAsync(linkedFields)) {
                                results = await this.mapItemsAsync(tmp, linkedFields);
                            }
                            else {
                                results = this.mapItemsSync(tmp);
                            }
                        }
                        resolve(results);
                    }
                    catch (error) {
                        reject(error);
                    }
                }
                else {
                    resolve([]);
                }
            });
            this.storePromise(promise, promiseKey);
        }
        return promise;
    }

    protected abstract addOrUpdateItem_Internal(item: T): Promise<T>;


    @trace(TraceLevel.Service)
    public async addOrUpdateItem(item: T): Promise<T> {
        item.error = undefined;
        this.updateInternalLinks(item);
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
                const converted = this.convertItemToDbFormat(itemResult);
                await this.dbService.addOrUpdateItem(converted);
                this.UpdateIdsLastLoad(converted.id);
                result = itemResult;


            } catch (error) {
                console.error(error);
                if (error.name === Constants.Errors.ItemVersionConfict) {
                    itemResult = await this.getItemById_Internal(item.id);
                    const converted = this.convertItemToDbFormat(itemResult);
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
            const dbItem = this.convertItemToDbFormat(item);
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
            isconnected = await UtilsService.CheckOnline();
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
            for (const item of results) {
                const converted = this.convertItemToDbFormat(item);
                await this.dbService.addOrUpdateItem(converted);
                this.UpdateIdsLastLoad(converted.id);
            }


        }
        else {
            // TODO: promise.All
            for (const item of items) {
                const copy = cloneDeep(item);
                const dbItem = this.convertItemToDbFormat(item);
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

    // TODO: remove cached ids
    @trace(TraceLevel.Service)
    public async deleteItem(item: T): Promise<T> {
        item.error = undefined;
        if (typeof (item.id) === "number" && item.id === -1) {
            item.deleted = true;
        }
        else {
            let isconnected = true;
            if (ServicesConfiguration.configuration.checkOnline) {
                isconnected = await UtilsService.CheckOnline();
            }
            if (isconnected) {
                if (typeof (item.id) !== "number" || item.id > -1) {
                    item = await this.deleteItem_Internal(item);
                }
                if (item.deleted || item.id < -1) {
                    item = await this.dbService.deleteItem(item);
                }
            }
            else {
                item = await this.dbService.deleteItem(item);

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
        items.filter(i => (typeof (i.id) === "number" && i.id === -1)).forEach(i => {
            i.error = undefined;
            i.deleted = true;
        });
        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            await this.deleteItems_Internal(items.filter(i => (typeof (i.id) !== "number" || i.id > -1)));
            await this.dbService.deleteItems(items.filter(i => i.deleted || i.id < -1));
        }
        else {
            await this.dbService.deleteItems(items.filter(i => i.id > -1));
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
        items.filter(i => (typeof (i.id) === "number" && i.id === -1)).forEach(i => {
            i.error = undefined;
            i.deleted = true;
        });
        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            await this.deleteItems_Internal(items.filter(i => (typeof (i.id) !== "number" || i.id > -1)));
            await this.dbService.deleteItems(items.filter(i => i.deleted || i.id < -1));
        }
        else {
            await this.dbService.deleteItems(items.filter(i => i.id > -1));
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
    public async persistItemData(data: any, linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem[] }): Promise<T> {
        let results: Array<T>;
        if (this.isPersistItemsDataAsync(linkedFields, preloaded)) {
            results = await this.persistItemsDataAsync_internal([data], linkedFields, preloaded);
        }
        else {
            results = this.persistItemsDataSync_internal([data]);
        }
        const result = results.shift();
        const convresult = this.convertItemToDbFormat(result);
        await this.dbService.addOrUpdateItem(convresult);
        this.UpdateIdsLastLoad(convresult.id);
        return result;
    }

    @trace(TraceLevel.Service)
    public async persistItemsData(data: any[], linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem[] }): Promise<T[]> {
        let result: Array<T>;
        if (this.isPersistItemsDataAsync(linkedFields, preloaded)) {
            result = await this.persistItemsDataAsync_internal(data, linkedFields, preloaded);
        }
        else {
            result = this.persistItemsDataSync_internal(data);
        }
        const convresult = result.map(r => this.convertItemToDbFormat(r));
        await this.dbService.addOrUpdateItems(convresult);
        this.UpdateIdsLastLoad(...convresult.map(cr => cr.id));
        return result;
    }

    protected isPersistItemsDataAsync(linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem[] }): boolean {
        return !this.initialized || (!preloaded && this.needsPersistInner(linkedFields)) || this.hasLookup(linkedFields);
    }

    @trace(TraceLevel.Internal)
    protected async persistItemsDataAsync_internal(data: any[], linkedFields?: Array<string>, preloaded?: { [modelName: string]: BaseItem[] }): Promise<T[]> {
        let result = null;
        if (data) {
            if (!this.initialized) {
                await this.Init();
            }
            if (!preloaded && this.needsPersistInner(linkedFields)) {
                preloaded = await this.persistInner(data, linkedFields);
            }
            result = data.map(d => this.populateItem(d));
            if (this.hasLookup(linkedFields)) {
                await this.populateLookups(result, linkedFields, preloaded);
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
            (fields[k].fieldType === FieldType.Lookup || fields[k].fieldType === FieldType.LookupMulti) &&
            fields[k].containsFullObject &&
            !stringIsNullOrEmpty(fields[k].modelName)
        );
        return keys.length > 0;
    }

    protected async persistInner(objects: any[], linkedFields?: Array<string>): Promise<{ [modelName: string]: BaseItem[] }> {
        let level = 0;
        let result: { [modelName: string]: BaseItem[] } = undefined;
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
                        if (descriptor.fieldType === FieldType.Lookup) {
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
    private linkedLookupFields(loadLookups?: Array<string>): Array<IFieldDescriptor> {
        const result: Array<IFieldDescriptor> = [];
        const fields = this.ItemFields;
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key];
                if ((fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) && !stringIsNullOrEmpty(fieldDesc.modelName)) {
                    if (!loadLookups || (loadLookups.length === 1 && loadLookups[0] === 'loadAll') || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                        result[key] = fieldDesc;
                    }
                }
            }
        }

        return result;
    }

    public hasLookup(linkedFields?: Array<string>): boolean {
        const lookupFields = this.linkedLookupFields(linkedFields);
        return Object.keys(lookupFields).filter(k => lookupFields.hasOwnProperty(k)).length > 0;
    }

    @trace(TraceLevel.ServiceUtilities)
    protected async populateLookups(items: Array<T>, loadLookups?: Array<string>, innerItems?: { [modelName: string]: BaseItem[] }): Promise<void> {
        if (!this.initialized) {
            await this.Init();
        }
        // get lookup fields
        const lookupFields = this.linkedLookupFields(loadLookups);

        // init values and retrieve all ids by model
        const allIds = {};
        const cachedIds = {};
        const innerResult = {};
        for (const key in lookupFields) {
            if (lookupFields.hasOwnProperty(key)) {
                const fieldDesc = lookupFields[key];
                allIds[fieldDesc.modelName] = allIds[fieldDesc.modelName] || [];
                cachedIds[fieldDesc.modelName] = cachedIds[fieldDesc.modelName] || [];
                const ids = allIds[fieldDesc.modelName];
                const cached = cachedIds[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const links = item.__getInternalLinks(key);
                    //init value 
                    if (fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) {
                        item[key] = fieldDesc.defaultValue;
                    }
                    if (fieldDesc.fieldType === FieldType.Lookup &&
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
                            if (this.isServiceLookupIdCached(fieldDesc.modelName, links)) {
                                cached.push(links);
                            }
                            else {
                                ids.push(links);
                            }
                        }
                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti &&
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
                                if (this.isServiceLookupIdCached(fieldDesc.modelName, id)) {
                                    cached.push(id);
                                }
                                else {
                                    ids.push(id);
                                }
                            }
                        });
                    }
                });

            }
        }
        // store preloaded ids
        for (const modelName in innerResult) {
            if (innerResult.hasOwnProperty(modelName)) {
                this.updateServiceCachedLookupIds(modelName, ...innerResult[modelName]);
            }
        }

        const resultItems: { [modelName: string]: BaseItem[] } = innerResult;
        // get from cache
        // Init queries       
        const cachedpromises: Array<() => Promise<BaseItem[]>> = [];
        for (const modelName in cachedIds) {
            if (cachedIds.hasOwnProperty(modelName)) {
                const ids = cachedIds[modelName];
                if (ids && ids.length > 0) {
                    const service = ServiceFactory.getServiceByModelName(modelName);
                    cachedpromises.push(() => service.getItemsFromCacheById(ids));
                }
            }
        }
        // execute and store
        const cachedresults = await UtilsService.executePromisesInStacks(cachedpromises, 3);
        cachedresults.forEach(itemsTab => {
            if (itemsTab.length > 0) {
                resultItems[itemsTab[0].constructor["name"]] = resultItems[itemsTab[0].constructor["name"]] || [];
                resultItems[itemsTab[0].constructor["name"]].push(...itemsTab);
            }
        });

        // Not cached
        // Init queries       
        const promises: Array<() => Promise<BaseItem[]>> = [];
        for (const modelName in allIds) {
            if (allIds.hasOwnProperty(modelName)) {
                const ids = allIds[modelName];
                if (ids && ids.length > 0) {
                    const service = ServiceFactory.getServiceByModelName(modelName);
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
                this.updateServiceCachedLookupIds(itemsTab[0].constructor["name"], ...itemsTab);
            }
        });

        // Associate to items
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName];
                const refCol = resultItems[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const links = item.__getInternalLinks(propertyName);
                    if (fieldDesc.fieldType === FieldType.Lookup &&
                        links &&
                        links !== -1) {
                        const litem = find(refCol, { id: links });
                        if (litem) {
                            item[propertyName] = litem;
                        }

                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti &&
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

    protected updateInternalLinks(item: T, loadLookups?: Array<string>): void {
        const lookupFields = this.linkedLookupFields(loadLookups);
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName];
                if (!loadLookups || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                    if (fieldDesc.fieldType === FieldType.Lookup) {
                        item.__deleteInternalLinks(propertyName);
                        if (item[propertyName] && item[propertyName].id > -1) {
                            item.__setInternalLinks(propertyName, item[propertyName].id);
                        }
                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti) {
                        item.__deleteInternalLinks(propertyName);
                        if (item[propertyName] && item[propertyName].length > 0) {
                            item.__setInternalLinks(propertyName, item[propertyName].filter(l => l.id !== -1).map(l => l.id));
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
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                //link defered
                                if (item[propertyName]) {
                                    result.__setInternalLinks(propertyName, (item[propertyName] as unknown as BaseItem).id);
                                }
                                delete result[propertyName];
                            }
                            break;
                        case FieldType.UserMulti:
                        case FieldType.TaxonomyMulti:
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const ids = [];
                                if (item[propertyName]) {
                                    (item[propertyName] as unknown as BaseItem[]).forEach(element => {
                                        if (element?.id) {
                                            if ((typeof (element.id) === "number" && element.id > 0) || (typeof (element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
                                                ids.push(element.id);
                                            }
                                        }
                                    });
                                }
                                if (ids.length > 0) {
                                    result.__setInternalLinks(propertyName, ids.length > 0 ? ids : []);
                                }
                                delete result[propertyName];
                            }
                            break;
                        case FieldType.Lookup:
                        case FieldType.LookupMulti:
                            // internal links allready updated before (used for rest calls)
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                delete result[propertyName];
                                result.__setInternalLinks(propertyName, item.__getInternalLinks(propertyName));
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
        return !this.initialized || this.hasLookup(linkedFields);
    }

    /**
     * populate item from db storage
     * @param item - db item with links in internalLinks fields
     */
    @trace(TraceLevel.ServiceUtilities)
    public async mapItemsAsync(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        let results: Array<T> = [];
        if (items && items.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
            results = this.mapItems_internal(items);
        }
        if (this.hasLookup(linkedFields)) {
            await this.populateLookups(results, linkedFields);
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
                        if (
                            fieldDescriptor.fieldType === FieldType.User ||
                            fieldDescriptor.fieldType === FieldType.Taxonomy) {
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                // get values from init values
                                const id: number = item.__getInternalLinks(propertyName);
                                if (id !== null) {
                                    const destElements = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                                    const existing = find(destElements, (destElement) => {
                                        return destElement.id === id;
                                    });
                                    result[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                                }
                                else {
                                    result[propertyName] = fieldDescriptor.defaultValue;
                                }
                            }
                            result.__deleteInternalLinks(propertyName);
                        }
                        else if (
                            fieldDescriptor.fieldType === FieldType.UserMulti ||
                            fieldDescriptor.fieldType === FieldType.TaxonomyMulti) {
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                // get values from init values
                                const ids = item.__getInternalLinks(propertyName) || [];
                                if (ids.length > 0) {
                                    const val = [];
                                    const targetItems = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                                    ids.forEach(id => {
                                        const existing = find(targetItems, (tmpitem) => {
                                            return tmpitem.id === id;
                                        });
                                        if (existing) {
                                            val.push(existing);
                                        }
                                    });
                                    result[propertyName] = val;
                                }
                                else {
                                    result[propertyName] = fieldDescriptor.defaultValue;
                                }
                            }
                            result.__deleteInternalLinks(propertyName);
                        }
                        else if (fieldDescriptor.fieldType === FieldType.Json && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
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



                    let service: BaseDataService<BaseItem>;

                    try {

                        service = ServiceFactory.getServiceByModelName(modelName);

                    } catch {

                        console.warn("No service found for '" + modelName + "'");

                    }



                    if (service) {

                        const allitems = await service.__getAllFromCache();
                        const updated = [];
                        allitems.forEach(element => {
                            let needUpdate = false;
                            lookupProperties.forEach(propertyName => {
                                const fieldDescription = modelFields[propertyName];
                                if (fieldDescription.fieldType === FieldType.Lookup) {
                                    if (fieldDescription.modelName) {
                                        // serch in internalLinks
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
                            await service.__updateCache(...updated);
                        }
                    }
                }
            }
        }
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
    public async refreshData(): Promise<void> {
        // Invalidate cache
        const cacheKey = this.getCacheKey(); // Default key is "ALL"
        window.sessionStorage.removeItem(cacheKey);
        // remove local cache
        this.initValues = {};
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