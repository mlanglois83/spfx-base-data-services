import { assign, cloneDeep, findIndex } from "@microsoft/sp-lodash-subset";
import { IBaseItem, IDataService, IQuery, ILogicalSequence, IPredicate } from "../../interfaces";
import { OfflineTransaction, TaxonomyTerm } from "../../models";
import { UtilsService } from "../UtilsService";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { Text } from "@microsoft/sp-core-library";
import { TransactionType, Constants, LogicalOperator, TestOperator, QueryToken } from "../../constants";
import { ServicesConfiguration } from "../../configuration";
import { stringIsNullOrEmpty } from "@pnp/common";


/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp 
 */
export abstract class BaseDataService<T extends IBaseItem> extends BaseService implements IDataService<T> {
    private itemModelType: (new (item?: any) => T);
    protected transactionService: TransactionService;
    protected dbService: BaseDbService<T>;
    protected cacheDuration = -1;

    public get ItemFields(): any {
        return {};
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
        return;
    }

    /**
     * 
     * @param type - type of items
     * @param context - context of the current wp
     * @param tableName - name of indexedDb table 
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration = -1) {
        super();
        this.itemModelType = type;
        this.cacheDuration = cacheDuration;
        this.dbService = new BaseDbService<T>(type, tableName);
        this.transactionService = new TransactionService();
    }

    protected getCacheKey(key = "all"): string {
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
    }

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

    protected abstract getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>>;

    /* 
     * Retrieve all elements from datasource depending on connection is enabled
     * If service is not configured as offline, an exception is thrown;
     */
    public async getAll(linkedFields?: Array<string>): Promise<Array<T>> {
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


    public async get(query: IQuery, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(query).toString();
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
                        result = await this.get_Internal(query, linkedFields);
                        //check if data exist for this query in database
                        let tmp = await this.dbService.get(query);
                        tmp = this.filterItems(query, tmp);

                        //if data exists trash them 
                        if (tmp && tmp.length > 0) {
                            await Promise.all(tmp.map((dbItem) => { return this.dbService.deleteItem(dbItem); }));
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

    public async getItemById(id: number, linkedFields?: Array<string>): Promise<T> {
        const promiseKey = "getById_" + id.toString();
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
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

    public async getItemsById(ids: Array<number | string>, linkedFields?: Array<string>): Promise<Array<T>> {
        const promiseKey = "getByIds_" + ids.join();
        let promise = this.getExistingPromise(promiseKey);
        if (promise) {
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

    public async addOrUpdateItems(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        let results: Array<T> = [];

        let isconnected = true;
        if (ServicesConfiguration.configuration.checkOnline) {
            isconnected = await UtilsService.CheckOnline();
        }
        if (isconnected) {
            results = await this.addOrUpdateItems_Internal(items, onItemUpdated);
            const versionErrors = results.filter((res) => {
                return res.error !== null || res.error !== undefined && res.error.name === Constants.Errors.ItemVersionConfict;
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


    protected abstract deleteItem_Internal(item: T): Promise<void>;

    public async deleteItem(item: T): Promise<void> {
        let isconnected = true;
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
            const ot: OfflineTransaction = new OfflineTransaction();
            const converted = await this.convertItemToDbFormat(item);
            ot.itemData = assign({}, converted);
            ot.itemType = item.constructor["name"];
            ot.title = TransactionType.Delete;
            await this.transactionService.addOrUpdateItem(ot);
        }

        return null;
    }


    protected async convertItemToDbFormat(item: T): Promise<T> {
        return item;
    }

    public mapItems(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> { // eslint-disable-line @typescript-eslint/no-unused-vars
        return Promise.resolve(items);
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


    ////////////////////////////// Queries ////////////////////////////////////
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
                        if (order.ascending === false) {
                            if ((aKey || "").localeCompare(bKey || "") < 0) {
                                return 1;
                            }
                            if ((aKey || "").localeCompare(bKey || "") > 0) {
                                return -1;
                            }
                        }
                        else {
                            if ((aKey || "").localeCompare(bKey || "") < 0) {
                                return -1;
                            }
                            if ((aKey || "").localeCompare(bKey || "") > 0) {
                                return 1;
                            }
                        }
                    }
                    else if (aKey.id && bKey.id) {
                        if (order.ascending === false) {
                            if ((aKey.title || "").localeCompare(bKey.title || "") < 0) {
                                return 1;
                            }
                            if ((aKey.title || "").localeCompare(bKey.title || "") > 0) {
                                return -1;
                            }
                        }
                        else {
                            if ((aKey.title || "").localeCompare(bKey.title || "") < 0) {
                                return -1;
                            }
                            if ((aKey.title || "").localeCompare(bKey.title || "") > 0) {
                                return 1;
                            }
                        }
                    }
                    else {
                        if (order.ascending === false) {
                            if (aKey < bKey) {
                                return 1;
                            }
                            if (aKey.title > bKey) {
                                return -1;
                            }
                        }
                        else {
                            if (aKey < bKey) {
                                return -1;
                            }
                            if (aKey.title > bKey) {
                                return 1;
                            }
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
                    result = value.wssids.indexOf(refVal) !== -1;
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
                    result = Array.isArray(refVal) && refVal.some(v => value.wssids.indexOf(v) !== -1);
                }
                else {
                    result = Array.isArray(refVal) && refVal.some(v => v === value);
                }
                break;
            case TestOperator.Includes:
                if (Array.isArray(value)) {
                    for (const lookup of value) {
                        if (predicate.lookupId) {
                            if (lookup && lookup.id) {
                                if (typeof (lookup.id) === "number") {
                                    result = lookup === refVal;
                                }
                                else if (lookup instanceof TaxonomyTerm) {
                                    result = lookup.wssids.indexOf(refVal) !== -1;
                                }

                            }
                        }
                        else if (lookup && lookup.id) {
                            result = lookup.title === refVal;
                        }
                        if (result) {
                            break;
                        }
                    }
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
                    result = value.wssids.indexOf(refVal) === -1;
                }
                else {
                    result = value !== refVal;
                }
                break;
            case TestOperator.NotIncludes:
                if (Array.isArray(value)) {
                    result = true;
                    for (const lookup of value) {
                        if (predicate.lookupId) {
                            if (lookup && lookup.id) {
                                if (typeof (lookup.id) === "number") {
                                    result = lookup === refVal;
                                }
                                else if (lookup instanceof TaxonomyTerm) {
                                    result = lookup.wssids.indexOf(refVal) !== -1;
                                }

                            }
                        }
                        else if (lookup && lookup.id) {
                            result = lookup.title === refVal;
                        }
                        if (result) {
                            break;
                        }
                        if (!result) {
                            break;
                        }
                    }
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