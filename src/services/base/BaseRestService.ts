import { isArray, stringIsNullOrEmpty } from "@pnp/core";
import { cloneDeep, find } from "lodash";
import * as mime from "mime-types";
import { ServicesConfiguration } from "../../configuration";
import { Constants, FieldType, TestOperator, TraceLevel } from "../../constants/index";
import { Decorators } from "../../decorators";
import { IEndPointBinding } from "../../interfaces/IEndPointBindings";
import { IBaseRestServiceOptions, IBaseSPServiceOptions, IEndPointBindings, IFieldDescriptor, ILogicalSequence, IOrderBy, IPredicate, IQuery, IRestLogicalSequence, IRestPredicate, IRestQuery } from "../../interfaces/index";
import { BaseItem, RestItem, RestResultMapping, User } from "../../models";
import { RestFile } from "../../models/base/RestFile";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { UserService } from "../graph/UserService";
import { BaseDataService } from "./BaseDataService";
import { BaseDbService } from "./cache/BaseDbService";

const trace = Decorators.trace;

/**
 * 
 * Base service for sp list items operations
 */
export class BaseRestService<T extends RestItem<string | number> | RestFile<string | number>> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/

    protected restMappingDb: BaseDbService<RestResultMapping<string | number>>;
    protected serviceOptions: IBaseRestServiceOptions;

    protected baseServiceUrl: string;

    public get Bindings(): IEndPointBindings {
        return this.constructor["serviceProps"].endpoints;
    }

    public get serviceUrl(): string {
        return this.baseServiceUrl + this.constructor["serviceProps"].relativeUrl;
    }

    public get disableVersionCheck(): boolean {
        return this.constructor["serviceProps"].disableVersionCheck === true;
    }

    /***************************** Constructor **************************************/
    /**
     * 
     * @param type - items type
     * @param baseServiceUrl - base url of rest api hosting the service
     * @param tableName - name of table in local db
     * @param cacheDuration - cache duration in minutes
     */
    constructor(itemType: (new (item?: any) => T), baseServiceUrl: string, options?: IBaseRestServiceOptions, ...args: any[]) {
        super(itemType, options, baseServiceUrl, ...args);
        this.baseServiceUrl = baseServiceUrl;
        this.restMappingDb = new BaseDbService(RestResultMapping, "RestMapping");
    }

    /****************************** get item methods ***********************************/
    protected populateItem(restItem: any): T {
        const item = super.populateItem(restItem);
        if (item instanceof RestFile) {
            item.mimeType = (mime.lookup(item.title) as string) || 'application/octet-stream';
        }
        return item;
    }

    protected getLookupId(value: any, fieldDescriptor: IFieldDescriptor): (string | number) {
        if(typeof(value) === "string" || typeof(value) === "number") {
            return value;
        }
        else if(value) {
            if(!stringIsNullOrEmpty(fieldDescriptor.lookupFieldName)){
                return value[fieldDescriptor.lookupFieldName];
            }
            else {
                const modelFields = stringIsNullOrEmpty(fieldDescriptor.modelName) ? {} : ServiceFactory.getModelFields(fieldDescriptor.modelName);
                const idField = modelFields[Constants.commonRestFields.id]?.fieldName || Constants.commonRestFields.id;
                return value[idField];
            }
        }
        return undefined;    
    }

    protected populateFieldValue(restItem: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
        super.populateFieldValue(restItem, destItem, propertyName, fieldDescriptor);
        const defaultValue = cloneDeep(fieldDescriptor.defaultValue);
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Lookup:                
                const lookupId = this.getLookupId(restItem[fieldDescriptor.fieldName], fieldDescriptor);
                if (lookupId !== undefined) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, lookupId);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = lookupId;
                    }

                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.LookupMulti:                
                const lookupIds: Array<string|number> = restItem[fieldDescriptor.fieldName] && Array.isArray(restItem[fieldDescriptor.fieldName]) ?
                    restItem[fieldDescriptor.fieldName].map(ri => this.getLookupId(ri, fieldDescriptor)).filter(id => id != undefined) :
                    [];
                if (lookupIds.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, lookupIds);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = lookupIds;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.User:
                const upn: string = restItem[fieldDescriptor.fieldName];
                if (!stringIsNullOrEmpty(upn)) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, upn);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = upn;
                    }

                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.UserMulti:
                const upns: Array<string> = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                if (upns.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, upns);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = upns;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                const conJsonId = !stringIsNullOrEmpty(restItem[fieldDescriptor.fieldName]) ? JSON.parse(restItem[fieldDescriptor.fieldName]) : null;
                const termid: string = conJsonId && conJsonId.length > 0 ? conJsonId[0].id : null;
               if (!stringIsNullOrEmpty(termid)) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, termid);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = termid;
                    }

                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                const conJsonIds = !stringIsNullOrEmpty(restItem[fieldDescriptor.fieldName]) ? JSON.parse(restItem[fieldDescriptor.fieldName]) : null;
                const tmterms = conJsonIds ? conJsonIds : [];
                const tids = tmterms.map(t => t.id); 
                if (tids.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, tids);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = tids;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            default: break;
        }
    }
    /****************************** Send item methods ***********************************/

    protected get ignoredFields(): string[] {
        return [
            Constants.commonRestFields.created,
            Constants.commonRestFields.author,
            Constants.commonRestFields.editor,
            Constants.commonRestFields.modified
        ];
    }

    protected async convertFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor: IFieldDescriptor): Promise<void> {
        await super.convertFieldValue(item, destItem, propertyName, fieldDescriptor);
        const itemValue = item[propertyName];
        if (!this.isFieldIgnored(item, propertyName, fieldDescriptor)) {
            switch (fieldDescriptor.fieldType) {
                case FieldType.Lookup:
                    const link = item.__getInternalLinks(propertyName);
                    if (itemValue) {
                        if (typeof (itemValue) === "number") {
                            destItem[fieldDescriptor.fieldName] = itemValue > 0 ? itemValue : null;
                        }
                        else {
                            destItem[fieldDescriptor.fieldName] = link && link > 0 ? link : null;
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = null;
                    }
                    break;
                case FieldType.LookupMulti:
                    break;
                case FieldType.User:
                    if (itemValue) {
                        if (typeof (itemValue) === "number") {
                            destItem[fieldDescriptor.fieldName] = itemValue > 0 ? itemValue : null;
                        }
                        else {
                            destItem[fieldDescriptor.fieldName] = await this.convertSingleUserFieldValue(itemValue);
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = null;
                    }
                    break;
                case FieldType.UserMulti:
                    if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                        const firstUserVal = itemValue[0];
                        if (typeof (firstUserVal) === "number") {
                            destItem[fieldDescriptor.fieldName] = itemValue;
                        }
                        else {
                            const userIds = await Promise.all(itemValue.map((user) => {
                                return this.convertSingleUserFieldValue(user);
                            }));
                            destItem[fieldDescriptor.fieldName] = userIds;
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = [];
                    }
                    break;
                case FieldType.Taxonomy:
                    destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify([{ id: itemValue.id }]) : null;
                    break;
                case FieldType.TaxonomyMulti:
                    if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                        destItem[fieldDescriptor.fieldName] = JSON.stringify(itemValue.map((t) => { return { id: t.id }; }));
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = null;
                    }
                    break;
                default:
                    break;
            }
        }
    }


    /********************** SP Fields conversion helpers *****************************/

    private async convertSingleUserFieldValue(value: User): Promise<string> {
        let result: string = null;
        if (value) {
            if (value.isLocal) {
                const userService: UserService = ServiceFactory.getService(User).cast<UserService>();
                value = await userService.linkToSpUser(value);
            }
            result = value.userPrincipalName;
        }
        return result;
    }


    /***************** SP Calls associated to service standard operations ********************/

    @trace(TraceLevel.Queries)
    protected async get_Query(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        const restQuery = this.getRestQuery(query);
        if (linkedFields && linkedFields.length === 1 && linkedFields[0] === 'loadAll') {
            restQuery.loadAll = true;
        }
        return this.executeRequest(`${this.serviceUrl}${this.Bindings.get.url}`, this.Bindings.get.method, restQuery);
    }

    /**
     * Get an item by id
     * @param {number} id - item id
     */
    @trace(TraceLevel.Queries)
    protected async getItemById_Query(id: number, linkedFields?: Array<string>): Promise<any> {// eslint-disable-line @typescript-eslint/no-unused-vars
        return this.executeRequest(`${this.serviceUrl}${this.Bindings.getItemById.url}/${id}`, this.Bindings.getItemById.method);
    }


    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    @trace(TraceLevel.Queries)
    protected async getItemsById_Query(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<any>> {
        const result: Array<T> = [];
        const promises: (() => Promise<Array<any>>)[] = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 2000);
            promises.push(() => this.get_Query({
                test: {
                    type: "predicate",
                    operator: TestOperator.In,
                    propertyName: "id",
                    value: sub
                },
                limit: 2000
            }, linkedFields));
        }
        const res = await UtilsService.executePromisesInStacks(promises, 3);
        for (const tmp of res) {
            result.push(...tmp.filter(i => { return i !== null && i !== undefined; }));
        }
        return result;
    }

    /**
     * Retrieve all items
     * 
     */
    @trace(TraceLevel.Queries)
    protected async getAll_Query(linkedFields?: Array<string>): Promise<Array<any>> {// eslint-disable-line @typescript-eslint/no-unused-vars
        return this.executeRequest(`${this.serviceUrl}${this.Bindings.getAll.url}`, this.Bindings.getAll.method);
    }

    /**
     * Add or update an item
     * @param item - SPItem derived object to be converted
     */
    @trace(TraceLevel.Internal)
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        const result = cloneDeep(item);
        if (item.isLocal) {
            const converted = await this.convertItem(item);
            const addResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
            await this.populateCommonFields(result, addResult);
            if (item.isCreatedOffline) {
                await this.updateLinksInDb(Number(item.id), Number(result.id));
            }
        }
        else {
            // check version (cannot update if newer)
            if (item.version && !this.disableVersionCheck) {
                const existing = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getItemById.url}/${item.id}`, this.Bindings.getItemById.method);
                if (parseFloat(existing[Constants.commonRestFields.version]) > item.version) {
                    const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    const converted = await this.convertItem(item);
                    const updateResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
                    await this.populateCommonFields(result, updateResult);
                }
            }
            else {
                const converted = await this.convertItem(item);
                try {
                    const updateResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
                    await this.populateCommonFields(result, updateResult);
                } catch (error) {
                    if (error.name === "409") {
                        const conflicterror = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                        conflicterror.name = Constants.Errors.ItemVersionConfict;
                        throw conflicterror;
                    }
                    else {
                        throw error;
                    }
                }
            }
        }
        return result;
    }

    /**
     * Add or update items in batch
     * @param items Array of model type to be added or updated
     * @param onItemUpdated callback function called when an item has been added or updated
     */
    @trace(TraceLevel.Internal)
    protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        const result = cloneDeep(items);
        const itemsToAdd = result.filter(item => item.isLocal);
        const versionedItems = result.filter((item) => {
            return !this.disableVersionCheck && item.version !== undefined && item.version !== null && !item.isLocal;
        });
        const updatedItems = result.filter((item) => {
            return (this.disableVersionCheck || item.version === undefined || item.version === null) && !item.isLocal;
        });

        // creation batch
        if (itemsToAdd.length > 0) {
            let idx = 0;
            // TODO call stacks
            while (itemsToAdd.length > 0) {
                const sub = itemsToAdd.splice(0, 100);
                const converted = await Promise.all(sub.map(item => this.convertItem(item)));
                try {
                    const addResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItems.url}`, this.Bindings.addOrUpdateItems.method, converted);
                    for (let index = 0; index < sub.length; index++) {
                        const item = sub[index];
                        const currentIdx = idx;
                        const itemId = item.id;
                        await this.populateCommonFields(item, addResult[index]);
                        if (item.isCreatedOffline) {
                            await this.updateLinksInDb(Number(itemId), Number(item.id));
                        }
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                        idx++;
                    }
                } catch (error) {
                    for (let index = 0; index < sub.length; index++) {
                        const currentIdx = idx;
                        const item = sub[index];
                        item.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                        idx++;
                    }
                }
            }
        }

        // versionned batch --> check conflicts
        if (versionedItems.length > 0) {
            let idx = 0;
            // TODO: Call stack
            while (versionedItems.length > 0) {
                const sub = versionedItems.splice(0, 100);
                // get items to check version
                try {
                    const restQuery = this.getRestQuery({
                        test: {
                            type: "predicate",
                            operator: TestOperator.In,
                            propertyName: "id",
                            value: sub.map(item => item.id)
                        },
                        limit: 2000
                    });
                    const versionitems = await this.executeRequest(`${this.serviceUrl}${this.Bindings.get.url}`, this.Bindings.get.method, restQuery);
                    for (const subitem of sub) {
                        const currentIdx = idx;
                        const existing = find(versionitems, i => { return i.id === subitem.id; });
                        if (parseFloat(existing[Constants.commonRestFields.version]) > subitem.version) {
                            const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                            error.name = Constants.Errors.ItemVersionConfict;
                            subitem.error = error;
                            if (onItemUpdated) {
                                onItemUpdated(items[currentIdx], subitem);
                            }
                        }
                        else {
                            updatedItems.push(subitem);
                        }
                        idx++;
                    }
                }
                catch (error) {
                    for (const subitem of sub) {
                        const currentIdx = idx;
                        subitem.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], subitem);
                        }
                        idx++;
                    }
                }
            }
        }
        // classical update + version checked
        if (updatedItems.length > 0) {
            let idx = 0;
            // TODO: Call stack
            while (updatedItems.length > 0) {
                const sub = updatedItems.splice(0, 100);
                try {
                    // TODO : Manage version conflicts in batch
                    const converted = await Promise.all(sub.map(item => this.convertItem(item)));
                    const results = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItems.url}`, this.Bindings.addOrUpdateItems.method, converted);
                    for (let index = 0; index < sub.length; index++) {
                        const subitem = sub[index];
                        const currentIdx = idx;
                        if (results[index]) {
                            await this.populateCommonFields(subitem, results[index]);
                        }
                        else {
                            // item is null --> conflict
                            const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                            error.name = Constants.Errors.ItemVersionConfict;
                            subitem.error = error;
                        }
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], subitem);
                        }
                        idx++;
                    }
                } catch (error) {
                    for (const subitem of sub) {
                        const currentIdx = idx;
                        subitem.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], subitem);
                        }
                        idx++;
                    }
                }
            }
        }
        return result;
    }

    /**
     * Delete an item
     * @param item - SPItem derived class to be deleted
     */
    @trace(TraceLevel.Internal)
    protected async deleteItem_Internal(item: T): Promise<T> {
        try {
            const result = await this.executeRequest(`${this.serviceUrl}${this.Bindings.deleteItem.url}/${item.id}`, this.Bindings.deleteItem.method);
            item.deleted = result;
        }
        catch (error) {
            item.error = error;
        }
        return item;
    }

    /**
     * Delete an item
     * @param item - SPItem derived class to be deleted
     */
    @trace(TraceLevel.Internal)
    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        try {
            const results = await this.executeRequest(`${this.serviceUrl}${this.Bindings.deleteItems.url}`, this.Bindings.deleteItems.method, items.map(i => i.id));
            for (let index = 0; index < items.length; index++) {
                items[index].deleted = results[index];
            }
        } catch (error) {
            items.forEach(i => i.error = error);
        }
        return items;
    }

    @trace(TraceLevel.Internal)
    protected recycleItem_Internal(items: T[]): Promise<T[]> {
        throw new Error("Method not implemented." + JSON.stringify(items));
    }

    @trace(TraceLevel.Internal)
    protected recycleItems_Internal(items: T[]): Promise<T[]> {
        throw new Error("Method not implemented." + JSON.stringify(items));
    }

    @trace(TraceLevel.Service)
    public async getByRestQuery(restQuery: IEndPointBinding, data?: any, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(restQuery).toString() + super.hashCode(data).toString() + super.hashCode(linkedFields).toString();
        return this.callAsyncWithPromiseManagement(async () => {
            let result = new Array<T>();
            //has to refresh cache
            let reloadData = this.needRefreshCache(keyCached);
            //if refresh is needed, test offline/online
            if (reloadData && ServicesConfiguration.configuration.checkOnline) {
                reloadData = navigator.onLine;
            }
            if(!reloadData) {
                try {
                    const mapping = await this.restMappingDb.getItemById(keyCached);
                    if (mapping && mapping.itemIds && mapping.itemIds.length > 0) {
                        const tmp = await this.cacheService.getItemsById(mapping.itemIds);
                        if (this.isMapItemsAsync(linkedFields)) {
                            result = await this.mapItemsAsync(tmp, linkedFields);
                        }
                        else {
                            result = this.mapItemsSync(tmp);
                        }
                    }
                } catch (error) {
                    console.error(error);
                    reloadData = !ServicesConfiguration.configuration.checkOnline || navigator.onLine;
                }
            }
            if (reloadData) {
                const json = await this.executeRequest(restQuery.url, restQuery.method, data);
                if (this.isPersistItemsDataAsync(linkedFields)) {
                    result = await this.persistItemsDataAsync_internal(json, linkedFields);
                }
                else {
                    result = this.persistItemsDataSync_internal(json);
                }
                if(this.hasCache) {
                    //check if data exist for this query in database
                    let mapping = await this.restMappingDb.getItemById(keyCached);
                    if (mapping) {
                        const tmp = await this.cacheService.getItemsById(mapping.itemIds);
                        //if data exists trash them 
                        if (tmp && tmp.length > 0) {
                            await this.cacheService.deleteItems(tmp);
                        }
                    }
                    if (result && result.length > 0) {
                        const convresult = result.map(res => this.convertItemToDbFormat(res));
                        await this.cacheService.addOrUpdateItems(convresult);
                        mapping = new RestResultMapping();
                        mapping.id = keyCached;
                        mapping.itemIds = convresult.map(r => r.id);
                        await this.restMappingDb.addOrUpdateItem(mapping);
                        this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        this.UpdateCacheData(keyCached);
                    }
                    else if (mapping) {
                        this.resetIdsLastLoad();
                        this.resetCacheData(keyCached);
                        await this.restMappingDb.deleteItem(mapping);
                    }
                }
            }
            
            return result;
        }, keyCached);
    }

    /********************** Overrides for user field **************************************************/
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
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                //link defered
                                if (item[propertyName]) {
                                    result.__setInternalLinks(propertyName, (item[propertyName] as unknown as User).userPrincipalName);
                                }
                                delete result[propertyName];
                            }
                            break;
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
                            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const upns = [];
                                if (item[propertyName]) {
                                    (item[propertyName] as unknown as User[]).forEach(element => {
                                        if (!stringIsNullOrEmpty(element?.userPrincipalName)) {
                                            upns.push(element.userPrincipalName);
                                        }
                                    });
                                }
                                if (upns.length > 0) {
                                    result.__setInternalLinks(propertyName, upns.length > 0 ? upns : []);
                                }
                                delete result[propertyName];
                            }
                            break;
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

    protected async populateLinkedFields(items: T[], loadLinked?: string[], innerItems?: { [modelName: string]: BaseItem<string | number>[]; }): Promise<void> {
        await super.populateLinkedFields(items, loadLinked, innerItems);        
        // get linked fields
        const linkedFields = this.linkedFields(loadLinked).filter(lf => lf.fieldType === FieldType.User || lf.fieldType === FieldType.UserMulti);
        // init values and retrieve all ids by model
        const allUpns = {};
        const innerResult = {};
        for (const key in linkedFields) {
            if (linkedFields.hasOwnProperty(key)) {
                const fieldDesc = linkedFields[key];
                allUpns[fieldDesc.modelName] = allUpns[fieldDesc.modelName] || [];
                const upns = allUpns[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const links = item.__getInternalLinks(key);
                    //init value 
                    item[key] = fieldDesc.defaultValue;
                    
                    if (fieldDesc.fieldType === FieldType.User &&
                        // lookup has value
                        !stringIsNullOrEmpty(links)) {
                        // check in preloaded
                        let inner = undefined;
                        if (innerItems && innerItems[fieldDesc.modelName]) {
                            inner = find(innerItems[fieldDesc.modelName], ii => (ii as User).userPrincipalName === links);
                        }
                        // inner found
                        if (inner) {
                            innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                            innerResult[fieldDesc.modelName].push(inner);
                        }
                        else {
                            upns.push(links);
                        }
                    }
                    else if (fieldDesc.fieldType === FieldType.UserMulti &&
                        links &&
                        links.length > 0) {
                        links.forEach((upn) => {
                            let inner = undefined;
                            if (innerItems && innerItems[fieldDesc.modelName]) {
                                inner = find(innerItems[fieldDesc.modelName], ii => (ii as User).userPrincipalName === upn);
                            }
                            // inner found
                            if (inner) {
                                innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                                innerResult[fieldDesc.modelName].push(inner);
                            }
                            else {
                                upns.push(upn);
                            }
                        });
                    }
                });

            }
        }
        const resultItems: { [modelName: string]: User[] } = innerResult;
        
        // Init queries       
        const promises: Array<() => Promise<User[]>> = [];
        for (const modelName in allUpns) {
            if (allUpns.hasOwnProperty(modelName)) {
                const upns = allUpns[modelName];
                if (upns && upns.length > 0) {
                    const options: IBaseSPServiceOptions = {};
                    // for sp services
                    if(this.serviceOptions.hasOwnProperty('baseUrl')) {
                        options.baseUrl = (this.serviceOptions as IBaseSPServiceOptions).baseUrl;
                    }
                    const service = ServiceFactory.getServiceByModelName(modelName, options);
                    promises.push(() => (service as UserService).getByEmails(upns));
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
                    if (fieldDesc.fieldType === FieldType.User &&
                        !stringIsNullOrEmpty(links)) {
                        const litem = find(refCol, { userPrincipalName: links });
                        if (litem) {
                            item[propertyName] = litem;
                        }

                    }
                    else if (fieldDesc.fieldType === FieldType.UserMulti &&
                        links &&
                        links.length > 0) {
                        item[propertyName] = [];
                        links.forEach((upn) => {
                            const litem = find(refCol, { userPrincipalName: upn });
                            if (litem) {
                                item[propertyName].push(litem);
                            }
                        });
                    }
                });
            }
        }
    }

    /************************** Query filters ***************************/

    protected async populateCommonFields(item: T, restItem): Promise<void> {
        if (item.isLocal) {
            // update ids
            item.id = restItem[Constants.commonRestFields.id];
            item.uniqueId = restItem[Constants.commonRestFields.uniqueid];
        }
        if (restItem[Constants.commonRestFields.version] !== undefined) {
            item.version = restItem[Constants.commonRestFields.version];
        }
        const fields = this.ItemFields;
        await Promise.all(Object.keys(fields).filter((propertyName) => {
            if (fields.hasOwnProperty(propertyName)) {
                const fieldName = fields[propertyName].fieldName;
                return (fieldName === Constants.commonRestFields.author ||
                    fieldName === Constants.commonRestFields.created ||
                    fieldName === Constants.commonRestFields.editor ||
                    fieldName === Constants.commonRestFields.modified);
            }
        }).map(async (prop) => {
            const fieldName = fields[prop].fieldName;
            switch (fields[prop].fieldType) {
                case FieldType.Date:
                    if (restItem[fieldName]) {
                        item[prop] = new Date(restItem[fieldName]);
                    }
                    else {
                        item[prop] = fields[prop].defaultValue;
                    }

                    break;
                case FieldType.User:
                    const upn: string = restItem[fieldName];
                    if (!stringIsNullOrEmpty(upn)) {
                        let user: User = null;                        
                        const userService: UserService = ServiceFactory.getService(User).cast<UserService>();
                        user = new User();
                        user.userPrincipalName = upn;
                        user = await userService.linkToSpUser(user);
                        item[prop] = user;
                    }
                    else {
                        item[prop] = fields[prop].defaultValue;
                    }
                    break;
                default:
                    item[prop] = restItem[fieldName];
                    break;
            }
        }));

    }

    protected getRestQuery(query: IQuery<T>): IRestQuery<T> {
        const result: IRestQuery<T> = {};
        if (query) {
            result.lastId = query.lastId as number;
            result.limit = query.limit;
            result.orderBy = this.getOrderBy(query.orderBy);
            if (query.test) {
                if (query.test.type === "sequence") {
                    result.test = this.getRestSequence(query.test);
                }
                else {
                    result.test = {
                        predicates: [this.getRestPredicate(query.test)]
                    };
                }
            }
        }
        return result;
    }

    private getOrderBy(orderby: IOrderBy<T, keyof T>[]): IOrderBy<T, keyof T>[] {
        const result = [];
        if (orderby) {
            orderby.forEach(ob => {
                const copy = cloneDeep(ob);
                copy.propertyName = this.ItemFields[ob.propertyName.toString()].fieldName as keyof T;
                result.push(copy);
            });
        }
        return result;
    }

    private getRestSequence(sequence: ILogicalSequence<T>): IRestLogicalSequence<T> {
        const result: IRestLogicalSequence<T> = {
            logicalOperator: sequence.operator,
            predicates: [],
            sequences: []
        };
        sequence.children.forEach((child) => {
            if (child.type === "predicate") {
                result.predicates.push(this.getRestPredicate(child));
            }
            else {
                const seq = this.getRestSequence(child);
                result.sequences.push(seq);
            }
        });
        return result;
    }
    private getRestPredicate(predicate: IPredicate<T, keyof T>): IRestPredicate<T, keyof T> {

        return {
            logicalOperator: predicate.operator,
            propertyName: this.ItemFields[predicate.propertyName.toString()].fieldName as keyof T,
            value: predicate.value,
            includeTimeValue: predicate.includeTimeValue,
            lookupId: predicate.lookupId
        };
    }

    public async getToken(): Promise<string> {
        const aadTokenProvider = await ServicesConfiguration.context.aadTokenProviderFactory.getTokenProvider();
        const token = await aadTokenProvider.getToken(ServicesConfiguration.configuration.aadAppId);
        if (stringIsNullOrEmpty(token)) {
            throw Error("Error while getting authentication token");
        }
        return `Bearer ${token}`;
    }


    private async initRequest(method: string, data?: any): Promise<RequestInit> {
        try {
            const aadTokenProvider = await ServicesConfiguration.context.aadTokenProviderFactory.getTokenProvider();
            const token = await aadTokenProvider.getToken(ServicesConfiguration.configuration.aadAppId);
            if (stringIsNullOrEmpty(token)) {
                throw Error("Error while getting authentication token");
            }
        } catch (error) {

        }

        const headers = {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': this.serviceUrl,
            'Access-Control-Allow-Headers': "*",
            'authorization': await this.getToken()
        };
        if (data != null) {
            const postData: string = JSON.stringify(data);
            return {
                method: method,
                body: postData,
                mode: 'cors',
                headers: headers,
                referrer: ServicesConfiguration.baseUrl,
                referrerPolicy: "no-referrer-when-downgrade"
            };
        }
        return {
            method: method,
            mode: 'cors',
            headers: headers,
            referrer: ServicesConfiguration.baseUrl,
            referrerPolicy: "no-referrer-when-downgrade"
        };
    }

    protected async executeRequest(url: string, method: string, data?: any): Promise<any> {
        const req = await this.initRequest(method, data);
        const response = await fetch(url, req);
        if (response.ok) {
            return response.json();
        }
        else {
            const error = new Error();
            error.message = "Error while executing request";
            error.name = response.status.toString();
            error.stack = await response.text();
            console.error(error.toString(), "\n", error.stack);
            throw error;
        }
    }
}
