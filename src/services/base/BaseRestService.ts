import { ServicesConfiguration } from "../..";
import { cloneDeep, find, assign, findIndex } from "@microsoft/sp-lodash-subset";
import { Constants, FieldType, TestOperator } from "../../constants/index";
import { IBaseItem, IFieldDescriptor, IQuery, ILogicalSequence, IRestQuery, IRestLogicalSequence, IEndPointBindings, IPredicate, IRestPredicate } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { UtilsService } from "..";
import { RestItem, User, OfflineTransaction } from "../../models";
import { BaseItem } from "../../models/base/BaseItem";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";

/**
 * 
 * Base service for sp list items operations
 */
export class BaseRestService<T extends IBaseItem> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/
    protected initValues: any = {};

    protected baseServiceUrl: string;

    /* AttachmentService */

    public get ItemFields(): any {
        const result = {};
        assign(result, this.itemType["Fields"][RestItem["name"]]);
        if (this.itemType["Fields"][this.itemType["name"]]) {
            assign(result, this.itemType["Fields"][this.itemType["name"]]);
        }
        return result;
    }

    public get Bindings(): IEndPointBindings {
        return this.constructor["serviceProps"].endpoints;
    }

    public get serviceUrl(): string {
        return this.baseServiceUrl + this.constructor["serviceProps"].relativeUrl;
    }

    /***************************** Constructor **************************************/
    /**
     * 
     * @param type - items type
     * @param baseServiceUrl - base url of rest api hosting the service
     * @param tableName - name of table in local db
     * @param cacheDuration - cache duration in minutes
     */
    constructor(type: (new (item?: any) => T), baseServiceUrl: string, tableName: string, cacheDuration?: number) {
        super(type, tableName, cacheDuration);
        this.baseServiceUrl = baseServiceUrl;
    }


    /***************************** External sources init and access **************************************/

    private initialized = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }
    private initPromise: Promise<void> = null;

    protected async init_internal(): Promise<void> {
        return;
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
                        const fields = this.ItemFields;
                        const models = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                // REM MLS : lookup removed from preload
                                if (fieldDescription.modelName &&
                                    models.indexOf(fieldDescription.modelName) === -1 &&
                                    fieldDescription.fieldType !== FieldType.Lookup &&
                                    fieldDescription.fieldType !== FieldType.LookupMulti) {
                                    models.push(fieldDescription.modelName);
                                }
                            }
                        }
                        await Promise.all(models.map(async (modelName) => {
                            if (!this.initValues[modelName]) {
                                const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                const values = await service.getAll();
                                this.initValues[modelName] = values;
                            }
                        }));
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
    
    private getServiceInitValues(modelName: string): any {
        return this.initValues[modelName];
    }

    /****************************** get item methods ***********************************/
    private getItemFromRest(restItem: any): T {
        const item = new this.itemType();
        Object.keys(this.ItemFields).map((propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            this.setFieldValue(restItem, item, propertyName, fieldDescription);
        });
        return item;
    }

    // TODO : test
    private setFieldValue(restItem: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
        const converted = destItem as unknown as RestItem;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:                
                converted[propertyName] = restItem[fieldDescriptor.fieldName] !== null && restItem[fieldDescriptor.fieldName] !== undefined ? restItem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                break;
            case FieldType.Date:
                converted[propertyName] = restItem[fieldDescriptor.fieldName] ? new Date(restItem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                const lookupId: number = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : -1;
                if (lookupId !== -1) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        converted.__setInternalLinks(propertyName, lookupId);
                        converted[propertyName] = fieldDescriptor.defaultValue;

                    }
                    else {
                        converted[propertyName] = lookupId;
                    }

                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }

                break;
            case FieldType.LookupMulti:
                const lookupIds: Array<number> = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                if (lookupIds.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // LOOKUPS --> links
                        converted.__setInternalLinks(propertyName, lookupIds);
                        converted[propertyName] = fieldDescriptor.defaultValue;
                    }
                    else {
                        converted[propertyName] = lookupIds;
                    }
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.User:
                const upn: string = restItem[fieldDescriptor.fieldName];
                if (!stringIsNullOrEmpty(upn)) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const users = this.getServiceInitValues(fieldDescriptor.modelName);
                        const existing = find(users, (user: User) => {
                            return !stringIsNullOrEmpty(user.userPrincipalName) && user.userPrincipalName.toLowerCase() === upn.toLowerCase();
                        });
                        converted[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        converted[propertyName] = upn;
                    }
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.UserMulti:
                const upns: Array<string> = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                if (upns.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const val = [];
                        const users = this.getServiceInitValues(fieldDescriptor.modelName);
                        upns.forEach(umupn => {
                            if(!stringIsNullOrEmpty(umupn)) {
                                const existing = find(users, (user: User) => {
                                    return !stringIsNullOrEmpty(user.userPrincipalName) && user.userPrincipalName.toLowerCase() === umupn.toLowerCase();
                                });
                                if (existing) {
                                    val.push(existing);
                                }
                            }                            
                        });
                        converted[propertyName] = val;
                    }
                    else {
                        converted[propertyName] = upns;
                    }
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                const termid: string = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : null;
                if (!stringIsNullOrEmpty(termid)) {
                    const tterms = this.getServiceInitValues(fieldDescriptor.modelName);
                    const existing = find(tterms, (term) => {
                        return term.id === termid;
                    });
                    converted[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                const tmterms = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                if (tmterms.length > 0) {
                    // get values from init values
                    const val = [];
                    const allterms = this.getServiceInitValues(fieldDescriptor.modelName);
                    tmterms.forEach(tmterm => {
                        const existing = find(allterms, (term) => {
                            return term.id === tmterm;
                        });
                        if (existing) {
                            val.push(existing);
                        }
                    });
                    converted[propertyName] = val;
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Json:
                converted[propertyName] = restItem[fieldDescriptor.fieldName] ? JSON.parse(restItem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
    }
    /****************************** Send item methods ***********************************/
    private async getRestItem(item: T): Promise<any> {
        const restItem = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            await this.setRestFieldValue(item, restItem, propertyName, fieldDescription);
        }));
        return restItem;
    }


    // TODO : test
    private async setRestFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor: IFieldDescriptor): Promise<void> {
        const converted = item as unknown as RestItem;
        const itemValue = converted[propertyName];
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
            case FieldType.Date:
                if (fieldDescriptor.fieldName !== Constants.commonFields.author &&
                    fieldDescriptor.fieldName !== Constants.commonFields.created &&
                    fieldDescriptor.fieldName !== Constants.commonFields.editor &&
                    fieldDescriptor.fieldName !== Constants.commonFields.modified &&
                    fieldDescriptor.fieldName !== Constants.commonFields.version) {
                    destItem[fieldDescriptor.fieldName] = itemValue;
                }
                break;
            case FieldType.Lookup:
                const link = converted.__getInternalLinks(propertyName);
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
                if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                    const links = converted.__getInternalLinks(propertyName);
                    const firstLookupVal = itemValue[0];
                    if (typeof (firstLookupVal) === "number") {
                        destItem[fieldDescriptor.fieldName] = itemValue;
                    }
                    else {
                        if (links && links.length > 0) {
                            destItem[fieldDescriptor.fieldName] = links;
                        }
                        else {
                            destItem[fieldDescriptor.fieldName] = [];
                        }
                    }
                }
                else {
                    destItem[fieldDescriptor.fieldName] = [];
                }
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
                destItem[fieldDescriptor.fieldName] = itemValue ? itemValue.id : null;
                break;
            case FieldType.TaxonomyMulti:
                if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                    destItem[fieldDescriptor.fieldName] = itemValue.map(t => t.id);
                }
                else {
                    destItem[fieldDescriptor.fieldName] = null;
                }
                break;
            case FieldType.Json:
                destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                break;
        }
    }

    /****************************** Lookup loading **************************************/

    /********************** SP Fields conversion helpers *****************************/
    
    private async convertSingleUserFieldValue(value: User): Promise<string> {
        let result: any = null;
        if (value) {
            if (value.id <= 0) {
                const userService: UserService = new UserService();
                value = await userService.linkToSpUser(value);

            }
            result = value.userPrincipalName;
        }
        return result;
    }

    /**********************************Service specific calls  *******************************/


    /********************************** Link to lookups  *************************************/
    private linkedLookupFields(loadLookups?: Array<string>): any {
        const result = [];
        const fields = this.ItemFields;
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key] as IFieldDescriptor;
                if ((fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) && !stringIsNullOrEmpty(fieldDesc.modelName)) {
                    if (!loadLookups || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                        result[key] = fieldDesc;
                    }
                }
            }
        }

        return result;
    }

    private async populateLookups(items: Array<T>, loadLookups?: Array<string>): Promise<void> {
        // get lookup fields
        const lookupFields = this.linkedLookupFields(loadLookups);
        // init values and retrieve all ids by model
        const allIds = {};
        for (const key in lookupFields) {
            if (lookupFields.hasOwnProperty(key)) {
                const fieldDesc = lookupFields[key] as IFieldDescriptor;
                allIds[fieldDesc.modelName] = allIds[fieldDesc.modelName] || [];
                const ids = allIds[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const converted = item as unknown as BaseItem;
                    const links = converted.__getInternalLinks(key);
                    //init value 
                    if (fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) {
                        converted[key] = fieldDesc.defaultValue;
                    }
                    if (fieldDesc.fieldType === FieldType.Lookup &&
                        // lookup has value
                        links &&
                        links !== -1 &&
                        // not allready loaded (local cache)
                        (!this.initValues[fieldDesc.modelName]
                            ||
                            !find(this.initValues[fieldDesc.modelName], { id: links })
                        ) &&
                        // not allready in load list
                        ids.indexOf(links) === -1
                    ) {

                        ids.push(links);
                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti &&
                        links &&
                        links.length > 0) {
                        links.forEach((id) => {
                            if (// not allready loaded (local cache)
                                (!this.initValues[fieldDesc.modelName]
                                    ||
                                    !find(this.initValues[fieldDesc.modelName], { id: id })
                                ) &&
                                // not allready in load list
                                ids.indexOf(id) === -1) {
                                ids.push(id);
                            }
                        });
                    }
                });
            }
        }
        // Init queries       
        const promises: Array<Promise<BaseItem[]>> = [];
        for (const modelName in allIds) {
            if (allIds.hasOwnProperty(modelName)) {
                const ids = allIds[modelName];
                if (ids && ids.length > 0) {
                    const service = ServicesConfiguration.configuration.serviceFactory.create(modelName) as BaseDataService<BaseItem>;
                    promises.push(service.getItemsById(ids));
                }
            }
        }
        // execute and store
        const results = await UtilsService.runPromisesInStacks(promises, 3);
        results.forEach(itemsTab => {
            if (itemsTab.length > 0) {
                const modelName = itemsTab[0].constructor["name"];
                this.initValues[modelName] = this.initValues[modelName] || [];
                this.initValues[modelName].push(...itemsTab);
            }
        });
        // Associate to items
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName] as IFieldDescriptor;
                const refCol = this.initValues[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const converted = item as unknown as BaseItem;
                    const links = converted.__getInternalLinks(propertyName);
                    if (fieldDesc.fieldType === FieldType.Lookup &&
                        links &&
                        links !== -1) {
                        const litem = find(refCol, { id: links });
                        if (litem) {
                            converted[propertyName] = litem;
                        }

                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti &&
                        links &&
                        links.length > 0) {
                        item[propertyName] = [];
                        links.forEach((id) => {
                            const litem = find(refCol, { id: id });
                            if (litem) {
                                converted[propertyName].push(litem);
                            }
                        });
                    }
                });
            }
        }
    }

    private updateInternalLinks(item: T, loadLookups?: Array<string>): void {
        const converted = item as unknown as BaseItem;
        const lookupFields = this.linkedLookupFields();
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName] as IFieldDescriptor;
                if (!loadLookups || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                    if (fieldDesc.fieldType === FieldType.Lookup) {
                        converted.__deleteInternalLinks(propertyName);
                        if (converted[propertyName] && converted[propertyName].id > -1) {
                            converted.__setInternalLinks(propertyName, converted[propertyName].id);
                        }
                    }
                    else if (fieldDesc.fieldType === FieldType.LookupMulti) {
                        converted.__deleteInternalLinks(propertyName);
                        if (converted[propertyName] && converted[propertyName].length > 0) {
                            converted.__setInternalLinks(propertyName, converted[propertyName].filter(l => l.id !== -1).map(l => l.id));
                        }
                    }
                }
            }
        }
    }
    /***************** SP Calls associated to service standard operations ********************/


    /**
     * Get items by query
     * @protected
     * @param {IQuery} query - query used to retrieve items
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected async get_Internal(query: IQuery, linkedFields?: Array<string>): Promise<Array<T>> { 
        const restQuery = this.getRestQuery(query);            
        let results = new Array<T>();
        const items = await this.executeRequest(`${this.serviceUrl}${this.Bindings.get.url}`, this.Bindings.get.method, restQuery);
        if (items && items.length > 0) {
            await this.Init();
            results = items.map((r) => {
                return this.getItemFromRest(r);
            });            
        }
        await this.populateLookups(results, linkedFields);
        return results;
    }

    /**
     * Get an item by id
     * @param {number} id - item id
     */
    protected async getItemById_Internal(id: number, linkedFields?: Array<string>): Promise<T> {
        let result = null;
        const temp = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getItemById.url}/${id}`, this.Bindings.getItemById.method);
        if (temp) {
            await this.Init();
            result = this.getItemFromRest(temp);
            await this.populateLookups([result], linkedFields);
        }
        return result;
    }


    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    protected async getItemsById_Internal(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<T>> {
        const result: Array<T> = [];
        const promises: Promise<Array<T>>[] = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 2000);
            promises.push(this.get_Internal({
                test: {
                    type: "predicate",
                    operator: TestOperator.In,
                    propertyName: "id",
                    value: sub
                },
                limit: 2000
            }));
        }
        const res = await UtilsService.runPromisesInStacks(promises, 3);
        for (const tmp of res) {
            result.push(...tmp.filter(i => { return i !== null && i !== undefined; }));
        }
        await this.populateLookups(result, linkedFields);
        return result;
    }

    /**
     * Retrieve all items
     * 
     */
    protected async getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>> {
        let results: Array<T> = [];
        const items = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getAll.url}`, this.Bindings.getAll.method);
        if (items && items.length > 0) {
            await this.Init();
            results = items.map((r) => {
                return this.getItemFromRest(r);
            });
        }
        await this.populateLookups(results, linkedFields);
        return results;
    }

    public async addOrUpdateItem(item: T, loadLookups?: Array<string>): Promise<T> {
        this.updateInternalLinks(item, loadLookups);
        return super.addOrUpdateItem(item);
    }

    public async addOrUpdateItems(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void, loadLookups?: Array<string>): Promise<Array<T>> {
        items.forEach(item => this.updateInternalLinks(item, loadLookups));
        return super.addOrUpdateItems(items, onItemUpdated);
    }

    /**
     * Add or update an item
     * @param item - SPItem derived object to be converted
     */
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        const result = cloneDeep(item);
        if (item.id < 0) {
            const converted = await this.getRestItem(item);            
            const addResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
            await this.populateCommonFields(result, addResult.data);
            if (item.id < -1) {
                await this.updateLinksInDb(Number(item.id), Number(result.id));
            }         
        }
        else {
            // check version (cannot update if newer)
            if (item.version) {                
                const existing = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getItemById.url}/${item.id}`, this.Bindings.getItemById.method);
                if (parseFloat(existing[Constants.commonFields.version]) > item.version) {
                    const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    const converted = await this.getRestItem(item);
                    const updateResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);                 
                    await this.populateCommonFields(result, updateResult); 
                }               
            }
            else {
                const converted = await this.getRestItem(item);
                const updateResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);                                   
                await this.populateCommonFields(result, updateResult);
                if (item.id < -1) {
                    await this.updateLinksInDb(Number(item.id), Number(result.id));
                }
            }
        }
        return result;
    }

    /*
    * TODO
    */
    protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        const result = cloneDeep(items);
        const itemsToAdd = result.filter((item) => {
            return item.id < 0;
        });
        const versionedItems = result.filter((item) => {
            return item.version !== undefined && item.version !== null && item.id > 0;
        });
        const updatedItems = result.filter((item) => {
            return (item.version === undefined || item.version === null) && item.id > 0;
        });

        // creation batch
        if (itemsToAdd.length > 0) {
            let idx = 0;
            // TODO call stacks
            while (itemsToAdd.length > 0) {
                const sub = itemsToAdd.splice(0, 100);
                const converted = await Promise.all(sub.map(item => this.getRestItem(item)));
                try {
                    const addResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItems.url}`, this.Bindings.addOrUpdateItems.method, converted);
                    for (let index = 0; index < sub.length; index++) {
                        const item = sub[index];
                        const currentIdx = idx;
                        const itemId = item.id;
                        await this.populateCommonFields(item, addResult[index]);
                        if (itemId < -1) {
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
                        if (parseFloat(existing[Constants.commonFields.version]) > subitem.version) {
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
                catch(error) {
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
                    const converted = await Promise.all(sub.map(item => this.getRestItem(item)));
                    const results = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItems.url}`, this.Bindings.addOrUpdateItems.method, converted);
                    for (let index = 0; index < sub.length; index++) {
                        const subitem = sub[index];
                        const currentIdx = idx;
                        await this.populateCommonFields(subitem, results[index]);
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
    protected async deleteItem_Internal(item: T): Promise<void> {
        return this.executeRequest(`${this.serviceUrl}${this.Bindings.deleteItem.url}/${item.id}`, this.Bindings.deleteItem.method);
    }

    /************************** Query filters ***************************/
 
    protected async populateCommonFields(item, restItem): Promise<void> {
        if (item.id < 0) {
            // update id
            item.id = restItem.Id;
        }
        if (restItem[Constants.commonFields.version]) {
            item.version = parseFloat(restItem[Constants.commonFields.version]);
        }
        const fields = this.ItemFields;
        await Promise.all(Object.keys(fields).filter((propertyName) => {
            if (fields.hasOwnProperty(propertyName)) {
                const fieldName = fields[propertyName].fieldName;
                return (fieldName === Constants.commonFields.author ||
                    fieldName === Constants.commonFields.created ||
                    fieldName === Constants.commonFields.editor ||
                    fieldName === Constants.commonFields.modified);
            }
        }).map(async (prop) => {
            const fieldName = fields[prop].fieldName;
            switch (fields[prop].fieldType) {
                case FieldType.Date:
                    item[prop] = new Date(restItem[fieldName]);
                    break;
                case FieldType.User:
                    const id = restItem[fieldName];
                    let user = null;
                    if (this.initialized) {
                        const users = this.getServiceInitValues(User["name"]);
                        user = find(users, (u) => { return u.id === id; });
                    }
                    else {
                        const userService: UserService = new UserService();
                        user = await userService.getItemById(id);
                    }
                    item[prop] = user;
                    break;
                default:
                    item[prop] = restItem[fieldName];
                    break;
            }
        }));

    }



    /**
     * convert full item to db format (with links only)
     * @param item - full provisionned item
     */
    protected async convertItemToDbFormat(item: T): Promise<T> {
        const converted = item as unknown as BaseItem;
        const result: T = cloneDeep(item);
        const convertedResult = result as unknown as BaseItem;
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                switch (fieldDescriptor.fieldType) {
                    case FieldType.User:
                    case FieldType.Taxonomy:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            if (converted[propertyName]) {
                                convertedResult.__setInternalLinks(propertyName, converted[propertyName].id);
                            }
                            delete convertedResult[propertyName];
                        }
                        break;
                    case FieldType.UserMulti:
                    case FieldType.TaxonomyMulti:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            const ids = [];
                            if (converted[propertyName]) {
                                converted[propertyName].forEach(element => {
                                    if (element.id) {
                                        if ((typeof (element.id) === "number" && element.id > 0) || (typeof (element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
                                            ids.push(element.id);
                                        }
                                    }
                                });
                            }
                            convertedResult.__setInternalLinks(propertyName, ids.length > 0 ? ids : []);
                            delete convertedResult[propertyName];
                        }
                        break;
                    case FieldType.Lookup:
                    case FieldType.LookupMulti:
                        // internal links allready updated before (used for rest calls)
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            delete convertedResult[propertyName];
                            convertedResult.__setInternalLinks(propertyName, converted.__getInternalLinks(propertyName));
                        }
                        break;
                    default:                        
                        break;
                }

            }
        }
        return result;
    }

    /**
     * populate item from db storage
     * @param item - db item with links in internalLinks fields
     */
    public async mapItems(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        const results: Array<T> = [];
        if (items && items.length > 0) {
            await this.Init();
            for (const item of items) {
                const converted = item as unknown as BaseItem;
                const result: T = cloneDeep(item);
                const convertedResult = result as unknown as BaseItem;
                if (item) {
                    for (const propertyName in this.ItemFields) {
                        if (this.ItemFields.hasOwnProperty(propertyName)) {
                            const fieldDescriptor = this.ItemFields[propertyName];
                            if (//fieldDescriptor.fieldType === FieldType.Lookup ||
                                fieldDescriptor.fieldType === FieldType.User ||
                                fieldDescriptor.fieldType === FieldType.Taxonomy) {
                                if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                    // get values from init values
                                    const id: number = converted.__getInternalLinks(propertyName);
                                    if (id !== null) {
                                        const destElements = this.getServiceInitValues(fieldDescriptor.modelName);
                                        const existing = find(destElements, (destElement) => {
                                            return destElement.id === id;
                                        });
                                        convertedResult[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                                    }
                                    else {
                                        convertedResult[propertyName] = fieldDescriptor.defaultValue;
                                    }
                                }
                                convertedResult.__deleteInternalLinks(propertyName);
                            }
                            else if (//fieldDescriptor.fieldType === FieldType.LookupMulti ||
                                fieldDescriptor.fieldType === FieldType.UserMulti ||
                                fieldDescriptor.fieldType === FieldType.TaxonomyMulti) {
                                if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                    // get values from init values
                                    const ids = converted.__getInternalLinks(propertyName) || [];
                                    if (ids.length > 0) {
                                        const val = [];
                                        const targetItems = this.getServiceInitValues(fieldDescriptor.modelName);
                                        ids.forEach(id => {
                                            const existing = find(targetItems, (tmpitem) => {
                                                return tmpitem.id === id;
                                            });
                                            if (existing) {
                                                val.push(existing);
                                            }
                                        });
                                        convertedResult[propertyName] = val;
                                    }
                                    else {
                                        convertedResult[propertyName] = fieldDescriptor.defaultValue;
                                    }
                                }
                                convertedResult.__deleteInternalLinks(propertyName);
                            }
                            else {
                                convertedResult[propertyName] = converted[propertyName];
                            }
                        }
                    }
                }
                convertedResult.__clearEmptyInternalLinks();
                results.push(result);
            }
        }
        await this.populateLookups(results, linkedFields);
        return results;
    }

    public async updateLinkedTransactions(oldId: number, newId: number, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        // Update items pointing to this in transactions
        nextTransactions.forEach(transaction => {
            let currentObject = null;
            let needUpdate = false;
            const service = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
            const fields = service.ItemFields;
            // search for lookup fields
            for (const propertyName in fields) {
                if (fields.hasOwnProperty(propertyName)) {
                    const fieldDescription: IFieldDescriptor = fields[propertyName];
                    if (fieldDescription.refItemName === this.itemType["name"] || fieldDescription.modelName === this.itemType["name"]) {
                        // get object if not done yet
                        if (!currentObject) {
                            const destType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                            currentObject = new destType();
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
                this.transactionService.addOrUpdateItem(transaction);
            }
        });
        return nextTransactions;
    }

    private async updateLinksInDb(oldId: number, newId: number): Promise<void> {
        const allFields = assign({}, this.itemType["Fields"]);
        delete allFields[RestItem["name"]];
        delete allFields[this.itemType["name"]];
        for (const modelName in allFields) {
            if (allFields.hasOwnProperty(modelName)) {
                const modelFields = allFields[modelName];
                const lookupProperties = Object.keys(modelFields).filter((prop) => {
                    return (modelFields[prop].refItemName &&
                        modelFields[prop].refItemName === this.itemType["name"] || modelFields[prop].modelName === this.itemType["name"]);
                });
                if (lookupProperties.length > 0) {
                    const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                    const allitems = await service.__getAllFromCache();
                    const updated = [];
                    allitems.forEach(element => {
                        const converted = element as unknown as BaseItem;
                        let needUpdate = false;
                        lookupProperties.forEach(propertyName => {
                            const fieldDescription = modelFields[propertyName];
                            if (fieldDescription.fieldType === FieldType.Lookup) {
                                if (fieldDescription.modelName) {
                                    // serch in internalLinks
                                    const link = converted.__getInternalLinks(propertyName);
                                    if (link && link === oldId) {
                                        converted.__setInternalLinks(propertyName, newId);
                                        needUpdate = true;
                                    }
                                }
                                else if (converted[propertyName] === oldId) {
                                    // change field
                                    converted[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                            else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                if (fieldDescription.modelName) {
                                    // search in internalLinks
                                    const links = converted.__getInternalLinks(propertyName);
                                    if (links && isArray(links)) {
                                        // find item
                                        const lookupidx = findIndex(links, (id) => { return id === oldId; });
                                        // change id
                                        if (lookupidx > -1) {
                                            converted.__setInternalLinks(propertyName, newId);
                                            needUpdate = true;
                                        }
                                    }
                                }
                                else if (converted[propertyName] && isArray(converted[propertyName])) {
                                    // find index
                                    const lookupidx = findIndex(converted[propertyName], (id) => { return id === oldId; });
                                    // change field
                                    // change id
                                    if (lookupidx > -1) {
                                        converted[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                            }
                        });
                        if (needUpdate) {
                            updated.push(converted);
                        }
                    });
                    if (updated.length > 0) {
                        await service.__updateCache(...updated);
                    }
                }
            }
        }
    }

    private getRestQuery(query: IQuery): IRestQuery {
        const result: IRestQuery = {
            lastId: query.lastId as number,
            limit: query.limit,
            orderBy: query.orderBy
        };
        if(query.test) {
            if(query.test.type === "sequence") {
                result.test = this.getRestSequence(query.test);
            }
            else {
                result.test = {
                    predicates: [this.getRestPredicate(query.test)]
                };
            }
        }
        return result;
    }

    private getRestSequence(sequence: ILogicalSequence): IRestLogicalSequence {
        const result: IRestLogicalSequence = {
            logicalOperator: sequence.operator,
            predicates: [],
            sequences: []
        };
        sequence.children.forEach((child) => {
            if(child.type === "predicate") {                
                result.predicates.push(this.getRestPredicate(child));
            }
            else {
                const seq = this.getRestSequence(child);
                result.sequences.push(seq);
            }
        });
        return result;
    }
    private getRestPredicate(predicate: IPredicate): IRestPredicate {
        return {
            logicalOperator: predicate.operator,
            propertyName: predicate.propertyName,
            value: predicate.value,
            includeTimeValue: predicate.value,
            lookupId: predicate.lookupId
        };
    }

    private async initRequest(method: string, data?: any): Promise<RequestInit>{  
        const aadTokenProvider = await ServicesConfiguration.context.aadTokenProviderFactory.getTokenProvider();
        const token = await aadTokenProvider.getToken(ServicesConfiguration.configuration.aadAppId);
        if(stringIsNullOrEmpty(token)) {
            throw Error("Error while getting authentication token");
        }
        const headers = {
            'Accept': 'application/json', 'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': "*", 'Access-Control-Allow-Headers': "*",
            'authorization': `Bearer ${token}`
          };
          if (data != null) {
            const postData: string = JSON.stringify(data);
            return { method: method, body: postData, mode: 'cors', headers: headers };
          }
          return { method: method, mode: 'cors', headers: headers };
    }   

    private async executeRequest(url: string, method: string, data?: any): Promise<any> {
        const req = await this.initRequest(method, data);
        const response = await fetch(url, req);
            if(response.ok) {                    
                return response.json();                    
            }
            else {
                const error = new Error();
                error.message = "Error while executing request";
                error.name = response.status.toString();
                error.stack = await response.text();
                throw error;
            }
    }
}
