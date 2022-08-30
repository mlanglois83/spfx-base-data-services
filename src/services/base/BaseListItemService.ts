import { cloneDeep, find, findIndex } from "lodash";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";
import { sp } from "@pnp/sp";
import "@pnp/sp/content-types/list";
import "@pnp/sp/fields/list";
import { IItemAddResult } from '@pnp/sp/items';
import "@pnp/sp/items/list";
import "@pnp/sp/lists";
import { ICamlQuery, IList } from "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import { Semaphore } from "async-mutex";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { Constants, FieldType, LogicalOperator, QueryToken, TestOperator, TraceLevel } from "../../constants/index";
import { Decorators } from "../../decorators";
import { IFieldDescriptor, ILogicalSequence, IOrderBy, IPredicate, IQuery } from "../../interfaces/index";
import { BaseItem, SPFile, SPItem, TaxonomyTerm, User } from "../../models";
import { UserService } from "../graph/UserService";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { BaseDataService } from "./BaseDataService";
import { BaseDbService } from "./BaseDbService";



const trace = Decorators.trace;

/**
 * 
 * Base service for sp list items operations
 */
export class BaseListItemService<T extends SPItem> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/
    protected listRelativeUrl: string;
    protected taxoMultiFieldNames: { [fieldName: string]: string } = {};
    protected checkLastModify = true;

    /* AttachmentService */
    protected attachmentsService: BaseDbService<SPFile>;


    /**
     * Associeted list (pnpjs)
     */
    protected get list(): IList {
        return sp.web.getList(this.listRelativeUrl);
    }
    /***************************** Constructor **************************************/
    /**
     * 
     * @param type - items type
     * @param listRelativeUrl - list web relative url
     * @param tableName - name of table in local db
     * @param cacheDuration - cache duration in minutes
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, cacheDuration?: number, checkLastModify?: boolean) {
        super(type, cacheDuration);
        this.listRelativeUrl = ServicesConfiguration.serverRelativeUrl + listRelativeUrl;
        if (this.hasAttachments) {
            this.attachmentsService = new BaseDbService<SPFile>(SPFile, "ListAttachments");
        }
        if (checkLastModify !== undefined) {
            this.checkLastModify = checkLastModify;
        }
    }

    /********** init for taxo multi ************/
    private fieldsInitialized = false;
    private initFieldsPromise: Promise<void> = null;
    @trace(TraceLevel.ServiceUtilities)
    private async initFields(): Promise<void> {
        if (!this.initFieldsPromise) {
            this.initFieldsPromise = new Promise<void>(async (resolve, reject) => {
                if (this.fieldsInitialized) {
                    resolve();
                }
                else {
                    this.taxoMultiFieldNames = {};
                    try {
                        const fields = this.ItemFields;
                        const taxofields = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                if (fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                                    if (stringIsNullOrEmpty(fieldDescription.hiddenFieldName)) {
                                        taxofields.push(fieldDescription.fieldName);
                                    }
                                    else {
                                        this.taxoMultiFieldNames[fieldDescription.fieldName] = fieldDescription.hiddenFieldName;
                                    }
                                }
                            }
                        }
                        await Promise.all(taxofields.map(async (tf) => {
                            const hiddenField = await this.list.fields.getByTitle(`${tf}_0`).select("InternalName").get();
                            this.taxoMultiFieldNames[tf] = hiddenField.InternalName;
                        }));
                        this.fieldsInitialized = true;
                        this.initFieldsPromise = null;
                        resolve();
                    }
                    catch (error) {
                        this.initFieldsPromise = null;
                        reject(error);
                    }
                }
            });
        }
        return this.initFieldsPromise;

    }

    /****************************** get item methods ***********************************/

    protected populateFieldValue(spitem: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
        super.populateFieldValue(spitem, destItem, propertyName, fieldDescriptor);        
        const defaultValue = cloneDeep(fieldDescriptor.defaultValue);
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if (fieldDescriptor.fieldName === Constants.commonFields.version) {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : defaultValue;
                }
                else if (fieldDescriptor.fieldName === Constants.commonFields.attachments) {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].map((fileobj) => { return new SPFile(fileobj); }) : defaultValue;
                }
                else if(fieldDescriptor.fieldName.indexOf("/") !== -1) {
                    const splitteed = fieldDescriptor.fieldName.split("/");
                    let current = spitem;
                    splitteed.forEach(s => {
                        current = current[s];
                    });
                    if(current) {
                        destItem[propertyName] = current;
                    }
                    else {
                        destItem[propertyName] = defaultValue;
                    }
                }
                break;
            case FieldType.Lookup:                
                if (fieldDescriptor.containsFullObject && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                    const obj = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : null;
                    if (obj && typeof (obj[Constants.commonFields.id]) === "number") {
                        // object allready persisted before, retrieve id and store like classical lookup
                        destItem.__setInternalLinks(propertyName, obj[Constants.commonRestFields.id]);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = defaultValue;
                    }                       
                }
                else {
                    const lookupId: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                    if (lookupId !== -1) {
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
                }
                break;
            case FieldType.LookupMulti:
                if (fieldDescriptor.containsFullObject && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                    const lookupIds: Array<number> = spitem[fieldDescriptor.fieldName] && Array.isArray(spitem[fieldDescriptor.fieldName]) ?
                    spitem[fieldDescriptor.fieldName].map(ri => ri[Constants.commonFields.id]).filter(objid => typeof (objid) === "number") :
                    [];
                    if (lookupIds.length > 0) {
                        // LOOKUPS --> links
                        destItem.__setInternalLinks(propertyName, lookupIds);
                        destItem[propertyName] = defaultValue;
                    }
                    else {
                        destItem[propertyName] = defaultValue;
                    }           
                }
                else {
                    const lookupIds: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results : spitem[fieldDescriptor.fieldName + "Id"]) : [];
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
                }
                break;
            case FieldType.User:
                const id: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if (id !== -1) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const users = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                        const existing = find(users, (user) => {
                            return user.id === id;
                        });
                        destItem[propertyName] = existing ? existing : defaultValue;
                    }
                    else {
                        destItem[propertyName] = id;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.UserMulti:
                const ids: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results : spitem[fieldDescriptor.fieldName + "Id"]) : [];
                if (ids.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const val = [];
                        const users = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                        ids.forEach(umid => {
                            const existing = find(users, (user) => {
                                return user.id === umid;
                            });
                            if (existing) {
                                val.push(existing);
                            }
                        });
                        destItem[propertyName] = val;
                    }
                    else {
                        destItem[propertyName] = ids;
                    }
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                const wssid: number = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if (wssid !== -1) {
                    const tterms = this.getServiceInitValuesByName<TaxonomyTerm>(fieldDescriptor.modelName);
                    destItem[propertyName] = this.getTaxonomyTermByWssId(wssid, tterms);
                }
                else {
                    destItem[propertyName] = defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                const tmterms = spitem[fieldDescriptor.fieldName] ? (spitem[fieldDescriptor.fieldName].results ? spitem[fieldDescriptor.fieldName].results : spitem[fieldDescriptor.fieldName]) : [];
                if (tmterms.length > 0) {
                    const allterms = this.getServiceInitValuesByName<TaxonomyTerm>(fieldDescriptor.modelName);
                    destItem[propertyName] = tmterms.map((term) => {
                        return this.getTaxonomyTermByWssId(term.WssId, allterms);
                    });
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
            Constants.commonFields.created,
            Constants.commonFields.author,
            Constants.commonFields.editor,
            Constants.commonFields.modified,
            Constants.commonFields.version
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
                            destItem[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                        }
                        else {
                            destItem[fieldDescriptor.fieldName + "Id"] = link && link > 0 ? link : null;
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName + "Id"] = null;
                    }
                    break;
                case FieldType.LookupMulti:
                    if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                        const links = item.__getInternalLinks(propertyName);
                        const firstLookupVal = itemValue[0];
                        if (typeof (firstLookupVal) === "number") {
                            destItem[fieldDescriptor.fieldName + "Id"] = { results: itemValue };
                        }
                        else {
                            if (links && links.length > 0) {
                                destItem[fieldDescriptor.fieldName + "Id"] = { results: links };
                            }
                            else {
                                destItem[fieldDescriptor.fieldName + "Id"] = { results: [] };
                            }
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName + "Id"] = { results: [] };
                    }
                    break;
                case FieldType.User:
                    if (itemValue) {
                        if (typeof (itemValue) === "number") {
                            destItem[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                        }
                        else {
                            destItem[fieldDescriptor.fieldName + "Id"] = await this.convertSingleUserFieldValue(itemValue);
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName + "Id"] = null;
                    }
                    break;
                case FieldType.UserMulti:
                    if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                        const firstUserVal = itemValue[0];
                        if (typeof (firstUserVal) === "number") {
                            destItem[fieldDescriptor.fieldName + "Id"] = { results: itemValue };
                        }
                        else {
                            const userIds = await Promise.all(itemValue.map((user) => {
                                return this.convertSingleUserFieldValue(user);
                            }));
                            destItem[fieldDescriptor.fieldName + "Id"] = { results: userIds };
                        }
                    }
                    else {
                        destItem[fieldDescriptor.fieldName + "Id"] = { results: [] };
                    }
                    break;
                case FieldType.Taxonomy:
                    destItem[fieldDescriptor.fieldName] = this.convertTaxonomyFieldValue(itemValue);
                    break;
                case FieldType.TaxonomyMulti:
                    const hiddenFieldName = this.taxoMultiFieldNames[fieldDescriptor.fieldName];
                    if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                        destItem[hiddenFieldName] = this.convertTaxonomyMultiFieldValue(itemValue);
                    }
                    else {
                        destItem[hiddenFieldName] = null;
                    }
                    break;
                default: break;
            }
        }
    }

    /****************************** Lookup loading **************************************/

    /********************** SP Fields conversion helpers *****************************/
    private convertTaxonomyFieldValue(value: TaxonomyTerm): any {
        let result: any = null;
        if (value) {
            result = {
                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                Label: value.title,
                TermGuid: value.id,
                WssId: -1 // fake
            };
        }
        return result;
    }
    private convertTaxonomyMultiFieldValue(value: Array<TaxonomyTerm>): string {
        let result: string = null;
        if (value) {
            result = value.map(term => `-1;#${term.title}|${term.id};#`).join("");
        }
        return result;
    }

    private async convertSingleUserFieldValue(value: User): Promise<User | number> {
        let result: User | number = null;
        if (value) {
            if (value.id <= 0) {
                const userService: UserService = ServiceFactory.getService(User).cast<UserService>();
                value = await userService.linkToSpUser(value);

            }
            result = value.id;
        }
        return result;
    }

    /**
     * 
     * @param wssid - wssid of term to retrieve
     * @param terms - terms list where term must be found
     */
    public getTaxonomyTermByWssId<TermType extends TaxonomyTerm>(wssid: number, terms: Array<TermType>): TermType {
        return find(terms, (term) => {
            return (term.wssids && term.wssids.indexOf(wssid) > -1);
        });
    }



    /******************************************* Cache Management *************************************************/


    /*******************************  store list last modified date***********************/
    private lastModifiedDate = "lastResultClassLifeTime";



    /*******************************  store last check from list last modified date***********************/
    private lastModifiedDateCheck = "lastResultClassLifeTimeCheck";

    protected set lastModifiedListCheck(newValue: Date) {
        const cacheKey = this.getCacheKey(this.lastModifiedDateCheck);
        window.sessionStorage.setItem(cacheKey, JSON.stringify(newValue));
    }

    protected get lastModifiedListCheck(): Date {

        const cacheKey = this.getCacheKey(this.lastModifiedDateCheck);

        const lastDataLoadString = window.sessionStorage.getItem(cacheKey);
        let lastDataLoad: Date = null;

        if (lastDataLoadString) {
            lastDataLoad = new Date(JSON.parse(window.sessionStorage.getItem(cacheKey)));
        }

        return lastDataLoad;
    }

    //perf issue with await
    // protected async needRefreshCache(key = "all"): Promise<boolean> {
    //     //get parent need refresh information
    //     let result: boolean = super.needRefreshCache(key);
    //     if (this.checkLastModify) {
    //         //if not need refresh cache, test, last modified list modified
    //         if (!result) {

    //             //check online
    //             const isconnected = ServicesConfiguration.configuration.lastConnectionCheckResult

    //             if (isconnected) {

    //                 //get last cache date
    //                 const cachedDataDate = super.getCachedData(key);
    //                 //if a date existe, check if renew necessary
    //                 //else load data
    //                 if (cachedDataDate) {

    //                     const lastModifiedDate = await this.LastModfiedList();

    //                     result = lastModifiedDate > cachedDataDate;
    //                 }
    //             }
    //         }
    //     }
    //     return result;
    // }


    /**
     * Cache has to be reloaded ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    protected async LastModfiedList(): Promise<Date> {

        //avoid fetchnig multiple same request as same time
        let promise = this.getExistingPromise(this.lastModifiedDate);
        if (promise) {
            console.log(this.serviceName + " needRefreshCache : load allready called before, sharing promise");
        }
        else {

            const semaphore = new Semaphore(1);

            const semacq = await semaphore.acquire();

            try {

                promise = new Promise<Date>(async (resolve, reject) => {
                    try {

                        //get last modified date store in cache, if exists
                        const cacheKey = this.getCacheKey(this.lastModifiedDate);

                        const lastDataLoadString = window.sessionStorage.getItem(cacheKey);
                        let lastModifiedSave: Date = null;

                        if (lastDataLoadString) {
                            lastModifiedSave = new Date(JSON.parse(window.sessionStorage.getItem(cacheKey)));
                        }


                        //to avoid send x request during 20 seconds
                        //get date when the last modified lsite date was checked
                        const temp = this.lastModifiedListCheck;
                        if (temp) {
                            //add 20 seconds, cache duration
                            temp.setSeconds(this.lastModifiedListCheck.getSeconds() + 20);
                        }

                        //if not previous result or last check is more than 20 seconds.
                        if (!lastModifiedSave || (!temp || (temp < new Date()))) {
                            try {
                                let tempList: any = undefined;
                                /*if(ServicesConfiguration.context) {
                                //send request
                                    const response = await ServicesConfiguration.context.spHttpClient.get(`${ServicesConfiguration.baseUrl}/_api/web/getList('${this.listRelativeUrl}')`,
                                    SPHttpClient.configurations.v1,
                                    {
                                        headers: {
                                            'Accept': 'application/json;odata.metadata=minimal',
                                            'Cache-Control': 'no-cache'
                                        }
                                    });
                                    //get response 
                                    tempList = await response.json();
                                }
                                else {*/
                                    const init: RequestInit = {
                                        headers: {
                                            'Accept': 'application/json;odata.metadata=minimal',
                                            'Cache-Control': 'no-cache'
                                        },
                                        credentials: 'same-origin',
                                        method: "GET"
                                    };
                                    const response = await fetch(`${ServicesConfiguration.baseUrl}/_api/web/getList('${this.listRelativeUrl}')`, init);
                                    tempList = await response.json();
                                //}

                                //store date when last modified date list is checked
                                this.lastModifiedListCheck = new Date();

                                lastModifiedSave = new Date(tempList.LastItemUserModifiedDate ? tempList.LastItemUserModifiedDate : tempList.d.LastItemUserModifiedDate);
                                //store last modified date list
                                window.sessionStorage.setItem(cacheKey, JSON.stringify(lastModifiedSave));

                            } catch (error) {
                                console.error(error);
                            }
                        }
                        resolve(lastModifiedSave);
                    } catch (error) {
                        reject(error);
                    }
                });


                this.storePromise(promise, this.lastModifiedDate);
            } finally {
                semacq[1](); // release
            }

        }



        return promise;
    }
    /**
     * Retrieve id of items to be reloaded
     * @param ids - id if items to check
     */
    protected async getExpiredIds(...ids: Array<number>): Promise<Array<number>> {
        let result: Array<number> = await super.getExpiredIds(...ids) as number[];
        if (this.checkLastModify) {
            if (result.length < ids.length) {

                const isconnected = await UtilsService.CheckOnline();
                if (isconnected) {

                    try {

                        const lastModifiedDate = await this.LastModfiedList();

                        result = [];
                        ids.forEach((id) => {
                            const lastLoad = this.getIdLastLoad(id);
                            if (!lastLoad || lastLoad < lastModifiedDate) {
                                result.push(id);
                            }
                        });


                    } catch (error) {
                        console.error(error);
                    }


                }
            }
        }

        return result;
    }

    /**********************************Service specific calls  *******************************/


    /***************** SP Calls associated to service standard operations ********************/


    protected async get_Query(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<any>> {
        const spQuery = this.getCamlQuery(query);
        const selectFields = this.getOdataFieldNames(linkedFields);
        const expandFields = this.getOdataExpandFieldNames(linkedFields);
        const itemsQuery = this.list.select(...selectFields).expand(...expandFields);
        return itemsQuery.getItemsByCAMLQuery(spQuery, ...expandFields);
    }

    /**
     * Get an item by id
     * @param {number} id - item id
     */
    @trace(TraceLevel.Queries)
    protected async getItemById_Query(id: number, linkedFields?: Array<string>): Promise<any> {
        const selectFields = this.getOdataFieldNames(linkedFields);
        const expandFields = this.getOdataExpandFieldNames(linkedFields);
        const itemsQuery = this.list.items.getById(id).select(...selectFields).expand(...expandFields);
        return itemsQuery.get();
    }


    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    @trace(TraceLevel.Queries)
    protected async getItemsById_Query(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<any>> {
        const result: Array<any> = [];
        const promises: Promise<Array<any>>[] = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 2000);
            promises.push(this.get_Query({
                test: {
                    type: "predicate",
                    operator: TestOperator.In,
                    propertyName: "id",
                    value: sub
                },
                limit: 2000
            }, linkedFields));
        }
        const res = await UtilsService.runPromisesInStacks(promises, 3);
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
    protected async getAll_Query(linkedFields?: Array<string>): Promise<Array<any>> {
        const selectFields = this.getOdataFieldNames(linkedFields);        
        const expandFields = this.getOdataExpandFieldNames(linkedFields);
        const itemsQuery = this.list.items.select(...selectFields).expand(...expandFields);
        return itemsQuery.getAll();
    }

    /**
     * Add or update an item
     * @param item - SPItem derived object to be converted
     */
    @trace(TraceLevel.Internal)
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        const result = cloneDeep(item);
        await this.initFields();
        const selectFields = this.getOdataCommonFieldNames();
        if (item.id < 0) {
            const converted = await this.convertItem(item);
            const addResult = await this.list.items.select(...selectFields).add(converted);
            await this.populateCommonFields(result, addResult.data);
            await this.updateWssIds(result, addResult.data);
            if (item.id < -1) {
                await this.updateLinksInDb(Number(item.id), Number(result.id));
            }
        }
        else {
            // check version (cannot update if newer)
            if (item.version) {
                const existing = await this.list.items.getById(item.id).select(Constants.commonFields.version).get();
                if (parseFloat(existing[Constants.commonFields.version]) > item.version) {
                    const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    const converted = await this.convertItem(item);
                    const updateResult = await this.list.items.getById(item.id).select(...selectFields).update(converted);
                    const version = await updateResult.item.select(...selectFields).get();
                    await this.populateCommonFields(result, version);
                    await this.updateWssIds(result, version);
                }
            }
            else {
                const converted = await this.convertItem(item);
                const updateResult = await this.list.items.getById(item.id).update(converted);
                const version = await updateResult.item.select(...selectFields).get();
                await this.populateCommonFields(result, version);
                await this.updateWssIds(result, version);
            }
        }
        return result;
    }

    @trace(TraceLevel.Internal)
    protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void, onItemRefreshed?: (index: number, length: number) => void): Promise<Array<T>> {
        const result: Array<T> = cloneDeep(items);
    
        const itemsToAdd = result.filter((item) => {
            return item.id < 0;
        });
        const versionedItems = result.filter((item) => {
            return item.version !== undefined && item.version !== null && item.id > 0;
        });
        const updatedItems = result.filter((item) => {
            return (item.version === undefined || item.version === null) && item.id > 0;
        });

        await this.initFields();
        const entityTypeFullName = await this.list.getListItemEntityTypeFullName();
        const selectFields = this.getOdataCommonFieldNames();
        // creation batch
        if (itemsToAdd.length > 0) {
            let idx = 0;
            const batches = [];
            while (itemsToAdd.length > 0) {
                const sub = itemsToAdd.splice(0, 100);
                const batch = sp.createBatch();
                for (const item of sub) {
                    const currentIdx = idx;
                    const itemId = item.id;
                    const converted = await this.convertItem(item);
                    this.list.items.select(...selectFields).inBatch(batch).add(converted, entityTypeFullName).then(async (addResult: IItemAddResult) => {
                        await this.populateCommonFields(item, addResult.data);
                        await this.updateWssIds(item, addResult.data);
                        if (itemId < -1) {
                            await this.updateLinksInDb(Number(itemId), Number(item.id));
                        }
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                    }).catch((error) => {
                        item.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                    });
                    idx++;
                }
                batches.push(batch);
            }
            await UtilsService.runBatchesInStacks(batches, 3);
        }
        // versionned batch --> check conflicts
        if (versionedItems.length > 0) {
            let idx = 0;
            const batches = [];
            while (versionedItems.length > 0) {
                const sub = versionedItems.splice(0, 100);
                const batch = sp.createBatch();
                for (const item of sub) {
                    const currentIdx = idx;
                    this.list.items.getById(item.id).select(Constants.commonFields.version).inBatch(batch).get().then(async (existing) => {
                        if (parseFloat(existing[Constants.commonFields.version]) > item.version) {
                            const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                            error.name = Constants.Errors.ItemVersionConfict;
                            item.error = error;
                            if (onItemUpdated) {
                                onItemUpdated(items[currentIdx], item);
                            }
                        }
                        else {
                            updatedItems.push(item);
                        }
                    }).catch((error) => {
                        item.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                    });
                    idx++;
                }
                batches.push(batch);
            }
            await UtilsService.runBatchesInStacks(batches, 3);
        }
        // 
        const resultItems: Array<T> = [];
        // classical update batch + version checked
        if (updatedItems.length > 0) {
            let idx = 0;
            const batches = [];
            while (updatedItems.length > 0) {
                const sub = updatedItems.splice(0, 100);
                const batch = sp.createBatch();
                for (const item of sub) {
                    const currentIdx = idx;
                    const converted = await this.convertItem(item);
                    this.list.items.getById(item.id).select(...selectFields).inBatch(batch).update(converted, '*', entityTypeFullName).then(async () => {                                            
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                        resultItems.push(item);
                    }).catch((error) => {
                        item.error = error;
                        if (onItemUpdated) {
                            onItemUpdated(items[currentIdx], item);
                        }
                    });
                    idx++;
                }
                batches.push(batch);
            }
            await UtilsService.runBatchesInStacks(batches, 3);
        }    
        // update properties
        const resultsLength = resultItems.length;
        if (resultItems.length > 0) {
            let idx = 0;
            const batches = [];
            while (resultItems.length > 0) {
                const sub = resultItems.splice(0, 100);
                const batch = sp.createBatch();
                for (const item of sub) {
                    const currentIdx = idx;
                    this.list.items.getById(item.id).select(...selectFields).inBatch(batch).get().then(async (version) => {
                        await this.populateCommonFields(item, version);
                        await this.updateWssIds(item, version);
                        if (onItemRefreshed) {
                            onItemRefreshed(currentIdx, resultsLength);
                        }
                    }).catch((error) => {
                        item.error = error;
                        if (onItemRefreshed) {
                            onItemRefreshed(currentIdx, resultsLength);
                        }
                    });
                    idx++;
                }
                batches.push(batch);
            }
            await UtilsService.runBatchesInStacks(batches, 3);
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
            await this.list.items.getById(item.id).recycle();
            item.deleted = true;
        }
        catch (error) {
            item.error = error;
        }
        return item;
    }

    @trace(TraceLevel.Internal)
    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        const batch = sp.createBatch();
        items.forEach(item => {
            this.list.items.getById(item.id).inBatch(batch).recycle().then(() => {
                item.deleted = true;
            }).catch((error) => {
                item.error = error;
            });
        });
        await batch.execute();
        return items;
    }


    @trace(TraceLevel.ServiceUtilities)
    private async getAttachmentContent(attachment: SPFile): Promise<void> {
        const content = await sp.web.getFileByServerRelativeUrl(attachment.serverRelativeUrl).getBuffer();
        attachment.content = content;
    }

    @trace(TraceLevel.Service)
    public async cacheAttachmentsContent(): Promise<void> {
        const prop = this.attachmentProperty;
        if (prop !== null) {
            let load = true;
            if (ServicesConfiguration.configuration.checkOnline) {
                load = await UtilsService.CheckOnline();
            }
            if (load) {
                const updatedItems: T[] = [];
                const operations: Promise<void>[] = [];
                const items = await this.dbService.getAll();
                for (const item of items) {                    
                    let mapped: Array<T>;
                    if(this.isMapItemsAsync()) {
                        mapped = await this.mapItemsAsync([item]);
                    }
                    else {
                        mapped = this.mapItemsSync([item]);
                    }
                    const converted = mapped.shift();
                    if (converted[prop] && converted[prop].length > 0) {
                        updatedItems.push(converted);
                        converted[prop].forEach(attachment => {
                            operations.push(this.getAttachmentContent(attachment));
                        });
                    }

                }
                operations.map(operation => {
                    return operation;
                }).reduce((chain, operation) => {
                    return chain.then(() => { return operation; });
                }, Promise.resolve()).then(async () => {

                    if (updatedItems.length > 0) {
                        const dbitems = updatedItems.map(u => this.convertItemToDbFormat(u));
                        await this.dbService.addOrUpdateItems(dbitems);
                    }
                });

            }
        }

    }
    /************************** Query filters ***************************/

    /**
     * Retrive all fields to include in odata setect parameter
     */
    private get hasAttachments(): boolean {
        return this.attachmentProperty !== null;
    }

    private get attachmentProperty(): string {
        let result: string = null;
        const fields = this.ItemFields;
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key];
                if (fieldDesc.fieldName === Constants.commonFields.attachments) {
                    result = key;
                    break;
                }
            }
        }
        return result;
    }

    /**
     * Retrive all fields to include in odata setect parameter
     */
    private getOdataFieldNames(linkedFields?: Array<string>): Array<string> {
        const fields = this.ItemFields;
        const fieldNames = Object.keys(fields).filter((propertyName) => {
            return fields.hasOwnProperty(propertyName) &&
                (!linkedFields || (linkedFields.length === 1 && linkedFields[0] === 'loadAll') || linkedFields.indexOf(fields[propertyName].fieldName) !== -1);
        }).map((prop) => {
            let result: string = fields[prop].fieldName;
            switch (fields[prop].fieldType) {
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                case FieldType.User:
                case FieldType.UserMulti:
                    result += "Id";
                    break;
                default:
                    break;
            }
            return result;
        });
        return fieldNames;
    }

    /**
     * Retrive all fields to include in odata setect parameter
     */
     private getOdataExpandFieldNames(linkedFields?: Array<string>): Array<string> {
        const fields = this.ItemFields;
        const fieldNames = Object.keys(fields).filter((propertyName) => {
            return fields.hasOwnProperty(propertyName) &&
                (!linkedFields || (linkedFields.length === 1 && linkedFields[0] === 'loadAll') || linkedFields.indexOf(fields[propertyName].fieldName) !== -1);
        }).filter(propertyName => {
            const result: string = fields[propertyName].fieldName;
            return result === Constants.commonFields.attachments || result.indexOf("/") !== -1;
        }).map((prop) => {
            const result: string = fields[prop].fieldName;
            return result.split("/").shift();
        });
        return fieldNames;
    }

    
    private getOdataCommonFieldNames(): Array<string> {
        const fields = this.ItemFields;
        const fieldNames = [Constants.commonFields.version];
        Object.keys(fields).filter((propertyName) => {
            return fields.hasOwnProperty(propertyName);
        }).forEach((prop) => {
            const fieldName = fields[prop].fieldName;
            if (fieldName === Constants.commonFields.author ||
                fieldName === Constants.commonFields.created ||
                fieldName === Constants.commonFields.editor ||
                fieldName === Constants.commonFields.modified) {
                let result: string = fields[prop].fieldName;
                switch (fields[prop].fieldType) {
                    case FieldType.Lookup:
                    case FieldType.LookupMulti:
                    case FieldType.User:
                    case FieldType.UserMulti:
                        result += "Id";
                        break;
                    default:
                        break;
                }
                fieldNames.push(result);
            }
        });
        return fieldNames;
    }

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
                    const id = restItem[fieldName + "Id"];
                    let user = null;
                    if (this.initialized) {
                        const users = this.getServiceInitValues(User["name"]);
                        user = find(users, (u) => { return u.id === id; });
                    }
                    else {
                        const userService = ServiceFactory.getService(User);
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


    @trace(TraceLevel.ServiceUtilities)
    private async updateWssIds(item: T, spItem: any): Promise<void> {
        // if taxonomy field, store wssid in db (add or update) --> service + this.init
        const fields = this.ItemFields;
        // serch for Taxonomy fields
        for (const propertyName in fields) {
            if (fields.hasOwnProperty(propertyName)) {

                const fieldDescription: IFieldDescriptor = fields[propertyName];
                if (fieldDescription.fieldType === FieldType.Taxonomy) {
                    let needUpdate = false;
                    // get wssid from item
                    const wssid = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].WssId : -1;
                    if (wssid !== -1) {
                        const id = item[propertyName].id;
                        // find corresponding object in service
                        const service = ServiceFactory.getServiceByModelName(fieldDescription.modelName);
                        const term = await service.__getFromCache(id);
                        if (term instanceof TaxonomyTerm) {
                            term.wssids = term.wssids || [];
                            if (term.wssids.indexOf(wssid) === -1) {
                                term.wssids.push(wssid);
                                needUpdate = true;
                            }
                        }
                        if (needUpdate) {
                            await service.__updateCache(term);
                            // update initValues
                            if (this.initialized) {
                                const idx = findIndex(this.initValues[fieldDescription.modelName], (t: BaseItem) => { return t.id === id; });
                                if (idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = term;
                                }
                            }
                        }
                    }
                }
                else if (fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                    const updated = [];
                    const terms = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].results : [];
                    const service = ServiceFactory.getServiceByModelName(fieldDescription.modelName);
                    if (terms && terms.length > 0) {
                        await Promise.all(terms.map(async (termitem) => {
                            const wssid = termitem.WssId;
                            const id = termitem.TermGuid;
                            // find corresponding object in allready updated
                            let term = find(updated, (u) => { return u.id === id; });
                            if (!term) {
                                term = await service.__getFromCache(id);
                            }
                            if (term instanceof TaxonomyTerm) {
                                term.wssids = term.wssids || [];
                                if (term.wssids.indexOf(wssid) === -1) {
                                    term.wssids.push(wssid);
                                    if (!find(updated, (u) => { return u.id === id; })) {
                                        updated.push(term);
                                    }
                                }
                            }
                        }));
                    }
                    if (updated.length > 0) {
                        await service.__updateCache(...updated);
                        // update initValues
                        if (this.initialized) {
                            updated.forEach((u) => {
                                const idx = findIndex(this.initValues[fieldDescription.modelName], (t: BaseItem) => { return t.id === u.id; });
                                if (idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = u;
                                }
                            });
                        }
                    }
                }
            }
        }
    }

    @trace(TraceLevel.Service)
    public async refreshData(): Promise<void> {
        this.initialized = false;
        this.initValues = {};
        return super.refreshData();
    }


    private getCamlQuery(query: IQuery<T>): ICamlQuery {
        const result: ICamlQuery = {
            ViewXml: `<View Scope="RecursiveAll">
                <Query>
                    ${this.getWhere(query)}
                    ${this.getOrderBy(query)}
                </Query>
                ${query.limit !== undefined ? `<RowLimit>${query.limit}</RowLimit>` : ""}
            </View>`,
            DatesInUtc: true
        };
        if (query.lastId !== undefined) {
            result.ListItemCollectionPosition = {
                "PagingInfo": "Paged=TRUE&p_ID=" + query.lastId
            };
        }
        return result;
    }

    private getOrderBy(query: IQuery<T>): string {
        let result = "";
        if (query.orderBy && query.orderBy.length > 0) {
            result = `<OrderBy>
                ${query.orderBy.map(ob => this.getFieldRef(ob)).join('')}
            </OrderBy>`;
        }
        return result;
    }
    private getWhere(query: IQuery<T>): string {
        let result = "";
        if (query.test) {
            result = `<Where>
                ${query.test.type === "predicate" ? this.getPredicate(query.test) : this.getLogicalSequence(query.test)}
            </Where>`;
        }
        return result;
    }
    private getLogicalSequence(sequence: ILogicalSequence<T>): string {

        const cloneSequence = cloneDeep(sequence);

        if (!cloneSequence.children || cloneSequence.children.length === 0) {
            return "";
        }
        if (cloneSequence.children.length === 1) {
            if (cloneSequence.children[0].type === "predicate") {
                return this.getPredicate(cloneSequence.children[0]);
            }
            else {
                return this.getLogicalSequence(cloneSequence.children[0]);
            }
        }
        else {
            // first part
            let result = `<${cloneSequence.operator}>`;
            if (cloneSequence.children[0].type === "predicate") {
                result += this.getPredicate(cloneSequence.children[0]);
            }
            else {
                result += this.getLogicalSequence(cloneSequence.children[0]);
            }
            cloneSequence.children.splice(0, 1);
            result += this.getLogicalSequence(cloneSequence);
            result += `</${cloneSequence.operator}>`;
            return result;
        }
    }

    private getPredicate(predicate: IPredicate<T, keyof T>): string {
        let result = "";
        switch (predicate.operator) {
            case TestOperator.IsNotNull:
            case TestOperator.IsNull:
                result = `<${predicate.operator}>
                    ${this.getFieldRef(predicate)}
                </${predicate.operator}>`;
                break;
            case TestOperator.In:
                if (predicate.value && isArray(predicate.value) && predicate.value.length > 0) {
                    if (predicate.value.length <= 500) {
                        return `<${predicate.operator}>
                            ${this.getFieldRef(predicate)}
                            <Values>
                                ${predicate.value.map(v => this.getValue(predicate, v, predicate.lookupId)).join('')}
                            </Values>
                        </${predicate.operator}>`;
                    }
                    else {
                        const transformed: ILogicalSequence<T> = {
                            type: "sequence",
                            operator: LogicalOperator.Or,
                            children: []
                        };
                        const copy = predicate.value;
                        while (copy.length) {
                            const subValues = copy.splice(0, 500);
                            transformed.children.push({
                                type: "predicate",
                                operator: TestOperator.In,
                                propertyName: predicate.propertyName,
                                value: subValues,
                                includeTimeValue: predicate.includeTimeValue,
                                lookupId: predicate.lookupId
                            });
                        }
                        result = this.getLogicalSequence(transformed);
                    }
                }
                else {
                    result = `<${predicate.operator}>
                        ${this.getFieldRef(predicate)}
                        <Values>
                            ${this.getValue(predicate, -1, predicate.lookupId)}
                        </Values>
                    </${predicate.operator}>`;
                }

                break;
            default:
                result = `<${predicate.operator}>
                    ${this.getFieldRef(predicate)}
                    ${this.getValue(predicate, predicate.value, predicate.lookupId)}
                </${predicate.operator}>`;
                break;
        }
        return result;
    }
    private getFieldRef(obj: IPredicate<T, keyof T> | IOrderBy<T, keyof T>): string {
        let result = "";
        const fields = this.ItemFields;
        const field = fields[obj.propertyName.toString()];
        if (field) {
            result = `<FieldRef Name="${field.fieldName}"${obj.type === "predicate" && obj.lookupId ? " LookupId=\"TRUE\"" : ""}${obj.type === "orderby" && obj.ascending !== undefined && !obj.ascending ? " Ascending=\"FALSE\"" : ""} />`;
        }
        else {
            throw new Error("Field was not found : " + obj.propertyName);
        }
        return result;
    }
    private getValue(obj: IPredicate<T, keyof T>, fieldValue: any, lookupID?: boolean): string {
        let result = "";
        const fields = this.ItemFields;
        const field = fields[obj.propertyName.toString()];
        if (field) {
            let type = "";
            let value = "";
            switch (field.fieldType) {
                case FieldType.Simple:
                    if (typeof (fieldValue) === "number") {
                        type = "Number";
                        value = fieldValue.toString();
                    }
                    else if (typeof (fieldValue) === "boolean") {
                        type = "Boolean";
                        value = fieldValue ? "1" : "0";
                    }
                    else {
                        type = "Text";
                        value = fieldValue.toString();
                    }
                    break;
                case FieldType.Date:
                    type = "DateTime";
                    if (fieldValue === QueryToken.Now || fieldValue === QueryToken.Today) {
                        value = `<${fieldValue}/>`;
                    }
                    else {
                        value = fieldValue.toISOString().replace(/\.\d+Z/g, "Z");
                    }
                    break;
                case FieldType.Json:
                    type = "Text";
                    value = JSON.stringify(fieldValue);
                    break;
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                    type = lookupID ? "Integer" : "Lookup";
                    value = fieldValue.toString();
                    break;
                case FieldType.Taxonomy:
                case FieldType.TaxonomyMulti:
                    type = lookupID ? "Integer" : "Text";
                    value = fieldValue.toString();
                    break;
                case FieldType.User:
                case FieldType.UserMulti:
                    type = fieldValue === QueryToken.UserID ? "Integer" : "User";
                    if (fieldValue === QueryToken.UserID) {
                        value = `<${fieldValue}/>`;
                    }
                    else {
                        value = fieldValue.toString();
                    }
                    break;
                default:
                    if (typeof (fieldValue) === "number") {
                        type = "Number";
                        value = fieldValue.toString();
                    }
                    else if (typeof (fieldValue) === "boolean") {
                        type = "Boolean";
                        value = fieldValue ? "1" : "0";
                    }
                    else {
                        type = "Text";
                        value = fieldValue.toString();
                    }
                    break;
            }
            result = `<Value Type="${type}" ${(type === "DateTime" && obj.includeTimeValue !== undefined ? (" IncludeTimeValue=\"" + (obj.includeTimeValue ? "TRUE" : "FALSE") + "\"") : "") + (type === "DateTime" ? " StorageTZ=\"TRUE\"" : "")}>${value}</Value>`;
        }
        else {
            throw new Error("Field was not found : " + obj.propertyName);
        }
        return result;
    }
}
