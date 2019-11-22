import { ServicesConfiguration } from "../..";
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep, find, assign } from "@microsoft/sp-lodash-subset";
import { CamlQuery, List, sp } from "@pnp/sp";
import { Constants, FieldType } from "../../constants/index";
import { IBaseItem, IFieldDescriptor } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { BaseService } from "./BaseService";
import { UtilsService } from "..";
import { SPItem, User, TaxonomyTerm, OfflineTransaction } from "../../models";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";

/**
 * 
 * Base service for sp list items operations
 */
export class BaseListItemService<T extends IBaseItem> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/
    protected listRelativeUrl: string;
    protected initValues: any = {};
    protected tardiveLinks: any = {};

    public get ItemFields(): any {
        let result = {}
        assign(result, this.itemType["Fields"][SPItem["name"]]);
        if(this.itemType["Fields"][this.itemType["name"]]) {
            assign(result, this.itemType["Fields"][this.itemType["name"]]);
        }
        return result;
    }

    /**
     * Associeted list (pnpjs)
     */
    protected get list(): List {
        return sp.web.getList(this.listRelativeUrl);
    }

    /***************************** Constructor **************************************/
    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param listRelativeUrl list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, tableName: string, cacheDuration?: number) {
        super(type, tableName, cacheDuration);
        this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;

    }

    
    /***************************** External sources init and access **************************************/
    
    private initialized: boolean = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }
    private initPromise: Promise<void> = null;

    protected async init_internal(): Promise<void>{};

    public async Init(): Promise<void> {
        if(!this.initPromise) {
            this.initPromise =  new Promise<void>(async (resolve, reject) => {
                if(this.initialized) {
                    resolve();
                }
                else {
                    try {
                        if(this.init_internal) {
                            await this.init_internal();
                        }
                        let fields = this.ItemFields;
                        let services = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                if(fieldDescription.serviceName && services.indexOf(fieldDescription.serviceName) === -1) {
                                    services.push(fieldDescription.serviceName);
                                }                
                            }
                        }
                        await Promise.all(services.map(async (serviceName) => {
                            if(!this.initValues[serviceName]) {
                                let service = ServicesConfiguration.configuration.serviceFactory.create(serviceName);
                                let values = await service.getAll();
                                this.initValues[serviceName] = values;
                            }
                        }));
                        this.initialized = true;
                        this.initPromise = null;
                        resolve();
                    }
                    catch(error) {
                        this.initPromise = null;
                        reject(error);
                    }
                }                
            });
        }
        return this.initPromise;
        
    }  

    private getServiceInitValues(serviceName: string): any {
        return this.initValues[serviceName];        
    }

    /****************************** get item methods ***********************************/
    private getItemFromRest(spitem: any): T {
        let item = new this.itemType();
        Object.keys(this.ItemFields).map((propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            item[propertyName] = this.getFieldValue(spitem, propertyName, fieldDescription);
        });
        return item;
    }
    // TODO : Tardive link
    private getFieldValue(spitem: T, propertyName: string,  fieldDescriptor:IFieldDescriptor): any {
        let value = fieldDescriptor.defaultValue;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch(fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if(fieldDescriptor.fieldName === "OData__UIVersionString") {
                    value = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                }
                else {
                    value = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                }                
                break;                
            case FieldType.Date:
                    value = spitem[fieldDescriptor.fieldName] ? new Date(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                let lookupId = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(lookupId !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.serviceName)) {
                        //link defered
                        spitem.__internalLinks[propertyName] = lookupId;
                    }
                    else {
                        value = lookupId;
                    }   
                }
                break;
            case FieldType.LookupMulti:
                    let lookupIds = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : [];
                    if(lookupIds.length > 0) {
                        if(!stringIsNullOrEmpty(fieldDescriptor.serviceName)) {                            
                            spitem.__internalLinks[propertyName] = lookupIds;
                        }
                        else {
                            value = lookupIds;
                        }
                    }
                    break;
            case FieldType.User:
                let id = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(id !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.serviceName)) {                         
                        spitem.__internalLinks[propertyName] = id;
                    }
                    else {
                        value = id;
                    }   
                }
                break;
            case FieldType.UserMulti:
                let ids = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : [];
                if(ids.length > 0) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.serviceName)) {                         
                        spitem.__internalLinks[propertyName] = ids;
                    }
                    else {
                        value = ids;
                    }
                }
                break;
            case FieldType.Taxonomy:
                let wssid = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if(id !== -1) {
                    let terms = this.getServiceInitValues(fieldDescriptor.serviceName);
                    value = this.getTaxonomyTermByWssId(wssid, terms);
                }
                break;
            case FieldType.TaxonomyMulti:
                    const terms = spitem[fieldDescriptor.fieldName];
                    if(terms) {
                        let allterms = this.getServiceInitValues(fieldDescriptor.serviceName);
                        value = terms.map((term) => {
                            return term.getTaxonomyTermByWssId(term.WssId, allterms);
                        });
                    }
                break;
            case FieldType.Json:
                    value = spitem[fieldDescriptor.fieldName] ? JSON.parse(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
        return value;
    }
    /****************************** Send item methods ***********************************/
    private async getSPRestItem(item: T): Promise<any> {
        let spitem = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            if(propertyName != "Version") {
                 let value = await this.convertFieldValueToRest(item[propertyName], fieldDescription);
                 assign(spitem[fieldDescription.fieldName], value);
            }
        }));
        return spitem;
    }
    private async convertFieldValueToRest(itemValue: any, fieldDescriptor:IFieldDescriptor): Promise<any> {
        let value = {};
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch(fieldDescriptor.fieldType) {
            case FieldType.Simple:
            case FieldType.Date:
                value[fieldDescriptor.fieldName] = itemValue;
                break;
            case FieldType.Lookup:
                if(itemValue) {
                    if(typeof(itemValue) === "number") {
                        value[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                    }
                    else {
                        value[fieldDescriptor.fieldName + "Id"] = itemValue.id > 0 ? itemValue.id : null;
                    }
                }
                else {
                    value[fieldDescriptor.fieldName + "Id"] = null;
                }
            case FieldType.LookupMulti:      
                if(itemValue && isArray(itemValue) && itemValue.length > 0){
                    let firstLookupVal = itemValue[0];
                    if(typeof(firstLookupVal) === "number") {
                        value[fieldDescriptor.fieldName + "Id"] = itemValue;
                    }
                    else {
                        value[fieldDescriptor.fieldName + "Id"] = itemValue.map((lookupMultiElt) => {return lookupMultiElt.id; });
                    }
                }      
                else {
                    value[fieldDescriptor.fieldName + "Id"] = null;
                }
                break;
            case FieldType.User:
                    if(itemValue) {
                        if(typeof(itemValue) === "number") {
                            value[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                        }
                        else {
                            value[fieldDescriptor.fieldName + "Id"] = await this.convertSingleUserFieldValue(itemValue);
                        }
                    }
                    else {
                        value[fieldDescriptor.fieldName + "Id"] = null;
                    }                
                break;
            case FieldType.UserMulti:
                if(itemValue && isArray(itemValue) && itemValue.length > 0) {
                    let firstUserVal = itemValue[0];
                    if(typeof(firstUserVal) === "number") {
                        value[fieldDescriptor.fieldName + "Id"] = itemValue;
                    }
                    else {
                        value[fieldDescriptor.fieldName + "Id"] = await Promise.all(itemValue.map((user) => {
                            return this.convertSingleUserFieldValue(user);
                        }));
                    }
                }
                else {
                    value[fieldDescriptor.fieldName + "Id"] = null;
                }
                break;
            case FieldType.Taxonomy:
                value[fieldDescriptor.fieldName] = this.convertTaxonomyFieldValue(itemValue);
                break;
            case FieldType.TaxonomyMulti:
                if(itemValue && isArray(itemValue) && itemValue.length > 0) {
                    value[fieldDescriptor.fieldName] = itemValue.map((term) => {
                        return this.convertTaxonomyFieldValue(term);
                    });
                }
                else {
                    value[fieldDescriptor.fieldName] = null;
                }
                break;
            case FieldType.Json:
                    value[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                break;
        }
        return value;
    }

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

    private async convertSingleUserFieldValue(value: User): Promise<any> {
        let result: any = null;
        if (value) {
            if(!value.spId || value.spId <=0) {
                let userService:UserService = new UserService();
                value = await userService.linkToSpUser(value);

            }
            result = value.spId;
        }
        return result;
    }

    /**
     * 
     * @param wssid 
     * @param terms 
     */
    public getTaxonomyTermByWssId<T extends TaxonomyTerm>(wssid: number, terms: Array<T>): T {
        return find(terms, (term) => {
            return (term.wssids && term.wssids.indexOf(wssid) > -1);
        });
    }

    /******************************************** DB conversion ***************************************************/
    private getDbItem(item: T): any {

    }

    /******************************************* Cache Management *************************************************/

    /**
     * Cache has to be reloaded ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    protected async  needRefreshCache(key: string = "all"): Promise<boolean> {
        let result: boolean = await super.needRefreshCache(key);

        if (!result) {

            let isconnected = await UtilsService.CheckOnline();
            if (isconnected) {

                let cachedDataDate = await super.getCachedData(key);
                if (cachedDataDate) {

                    try {
                        let response = await ServicesConfiguration.context.spHttpClient.get(`${ServicesConfiguration.context.pageContext.web.absoluteUrl}/_api/web/getList('${this.listRelativeUrl}')`,
                            SPHttpClient.configurations.v1,
                            {
                                headers: {
                                    'Accept': 'application/json;odata.metadata=minimal',
                                    'Cache-Control': 'no-cache'
                                }
                            });

                        let tempList = await response.json();
                        let lastModifiedDate = new Date(tempList.LastItemUserModifiedDate ? tempList.LastItemUserModifiedDate : tempList.d.LastItemUserModifiedDate);
                        result = lastModifiedDate > cachedDataDate;


                    } catch (error) {
                        console.error(error);
                    }


                }
            }
        }

        return result;
    }
    /***************** SP Calls associated to service standard operations ********************/

    /**
     * Get items by query
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected async get_Internal(query: any): Promise<Array<T>> {
        let results = new Array<T>();
        let selectFields = this.getOdataFieldNames();
        let items = await this.list.select(...selectFields).getItemsByCAMLQuery({
            ViewXml: `<View Scope="RecursiveAll"><Query>${query}</Query></View>`
        } as CamlQuery);
        if(items && items.length > 0) {
            await this.Init();
            results = items.map((r) => { 
                return this.getItemFromRest(r); 
            });
        }
        return results;
    }




    /**
     * Get an item by id
     * @param id item id
     */
    protected async getItemById_Internal(id: number): Promise<T> {
        let result = null;
        let selectFields = this.getOdataFieldNames();
        let temp = await this.list.items.getById(id).select(...selectFields).get();
        if (temp) {
            await this.Init();
            result = this.getItemFromRest(temp);
            return result;
        }

        return result;
    }

    /**
     * Get a list of items by id
     * @param id item id
     */
    protected async getItemsById_Internal(ids: Array<number>): Promise<Array<T>> {
        let results: Array<T> = [];
        let selectFields = this.getOdataFieldNames();
        let batch = sp.createBatch();
        ids.forEach((id) => {
            this.list.items.getById(id).select(...selectFields).inBatch(batch).get().then((item)=> {
                results.push(this.getItemFromRest(item));
            })
        });
        await batch.execute();
        return results;   
    }

    /**
     * Retrieve all items
     * 
     */
    protected async getAll_Internal(): Promise<Array<T>> {
        let results: Array<T> = [];
        let selectFields = this.getOdataFieldNames();
        let items = await this.list.items.select(...selectFields).getAll();
        if(items && items.length > 0) {
            await this.Init();
            results = items.map((r) => { 
                return this.getItemFromRest(r); 
            });
        }
        return results;
    }

    /**
     * Add or update an item
     * @param item SPItem derived object to be converted
     */
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        let result = cloneDeep(item);
        if (item.id < 0) {
            let converted = await this.getSPRestItem(item);
            let addResult = await this.list.items.add(converted);
            if(addResult.data["OData__UIVersionString"]) {
                result.version = parseFloat(addResult.data["OData__UIVersionString"]);
            }
            // TODO: update lookups + new wssids + users spid
        }
        else {
            // check version (cannot update if newer)
            if (item.version) {
                let existing = await this.list.items.getById(<number>item.id).select("OData__UIVersionString").get();
                if (parseFloat(existing["OData__UIVersionString"]) > item.version) {
                    let error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    let converted = await this.getSPRestItem(item);
                    let updateResult = await this.list.items.getById(<number>item.id).update(converted);
                    let version = await updateResult.item.select("OData__UIVersionString").get();
                    if(version["OData__UIVersionString"]) {
                        result.version = parseFloat(version["OData__UIVersionString"]);
                    }
                    // TODO: new wssids + users spid
                }
            }
            else {
                let converted = await this.getSPRestItem(item);
                let updateResult = await this.list.items.getById(<number>item.id).update(converted);
                let version = await updateResult.item.select("OData__UIVersionString").get();
                if(version["OData__UIVersionString"]) {
                    result.version = parseFloat(version["OData__UIVersionString"]);
                }
                // TODO: new wssids + users spid
            }
        }
        return result;
    }

    /**
     * Delete an item
     * @param item SPItem derived class to be deletes
     */
    protected async deleteItem_Internal(item: T): Promise<void> {
        await this.list.items.getById(<number>item.id).delete();
    }

    /************************** Query filters ***************************/


    /**
     * Retrive all fields to include in odata setect parameter
     */
    private getOdataFieldNames(): Array<string> {
        let fields = this.ItemFields;
        let fieldNames = Object.keys(fields).filter((propertyName) => { 
            return fields.hasOwnProperty(propertyName); 
        }).map((prop) => {
            let result: string = fields[prop].fieldName;
            switch(fields[prop].fieldType) {
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                case FieldType.User:
                case FieldType.UserMulti:
                    result += "Id";
                default:
                    break;
            }
            return result;
        });
        return fieldNames;
    }

    protected convertItemToDbFormat(item: T): T {
        // TODO: store object with minimal value
        let result: T = new this.itemType();
        for (const key in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(key)) {
                const fieldDescriptor = this.ItemFields[key];
                //switch
                
            }
        }
        return item;
    }

    public mapItem(item: T): T {
        // TODO: get lazy loading values
        return item;
    }
    
    public async updateLinkedTransactions(oldId: number, newId: number, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        // TODO: update items pointing to this + user ids on transactions
        return nextTransactions;
    }
    
    private async updateLinksInDb(item: T): Promise<void>{
        // TODO: update items pointing to this (create only + id < -1)
        
    }

    
    private async updateWssIdsAndUsersSpIds(item: T): Promise<void> {
        //TODO: if taxonomy field, store wssid in db (add or update) --> service + taxohidden
    }
}
