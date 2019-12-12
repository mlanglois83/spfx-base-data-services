import { ServicesConfiguration } from "../..";
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep, find, assign, findIndex, update } from "@microsoft/sp-lodash-subset";
import { CamlQuery, List, sp } from "@pnp/sp";
import { Constants, FieldType } from "../../constants/index";
import { IBaseItem, IFieldDescriptor } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { BaseService } from "./BaseService";
import { UtilsService } from "..";
import { SPItem, User, TaxonomyTerm, OfflineTransaction } from "../../models";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";
import { BaseTermsetService } from "./BaseTermsetService";

/**
 * 
 * Base service for sp list items operations
 */
export class BaseListItemService<T extends IBaseItem> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/
    protected listRelativeUrl: string;
    protected initValues: any = {};
    protected taxoMultiFieldNames: any = {};

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
                    this.initValues = {};
                    try {
                        if(this.init_internal) {
                            await this.init_internal();
                        }
                        let fields = this.ItemFields;
                        let models = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                if(fieldDescription.modelName && models.indexOf(fieldDescription.modelName) === -1) {
                                    models.push(fieldDescription.modelName);
                                }                                            
                            }
                        }
                        await Promise.all(models.map(async (modelName) => {
                            if(!this.initValues[modelName]) {
                                let service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                let values = await service.getAll();
                                this.initValues[modelName] = values;
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
    /********** init for taxo multi ************/
    private fieldsInitialized: boolean = false;
    private initFieldsPromise: Promise<void> = null;
    private async initFields(): Promise<void> {
        if(!this.initFieldsPromise) {
            this.initFieldsPromise =  new Promise<void>(async (resolve, reject) => {
                if(this.fieldsInitialized) {
                    resolve();
                }
                else {
                    this.taxoMultiFieldNames = {};
                    try {
                        let fields = this.ItemFields;
                        let taxofields = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                if(fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                                    taxofields.push(fieldDescription.fieldName);
                                }                                      
                            }
                        }
                        await Promise.all(taxofields.map(async (tf) => {
                            let hiddenField = await this.list.fields.getByTitle(`${tf}_0`).select("InternalName").get();
                            this.taxoMultiFieldNames[tf] = hiddenField.InternalName;
                        }));
                        this.fieldsInitialized = true;
                        this.initFieldsPromise = null;
                        resolve();
                    }
                    catch(error) {
                        this.initFieldsPromise = null;
                        reject(error);
                    }
                }                
            });
        }
        return this.initFieldsPromise;
        
    }  

    private getServiceInitValues(modelName: string): any {
        return this.initValues[modelName];        
    }

    /****************************** get item methods ***********************************/
    private getItemFromRest(spitem: any): T {
        let item = new this.itemType();
        Object.keys(this.ItemFields).map((propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            this.setFieldValue(spitem, item, propertyName, fieldDescription);
        });
        return item;
    }

    private setFieldValue(spitem: any, destItem: T, propertyName: string,  fieldDescriptor: IFieldDescriptor): void {
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch(fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if(fieldDescriptor.fieldName === Constants.commonFields.version) {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                }
                else {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                }                
                break;                
            case FieldType.Date:
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? new Date(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                let lookupId: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(lookupId !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        let destElements = this.getServiceInitValues(fieldDescriptor.modelName);                        
                        let existing = find(destElements, (destElement) => {
                            return destElement.id === lookupId;
                        });
                        destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                        
                    }
                    else {
                        destItem[propertyName] = lookupId;
                    } 
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                  
                break;
            case FieldType.LookupMulti:
                    let lookupIds: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : [];
                    if(lookupIds.length > 0) {
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {    
                            // get values from init values
                            let val = [];
                            let targetItems = this.getServiceInitValues(fieldDescriptor.modelName);
                            lookupIds.forEach(id => {
                                let existing = find(targetItems, (item) => {
                                    return item.id === id;
                                });
                                if(existing) {
                                    val.push(existing);
                                } 
                            });
                            destItem[propertyName] = val;
                        }
                        else {
                            destItem[propertyName] = lookupIds;
                        }
                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                    break;
            case FieldType.User:
                let id: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(id !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {                         
                        // get values from init values
                        let users = this.getServiceInitValues(fieldDescriptor.modelName);                        
                        let existing = find(users, (user) => {
                            return user.spId === id;
                        });
                        destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        destItem[propertyName] = id;
                    } 
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }                      
                break;
            case FieldType.UserMulti:
                let ids: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : [];                
                if(ids.length > 0) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {    
                        // get values from init values
                        let val = [];
                        let users = this.getServiceInitValues(fieldDescriptor.modelName);
                        ids.forEach(id => {
                            let existing = find(users, (user) => {
                                return user.spId === id;
                            });
                            if(existing) {
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
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                let wssid: number = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if(id !== -1) {
                    let terms = this.getServiceInitValues(fieldDescriptor.modelName);
                    destItem[propertyName] = this.getTaxonomyTermByWssId(wssid, terms);
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                    const terms = spitem[fieldDescriptor.fieldName];
                    if(terms && terms.results) {
                        let allterms = this.getServiceInitValues(fieldDescriptor.modelName);
                        destItem[propertyName] = terms.results.map((term) => {
                            return this.getTaxonomyTermByWssId(term.WssId, allterms);
                        });
                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                break;
            case FieldType.Json:
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? JSON.parse(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
    }
    /****************************** Send item methods ***********************************/
    private async getSPRestItem(item: T): Promise<any> {
        let spitem = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            await this.setRestFieldValue(item, spitem, propertyName, fieldDescription);
        }));
        return spitem;
    }
    private async setRestFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor:IFieldDescriptor): Promise<void> {
        let itemValue = item[propertyName];
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch(fieldDescriptor.fieldType) {
            case FieldType.Simple:
            case FieldType.Date:
                    if(fieldDescriptor.fieldName !== Constants.commonFields.author && 
                        fieldDescriptor.fieldName !== Constants.commonFields.created && 
                        fieldDescriptor.fieldName !== Constants.commonFields.editor &&
                        fieldDescriptor.fieldName !== Constants.commonFields.modified &&
                        fieldDescriptor.fieldName !== Constants.commonFields.version) {
                            
                        destItem[fieldDescriptor.fieldName] = itemValue;
                    }
                break;
            case FieldType.Lookup:                
                if(itemValue) {
                    if(typeof(itemValue) === "number") {
                        destItem[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                    }
                    else {
                        destItem[fieldDescriptor.fieldName + "Id"] = itemValue.id > 0 ? itemValue.id : null;
                    }
                }
                else {
                    destItem[fieldDescriptor.fieldName + "Id"] = null;
                }
                break;
            case FieldType.LookupMulti:      
                if(itemValue && isArray(itemValue) && itemValue.length > 0){
                    let firstLookupVal = itemValue[0];
                    if(typeof(firstLookupVal) === "number") {
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: itemValue};
                    }
                    else {
                        let idArray = 
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: itemValue.map((lookupMultiElt) => {return lookupMultiElt.id; })};
                    }
                }      
                else {
                    destItem[fieldDescriptor.fieldName + "Id"] = {results: []};
                }
                break;
            case FieldType.User:
                    if(itemValue) {
                        if(typeof(itemValue) === "number") {
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
                if(itemValue && isArray(itemValue) && itemValue.length > 0) {
                    let firstUserVal = itemValue[0];
                    if(typeof(firstUserVal) === "number") {
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: itemValue};
                    }
                    else {
                        let userIds = await Promise.all(itemValue.map((user) => {
                            return this.convertSingleUserFieldValue(user);
                        }));
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: userIds};
                    }
                }
                else {
                    destItem[fieldDescriptor.fieldName + "Id"] = {results: []};
                }
                break;
            case FieldType.Taxonomy:
                destItem[fieldDescriptor.fieldName] = this.convertTaxonomyFieldValue(itemValue);
                break;
            case FieldType.TaxonomyMulti:
                let hiddenFieldName = this.taxoMultiFieldNames[fieldDescriptor.fieldName];
                if(itemValue && isArray(itemValue) && itemValue.length > 0) {
                    destItem[hiddenFieldName] = this.convertTaxonomyMultiFieldValue(itemValue);
                }
                else {
                    destItem[hiddenFieldName] = null;
                }
                break;
            case FieldType.Json:
                    destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                break;
        }
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
    private convertTaxonomyMultiFieldValue(value: Array<TaxonomyTerm>): any {
        let result: any = null;
        if (value) {
            result = value.map(term => `-1;#${term.title}|${term.id};#`).join("");            
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

    /**********************************Service specific calls  *******************************/
    
    /**
     * Get items by caml query
     * @param query caml query (<Where></Where>)
     * @param orderBy array of <FieldRef Name='Field1' Ascending='TRUE'/>
     * @param limit  number of lines
     * @param lastId last id for paged queries
     */
    public getByCamlQuery(query: string, orderBy?: string[], limit?: number, lastId?: number): Promise<Array<T>> {
        let queryXml = this.getQuery(query, orderBy,limit);
        let camlQuery = {
            ViewXml: queryXml
        } as CamlQuery;
        if(lastId !== undefined) {
            camlQuery.ListItemCollectionPosition = {
                "PagingInfo": "Paged=TRUE&p_ID=" + lastId
            }
        }
        return this.get(camlQuery);
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
        let items = await this.list.select(...selectFields).getItemsByCAMLQuery(query as CamlQuery);
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
        // TODO: created + modified + users
        let result = cloneDeep(item);
        await this.initFields();
        let selectFields = this.getOdataCommonFieldNames();
        if (item.id < 0) {
            let converted = await this.getSPRestItem(item);
            let addResult = await this.list.items.select(...selectFields).add(converted);                    
            await this.populateCommonFields(result, addResult.data);                   
            await this.updateWssIds(result, addResult.data); 
            if(item.id < -1) {
                await this.updateLinksInDb(Number(item.id), Number(result.id));
            }
        }
        else {            
            // check version (cannot update if newer)
            if (item.version) {
                let existing = await this.list.items.getById(<number>item.id).select(Constants.commonFields.version).get();
                if (parseFloat(existing[Constants.commonFields.version]) > item.version) {
                    let error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    let converted = await this.getSPRestItem(item);
                    let updateResult = await this.list.items.getById(<number>item.id).select(...selectFields).update(converted);
                    let version = await updateResult.item.select(...selectFields).get();                    
                    await this.populateCommonFields(result, version);                    
                    await this.updateWssIds(result, version);
                }
            }
            else {
                let converted = await this.getSPRestItem(item);
                let updateResult = await this.list.items.getById(<number>item.id).update(converted);
                let version = await updateResult.item.select(...selectFields).get();
                await this.populateCommonFields(result, version);                
                await this.updateWssIds(result, version);
            }
        }
        return result;
    }


    /**
     * Delete an item
     * @param item SPItem derived class to be deleted
     */
    protected async deleteItem_Internal(item: T): Promise<void> {
        await this.list.items.getById(<number>item.id).recycle();
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

    private getOdataCommonFieldNames(): Array<string> {
        let fields = this.ItemFields;
        let fieldNames = [Constants.commonFields.version];
        Object.keys(fields).filter((propertyName) => { 
            return fields.hasOwnProperty(propertyName); 
        }).forEach((prop) => {
            let fieldName = fields[prop].fieldName;
            if(fieldName  === Constants.commonFields.author ||
                fieldName  === Constants.commonFields.created ||
                fieldName  === Constants.commonFields.editor ||
                fieldName  === Constants.commonFields.modified) {
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
                    fieldNames.push(result);
                }
        });
        return fieldNames;
    }

    protected async populateCommonFields(item, restItem): Promise<void> {
        if(item.id < 0) {
            // update id
            item.id = restItem.Id;
        }
        if(restItem[Constants.commonFields.version]) {
            item.version = parseFloat(restItem[Constants.commonFields.version]);
        }
        let fields = this.ItemFields;
        await Promise.all(Object.keys(fields).filter((propertyName) => {
            let result = false;
            if(fields.hasOwnProperty(propertyName)) {                
                let fieldName = fields[propertyName].fieldName;
                return (fieldName  === Constants.commonFields.author ||
                    fieldName  === Constants.commonFields.created ||
                    fieldName  === Constants.commonFields.editor ||
                    fieldName  === Constants.commonFields.modified);
            }
        }).map(async (prop) => {
            let fieldName = fields[prop].fieldName;            
            switch(fields[prop].fieldType) {
                case FieldType.Date:
                    item[prop] = new Date(restItem[fieldName]);
                    break;
                case FieldType.User:
                    let id = restItem[fieldName + "Id"];
                    let user = null;
                    if(this.initialized) {
                        let users = this.getServiceInitValues(User["name"]);
                        user = find(users, (u) => { return u.spId === id; });
                    }
                    else {
                        let userService: UserService = new UserService();
                        user = await userService.getBySpId(id);
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
     * @param item full provisionned item
     */
    protected convertItemToDbFormat(item: T): T {
        let result: T = cloneDeep(item);
        delete result.__internalLinks;
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                switch(fieldDescriptor.fieldType) {
                    case FieldType.Lookup:  
                    case FieldType.User:           
                    case FieldType.Taxonomy:       
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            result.__internalLinks = result.__internalLinks || {};
                            result.__internalLinks[propertyName] = item[propertyName] ? item[propertyName].id : undefined;
                            delete result[propertyName];
                        }
                        break;
                    case FieldType.LookupMulti:
                    case FieldType.UserMulti:            
                    case FieldType.TaxonomyMulti:                      
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {  
                            let ids = [];
                            if(item[propertyName]) {
                                item[propertyName].forEach(element => {
                                    if(element.id) {
                                        if((typeof(element.id) === "number" && element.id > 0) || (typeof(element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
                                            ids.push(element.id);
                                        }
                                    }
                                });
                            }                          
                            result.__internalLinks = result.__internalLinks || {};
                            result.__internalLinks[propertyName] = ids.length > 0 ? ids : [];                            
                            delete result[propertyName];
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
     * @param item db item with links in __internalLinks fields
     */
    public async mapItem(item: T): Promise<T> {
        let result: T = cloneDeep(item);
        await this.Init();
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                switch(fieldDescriptor.fieldType) {
                    case FieldType.Lookup:
                    case FieldType.User:
                    case FieldType.Taxonomy:                    
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            // get values from init values
                            let id: number = item.__internalLinks[propertyName] ? item.__internalLinks[propertyName] : null;
                            if(id !== null) {
                                let destElements = this.getServiceInitValues(fieldDescriptor.modelName);                        
                                let existing = find(destElements, (destElement) => {
                                    return destElement.id === id;
                                });
                                result[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                            }
                            else {
                                result[propertyName] = fieldDescriptor.defaultValue;
                            }
                        }                    
                        break;
                    case FieldType.LookupMulti:  
                    case FieldType.UserMulti:
                    case FieldType.TaxonomyMulti:                      
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {    
                            // get values from init values
                            let ids = item.__internalLinks[propertyName] ? item.__internalLinks[propertyName] : [];
                            if(ids.length > 0) {
                                let val = [];
                                let targetItems = this.getServiceInitValues(fieldDescriptor.modelName);
                                ids.forEach(id => {
                                    let existing = find(targetItems, (item) => {
                                        return item.id === id;
                                    });
                                    if(existing) {
                                        val.push(existing);
                                    } 
                                });
                                result[propertyName] = val;
                            }
                            else {
                                result[propertyName] = fieldDescriptor.defaultValue;
                            }
                        }
                        break;                    
                    default:                        
                        result[propertyName] = item[propertyName] ;
                        break;                    
                }                
            }
        }
        delete result.__internalLinks;
        return result;
    }
    
    public async updateLinkedTransactions(oldId: number, newId: number, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        // Update items pointing to this in transactions
        nextTransactions.forEach(transaction => {
            let currentObject = null;
            let needUpdate: boolean = false;
            let service = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
            let fields = service.ItemFields;
            // search for lookup fields
            for (const propertyName in fields) {
                if (fields.hasOwnProperty(propertyName)) {
                    const fieldDescription: IFieldDescriptor = fields[propertyName];
                    if(fieldDescription.refItemName === this.itemType["name"]) {
                        // get object if not done yet
                        if(!currentObject) {
                            let destType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                            let currentObject = new destType();
                            assign(currentObject, transaction.itemData);
                        }                        
                        if(fieldDescription.fieldType === FieldType.Lookup) {
                            if(fieldDescription.modelName) {
                                // search in __internalLinks
                                if(currentObject.__internalLinks && currentObject.__internalLinks[propertyName] === oldId) {
                                    currentObject.__internalLinks[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                            else if(currentObject[propertyName] === oldId){
                                // change field
                                currentObject[propertyName] = newId;
                                needUpdate = true;
                            }
                        }
                        else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                            if(fieldDescription.modelName) {
                                // serch in __internalLinks
                                if(currentObject.__internalLinks && currentObject.__internalLinks[propertyName] && isArray(currentObject.__internalLinks[propertyName])) {
                                    // find item
                                    let lookupidx = findIndex(currentObject.__internalLinks[propertyName], (id) => {return id === oldId});
                                    // change id
                                    if(lookupidx > -1) {
                                        currentObject.__internalLinks[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                            }
                            else if(currentObject[propertyName] && isArray(currentObject[propertyName])){
                                // find index
                                let lookupidx = findIndex(currentObject[propertyName], (id) => {return id === oldId});
                                // change field
                                // change id
                                if(lookupidx > -1) {
                                    currentObject[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                        }

                    }
                    
                }
            }
            if(needUpdate) {
                transaction.itemData = assign({}, currentObject);
                this.transactionService.addOrUpdateItem(transaction); 
            }
        });
        return nextTransactions;
    }
    
    private async updateLinksInDb(oldId: number, newId: number): Promise<void>{
        let allFields = assign({}, this.itemType["Fields"]);
        delete allFields[SPItem["name"]];
        delete allFields[this.itemType["name"]];
        for (const modelName in allFields) {
            if (allFields.hasOwnProperty(modelName)) {     
                const modelFields =  allFields[modelName];    
                let lookupProperties = Object.keys(modelFields).filter((prop) => {
                    return modelFields[prop].refItemName &&
                        modelFields[prop].refItemName === this.itemType["name"];
                });
                if(lookupProperties.length > 0) {
                    let service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                    let allitems = await service.__getAllFromCache();
                    let updated = [];
                    allitems.forEach(element => { 
                        let needUpdate: boolean = false;                       
                        lookupProperties.forEach(propertyName => {
                            let fieldDescription = modelFields[propertyName];                
                            if(fieldDescription.fieldType === FieldType.Lookup) {
                                if(fieldDescription.modelName) {
                                    // serch in __internalLinks
                                    if(element.__internalLinks && element.__internalLinks[propertyName] === oldId) {
                                        element.__internalLinks[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                                else if(element[propertyName] === oldId){
                                    // change field
                                    element[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                            else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                if(fieldDescription.modelName) {
                                    // serch in __internalLinks
                                    if(element.__internalLinks && element.__internalLinks[propertyName] && isArray(element.__internalLinks[propertyName])) {
                                        // find item
                                        let lookupidx = findIndex(element.__internalLinks[propertyName], (id) => {return id === oldId});
                                        // change id
                                        if(lookupidx > -1) {
                                            element.__internalLinks[propertyName] = newId;
                                            needUpdate = true;
                                        }
                                    }
                                }
                                else if(element[propertyName] && isArray(element[propertyName])){
                                    // find index
                                    let lookupidx = findIndex(element[propertyName], (id) => {return id === oldId});
                                    // change field
                                    // change id
                                    if(lookupidx > -1) {
                                        element[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                            }
                        });
                        if(needUpdate) {
                            updated.push(element);
                        }
                    });
                    if(updated.length > 0) {
                        await service.__updateCache(...updated);
                    }
                }
            }
        }
        
    }

    
    private async updateWssIds(item: T, spItem: any): Promise<void> {
        // if taxonomy field, store wssid in db (add or update) --> service + this.init
        let fields = this.ItemFields;
        // serch for Taxonomy fields
        for (const propertyName in fields) {
            if (fields.hasOwnProperty(propertyName)) {     
                           
                const fieldDescription: IFieldDescriptor = fields[propertyName];
                if(fieldDescription.fieldType === FieldType.Taxonomy) {
                    let needUpdate = false;
                    // get wssid from item
                    let wssid = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].WssId : -1;
                    if(wssid !== -1) {
                        let id = item[propertyName].id;
                        // find corresponding object in service
                        let service = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription.modelName);
                        let term = await service.__getFromCache(id);
                        if(term instanceof TaxonomyTerm) {
                        term.wssids = term.wssids || [];
                            if(term.wssids.indexOf(wssid) === -1) {
                                term.wssids.push(wssid);
                                needUpdate = true;
                            }
                        }
                        if(needUpdate) {
                            await service.__updateCache(term);
                            // update initValues
                            if(this.initialized) {
                                let idx = findIndex(this.initValues[fieldDescription.modelName], (t: any) => { return t.id === id; });
                                if(idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = term;
                                }
                            }
                        }
                    }
                }
                else if (fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                    let updated = [];
                    let terms = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].results : [];     
                    let service = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription.modelName);              
                    if(terms && terms.length > 0) {
                        await Promise.all(terms.map(async (termitem) => {
                            let wssid = termitem.WssId;
                            let id = termitem.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1")
                            // find corresponding object in allready updated
                            let term = find(updated, (u) => {return u.id === id});
                            if(!term) {
                                term = await service.__getFromCache(id);
                            }
                            if(term instanceof TaxonomyTerm) {                           
                                term.wssids = term.wssids || [];
                                if(term.wssids.indexOf(wssid) === -1) {
                                    term.wssids.push(wssid);
                                    if(!find(updated, (u) => {return u.id === id})) {
                                        updated.push(term);
                                    }
                                }
                            }
                        }));
                    }
                    if(updated.length > 0) {
                        await service.__updateCache(...updated)
                        // update initValues
                        if(this.initialized) {
                            updated.forEach((u) => {
                                let idx = findIndex(this.initValues[fieldDescription.modelName], (t: any) => { return t.id === u.id; });
                                if(idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = u;
                                }
                            })      
                        }                  
                    }
                }
            }
        }
    }
    /**
     * 
     * @param query caml query (<Where></Where>)
     * @param orderBy array of <FieldRef Name='Field1' Ascending='TRUE'/>
     * @param limit  number of lines
     */
    private getQuery(query: string, orderBy?: string[], limit?: number): string {
        return`<View Scope="RecursiveAll">
            <Query>
                ${query}
                ${orderBy ? `<OrderBy>${orderBy.join('')}</OrderBy>` : ""}
            </Query>            
            ${limit !== undefined ? `<RowLimit>${limit}</RowLimit>` : ""}
        </View>`;
    }
}
