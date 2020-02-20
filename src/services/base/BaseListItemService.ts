import { ServicesConfiguration } from "../..";
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep, find, assign, findIndex } from "@microsoft/sp-lodash-subset";
import { CamlQuery, List, sp } from "@pnp/sp";
import { Constants, FieldType, TestOperator, QueryToken, LogicalOperator } from "../../constants/index";
import { IBaseItem, IFieldDescriptor, IAddOrUpdateResult, IQuery, IPredicate, ILogicalSequence, IOrderBy } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { UtilsService } from "..";
import { SPItem, User, TaxonomyTerm, OfflineTransaction, SPFile } from "../../models";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";
import { BaseDbService } from "./BaseDbService";

/**
 * 
 * Base service for sp list items operations
 */
export class BaseListItemService<T extends IBaseItem> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/
    protected listRelativeUrl: string;
    protected initValues: any = {};
    protected taxoMultiFieldNames: any = {};

    protected attachmentsService: BaseDbService<SPFile>;
    /* AttachmentService */

    public get ItemFields(): any {
        const result = {};
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
        this.attachmentsService = new BaseDbService<SPFile>(SPFile, tableName + "_Attachments");

    }

    
    /***************************** External sources init and access **************************************/
    
    private initialized = false;
    protected get isInitialized(): boolean {
        return this.initialized;
    }
    private initPromise: Promise<void> = null;

    protected async init_internal(): Promise<void>{
        return;
    }

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
                        const fields = this.ItemFields;
                        const models = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                // REM MLS : lookup removed from preload
                                if(fieldDescription.modelName && 
                                    models.indexOf(fieldDescription.modelName) === -1 && 
                                    fieldDescription.fieldType !== FieldType.Lookup && 
                                    fieldDescription.fieldType !== FieldType.LookupMulti) {
                                    models.push(fieldDescription.modelName);
                                }                                            
                            }
                        }
                        await Promise.all(models.map(async (modelName) => {
                            if(!this.initValues[modelName]) {
                                const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                const values = await service.getAll();
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
    private fieldsInitialized = false;
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
                        const fields = this.ItemFields;
                        const taxofields = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                if(fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                                    if(stringIsNullOrEmpty(fieldDescription.hiddenFieldName)) {
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
        const item = new this.itemType();
        Object.keys(this.ItemFields).map((propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            this.setFieldValue(spitem, item, propertyName, fieldDescription);
        });
        return item;
    }

    private setFieldValue(spitem: any, destItem: T, propertyName: string,  fieldDescriptor: IFieldDescriptor): void {
        const converted = destItem as unknown as SPItem;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch(fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if(fieldDescriptor.fieldName === Constants.commonFields.version) {
                    converted[propertyName] = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                }
                else if(fieldDescriptor.fieldName === Constants.commonFields.attachments) {
                    converted[propertyName] = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].map((fileobj) => {return new SPFile(fileobj);}) : fieldDescriptor.defaultValue;
                }
                else {
                    converted[propertyName] = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                }                
                break;                
            case FieldType.Date:
                converted[propertyName] = spitem[fieldDescriptor.fieldName] ? new Date(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                const lookupId: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(lookupId !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
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
                    const lookupIds: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results: spitem[fieldDescriptor.fieldName + "Id"]) : [];
                    if(lookupIds.length > 0) {
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {   
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
                const id: number = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if(id !== -1) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {                         
                        // get values from init values
                        const users = this.getServiceInitValues(fieldDescriptor.modelName);                        
                        const existing = find(users, (user) => {
                            return user.id === id;
                        });
                        converted[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        converted[propertyName] = id;
                    } 
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }                      
                break;
            case FieldType.UserMulti:
                const ids: Array<number> = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results: spitem[fieldDescriptor.fieldName + "Id"]) : [];                
                if(ids.length > 0) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {    
                        // get values from init values
                        const val = [];
                        const users = this.getServiceInitValues(fieldDescriptor.modelName);
                        ids.forEach(umid => {
                            const existing = find(users, (user) => {
                                return user.id === umid;
                            });
                            if(existing) {
                                val.push(existing);
                            } 
                        });
                        converted[propertyName] = val;
                    }
                    else {
                        converted[propertyName] = ids;
                    }
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                const wssid: number = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if(id !== -1) {
                    const tterms = this.getServiceInitValues(fieldDescriptor.modelName);
                    converted[propertyName] = this.getTaxonomyTermByWssId(wssid, tterms);
                }
                else {
                    converted[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                    const tmterms = spitem[fieldDescriptor.fieldName] ? (spitem[fieldDescriptor.fieldName].results ? spitem[fieldDescriptor.fieldName].results: spitem[fieldDescriptor.fieldName]) : [];
                    if(tmterms.length > 0) {
                        const allterms = this.getServiceInitValues(fieldDescriptor.modelName);
                        converted[propertyName] = tmterms.map((term) => {
                            return this.getTaxonomyTermByWssId(term.WssId, allterms);
                        });
                    }
                    else {
                        converted[propertyName] = fieldDescriptor.defaultValue;
                    }
                break;
            case FieldType.Json:
                converted[propertyName] = spitem[fieldDescriptor.fieldName] ? JSON.parse(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
    }
    /****************************** Send item methods ***********************************/
    private async getSPRestItem(item: T): Promise<any> {
        const spitem = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            await this.setRestFieldValue(item, spitem, propertyName, fieldDescription);
        }));
        return spitem;
    }
    private async setRestFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor: IFieldDescriptor): Promise<void> {
        const converted = item as unknown as SPItem;
        const itemValue = converted[propertyName];
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
                const link = converted.__getInternalLinks(propertyName);        
                if(itemValue) {
                    if(typeof(itemValue) === "number") {
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
                if(itemValue && isArray(itemValue) && itemValue.length > 0){
                    const links = converted.__getInternalLinks(propertyName);
                    const firstLookupVal = itemValue[0];
                    if(typeof(firstLookupVal) === "number") {
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: itemValue};
                    }
                    else {
                        if(links && links.length > 0) {
                            destItem[fieldDescriptor.fieldName + "Id"] = {results: links};
                        }
                        else {
                            destItem[fieldDescriptor.fieldName + "Id"] = {results: []};
                        }
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
                    const firstUserVal = itemValue[0];
                    if(typeof(firstUserVal) === "number") {
                        destItem[fieldDescriptor.fieldName + "Id"] = {results: itemValue};
                    }
                    else {
                        const userIds = await Promise.all(itemValue.map((user) => {
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
                const hiddenFieldName = this.taxoMultiFieldNames[fieldDescriptor.fieldName];
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
            if(value.id <=0) {
                const userService: UserService = new UserService();
                value = await userService.linkToSpUser(value);

            }
            result = value.id;
        }
        return result;
    }

    /**
     * 
     * @param wssid 
     * @param terms 
     */
    public getTaxonomyTermByWssId<TermType extends TaxonomyTerm>(wssid: number, terms: Array<TermType>): TermType {
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
    protected async  needRefreshCache(key = "all"): Promise<boolean> {
        let result: boolean = await super.needRefreshCache(key);

        if (!result) {

            const isconnected = await UtilsService.CheckOnline();
            if (isconnected) {

                const cachedDataDate = await super.getCachedData(key);
                if (cachedDataDate) {

                    try {
                        const response = await ServicesConfiguration.context.spHttpClient.get(`${ServicesConfiguration.context.pageContext.web.absoluteUrl}/_api/web/getList('${this.listRelativeUrl}')`,
                            SPHttpClient.configurations.v1,
                            {
                                headers: {
                                    'Accept': 'application/json;odata.metadata=minimal',
                                    'Cache-Control': 'no-cache'
                                }
                            });

                        const tempList = await response.json();
                        const lastModifiedDate = new Date(tempList.LastItemUserModifiedDate ? tempList.LastItemUserModifiedDate : tempList.d.LastItemUserModifiedDate);
                        result = lastModifiedDate > cachedDataDate;


                    } catch (error) {
                        console.error(error);
                    }


                }
            }
        }

        return result;
    }
    /**
     * Retrieve id of items to be reloaded
     * @param ids Id if items to check
     */
    protected async getExpiredIds(...ids: Array<number | string>): Promise<Array<number | string>> {
        let result: Array<number | string> = await super.getExpiredIds(...ids);

        if (result.length < ids.length) {

            const isconnected = await UtilsService.CheckOnline();
            if (isconnected) {               

                try {
                    const response = await ServicesConfiguration.context.spHttpClient.get(`${ServicesConfiguration.context.pageContext.web.absoluteUrl}/_api/web/getList('${this.listRelativeUrl}')`,
                        SPHttpClient.configurations.v1,
                        {
                            headers: {
                                'Accept': 'application/json;odata.metadata=minimal',
                                'Cache-Control': 'no-cache'
                            }
                        });
                        
                    const tempList = await response.json();
                    const lastModifiedDate = new Date(tempList.LastItemUserModifiedDate ? tempList.LastItemUserModifiedDate : tempList.d.LastItemUserModifiedDate);
                    result = [];
                    ids.forEach((id) => {
                        const lastLoad = this.getIdLastLoad(id);
                        if(!lastLoad || lastLoad < lastModifiedDate) {
                            result.push(id);
                        }
                    });


                } catch (error) {
                    console.error(error);
                }


            }
        }

        return result;
    }

    /**********************************Service specific calls  *******************************/
    
    /**
     * Get items by caml query
     * @param query caml query (<Where></Where>)
     * @param orderBy array of <FieldRef Name='Field' Ascending='TRUE'/>
     * @param limit  number of lines
     * @param lastId last id for paged queries
     */
    /*public getByCamlQuery(query: string, orderBy?: string[], limit?: number, lastId?: number): Promise<Array<T>> {
        const queryXml = this.getQuery(query, orderBy,limit);
        const camlQuery = {
            ViewXml: queryXml
        } as CamlQuery;
        if(lastId !== undefined) {
            camlQuery.ListItemCollectionPosition = {
                "PagingInfo": "Paged=TRUE&p_ID=" + lastId
            };
        }
        return this.get(camlQuery);
    }*/

    /********************************** Link to lookups  *************************************/
    private linkedLookupFields(loadLookups?: Array<string>): any {
        const result =[];
        const fields = this.ItemFields;
        for (const key in fields) {
            if (fields.hasOwnProperty(key)) {
                const fieldDesc = fields[key] as IFieldDescriptor;
                if((fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) && !stringIsNullOrEmpty(fieldDesc.modelName)) {
                    if(!loadLookups || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
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
                allIds[fieldDesc.modelName] = allIds[fieldDesc.modelName] ||[];
                const ids = allIds[fieldDesc.modelName];
                items.forEach((item: T) => {
                    const converted = item as unknown as SPItem;
                    const links = converted.__getInternalLinks(key);
                    //init value 
                    if(fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) {
                        converted[key] = fieldDesc.defaultValue;
                    }
                    if(fieldDesc.fieldType === FieldType.Lookup && 
                        // lookup has value
                        links && 
                        links !== -1 &&
                        // not allready loaded (local cache)
                        (!this.initValues[fieldDesc.modelName]
                            ||
                        !find(this.initValues[fieldDesc.modelName], {id: links})
                        ) &&
                        // not allready in load list
                        ids.indexOf(links) === -1
                        ) {
                        
                        ids.push(links);
                    } 
                    else if(fieldDesc.fieldType === FieldType.LookupMulti &&
                        links && 
                        links.length > 0) {                        
                            links.forEach((id) =>{
                            if(// not allready loaded (local cache)
                            (!this.initValues[fieldDesc.modelName]
                                ||
                            !find(this.initValues[fieldDesc.modelName], {id: id})
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
        const promises = [];
        for (const modelName in allIds) {
            if (allIds.hasOwnProperty(modelName)) {
                const ids = allIds[modelName];
                if(ids && ids.length > 0) {
                    const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                    promises.push(service.getItemsById(ids));
                }
            }
        }
        // execute and store
        const results = await Promise.all(promises);
        results.forEach(itemsTab => {
            if(itemsTab.length > 0) {
                const modelName = itemsTab[0].constructor.name;
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
                    const converted = item as unknown as SPItem;
                    const links = converted.__getInternalLinks(propertyName);
                    if(fieldDesc.fieldType === FieldType.Lookup && 
                        links && 
                        links !== -1) {
                        const litem = find(refCol, {id: links});
                        if(litem) {
                            converted[propertyName] = litem;
                        }
                        
                    } 
                    else if(fieldDesc.fieldType === FieldType.LookupMulti &&
                        links && 
                        links.length > 0) {     
                        item[propertyName] = [];                   
                        links.forEach((id) =>{
                            const litem = find(this.initValues[propertyName], {id: id});
                            if(litem) {
                                converted[propertyName].push(litem);
                            }
                        });                        
                    }
                });  
            }
        }
    }

    private updateInternalLinks(item: T, loadLookups?: Array<string>): void {
        const converted = item as unknown as SPItem;
        const lookupFields = this.linkedLookupFields();
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName] as IFieldDescriptor;
                if(!loadLookups || loadLookups.indexOf(fieldDesc.fieldName) !== -1) {
                    if(fieldDesc.fieldType === FieldType.Lookup) {
                        converted.__deleteInternalLinks(propertyName);
                        if(converted[propertyName] && converted[propertyName].id > -1) {
                            converted.__setInternalLinks(propertyName, converted[propertyName].id);
                        }
                    }
                    else if(fieldDesc.fieldType === FieldType.LookupMulti) {
                        converted.__deleteInternalLinks(propertyName);
                        if(converted[propertyName] && converted[propertyName].length > 0) {
                            converted.__setInternalLinks(propertyName, converted[propertyName].filter(l => l.id !== -1).map(l => l.id));
                        }
                    }
                }                
            }
        }
    }
    /***************** SP Calls associated to service standard operations ********************/
    
    public async get(query: IQuery, loadLookups?: Array<string>): Promise<Array<T>>{
        const results = await super.get(query);
        await this.populateLookups(results, loadLookups);
        return results;
    }

    /**
     * Get items by query
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected async get_Internal(query: IQuery): Promise<Array<T>> {
        const spQuery = this.getCamlQuery(query);
        let results = new Array<T>();
        const selectFields = this.getOdataFieldNames();
        let itemsQuery = this.list.select(...selectFields);
        if(this.hasAttachments) {
            itemsQuery = itemsQuery.expand(Constants.commonFields.attachments);
        }
        const items = await itemsQuery.getItemsByCAMLQuery(spQuery);
        if(items && items.length > 0) {
            await this.Init();
            results = items.map((r) => { 
                return this.getItemFromRest(r); 
            });
        }
        return results;
    }

    public async getItemById(id: number, loadLookups?: Array<string>): Promise<T>{
        const result = await super.getItemById(id);
        await this.populateLookups([result], loadLookups);
        return result;
    }

    /**
     * Get an item by id
     * @param id item id
     */
    protected async getItemById_Internal(id: number): Promise<T> {
        let result = null;
        const selectFields = this.getOdataFieldNames();
        let itemsQuery = this.list.items.getById(id).select(...selectFields);
        if(this.hasAttachments) {
            itemsQuery = itemsQuery.expand(Constants.commonFields.attachments);
        }
        const temp = await itemsQuery.get();
        if (temp) {
            await this.Init();
            result = this.getItemFromRest(temp);
            return result;
        }

        return result;
    }


    
    public async getItemsById(ids: Array<number>, loadLookups?: Array<string>): Promise<Array<T>>{
        const results = await super.getItemsById(ids);
        await this.populateLookups(results, loadLookups);
        return results;
    }

    /**
     * Get a list of items by id
     * @param id item id
     */
    protected async getItemsById_Internal(ids: Array<number>): Promise<Array<T>> {
        const results: Array<T> = [];
        const selectFields = this.getOdataFieldNames();
        const batch = sp.createBatch();
        ids.forEach((id) => {
            let itemsQuery = this.list.items.getById(id).select(...selectFields);
            if(this.hasAttachments) {
                itemsQuery = itemsQuery.expand(Constants.commonFields.attachments);
            }
            itemsQuery.inBatch(batch).get().then((item)=> {
                results.push(this.getItemFromRest(item));
            });
        });
        await batch.execute();
        return results;   
    }


    public async getAll(loadLookups?: Array<string>): Promise<Array<T>>{
        const results = await super.getAll();
        await this.populateLookups(results, loadLookups);
        return results;
    }

    /**
     * Retrieve all items
     * 
     */
    protected async getAll_Internal(): Promise<Array<T>> {
        let results: Array<T> = [];
        const selectFields = this.getOdataFieldNames();
        let itemsQuery = this.list.items.select(...selectFields);
        if(this.hasAttachments) {
            itemsQuery = itemsQuery.expand(Constants.commonFields.attachments);
        }
        const items = await itemsQuery.getAll();
        if(items && items.length > 0) {
            await this.Init();
            results = items.map((r) => { 
                return this.getItemFromRest(r); 
            });
        }
        return results;
    }

    public async addOrUpdateItem(item: T, loadLookups?: Array<string>): Promise<IAddOrUpdateResult<T>>{        
        this.updateInternalLinks(item, loadLookups);
        return super.addOrUpdateItem(item);
    }

    /**
     * Add or update an item
     * @param item SPItem derived object to be converted
     */
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        const result = cloneDeep(item);
        await this.initFields();
        const selectFields = this.getOdataCommonFieldNames();
        if (item.id < 0) {
            const converted = await this.getSPRestItem(item);
            const addResult = await this.list.items.select(...selectFields).add(converted);                    
            await this.populateCommonFields(result, addResult.data);                   
            await this.updateWssIds(result, addResult.data); 
            if(item.id < -1) {
                await this.updateLinksInDb(Number(item.id), Number(result.id));
            }
        }
        else {            
            // check version (cannot update if newer)
            if (item.version) {
                const existing = await this.list.items.getById(item.id as number).select(Constants.commonFields.version).get();
                if (parseFloat(existing[Constants.commonFields.version]) > item.version) {
                    const error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                    error.name = Constants.Errors.ItemVersionConfict;
                    throw error;
                }
                else {
                    const converted = await this.getSPRestItem(item);
                    const updateResult = await this.list.items.getById(item.id as number).select(...selectFields).update(converted);
                    const version = await updateResult.item.select(...selectFields).get();                    
                    await this.populateCommonFields(result, version);                    
                    await this.updateWssIds(result, version);
                }
            }
            else {
                const converted = await this.getSPRestItem(item);
                const updateResult = await this.list.items.getById(item.id as number).update(converted);
                const version = await updateResult.item.select(...selectFields).get();
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
        await this.list.items.getById(item.id as number).recycle();
    }

    private async getAttachmentContent(attachment: SPFile): Promise<void> {
        const content = await sp.web.getFileByServerRelativeUrl(attachment.serverRelativeUrl).getBuffer();
        attachment.content = content;
    }

    public async cacheAttachmentsContent(): Promise<void> {
        const prop = this.attachmentProperty;
        if(prop !== null) {
            let load = true;
            if (ServicesConfiguration.configuration.checkOnline) {
                load = await UtilsService.CheckOnline();
            }
            if(load) {
                const updatedItems: T[] = [];
                const operations: Promise<void>[] = [];
                const items = await this.dbService.getAll();
                for (const item of items) {
                    const converted = await this.mapItem(item);
                    if(converted[prop] && converted[prop].length > 0) {
                        updatedItems.push(converted);
                        converted[prop].forEach(attachment => {
                            operations.push(this.getAttachmentContent(attachment));
                        });
                    }
                    
                }
                operations.map(operation => {
                    return operation;                  
                }).reduce((chain, operation) => {                  
                    return chain.then(() => {return operation;});                  
                }, Promise.resolve()).then(async() => {

                    if(updatedItems.length > 0) {
                        const dbitems = await Promise.all(updatedItems.map((u) => {
                            return this.convertItemToDbFormat(u);
                        }));
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
                if(fieldDesc.fieldName === Constants.commonFields.attachments) {
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
    private getOdataFieldNames(): Array<string> {
        const fields = this.ItemFields;
        const fieldNames = Object.keys(fields).filter((propertyName) => { 
            return fields.hasOwnProperty(propertyName); 
        }).map((prop) => {
            let result: string = fields[prop].fieldName;
            switch(fields[prop].fieldType) {
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

    private getOdataCommonFieldNames(): Array<string> {
        const fields = this.ItemFields;
        const fieldNames = [Constants.commonFields.version];
        Object.keys(fields).filter((propertyName) => { 
            return fields.hasOwnProperty(propertyName); 
        }).forEach((prop) => {
            const fieldName = fields[prop].fieldName;
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
        if(item.id < 0) {
            // update id
            item.id = restItem.Id;
        }
        if(restItem[Constants.commonFields.version]) {
            item.version = parseFloat(restItem[Constants.commonFields.version]);
        }
        const fields = this.ItemFields;
        await Promise.all(Object.keys(fields).filter((propertyName) => {
            if(fields.hasOwnProperty(propertyName)) {                
                const fieldName = fields[propertyName].fieldName;
                return (fieldName  === Constants.commonFields.author ||
                    fieldName  === Constants.commonFields.created ||
                    fieldName  === Constants.commonFields.editor ||
                    fieldName  === Constants.commonFields.modified);
            }
        }).map(async (prop) => {
            const fieldName = fields[prop].fieldName;            
            switch(fields[prop].fieldType) {
                case FieldType.Date:
                    item[prop] = new Date(restItem[fieldName]);
                    break;
                case FieldType.User:
                    const id = restItem[fieldName + "Id"];
                    let user = null;
                    if(this.initialized) {
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
     * @param item full provisionned item
     */
    protected async convertItemToDbFormat(item: T): Promise<T> {
        const converted = item as unknown as SPItem;
        const result: T = cloneDeep(item);
        const convertedResult = result as unknown as SPItem;
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                switch(fieldDescriptor.fieldType) {
                    case FieldType.User:           
                    case FieldType.Taxonomy:       
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            if(converted[propertyName]) {
                                convertedResult.__setInternalLinks(propertyName, converted[propertyName].id);
                            }
                            delete convertedResult[propertyName];
                        }
                        break;
                    case FieldType.UserMulti:            
                    case FieldType.TaxonomyMulti:                      
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {  
                            const ids = [];
                            if(converted[propertyName]) {
                                converted[propertyName].forEach(element => {
                                    if(element.id) {
                                        if((typeof(element.id) === "number" && element.id > 0) || (typeof(element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
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
                        if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            delete convertedResult[propertyName];
                            convertedResult.__setInternalLinks(propertyName, converted.__getInternalLinks(propertyName));
                        }
                        break;
                    default:
                        if(fieldDescriptor.fieldName === Constants.commonFields.attachments) {
                            let ids = [];
                            if(converted[propertyName] && converted[propertyName].length > 0) {
                                const files = await this.attachmentsService.addOrUpdateItems(converted[propertyName]);
                                ids = files.map((f) => {
                                    return f.id;
                                });
                            }                          
                            convertedResult.__setInternalLinks(propertyName, ids.length > 0 ? ids : []);                            
                            delete convertedResult[propertyName];
                        }
                        break;                    
                }
                
            }
        }
        return result;
    }

    /**
     * populate item from db storage
     * @param item db item with links in internalLinks fields
     */
    public async mapItem(item: T): Promise<T> {
        const converted = item as unknown as SPItem;
        const result: T = cloneDeep(item);
        const convertedResult= result as unknown as SPItem;
        await this.Init();
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                if(//fieldDescriptor.fieldType === FieldType.Lookup ||
                    fieldDescriptor.fieldType === FieldType.User ||
                    fieldDescriptor.fieldType === FieldType.Taxonomy) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const id: number = converted.__getInternalLinks(propertyName);
                        if(id !== null) {
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
                else if(//fieldDescriptor.fieldType === FieldType.LookupMulti ||
                    fieldDescriptor.fieldType === FieldType.UserMulti ||
                    fieldDescriptor.fieldType === FieldType.TaxonomyMulti) {
                    if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {    
                        // get values from init values
                        const ids = converted.__getInternalLinks(propertyName) || [];
                        if(ids.length > 0) {
                            const val = [];
                            const targetItems = this.getServiceInitValues(fieldDescriptor.modelName);
                            ids.forEach(id => {
                                const existing = find(targetItems, (tmpitem) => {
                                    return tmpitem.id === id;
                                });
                                if(existing) {
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
                    if(fieldDescriptor.fieldName === Constants.commonFields.attachments) {
                        // get values from init values
                        const urls = converted.__getInternalLinks(propertyName) || [];
                        if(urls.length > 0) {
                            const files = await this.attachmentsService.getItemsById(urls);
                            convertedResult[propertyName] = files;
                        }                            
                        else {
                            convertedResult[propertyName] = fieldDescriptor.defaultValue;
                        }                        
                        convertedResult.__deleteInternalLinks(propertyName);
                    }
                    else {
                        convertedResult[propertyName] = converted[propertyName] ;
                    }
                }       
            }
        }        
        convertedResult.__clearEmptyInternalLinks();
        return result;
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
                    if(fieldDescription.refItemName === this.itemType["name"] || fieldDescription.modelName === this.itemType["name"]) {
                        // get object if not done yet
                        if(!currentObject) {
                            const destType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                            currentObject = new destType();
                            assign(currentObject, transaction.itemData);
                        }                        
                        if(fieldDescription.fieldType === FieldType.Lookup) {
                            if(fieldDescription.modelName) {
                                // search in internalLinks
                                const link = currentObject.__getInternalLinks(propertyName);
                                if(link && link === oldId) {
                                    currentObject.__setInternalLinks(propertyName, newId);
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
                                // serch in internalLinks
                                const links = currentObject.__getInternalLinks(propertyName);
                                if(links && isArray(links)) {
                                    // find item
                                    const lookupidx = findIndex(links, (id) => {return id === oldId;});
                                    // change id
                                    if(lookupidx > -1) {
                                        currentObject.__setInternalLinks(propertyName, newId);
                                        needUpdate = true;
                                    }
                                }
                            }
                            else if(currentObject[propertyName] && isArray(currentObject[propertyName])){
                                // find index
                                const lookupidx = findIndex(currentObject[propertyName], (id) => {return id === oldId;});
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
        const allFields = assign({}, this.itemType["Fields"]);
        delete allFields[SPItem["name"]];
        delete allFields[this.itemType["name"]];
        for (const modelName in allFields) {
            if (allFields.hasOwnProperty(modelName)) {     
                const modelFields =  allFields[modelName];    
                const lookupProperties = Object.keys(modelFields).filter((prop) => {
                    return (modelFields[prop].refItemName &&
                        modelFields[prop].refItemName === this.itemType["name"] || modelFields[prop].modelName === this.itemType["name"]);
                });
                if(lookupProperties.length > 0) {
                    const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                    const allitems = await service.__getAllFromCache();
                    const updated = [];
                    allitems.forEach(element => { 
                        const converted = element as unknown as SPItem;
                        let needUpdate = false;                       
                        lookupProperties.forEach(propertyName => {
                            const fieldDescription = modelFields[propertyName];                
                            if(fieldDescription.fieldType === FieldType.Lookup) {
                                if(fieldDescription.modelName) {
                                    // serch in internalLinks
                                    const link = converted.__getInternalLinks(propertyName);
                                    if(link && link === oldId) {
                                        converted.__setInternalLinks(propertyName, newId);
                                        needUpdate = true;
                                    }
                                }
                                else if(converted[propertyName] === oldId){
                                    // change field
                                    converted[propertyName] = newId;
                                    needUpdate = true;
                                }
                            }
                            else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                if(fieldDescription.modelName) {
                                    // search in internalLinks
                                    const links = converted.__getInternalLinks(propertyName);
                                    if(links && isArray(links)) {
                                        // find item
                                        const lookupidx = findIndex(links, (id) => {return id === oldId;});
                                        // change id
                                        if(lookupidx > -1) {
                                            converted.__setInternalLinks(propertyName, newId);
                                            needUpdate = true;
                                        }
                                    }
                                }
                                else if(converted[propertyName] && isArray(converted[propertyName])){
                                    // find index
                                    const lookupidx = findIndex(converted[propertyName], (id) => {return id === oldId;});
                                    // change field
                                    // change id
                                    if(lookupidx > -1) {
                                        converted[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                            }
                        });
                        if(needUpdate) {
                            updated.push(converted);
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
        const fields = this.ItemFields;
        // serch for Taxonomy fields
        for (const propertyName in fields) {
            if (fields.hasOwnProperty(propertyName)) {     
                           
                const fieldDescription: IFieldDescriptor = fields[propertyName];
                if(fieldDescription.fieldType === FieldType.Taxonomy) {
                    let needUpdate = false;
                    // get wssid from item
                    const wssid = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].WssId : -1;
                    if(wssid !== -1) {
                        const id = item[propertyName].id;
                        // find corresponding object in service
                        const service = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription.modelName);
                        const term = await service.__getFromCache(id);
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
                                const idx = findIndex(this.initValues[fieldDescription.modelName], (t: any) => { return t.id === id; });
                                if(idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = term;
                                }
                            }
                        }
                    }
                }
                else if (fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                    const updated = [];
                    const terms = spItem[fieldDescription.fieldName] ? spItem[fieldDescription.fieldName].results : [];     
                    const service = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription.modelName);              
                    if(terms && terms.length > 0) {
                        await Promise.all(terms.map(async (termitem) => {
                            const wssid = termitem.WssId;
                            const id = termitem.TermGuid;
                            // find corresponding object in allready updated
                            let term = find(updated, (u) => {return u.id === id;});
                            if(!term) {
                                term = await service.__getFromCache(id);
                            }
                            if(term instanceof TaxonomyTerm) {                           
                                term.wssids = term.wssids || [];
                                if(term.wssids.indexOf(wssid) === -1) {
                                    term.wssids.push(wssid);
                                    if(!find(updated, (u) => {return u.id === id;})) {
                                        updated.push(term);
                                    }
                                }
                            }
                        }));
                    }
                    if(updated.length > 0) {
                        await service.__updateCache(...updated);
                        // update initValues
                        if(this.initialized) {
                            updated.forEach((u) => {
                                const idx = findIndex(this.initValues[fieldDescription.modelName], (t: any) => { return t.id === u.id; });
                                if(idx !== -1) {
                                    this.initValues[fieldDescription.modelName][idx] = u;
                                }
                            });      
                        }                  
                    }
                }
            }
        }
    }
    

    private getCamlQuery(query: IQuery): CamlQuery {
        const result: CamlQuery = {
            ViewXml: `<View Scope="RecursiveAll">
                <Query>
                    ${this.getWhere(query)}
                    ${this.getOrderBy(query)}
                </Query>
                ${query.limit !== undefined ? `<RowLimit>${query.limit}</RowLimit>` : ""}
            </View>`
        };
        if(query.lastId !== undefined) {
            result.ListItemCollectionPosition = {
                "PagingInfo": "Paged=TRUE&p_ID=" + query.lastId
            };
        }
        return result;
    }

    private getOrderBy(query: IQuery): string {
        let result = "";
        if(query.orderBy && query.orderBy.length > 0) {
            result = 
            result = `<OrderBy>
                ${query.orderBy.map(ob => this.getFieldRef(ob)).join('')}
            </OrderBy>`;
        }
        return result;
    }
    private getWhere(query: IQuery): string {
        let result = "";
        if(query.test) {
            result = `<Where>
                ${query.test.type === "predicate" ? this.getPredicate(query.test) : this.getLogicalSequence(query.test)}
            </Where>`;
        }
        return result;
    }
    private getLogicalSequence(sequence: ILogicalSequence): string {
        if(!sequence.children || sequence.children.length === 0) {
            return "";
        }
        if(sequence.children.length === 1) {
            if(sequence.children[0].type === "predicate") {
                return this.getPredicate(sequence.children[0]);
            }
            else {
                return this.getLogicalSequence(sequence.children[0]);
            }
        }
        else {
            // first part
            let result = `<${sequence.operator}>`;
            if(sequence.children[0].type === "predicate") {
                result += this.getPredicate(sequence.children[0]);
            }
            else {
                result += this.getLogicalSequence(sequence.children[0]);
            }
            sequence.children.splice(0,1);
            result += this.getLogicalSequence(sequence);
            result += `</${sequence.operator}>`;
            return result;
        }
    }

    private getPredicate(predicate: IPredicate): string {
        let result = "";
        switch (predicate.operator) {
            case TestOperator.IsNotNull:
            case TestOperator.IsNull:
                result = `<${predicate.operator}>
                    ${this.getFieldRef(predicate)}
                </${predicate.operator}>`;
                break;    
            case TestOperator.In:
                if(predicate.value && isArray(predicate.value) && predicate.value.length > 0) {
                    if(predicate.value.length <= 500) {
                        return `<${predicate.operator}>
                            ${this.getFieldRef(predicate)}
                            <Values>
                                ${predicate.value.map(v => this.getValue(predicate.propertyName, v, predicate.lookupId)).join('')}
                            </Values>
                        </${predicate.operator}>`;
                    }
                    else {
                        const transformed: ILogicalSequence = {
                            type: "sequence",
                            operator: LogicalOperator.Or,
                            children: []
                        };
                        const copy = predicate.value
                        while(copy.length) {
                            const subValues = copy.splice(0,500);
                            transformed.children.push({
                                type: "predicate",
                                operator: TestOperator.In,
                                propertyName: predicate.propertyName,
                                value: subValues,
                                includeTimeValue: predicate.includeTimeValue,
                                lookupId: predicate.lookupId
                            });
                        }
                        return this.getLogicalSequence(transformed);
                    }
                }
                else {
                    return `<${predicate.operator}>
                        ${this.getFieldRef(predicate)}
                        <Values>
                            ${this.getValue(predicate.propertyName, -1, predicate.lookupId)}
                        </Values>
                    </${predicate.operator}>`;
                }

                break;    
            default:
                result = `<${predicate.operator}>
                    ${this.getFieldRef(predicate)}
                    ${this.getValue(predicate.propertyName, predicate.value, predicate.lookupId)}
                </${predicate.operator}>`;
                break;
        }
        return result;
    }
    private getFieldRef(obj: IPredicate | IOrderBy): string {
        let result = "";
        const fields = this.ItemFields;
        const field = fields[obj.propertyName];
        if(field) {    
            result = `<FieldRef Name="${field.fieldName}"${obj.type === "predicate" && obj.lookupId ? " LookupId=\"TRUE\"" : ""}${obj.type === "predicate" && obj.includeTimeValue !== undefined ? (" IncludeTimeValue=\"" + (obj.includeTimeValue ? "TRUE" : "FALSE") + "\"") : ""}${obj.type === "orderby" && obj.ascending !== undefined && ! obj.ascending ? " Ascending=\"FALSE\"" : ""} />`;
        }
        else {
            throw new Error("Field was not found : " + obj.propertyName);
        }
        return result;
    }
    private getValue(propertyName: string, fieldValue: any, lookupID?: boolean): string {
        let result = "";
        const fields = this.ItemFields;
        const field = fields[propertyName];
        if(field) {    
            let type = "";
            let value = "";
            switch (field.fieldType) {
                case FieldType.Simple:
                    if(typeof(fieldValue) === "number") {
                        type = "Number";       
                        value = fieldValue.toString();       
                    }
                    else if(typeof(fieldValue) === "boolean") {
                        type="Boolean";
                        value = fieldValue ? "1" : "0";
                    }          
                    else {
                        type="Text"; 
                        value = fieldValue.toString(); 
                    }                     
                    break;
                case FieldType.Date:
                    type = "DateTime";
                    if(fieldValue === QueryToken.Now || fieldValue === QueryToken.Today) {
                        value = `<${fieldValue}/>`;
                    }
                    else {
                        value = fieldValue.toISOString().replace(/\.\d+Z/g, "Z");
                    }
                    break;                
                case FieldType.Json:
                    type="Text";
                    value = JSON.stringify(fieldValue);
                    break;               
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                    type = "Lookup";        
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
                    if(fieldValue === QueryToken.UserID) {
                        value = `<${fieldValue}/>`;
                    }
                    else {                        
                        value = fieldValue.toString();   
                    }
                    break;
                default:
                    break;
            }
            result = `<Value Type="${type}">${value}</Value>`;
        }
        else {
            throw new Error("Field was not found : " + propertyName);
        }
        return result;
    }
}
