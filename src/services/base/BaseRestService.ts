import { ServicesConfiguration } from "../../configuration";
import { cloneDeep, find, assign, findIndex } from "@microsoft/sp-lodash-subset";
import { Constants, FieldType, TestOperator } from "../../constants/index";
import { IFieldDescriptor, IQuery, ILogicalSequence, IRestQuery, IRestLogicalSequence, IEndPointBindings, IPredicate, IRestPredicate, IOrderBy, IPreloadedData } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { UtilsService } from "../UtilsService";
import { RestItem, User, OfflineTransaction, RestResultMapping } from "../../models";
import { BaseItem } from "../../models/base/BaseItem";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";
import { RestFile } from "../../models/base/RestFile";
import * as mime from "mime-types";
import { ServiceFactory } from "../ServiceFactory";
import { IEndPointBinding } from "../../interfaces/IEndPointBindings";
import { BaseDbService } from "../base/BaseDbService";
import { Decorators } from "../../decorators";

const trace = Decorators.trace;

/**
 * 
 * Base service for sp list items operations
 */
export class BaseRestService<T extends (RestItem | RestFile)> extends BaseDataService<T>{

    /***************************** Fields and properties **************************************/

    protected restMappingDb: BaseDbService<RestResultMapping>;  


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
    constructor(type: (new (item?: any) => T), baseServiceUrl: string, cacheDuration?: number) {
        super(type, cacheDuration);
        this.baseServiceUrl = baseServiceUrl;
        this.restMappingDb = new BaseDbService(RestResultMapping, "RestMapping");
    }   

    /****************************** get item methods ***********************************/
    protected setPreloaded(preloaded: IPreloadedData ,modelName: string, itemId: string, propertyName: string, data: any): void {
        preloaded = preloaded || {};
        preloaded[modelName] = preloaded[modelName] || {};
        preloaded[modelName][itemId] = preloaded[modelName][itemId] || {};
        preloaded[modelName][itemId][propertyName] = data;
    }

    protected async getItemFromRest(restItem: any, preloaded: IPreloadedData): Promise<T> {
        const item = new this.itemType();
        const allProperties = Object.keys(this.ItemFields);
        // id used for links, should be populated first
        const idx = allProperties.indexOf(Constants.commonRestFields.id);
        if(idx !== -1) {
            allProperties.splice(idx, 1);
            allProperties.unshift(Constants.commonRestFields.id);
        }
        // set field values
        for (const propertyName of allProperties) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescription = this.ItemFields[propertyName];
                await this.setFieldValue(restItem, item, propertyName, fieldDescription, preloaded);
            }
        }
        // 
        if (item instanceof RestFile) {
            item.mimeType = (mime.lookup(item.title) as string) || 'application/octet-stream';
        }
        return item;
    }

    // TODO : test
    private async setFieldValue(restItem: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor, preloaded: IPreloadedData): Promise<void> {
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
                destItem[propertyName] = restItem[fieldDescriptor.fieldName] !== null && restItem[fieldDescriptor.fieldName] !== undefined ? restItem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                break;
            case FieldType.Date:
                destItem[propertyName] = restItem[fieldDescriptor.fieldName] ? new Date(restItem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                if (fieldDescriptor.containsFullObject && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                    const obj = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : null;
                    if (obj) {                        
                        this.setPreloaded(preloaded, fieldDescriptor.modelName, destItem.id.toString(), propertyName, obj);
                        // get service
                        /*const tmpservice = ServiceFactory.getServiceByModelName(fieldDescriptor.modelName);
                        const conv = await tmpservice.persistItemData(obj);
                        if (conv) {
                            this.updateInitValues(fieldDescriptor.modelName, conv);
                            destItem.__setInternalLinks(propertyName, conv.id);
                            destItem[propertyName] = conv;
                        }
                        else {
                            destItem[propertyName] = fieldDescriptor.defaultValue;
                        }*/

                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                }
                else {
                    const lookupId: number = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : -1;
                    if (lookupId !== -1) {
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            // LOOKUPS --> links
                            destItem.__setInternalLinks(propertyName, lookupId);
                            destItem[propertyName] = fieldDescriptor.defaultValue;

                        }
                        else {
                            destItem[propertyName] = lookupId;
                        }

                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                }
                break;
            case FieldType.LookupMulti: // TODO : in loadlookup
                if (fieldDescriptor.containsFullObject && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                    //let convertedObjects = [];
                    const values = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                    if (values.length > 0) {
                        // store localy (persist in populatelookups)
                        this.setPreloaded(preloaded, fieldDescriptor.modelName, destItem.id.toString(), propertyName, values);
                        /*
                        // get service
                        const tmpservice = ServiceFactory.getServiceByModelName(fieldDescriptor.modelName);
                        convertedObjects = await tmpservice.persistItemsData(values);                        
                        this.updateInitValues(fieldDescriptor.modelName, ...convertedObjects);
                        destItem.__setInternalLinks(propertyName, convertedObjects.map(c => c.id));
                        destItem[propertyName] = convertedObjects;
                        */
                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                }
                else {
                    const lookupIds: Array<number> = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName].map(ri => ri.id) : [];
                    if (lookupIds.length > 0) {
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            // LOOKUPS --> links
                            destItem.__setInternalLinks(propertyName, lookupIds);
                            destItem[propertyName] = fieldDescriptor.defaultValue;
                        }
                        else {
                            destItem[propertyName] = lookupIds;
                        }
                    }
                    else {
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                }
                break;
            case FieldType.User:
                const upn: string = restItem[fieldDescriptor.fieldName];
                if (!stringIsNullOrEmpty(upn)) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const users = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                        const existing = find(users, (user: User) => {
                            return !stringIsNullOrEmpty(user.userPrincipalName) && user.userPrincipalName.toLowerCase() === upn.toLowerCase();
                        });
                        destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        destItem[propertyName] = upn;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.UserMulti:
                const upns: Array<string> = restItem[fieldDescriptor.fieldName] ? restItem[fieldDescriptor.fieldName] : [];
                if (upns.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        const val = [];
                        const users = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                        upns.forEach(umupn => {
                            if (!stringIsNullOrEmpty(umupn)) {
                                const existing = find(users, (user: User) => {
                                    return !stringIsNullOrEmpty(user.userPrincipalName) && user.userPrincipalName.toLowerCase() === umupn.toLowerCase();
                                });
                                if (existing) {
                                    val.push(existing);
                                }
                            }
                        });
                        destItem[propertyName] = val;
                    }
                    else {
                        destItem[propertyName] = upns;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                const conJsonId = !stringIsNullOrEmpty(restItem[fieldDescriptor.fieldName]) ? JSON.parse(restItem[fieldDescriptor.fieldName]) : null;
                const termid: string = conJsonId && conJsonId.length > 0 ? conJsonId[0].id : null;
                if (!stringIsNullOrEmpty(termid)) {
                    const tterms = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                    const existing = find(tterms, (term) => {
                        return term.id === termid;
                    });
                    destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                const conJsonIds = !stringIsNullOrEmpty(restItem[fieldDescriptor.fieldName]) ? JSON.parse(restItem[fieldDescriptor.fieldName]) : null;
                const tmterms = conJsonIds ? conJsonIds : [];
                if (tmterms.length > 0) {
                    // get values from init values
                    const val = [];
                    const allterms = this.getServiceInitValuesByName(fieldDescriptor.modelName);
                    tmterms.forEach(tmterm => {
                        const existing = find(allterms, (term) => {
                            return term.id === tmterm.id;
                        });
                        if (existing) {
                            val.push(existing);
                        }
                    });
                    destItem[propertyName] = val;
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Json:
                if (restItem[fieldDescriptor.fieldName]) {
                    try {
                        if(fieldDescriptor.containsFullObject) {
                            if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                const itemType = ServiceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                                destItem[propertyName] = assign(new itemType(), restItem[fieldDescriptor.fieldName]);
                            }
                            else {
                                destItem[propertyName] = restItem[fieldDescriptor.fieldName];
                            }
                        }
                        else {
                            const jsonObj = JSON.parse(restItem[fieldDescriptor.fieldName]);
                            if(!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
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
                        destItem[propertyName] = fieldDescriptor.defaultValue;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
        }
    }
    /****************************** Send item methods ***********************************/
    protected async getRestItem(item: T): Promise<any> {
        const restItem = {};
        await Promise.all(Object.keys(this.ItemFields).map(async (propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            await this.setRestFieldValue(item, restItem, propertyName, fieldDescription);
        }));
        return restItem;
    }


    // TODO : test
    private async setRestFieldValue(item: T, destItem: any, propertyName: string, fieldDescriptor: IFieldDescriptor): Promise<void> {
        const itemValue = item[propertyName];
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;

        if (fieldDescriptor.fieldName !== Constants.commonRestFields.created &&
            fieldDescriptor.fieldName !== Constants.commonRestFields.author &&
            fieldDescriptor.fieldName !== Constants.commonRestFields.editor &&
            fieldDescriptor.fieldName !== Constants.commonRestFields.modified &&
            fieldDescriptor.fieldName !== Constants.commonRestFields.version &&
            (fieldDescriptor.fieldName !== Constants.commonRestFields.id || itemValue > 0) &&
            (fieldDescriptor.fieldName !== Constants.commonRestFields.uniqueid || item.id <= 0)
        ) {
            switch (fieldDescriptor.fieldType) {
                case FieldType.Simple:
                case FieldType.Date:
                    destItem[fieldDescriptor.fieldName] = itemValue;
                    break;
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
                case FieldType.Json:
                    if (fieldDescriptor.containsFullObject) {
                        destItem[fieldDescriptor.fieldName] = itemValue;
                    }
                    else {
                        destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                    }
                    break;
            }
        }
    }


    /********************** SP Fields conversion helpers *****************************/

    private async convertSingleUserFieldValue(value: User): Promise<string> {
        let result: string = null;
        if (value) {
            if (value.id <= 0) {
                const userService: UserService = ServiceFactory.getService(User).cast<UserService>();
                value = await userService.linkToSpUser(value);
            }
            result = value.userPrincipalName;
        }
        return result;
    }

    /**********************************Service specific calls  *******************************/


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

    private async persistPreloadedModelData(items: T[],modelName: string, modelData: {[itemId: string]: { [propertyName: string]: any }}): Promise<void> {
        // construct a list with elements to persist
        const dataObjects = [];                 
        for (const itemId in modelData) {
            if (modelData.hasOwnProperty(itemId)) {
                const itemData = modelData[itemId];
                for (const propertyName in itemData) {

                    if (itemData.hasOwnProperty(propertyName)) {
                        const dataObject = itemData[propertyName];
                        // get field description
                        const fieldDescription = this.ItemFields[propertyName];
                        if(fieldDescription.fieldType === FieldType.Lookup) {
                            dataObjects.push(dataObject);
                        }
                        else { // lookup multi
                            dataObjects.push(...dataObject);
                        }

                    }
                }
            }
        }
        if(dataObjects.length > 0) {
            // get service
            const modelService = ServiceFactory.getServiceByModelName(modelName);                        
            //persit items
            const resultItems = await modelService.persistItemsData(dataObjects);
            this.updateInitValues(modelName, ...resultItems);    
            // update items 
            for (const itemId in modelData) {
                if (modelData.hasOwnProperty(itemId)) {
                    const itemData = modelData[itemId];
                    for (const propertyName in itemData) {

                        if (itemData.hasOwnProperty(propertyName)) {
                            const dataObject = itemData[propertyName];
                            // get field description
                            const fieldDescription = this.ItemFields[propertyName];
                            if(fieldDescription.fieldType === FieldType.Lookup) {
                                const linkedItem =  find(resultItems, i => i.id === dataObject.id);
                                const item = find(items, i => i.id.toString() === itemId);
                                if(item) {
                                    item[propertyName]= linkedItem;
                                    item.__setInternalLinks(propertyName, linkedItem.id);
                                }
                            }
                            else { // lookup multi
                                const linkedItems =  resultItems.filter(i => dataObject.some(d => d.id === i.id));
                                const item = find(items, i => i.id.toString() === itemId);
                                if(item) {
                                    item[propertyName]= linkedItems;
                                    item.__setInternalLinks(propertyName, linkedItems.map(l => l.id));
                                }
                            }    
                        }
                    }
                }
            }

        }  
    }

    @trace()
    protected async populateLookups(items: Array<T>, loadEmbeded: boolean, preloaded: IPreloadedData, loadLookups?: Array<string>): Promise<void> {
        await this.Init();
        // get lookup fields
        const lookupFields = this.linkedLookupFields(loadLookups);
        // persitst preloaded
        preloaded = preloaded || {};
        const keys = Object.keys(preloaded).filter(modelName => preloaded.hasOwnProperty(modelName));
        if(keys.length > 0) {
            await Promise.all(keys.map(modelName => this.persistPreloadedModelData(items,modelName, preloaded[modelName])));
        }
        
        // init values and retrieve all ids by model
        const allIds = {};
        for (const key in lookupFields) {
            if (lookupFields.hasOwnProperty(key)) {
                const fieldDesc = lookupFields[key];
                // if online & containsfullobject, initvalues allready updated before by internal getter
                if(!fieldDesc.containsFullObject || loadEmbeded) {
                    allIds[fieldDesc.modelName] = allIds[fieldDesc.modelName] || [];
                    const ids = allIds[fieldDesc.modelName];
                    items.forEach((item: T) => {
                        const links = item.__getInternalLinks(key);
                        //init value 
                        if (fieldDesc.fieldType === FieldType.Lookup || fieldDesc.fieldType === FieldType.LookupMulti) {
                            item[key] = fieldDesc.defaultValue;
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
        }
        // Init queries       
        const promises: Array<Promise<BaseItem[]>> = [];
        for (const modelName in allIds) {
            if (allIds.hasOwnProperty(modelName)) {
                const ids = allIds[modelName];
                if (ids && ids.length > 0) {
                    const service = ServiceFactory.getServiceByModelName(modelName);
                    promises.push(service.getItemsById(ids));
                }
            }
        }
        // execute and store
        const results = await UtilsService.runPromisesInStacks(promises, 3);
        results.forEach(itemsTab => {
            if (itemsTab.length > 0) {
                this.updateInitValues(itemsTab[0].constructor["name"], ...itemsTab);
            }
        });
        // Associate to items
        for (const propertyName in lookupFields) {
            if (lookupFields.hasOwnProperty(propertyName)) {
                const fieldDesc = lookupFields[propertyName];
                if(!fieldDesc.containsFullObject || loadEmbeded){
                    const refCol = this.initValues[fieldDesc.modelName];
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
    /***************** SP Calls associated to service standard operations ********************/
    
    /**
     * Get items by query
     * @protected
     * @param {IQuery} query - query used to retrieve items
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    @trace()
    protected async get_Internal(query: IQuery, linkedFields?: Array<string>, preloaded?: IPreloadedData): Promise<Array<T>> {
        const restQuery = this.getRestQuery(query);
        if (linkedFields && linkedFields.length === 1 && linkedFields[0] === 'loadAll') {
            restQuery.loadAll = true;
        }
        let results = new Array<T>();
        const items = await this.executeRequest(`${this.serviceUrl}${this.Bindings.get.url}`, this.Bindings.get.method, restQuery);
        if (items && items.length > 0) {
            await this.Init();
            preloaded = preloaded || {};
            results = await Promise.all(items.map((r) => {
                return this.getItemFromRest(r, preloaded);
            }));
            await this.populateLookups(results, false, preloaded, linkedFields);
        }
        return results;
    }

    /**
     * Get an item by id
     * @param {number} id - item id
     */
    @trace()
    protected async getItemById_Internal(id: number, linkedFields?: Array<string>): Promise<T> {
        let result = null;
        const temp = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getItemById.url}/${id}`, this.Bindings.getItemById.method);
        if (temp) {
            await this.Init();
            const preloaded: IPreloadedData = {};
            result = await this.getItemFromRest(temp, preloaded);
            await this.populateLookups([result], false, preloaded, linkedFields);
        }
        return result;
    }


    /**
     * Get a list of items by id
     * @param ids - array of item id to retrieve
     */
    @trace()
    protected async getItemsById_Internal(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<T>> {
        const result: Array<T> = [];
        const preloaded: IPreloadedData = {};
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
            }, linkedFields, preloaded));
        }
        const res = await UtilsService.runPromisesInStacks(promises, 3);
        for (const tmp of res) {
            result.push(...tmp.filter(i => { return i !== null && i !== undefined; }));
        }
        if(result.length > 0) {
            await this.populateLookups(result, false, preloaded, linkedFields);
        }
        return result;
    }

    /**
     * Retrieve all items
     * 
     */
    @trace()
    protected async getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>> {
        let results: Array<T> = [];
        const items = await this.executeRequest(`${this.serviceUrl}${this.Bindings.getAll.url}`, this.Bindings.getAll.method);
        if (items && items.length > 0) {
            await this.Init();
            const preloaded: IPreloadedData = {};
            results = await Promise.all(items.map((r) => {
                return this.getItemFromRest(r, preloaded);
            }));
            await this.populateLookups(results, false, preloaded, linkedFields);
        }
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
    @trace()
    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        const result = cloneDeep(item);
        if (item.id < 0) {
            const converted = await this.getRestItem(item);
            const addResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
            await this.populateCommonFields(result, addResult);
            if (item.id < -1) {
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
                    const converted = await this.getRestItem(item);
                    const updateResult = await this.executeRequest(`${this.serviceUrl}${this.Bindings.addOrUpdateItem.url}`, this.Bindings.addOrUpdateItem.method, converted);
                    await this.populateCommonFields(result, updateResult);
                }
            }
            else {
                const converted = await this.getRestItem(item);
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
    @trace()
    protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        const result = cloneDeep(items);
        const itemsToAdd = result.filter((item) => {
            return item.id < 0;
        });
        const versionedItems = result.filter((item) => {
            return !this.disableVersionCheck && item.version !== undefined && item.version !== null && item.id > 0;
        });
        const updatedItems = result.filter((item) => {
            return (this.disableVersionCheck || item.version === undefined || item.version === null) && item.id > 0;
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
                    const converted = await Promise.all(sub.map(item => this.getRestItem(item)));
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
    @trace()
    protected async deleteItem_Internal(item: T): Promise<T> {
        try {
            await this.executeRequest(`${this.serviceUrl}${this.Bindings.deleteItem.url}/${item.id}`, this.Bindings.deleteItem.method);
            item.deleted = true;
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
    @trace()
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

    @trace()
    protected async persistItemData_internal(data: any, linkedFields?: Array<string>): Promise<T> {
        let result = null;
        if (data) {
            await this.Init();
            const preloaded: IPreloadedData = {};
            result = await this.getItemFromRest(data, preloaded);
            await this.populateLookups([result], false, preloaded, linkedFields);
        }
        return result;
    }

    @trace()
    protected async persistItemsData_internal(data: any[], linkedFields?: Array<string>): Promise<T[]> {
        let result = null;
        if (data) {
            await this.Init();
            const preloaded: IPreloadedData = {};
            result = await Promise.all(data.map(d => this.getItemFromRest(d, preloaded)));
            await this.populateLookups(result, false, preloaded, linkedFields);
        }
        return result;
    }


    @trace()
    public async getByRestQuery(restQuery: IEndPointBinding, data?: any, linkedFields?: Array<string>): Promise<Array<T>> {
        const keyCached = super.hashCode(restQuery).toString() + super.hashCode(data).toString() + super.hashCode(linkedFields).toString();
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
                        const json = await this.executeRequest(restQuery.url, restQuery.method, data);
                        result = await this.persistItemsData_internal(json, linkedFields);

                        //check if data exist for this query in database
                        let mapping = await this.restMappingDb.getItemById(keyCached);
                        if(mapping) {
                            const tmp = await this.dbService.getItemsById(mapping.itemIds);
                            //if data exists trash them 
                            if (tmp && tmp.length > 0) {
                                await this.dbService.deleteItems(tmp);
                            }
                        }
                        if(result && result.length > 0) {
                            const convresult = await Promise.all(result.map((res) => {
                                return this.convertItemToDbFormat(res);
                            }));
                            await this.dbService.addOrUpdateItems(convresult);
                            mapping = new RestResultMapping();
                            mapping.id = keyCached;
                            mapping.itemIds = convresult.map(r => r.id);
                            await this.restMappingDb.addOrUpdateItem(mapping);
                            this.UpdateIdsLastLoad(...convresult.map(e => e.id));
                        }
                        else if(mapping) {
                            await this.restMappingDb.deleteItem(mapping);
                        }                       
                        this.UpdateCacheData(keyCached);
                    }
                    else {
                        const mapping = await this.restMappingDb.getItemById(keyCached);
                        if(mapping && mapping.itemIds && mapping.itemIds.length > 0) {
                            const tmp = await this.dbService.getItemsById(mapping.itemIds);
                            result = await this.mapItems(tmp, linkedFields);
                        }
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


    /************************** Query filters ***************************/

    protected async populateCommonFields(item: T, restItem): Promise<void> {
        if (item.id < 0) {
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
                        if (this.isInitialized) {
                            const users = this.getServiceInitValues(User);
                            user = find(users, (u) => { return u.userPrincipalName?.toLowerCase() === upn?.toLowerCase(); });
                        }
                        else {
                            const userService: UserService = ServiceFactory.getService(User).cast<UserService>();
                            user = new User();
                            user.userPrincipalName = upn;
                            user = await userService.linkToSpUser(user);
                        }
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



    /**
     * convert full item to db format (with links only)
     * @param item - full provisionned item
     */
    protected async convertItemToDbFormat(item: T): Promise<T> {
        const result: T = await super.convertItemToDbFormat(item);
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
                                        if (element.id) {
                                            if ((typeof (element.id) === "number" && element.id > 0) || (typeof (element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
                                                ids.push(element.id);
                                            }
                                        }
                                    });
                                }
                                if(ids.length > 0) {
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

    /**
     * populate item from db storage
     * @param item - db item with links in internalLinks fields
     */
    @trace()
    public async mapItems(items: Array<T>, linkedFields?: Array<string>): Promise<Array<T>> {
        const results: Array<T> = [];
        if (items && items.length > 0) {
            await this.Init();
            for (const item of items) {
                const result: T = cloneDeep(item);
                if (item) {
                    for (const propertyName in this.ItemFields) {
                        if (this.ItemFields.hasOwnProperty(propertyName)) {
                            const fieldDescriptor = this.ItemFields[propertyName];
                            if (//fieldDescriptor.fieldType === FieldType.Lookup ||
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
                            else if (//fieldDescriptor.fieldType === FieldType.LookupMulti ||
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
                            else if(fieldDescriptor.fieldType === FieldType.Json && !stringIsNullOrEmpty(fieldDescriptor.modelName)) {
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
        }
        await this.populateLookups(results, true, {}, linkedFields);
        return results;
    }

    public async updateLinkedTransactions(oldId: number, newId: number, nextTransactions: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
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
                this.transactionService.addOrUpdateItem(transaction);
            }
        });
        return nextTransactions;
    }

    @trace()
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
                    const service = ServiceFactory.getServiceByModelName(modelName);
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

    protected getRestQuery(query: IQuery): IRestQuery {
        const result: IRestQuery = {};
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

    private getOrderBy(orderby: IOrderBy[]): IOrderBy[] {
        const result = [];
        if (orderby) {
            orderby.forEach(ob => {
                const copy = cloneDeep(ob);
                copy.propertyName = this.ItemFields[ob.propertyName].fieldName;
                result.push(copy);
            });
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
    private getRestPredicate(predicate: IPredicate): IRestPredicate {

        return {
            logicalOperator: predicate.operator,
            propertyName: this.ItemFields[predicate.propertyName].fieldName,
            value: predicate.value,
            includeTimeValue: predicate.includeTimeValue,
            lookupId: predicate.lookupId
        };
    }

    private async initRequest(method: string, data?: any): Promise<RequestInit> {
        const aadTokenProvider = await ServicesConfiguration.context.aadTokenProviderFactory.getTokenProvider();
        const token = await aadTokenProvider.getToken(ServicesConfiguration.configuration.aadAppId);
        if (stringIsNullOrEmpty(token)) {
            throw Error("Error while getting authentication token");
        }
        const headers = {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': "*",
            'Access-Control-Allow-Headers': "*",
            'authorization': `Bearer ${token}`
        };
        if (data != null) {
            const postData: string = JSON.stringify(data);
            return {
                method: method,
                body: postData,
                mode: 'cors',
                headers: headers,
                referrer: ServicesConfiguration.context.pageContext.web.absoluteUrl,
                referrerPolicy: "no-referrer-when-downgrade"
            };
        }
        return {
            method: method,
            mode: 'cors',
            headers: headers,
            referrer: ServicesConfiguration.context.pageContext.web.absoluteUrl,
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
