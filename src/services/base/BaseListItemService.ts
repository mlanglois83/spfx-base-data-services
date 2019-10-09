import { ServicesConfiguration } from "../..";
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { CamlQuery, List, sp } from "@pnp/sp";
import { Constants } from "../../constants/index";
import { IBaseItem } from "../../interfaces/index";
import { BaseDataService } from "./BaseDataService";
import { BaseService } from "./BaseService";
import { UtilsService } from "..";

/**
 * 
 * Base service for sp list items operations
 */
export class BaseListItemService<T extends IBaseItem> extends BaseDataService<T>{
    protected itemType: (new (item?: any) => T);
    protected listRelativeUrl: string;

    public get listItemType(): (new (item?: any) => T) {
        return this.itemType;
    }

    /**
     * Associeted list (pnpjs)
     */
    protected get list(): List {
        return sp.web.getList(this.listRelativeUrl);
    }

    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param listRelativeUrl list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, tableName: string, cacheDuration?: number) {
        super(type, tableName, cacheDuration);
        this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
        this.itemType = type;

    }


    /**
     * Cache has to be relaod ?
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

    /**
     *
     * TODO avoid getting all fields
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    protected async get_Internal(query: any): Promise<Array<T>> {
        let results = new Array<T>();

        let items = await this.list.getItemsByCAMLQuery({
            ViewXml: '<View Scope="RecursiveAll"><Query>' + query + '</Query></View>'
        } as CamlQuery, 'FieldValuesAsText');

        return items.map(r => { return new this.itemType(r); });
    }




    /**
     * 
     * @param id 
     */
    protected async getById_Internal(id: number): Promise<T> {
        let result = null;
        let temp = await this.list.items.getById(id).get();

        if (temp) {
            result = new this.itemType(temp);
        }

        return result;
    }

    /**
     * Retrieve all items
     * 
     * TODO avoid getting all fields
     */
    protected async getAll_Internal(): Promise<Array<T>> {

        let items = await this.list.items.getAll();
        return items.map(r => { return new this.itemType(r); });
    }

    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        let result = cloneDeep(item);
        if (item.id < 0) {
            let addResult = await this.list.items.add(item.convert());

            if (result.onAddCompleted) {
                result.onAddCompleted(addResult.data);

            }
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
                    await this.list.items.getById(<number>item.id).update(item.convert());
                    let version = await this.list.items.getById(<number>item.id).get();
                    if (result.onUpdateCompleted) {
                        result.onUpdateCompleted(version);
                    }
                }
            }
            else {
                let updateResult = await this.list.items.getById(<number>item.id).update(item.convert());

                if (result.onUpdateCompleted) {
                    result.onUpdateCompleted(updateResult.data);
                }
            }
        }
        return result;
    }

    protected async deleteItem_Internal(item: T): Promise<void> {
        await this.list.items.getById(<number>item.id).delete();
    }
}
