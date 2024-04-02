import { assign, cloneDeep } from "lodash";
import { ServicesConfiguration } from "../../../configuration";
import { IBaseFile, IBaseItem, IQuery } from "../../../interfaces";
import { UtilsService } from "../../UtilsService";
import { BaseCacheService } from "./BaseCacheService";

import { Constants, TraceLevel } from "../../../constants";
import { Decorators } from "../../../decorators";

const trace = Decorators.trace;

/**
 * Base classe for local storage cache interraction using SP repository
 */
export class BaseLocalStorageService<T extends IBaseItem<string | number>> extends BaseCacheService<T> {
    

    private get cacheKey() {
        return UtilsService.formatText(Constants.cacheKeys.localStorageTableFormat, ServicesConfiguration.configuration.serviceKey, ServicesConfiguration.baseUrl, this.tableName);
    }

    protected async getAllKeysInternal<TKey extends number | string>(): Promise<Array<TKey>> {
        const allitems = await this.getAll();
        return allitems.map(i => i.id as TKey);        
    }

    protected async getNextAvailableKey(): Promise<string | number> {
        let result: string | number;
        const tmp = new this.itemType();
        if (typeof (tmp.id) === "number") {
            const keys = await this.getAllKeysInternal() as number[];
            if (keys.length > 0) {
                const minKey = Math.min(...keys);
                result = Math.min(-1, minKey - 1);
            }
            else {
                result = -1;
            }
        }
        else {
            return Constants.models.offlineCreatedPrefix +  UtilsService.generateGuid();
        }
        return result;
    }


    /**
     * Add or update an item in DB and returns updated item
     * @param item - item to add or update
     */
    @trace(TraceLevel.DataBase)
    public async addOrUpdateItem(item: T): Promise<T> {        
        try {
            if (item.id === item.defaultKey) {
                const nextid = await this.getNextAvailableKey();
                item.id = nextid;
            }
            const allitems = await this.getAll();
            const existingIdx = allitems.findIndex(i => i.id === item.id);
            if(existingIdx !== -1) {
                allitems[existingIdx] = item;
            }
            else {
                allitems.push(item);
            }
            localStorage.setItem(this.cacheKey, JSON.stringify(allitems));
            return item;

        } catch (error) {
            console.error(error.message + " - " + error.Name);            
            item.error = error;
            return item;
        }
    }

    @trace(TraceLevel.DataBase)
    public async deleteItem(item: T): Promise<T> {    
        try {
            const allitems = await this.getAll();
            const existingIdx = allitems.findIndex(i => i.id === item.id);
            if(existingIdx !== -1) {
                allitems.splice(existingIdx, 1);
                item.deleted = true;
            }
            localStorage.setItem(this.cacheKey, JSON.stringify(allitems));
        } catch (error) {
            console.error(error.message + " - " + error.Name);            
            throw error;
        }
        return item;
    }

    @trace(TraceLevel.DataBase)
    public async deleteItems(items: Array<T>): Promise<Array<T>> {
        try {
            const allitems = await this.getAll();
            for (const item of items) {
                const existingIdx = allitems.findIndex(i => i.id === item.id);
                if(existingIdx !== -1) {
                    allitems.splice(existingIdx, 1);
                    item.deleted = true;
                }
            }            
            localStorage.setItem(this.cacheKey, JSON.stringify(allitems));
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            throw error;
        }
        return items;
    }


    @trace(TraceLevel.DataBase)
    public async get(query: IQuery<T>): Promise<Array<T>> { // eslint-disable-line @typescript-eslint/no-unused-vars
        const items = await this.getAll();
        return items;
    }


    /**
     * add items in table (ids updated)
     * @param newItems - items to add or update
     */
    @trace(TraceLevel.DataBase)
    public async addOrUpdateItems(newItems: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        let nextid = undefined;
        const copy = cloneDeep(newItems);
        try {            
            const allitems = await this.getAll();
            await Promise.all(copy.map(async (item, itemIdx) => {
                if (typeof (item.typedKey) === "number" && item.id === item.defaultKey) {
                    if (nextid === undefined) {
                        nextid = await this.getNextAvailableKey();
                    }
                    (item as IBaseItem<number>).id = nextid--;
                }
                else if(typeof (item.typedKey) === "string" && item.id === item.defaultKey) {
                    item.id = await this.getNextAvailableKey();
                }
                const existingIdx = allitems.findIndex(i => i.id === item.id);
                if(existingIdx !== -1) {
                    allitems[existingIdx] = item;
                }
                else {
                    allitems.push(item);
                }
                if (onItemUpdated) {
                    onItemUpdated(newItems[itemIdx], item);
                }
            }));
            localStorage.setItem(this.cacheKey, JSON.stringify(allitems));
            return copy;
        } catch (error) {
            console.error(error.message + " - " + error.Name);            
            throw error;
        }
    }


    /**
     * Retrieve all items from db table
     */
    @trace(TraceLevel.DataBase)
    public async getAll(): Promise<Array<T>> {
        const result = new Array<T>();   
        try {
            const stringValue = localStorage.getItem(this.cacheKey);
            let rows: (IBaseItem<string | number> | IBaseFile<string | number>)[];
            try {
                rows = JSON.parse(stringValue);
                if(!Array.isArray(rows)) {
                    rows = [];
                }
            }
            catch(error) {
                console.warn(error);
                rows = [];
            }
            rows.forEach((r) => {
                const item = new this.itemType();
                const resultItem = assign(item, r);
                result.push(resultItem);
            });
            return result;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            throw error;
        }
    }



    

    /**
     * Clear table
     */
    @trace(TraceLevel.DataBase)
    public async clear(): Promise<void> {
        try {
            localStorage.removeItem(this.cacheKey);
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            throw error;
        }
    }

    /**
     * Clear table and insert new items
     * @param newItems - items to insert in place of existing
     */
    @trace(TraceLevel.DataBase)
    public async replaceAll(newItems: Array<T>): Promise<void> {
        await this.clear();
        await this.addOrUpdateItems(newItems);
    }

    @trace(TraceLevel.DataBase)
    public async getItemById(id: number | string): Promise<T> {
        let result: T = null;    
        try {
            const all = await this.getAll();
            result = all.find(o => o.id === id);
            return result;
        } catch (error) {
            // key not found
            return null;
        }
    }

    @trace(TraceLevel.DataBase)
    public async getItemsById(ids: Array<number | string>): Promise<Array<T>> {
        const results: T[] = [];
        try {
            const all = await this.getAll();
            ids.forEach(id => {
                const obj = all.find(o => o.id === id);
                if(obj) {
                    results.push(obj);
                }
            });            
            return results;
        } catch (error) {
            // key not found
            return [];
        }
    }
}