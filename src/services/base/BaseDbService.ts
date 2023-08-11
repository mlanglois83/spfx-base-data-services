import { assign, cloneDeep } from "lodash";
import {  DB, ObjectStore, openDb } from "idb";
import { IBaseItem, IDataService, IQuery } from "../../interfaces";
import { BaseService } from "./BaseService";
import { UtilsService } from "../UtilsService";
import { ServicesConfiguration } from "../../configuration";

import { Mutex, Semaphore } from 'async-mutex';
import { BaseFile } from "../../models";
import { Decorators } from "../../decorators";
import { Constants, TraceLevel } from "../../constants";

const trace = Decorators.trace;

/**
 * Base classe for indexedDB interraction using SP repository
 */
export class BaseDbService<T extends IBaseItem<string | number>> extends BaseService implements IDataService<T> {
    protected tableName: string;

    protected itemType: (new (item?: any) => T);
    protected static openPromise: Promise<void> = undefined;
    protected static db: DB;
    protected static semaphore: Semaphore = undefined;

    protected static _maxSimultaneousDbAccess: number = undefined;
    public static set maxSimultaneousDbAccess(value: number) {
        if (BaseDbService._maxSimultaneousDbAccess !== value) {
            BaseDbService._maxSimultaneousDbAccess = value;
            if (value > 0) {
                BaseDbService.semaphore = new Semaphore(value);
            }
            else {
                BaseDbService.semaphore = undefined;
            }
        }
    }


    protected get hasSemaphore(): boolean {
        if (ServicesConfiguration.configuration.maxSimultaneousDbAccess !== BaseDbService._maxSimultaneousDbAccess) {
            BaseDbService.maxSimultaneousDbAccess = ServicesConfiguration.configuration.maxSimultaneousDbAccess;
        }
        return BaseDbService.semaphore !== undefined;
    }

    protected static async acquire(): Promise<() => void> {
        let func: () => void = (): void => { return; };
        if (ServicesConfiguration.configuration.maxSimultaneousDbAccess !== BaseDbService._maxSimultaneousDbAccess) {
            BaseDbService.maxSimultaneousDbAccess = ServicesConfiguration.configuration.maxSimultaneousDbAccess;
        }
        if (BaseDbService.semaphore) {
            const [, release] = await BaseDbService.semaphore.acquire();
            func = release;
        }
                return func;
    }


    protected get logFormat(): string {
        return "%Time% - [%ClassName%<%Property:itemType.name%> (%Property:tableName%)] --> %Function%: %Duration%ms";
    }




    public get serviceName(): string {
        return this.constructor["name"] + "<" + this.itemType["name"] + ">";
    }

    /**
     * 
     * @param tableName - name of the db table the service interracts with
     */
    constructor(type: (new (item?: any) => T), tableName: string) {
        super();
        this.tableName = tableName;
        //BaseDbService.db = null;
        this.itemType = type;
    }

    protected getChunksRegexp(fileId: number | string): RegExp {
        const escapedUrl = UtilsService.escapeRegExp(fileId.toString());
        return new RegExp("^" + escapedUrl + "_chunk_\\d+$", "g");
    }

    protected async getAllKeysInternal<TKey extends number | string>(store: ObjectStore<T, TKey>): Promise<Array<TKey>> {
        let result: Array<TKey> = [];
        if (store.getAllKeys) {
            result = await store.getAllKeys();
        }
        else {
            let cursor = await store.openCursor();
            while (cursor) {
                result.push(cursor.primaryKey);
                cursor = await cursor.continue();
            }
        }
        return result;
    }

    private static mutex = new Mutex();

    protected async getNextAvailableKey(store: ObjectStore<T, string | number>): Promise<string | number> {
        let result: string | number;
        const tmp = new this.itemType();
        if (typeof (tmp.id) === "number") {
            const release = await BaseDbService.mutex.acquire();
            try {
                const keys = await (this as BaseDbService<IBaseItem<number>>).getAllKeysInternal(store as ObjectStore<IBaseItem<number>, number>);
                if (keys.length > 0) {
                    const minKey = Math.min(...keys);
                    result = Math.min(-1, minKey - 1);
                }
                else {
                    result = -1;
                }
            } finally {
                release();
            }
        }
        else {
            return Constants.models.offlineCreatedPrefix +  UtilsService.generateGuid();
        }
        return result;
    }

    protected get dbInitialized(): boolean {
        return BaseDbService.db !== undefined;
    }

    /**
     * Opens indexed db, update structure if needed
     */
    @trace(TraceLevel.DataBase)
    protected async OpenDb(): Promise<void> {
        if (!this.dbInitialized) {
            if (BaseDbService.openPromise) {
                if (this.debug)
                    console.log("BaseDbService openDb : load allready called before, sharing promise");
            }
            else {
                BaseDbService.openPromise = new Promise<void>(async (resolve, reject) => {
                    console.warn("Open database");
                    if (!('indexedDB' in window)) {
                        throw new Error(ServicesConfiguration.configuration.translations.IndexedDBNotDefined);
                    }
                    const dbName = UtilsService.formatText(ServicesConfiguration.configuration.dbName, ServicesConfiguration.serverRelativeUrl);
                    try {
                        BaseDbService.db = await openDb(dbName, ServicesConfiguration.configuration.dbVersion, (UpgradeDB) => {
                            // remove old tables
                            for (let index = 0; index < UpgradeDB.objectStoreNames.length; index++) {
                                const element = UpgradeDB.objectStoreNames.item(index);
                                if (ServicesConfiguration.configuration.tableNames.indexOf(element) === -1) {
                                    UpgradeDB.deleteObjectStore(element);
                                }
                            }
                            // add new tables
                            for (const tableName of ServicesConfiguration.configuration.tableNames) {
                                if (!UpgradeDB.objectStoreNames.contains(tableName)) {
                                    UpgradeDB.createObjectStore(tableName, { keyPath: 'id', autoIncrement: tableName === "OfflineTransaction" });
                                }
                            }
                        });

                        resolve();
                    }
                    catch (error) {
                        reject(error);
                    }
                });
                BaseDbService.openPromise.then(() => {
                    BaseDbService.openPromise = null;
                }).catch(() => {
                    BaseDbService.openPromise = null;
                });
            }
            return BaseDbService.openPromise;
        }
    }

    /**
     * Add or update an item in DB and returns updated item
     * @param item - item to add or update
     */
    @trace(TraceLevel.DataBase)
    public async addOrUpdateItem(item: T): Promise<T> {
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }

        const tx = BaseDbService.db.transaction(this.tableName, 'readwrite');
        const store = tx.objectStore(this.tableName);
        try {
            if (item.id === item.defaultKey && !store.autoIncrement) {
                const nextid = await this.getNextAvailableKey(store as ObjectStore<T, string | number>);
                item.id = nextid;
            }
            if (item instanceof BaseFile && item.content && item.content.byteLength >= 10485760) {
                // remove existing chunks
                const keys = await this.getAllKeysInternal(store);
                const chunkRegex = this.getChunksRegexp(item.id);
                const chunkkeys = keys.filter((k) => {
                    const match = k.toString().match(chunkRegex);
                    return match && match.length > 0;
                });
                await Promise.all(chunkkeys.map((k) => {
                    return store.delete(k);
                }));
                // add chunked file
                let idx = 0;
                let size = 0;
                while (size < item.content.byteLength) {
                    const firstidx = idx * 10485760;
                    const lastidx = Math.min(item.content.byteLength, firstidx + 10485760);
                    const chunk = item.content.slice(firstidx, lastidx);
                    // create file object
                    const chunkitem = new this.itemType() as unknown as BaseFile<string>;
                    chunkitem.id = item.id.toString() + (idx === 0 ? "" : "_chunk_" + idx);
                    chunkitem.title = item.title;
                    chunkitem.mimeType = item.mimeType;
                    chunkitem.content = chunk;
                    await store.put(assign({}, chunkitem));
                    idx++;
                    size += chunk.byteLength;
                }

            }
            else {
                await store.put(assign({}, item)); // store simple object with data only 
            }
            await tx.complete;
            return item;

        } catch (error) {
            console.error(error.message + " - " + error.Name);
            try {
                tx.abort();
            } catch {
                // error allready thrown
            }
            item.error = error;
            return item;
        }
        finally {
            release();
        }
    }

    @trace(TraceLevel.DataBase)
    public async deleteItem(item: T): Promise<T> {
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const tx = BaseDbService.db.transaction(this.tableName, 'readwrite');
        const store = tx.objectStore(this.tableName);
        try {
            const deleteKeys = [item.id];
            if (item instanceof BaseFile) {
                const keys = await this.getAllKeysInternal(store);
                const chunkRegex = this.getChunksRegexp(item.id);
                const chunkkeys = keys.filter((k) => {
                    const match = k.toString().match(chunkRegex);
                    return match && match.length > 0;
                });
                deleteKeys.push(...chunkkeys);
            }
            await Promise.all(deleteKeys.map(async (k) => {
                await store.delete(k);
                item.deleted = true;
            }));
            await tx.complete;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            try {
                tx.abort();
            } catch {
                // error allready thrown
            }
            throw error;
        }
        finally {
            release();
        }
        return item;
    }

    @trace(TraceLevel.DataBase)
    public async deleteItems(items: Array<T>): Promise<Array<T>> {
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const tx = BaseDbService.db.transaction(this.tableName, 'readwrite');
        const store = tx.objectStore(this.tableName);
        try {
            for (const item of items) {
                const deleteKeys = [item.id];
                if (item instanceof BaseFile) {
                    const keys = await this.getAllKeysInternal(store);
                    const chunkRegex = this.getChunksRegexp(item.id);
                    const chunkkeys = keys.filter((k) => {
                        const match = k.toString().match(chunkRegex);
                        return match && match.length > 0;
                    });
                    deleteKeys.push(...chunkkeys);
                }
                await Promise.all(deleteKeys.map(async (k) => {
                    await store.delete(k);
                    item.deleted = true;
                }));
            }
            await tx.complete;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            try {
                tx.abort();
            } catch {
                // error allready thrown
            }
            throw error;
        }
        finally {
            release();
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
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        let nextid = undefined;
        const tx = BaseDbService.db.transaction(this.tableName, 'readwrite');
        const store = tx.objectStore(this.tableName);
        const copy = cloneDeep(newItems);
        try {
            await Promise.all(copy.map(async (item, itemIdx) => {
                if (typeof (item.id) === "number" && !store.autoIncrement && item.id === -1) {
                    if (nextid === undefined) {
                        nextid = await this.getNextAvailableKey(store as ObjectStore<T, number>);
                    }
                    (item as IBaseItem<number>).id = nextid--;
                }
                if (item instanceof BaseFile && item.content && item.content.byteLength >= 10485760) {
                    // remove existing chunks
                    const keys = await this.getAllKeysInternal(store);
                    const chunkRegex = this.getChunksRegexp(item.id);
                    const chunkkeys = keys.filter((k) => {
                        const match = k.toString().match(chunkRegex);
                        return match && match.length > 0;
                    });
                    await Promise.all(chunkkeys.map((k) => {
                        return store.delete(k);
                    }));
                    // add chunked file
                    let idx = 0;
                    let size = 0;
                    while (size < item.content.byteLength) {
                        const firstidx = idx * 10485760;
                        const lastidx = Math.min(item.content.byteLength, firstidx + 10485760);
                        const chunk = item.content.slice(firstidx, lastidx);
                        // create file object
                        const chunkitem = new this.itemType() as unknown as BaseFile<string>;
                        chunkitem.id = item.id.toString() + (idx === 0 ? "" : "_chunk_" + idx);
                        chunkitem.title = item.title;
                        chunkitem.mimeType = item.mimeType;
                        chunkitem.content = chunk;
                        await store.put(assign({}, chunkitem));
                        idx++;
                        size += chunk.byteLength;
                    }

                }
                else {
                    await store.put(assign({}, item)); // store simple object with data only 
                }
                if (onItemUpdated) {
                    onItemUpdated(newItems[itemIdx], item);
                }
            }));
            await tx.complete;

            return copy;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            try {
                tx.abort();
            } catch {
                // error allready thrown
            }
            throw error;
        }
        finally {
            release();
        }
    }


    /**
     * Retrieve all items from db table
     */
    @trace(TraceLevel.DataBase)
    public async getAll(): Promise<Array<T>> {
        const result = new Array<T>();
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const transaction = BaseDbService.db.transaction(this.tableName, 'readonly');
        const store = transaction.objectStore(this.tableName);
        try {
            const rows = await store.getAll();
            rows.forEach((r) => {
                const item = new this.itemType();
                const resultItem = assign(item, r);
                if (item instanceof BaseFile) {
                    // item is a part of another file
                    const chunkparts = (/^.*_chunk_\d+$/g).test(item.id.toString());
                    if (!chunkparts) {
                        // verify if there are other parts
                        const chunkRegex = this.getChunksRegexp(item.id);
                        const chunks = rows.filter((chunkedrow) => {
                            const match = chunkedrow.id.match(chunkRegex);
                            return match && match.length > 0;
                        });
                        if (chunks.length > 0) {
                            chunks.sort((a, b) => {
                                return parseInt(a.id.replace(/^.*_chunk_(\d+)$/g, "$1")) - parseInt(b.id.replace(/^.*_chunk_(\d+)$/g, "$1"));
                            });
                            resultItem.content = UtilsService.concatArrayBuffers(resultItem.content, ...chunks.map(c => {
                                const file = assign(new this.itemType(), c);
                                return file.content;
                            }));
                        }
                        result.push(resultItem);
                    }
                }
                else {
                    result.push(resultItem);
                }
            });
            //await transaction.complete;
            return result;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            throw error;
        }
        finally {
            release();
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

    /**
     * Clear table
     */
    @trace(TraceLevel.DataBase)
    public async clear(): Promise<void> {
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const tx = BaseDbService.db.transaction(this.tableName, 'readwrite');
        const store = tx.objectStore(this.tableName);
        try {
            await store.clear();
            await tx.complete;
        } catch (error) {
            console.error(error.message + " - " + error.Name);
            try {
                tx.abort();
            } catch {
                // error allready thrown
            }
            throw error;
        }
        finally {
            release();
        }
    }

    @trace(TraceLevel.DataBase)
    public async getItemById(id: number | string): Promise<T> {
        let result: T = null;
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const tx = BaseDbService.db.transaction(this.tableName, 'readonly');
        const store = tx.objectStore(this.tableName);
        try {
            const obj = await store.get(id);
            if (obj) {
                result = assign(new this.itemType(), obj);
                if (result instanceof BaseFile) {
                    // item is a part of another file
                    const chunkparts = (/^.*_chunk_\d+$/g).test(result.id.toString());
                    if (!chunkparts) {
                        // verify if there are other parts
                        const keys = await this.getAllKeysInternal(store);
                        const chunkRegex = this.getChunksRegexp(result.id);
                        const chunkkeys = keys.filter((k) => {
                            const match = k.toString().match(chunkRegex);
                            return match && match.length > 0;
                        });
                        const chunks = await Promise.all(chunkkeys.map((key) => store.get(key)));
                        await Promise.all(chunkkeys.map((k) => {
                            return store.delete(k);
                        }));


                        if (chunks.length > 0) {
                            chunks.sort((a: any, b: any) => {
                                return parseInt(a.id.replace(/^.*_chunk_(\d+)$/g, "$1")) - parseInt(b.id.replace(/^.*_chunk_(\d+)$/g, "$1"));
                            });
                            result.content = UtilsService.concatArrayBuffers(result.content, ...chunks.map(c => {
                                const file: BaseFile<string | number> = assign(new this.itemType(), c);
                                return file.content;
                            }));
                        }
                    }
                    else {
                        // no chunked parts here
                        result = null;
                    }
                }
            }
            return result;
        } catch (error) {
            // key not found
            return null;
        }
        finally {
            release();
        }
    }

    @trace(TraceLevel.DataBase)
    public async getItemsById(ids: Array<number | string>): Promise<Array<T>> {
        const results: T[] = [];
        if(!this.dbInitialized) {
            await this.OpenDb();
        }
        let release = (): void => { return; };
        if(this.hasSemaphore) {
            release = await BaseDbService.acquire();
        }
        const tx = BaseDbService.db.transaction(this.tableName, 'readonly');
        const store = tx.objectStore(this.tableName);
        try {
            await Promise.all(ids.map(async (id) => {
                let result = null;
                const obj = await store.get(id);
                if (obj) {
                    result = assign(new this.itemType(), obj);
                    if (result instanceof BaseFile) {
                        // item is a part of another file
                        const chunkparts = (/^.*_chunk_\d+$/g).test(result.id.toString());
                        if (!chunkparts) {
                            // verify if there are other parts
                            const keys = await this.getAllKeysInternal(store);
                            const chunkRegex = this.getChunksRegexp(result.id);
                            const chunkkeys = keys.filter((k) => {
                                const match = k.toString().match(chunkRegex);
                                return match && match.length > 0;
                            });
                            const chunks = await Promise.all(chunkkeys.map((key) => store.get(key)));
                            if (chunks.length > 0) {
                                chunks.sort((a: any, b: any) => {
                                    return parseInt(a.id.replace(/^.*_chunk_(\d+)$/g, "$1")) - parseInt(b.id.replace(/^.*_chunk_(\d+)$/g, "$1"));
                                });
                                result.content = UtilsService.concatArrayBuffers(result.content, ...chunks.map(c => {
                                    const file = assign(new this.itemType(), c);
                                    return file.content;
                                }));
                            }
                        }
                        else {
                            // no chunked parts here
                            result = null;
                        }
                    }
                }
                if (result) {
                    results.push(result);
                }
            }));
            return results;
        } catch (error) {
            // key not found
            return [];
        }
        finally {
            release();
        }
    }
}