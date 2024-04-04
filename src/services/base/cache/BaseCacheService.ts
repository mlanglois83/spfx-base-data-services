import { stringIsNullOrEmpty } from "@pnp/core";
import { IBaseItem, IDataService, IQuery } from "../../../interfaces";
import { BaseService } from "../BaseService";
import { ServicesConfiguration } from "../../../configuration";


/**
 * Base classe for indexedDB interraction using SP repository
 */
export abstract class BaseCacheService<T extends IBaseItem<string | number>> extends BaseService implements IDataService<T> {
    protected tableName: string;

    protected itemType: (new (item?: any) => T);
    
    protected get logFormat(): string {
        return "%Time% - [%ClassName%<%Property:itemType.name%> (%Property:tableName%)] --> %Function%: %Duration%ms";
    }

    public get serviceName(): string {
        return this.constructor["name"] + "<" + this.itemType["name"] + ">";
    }

    private internalCacheUrl: string;
    protected get cacheUrl(): string {
        if(stringIsNullOrEmpty(this.internalCacheUrl)) {
            return ServicesConfiguration.serverRelativeUrl;
        }
        else {
            return this.internalCacheUrl;
        }
    }

    /**
     * 
     * @param tableName - name of the db table the service interracts with
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheUrl?: string) {
        super();
        this.tableName = tableName;
        this.internalCacheUrl = cacheUrl;
        this.itemType = type;
    }
    public abstract getAll(): Promise<T[]>;
    public abstract get(query: IQuery<T>): Promise<T[]>;
    public abstract addOrUpdateItem(item: T): Promise<T>;
    public abstract addOrUpdateItems(items: T[], onItemUpdated?: (oldItem: T, newItem: T) => void, onRefreshItems?: (index: number, length: number) => void): Promise<T[]>;
    public abstract deleteItem(item: T): Promise<T>;
    public abstract deleteItems(items: T[]): Promise<T[]>;
    public abstract getItemById(id: string | number): Promise<T>;
    public abstract getItemsById(ids: (string | number)[]): Promise<T[]>;
    public abstract replaceAll(newItems: Array<T>);
}