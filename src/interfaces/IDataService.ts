import { IBaseItem, IAddOrUpdateResult } from ".";
import { OfflineTransaction } from "..";
/**
 * Contract interface for all dataservices
 */
export interface IDataService<T extends IBaseItem> {
    /**
     * Retrieve all available items
     */
    getAll(): Promise<Array<T>>;
    /**
     * Retrieve items using query
     * @param query query element (ie CAML for SP)
     */
    get(query: any): Promise<Array<T>>;
    /**
     * Adds or updates an item
     * @param item Instance of a Model that has to be sent
     */
    addOrUpdateItem(item: T): Promise<IAddOrUpdateResult<T>>;
    /**
     * Removes an item
     * @param item Instance of a Model that has to deleted
     */
    deleteItem(item: T): Promise<void>;
    /**
     * Retrieve item by id
     * @param id item id
     */
    getItemById(id: string | number): Promise<T>;
    /**
     * Retrieve items by ids
     * @param ids Array of ids
     */
    getItemsById(ids: Array<string | number>): Promise<Array<T>>;
}