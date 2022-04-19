
import { IBaseItem, IQuery } from ".";
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
     * @param query - query element
     */
    get(query: IQuery<T>): Promise<Array<T>>;
    /**
     * Adds or updates an item
     * @param item - instance of a Model that has to be sent
     */
    addOrUpdateItem(item: T): Promise<T>;
    /**
     * Adds or updates items
     * @param items - instances of a Model that has to be sent
     * @param onItemUpdated - function called when an item has benn updated
     */
    addOrUpdateItems(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>>;
    /**
     * Removes an item
     * @param item - instance of a Model that has to deleted
     */
    deleteItem(item: T): Promise<T>;
    /**
     * Removes an item
     * @param item - instances of a Model that has to deleted
     */
    deleteItems(items: Array<T>): Promise<Array<T>>;
    /**
     * Retrieve item by id
     * @param id - item id
     */
    getItemById(id: string | number): Promise<T>;
    /**
     * Retrieve items by ids
     * @param ids - array of ids
     */
    getItemsById(ids: Array<string | number>): Promise<Array<T>>;
}