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
    get(query: any): Promise<Array<T>>;
    addOrUpdateItem(item: T): Promise<IAddOrUpdateResult<T>>;
    deleteItem(item: T): Promise<void>;updateLinkedItems?: (oldId: number | string, newId: number | string, transactions: Array<OfflineTransaction>) => Promise<Array<OfflineTransaction>>;
}