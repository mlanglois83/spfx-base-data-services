import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction } from "../../models/index";
import { IAddOrUpdateResult } from "../../interfaces";
export declare class TransactionService extends BaseDbService<OfflineTransaction> {
    private transactionFileService;
    constructor();
    /**
     * Add or update an item in DB and returns updated item
     * @param item Item to add or update
     */
    addOrUpdateItem(item: OfflineTransaction): Promise<IAddOrUpdateResult<OfflineTransaction>>;
    deleteItem(item: OfflineTransaction): Promise<void>;
    /**
     * add items in table (ids updated)
     * @param newItems
     */
    addOrUpdateItems(newItems: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>>;
    /**
     * Retrieve all items from db table
     */
    getAll(): Promise<Array<OfflineTransaction>>;
    /**
     * Get a transaction given its id
     * @param id transaction id
     */
    getItemById(id: number): Promise<OfflineTransaction>;
    /**
     * Clear table
     */
    clear(): Promise<void>;
    private isFile;
}
