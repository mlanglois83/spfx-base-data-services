import { TransactionType } from "../constants";
import { IBaseItem } from "../interfaces";

/**
 * Offline transaction abstraction class
 */
export class OfflineTransaction implements IBaseItem {
    /**
     * Id of the transaction (auto increment from idb)
     */
    public id: number;
    /**
     * Type of the transaction (see TransactionType)
     */
    public title: TransactionType;
    /**
     * Type name of data item
     */
    public itemType: string;
    /**
     * Data item content (as simple object)
     */
    public itemData: any;
}