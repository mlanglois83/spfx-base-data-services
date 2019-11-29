import { TransactionType } from "../constants/index";
import { IBaseItem } from "../interfaces/index";

/**
 * Offline transaction abstraction class
 */
export class OfflineTransaction implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks: any = undefined;
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