import { TransactionType } from "../constants/index";
import { IBaseItem } from "../interfaces/index";

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
     * Name of the service that has to perform the operation
     */
    public serviceName: string;
    /**
     * Type name of data item
     */
    public itemType: string;
    /**
     * Data item content (as simple object)
     */
    public itemData: any;
}