import { TransactionType } from "../constants/index";
import { IBaseItem } from "../interfaces/index";
/**
 * Offline transaction abstraction class
 */
export declare class OfflineTransaction implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    __internalLinks: any;
    /**
     * Id of the transaction (auto increment from idb)
     */
    id: number;
    /**
     * Type of the transaction (see TransactionType)
     */
    title: TransactionType;
    /**
     * Name of the service that has to perform the operation
     */
    serviceName: string;
    /**
     * Type name of data item
     */
    itemType: string;
    /**
     * Data item content (as simple object)
     */
    itemData: any;
}
