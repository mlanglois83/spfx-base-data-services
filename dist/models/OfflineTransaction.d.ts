import { TransactionType } from "../constants/index";
import { IBaseItem } from "../interfaces/index";
export declare class OfflineTransaction implements IBaseItem {
    id: number;
    title: TransactionType;
    serviceName: string;
    itemType: string;
    itemData: any;
}
