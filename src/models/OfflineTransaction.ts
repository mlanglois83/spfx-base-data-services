import { TransactionType } from "../constants/index";
import { IBaseItem } from "../interfaces/index";

export class OfflineTransaction implements IBaseItem {
    public id: number;
    public title: TransactionType;
    public serviceName: string;
    public itemType: string;
    public itemData: any;
}