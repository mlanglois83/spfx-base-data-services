import { IBaseItem } from ".";
import { TransactionType } from "../constants";

export interface IItemSynchronized {
    oldId?: string|number;
    item: IBaseItem<string | number>;
    operation: TransactionType;
}