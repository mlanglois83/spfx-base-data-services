import { IBaseItem } from "./IBaseItem";
import { TransactionType } from "../constants";

export interface IItemSynchronized {
    oldId?: string|number;
    item: IBaseItem;
    operation: TransactionType;
}