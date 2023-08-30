import { LogicalOperator } from "../constants";
import { IBaseItem, IPredicate } from ".";

export interface ILogicalSequence<T extends IBaseItem<string | number>> {
    type: "sequence";
    operator: LogicalOperator;
    children: Array<ILogicalSequence<T> | IPredicate<T, keyof T>>;
}