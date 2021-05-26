import { IOrderBy, LogicalOperator, TestOperator } from "..";
import { IBaseItem } from ".";

export interface IRestQuery<T extends IBaseItem> {
    test?: IRestLogicalSequence<T>;
    orderBy?: Array<IOrderBy<T, keyof T>>;
    limit?: number;
    lastId?: number;
    loadAll?: boolean;
}
export interface IRestLogicalSequence<T extends IBaseItem> {
    logicalOperator?: LogicalOperator;
    predicates?: Array<IRestPredicate<T, keyof T>>;
    sequences?: Array<IRestLogicalSequence<T>>;
}
export interface IRestPredicate<T extends IBaseItem, K extends keyof T> { 
    logicalOperator: TestOperator;
    propertyName: K;
    value: any;
    lookupId?: boolean;
    includeTimeValue?: boolean;
}