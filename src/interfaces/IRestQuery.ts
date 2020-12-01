import { IOrderBy, LogicalOperator, TestOperator } from "..";

export interface IRestQuery {
    test?: IRestLogicalSequence;
    orderBy?: Array<IOrderBy>;
    limit?: number;
    lastId?: number;
    loadAll?: boolean;
}
export interface IRestLogicalSequence {
    logicalOperator?: LogicalOperator;
    predicates?: Array<IRestPredicate>;
    sequences?: Array<IRestLogicalSequence>;
}
export interface IRestPredicate { 
    logicalOperator: TestOperator;
    propertyName: string;
    value: any;
    lookupId?: boolean;
    includeTimeValue?: boolean;
}