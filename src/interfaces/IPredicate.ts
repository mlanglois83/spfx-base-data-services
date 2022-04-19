import { IBaseItem } from ".";
import { TestOperator } from "../constants";

export interface IPredicate<T extends IBaseItem, K extends keyof T> {    
    type: "predicate";
    operator: TestOperator;
    propertyName: K;
    value?: any;
    lookupId?: boolean;
    includeTimeValue?: boolean;
}