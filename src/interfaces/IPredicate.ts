import { IBaseItem } from ".";
import { TestOperator } from "../constants";

export interface IPredicate<T extends IBaseItem<string | number>, K extends keyof T> {    
    type: "predicate";
    operator: TestOperator;
    propertyName: K;
    value?: any;
    lookupId?: boolean;
    includeTimeValue?: boolean;
}