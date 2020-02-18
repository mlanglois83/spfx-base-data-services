import { TestOperator } from "../constants";

export interface IPredicate {    
    type: "predicate";
    operator: TestOperator;
    propertyName: string;
    value: any;
    lookupId?: boolean;
    includeTimeValue?: boolean;
}