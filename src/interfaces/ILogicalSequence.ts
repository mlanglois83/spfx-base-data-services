import { LogicalOperator } from "../constants";
import { IPredicate } from ".";

export interface ILogicalSequence {
    type: "sequence";
    operator: LogicalOperator;
    children: Array<ILogicalSequence | IPredicate>;
}