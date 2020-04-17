import { ILogicalSequence } from "./ILogicalSequence";
import { IPredicate } from "./IPredicate";
import { IOrderBy } from "./IOrderBy";

export interface IQuery {
    test?: ILogicalSequence | IPredicate;

    orderBy?: Array<IOrderBy>;
    limit?: number;
    lastId?: number | Text;
}