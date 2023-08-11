import { ILogicalSequence } from "./ILogicalSequence";
import { IPredicate } from "./IPredicate";
import { IOrderBy } from "./IOrderBy";
import { IBaseItem } from ".";

export interface IQuery<T extends IBaseItem<string | number>> {
    test?: ILogicalSequence<T> | IPredicate<T, keyof T>;
    orderBy?: Array<IOrderBy<T, keyof T>>;
    limit?: number;
    lastId?: number | Text;
}