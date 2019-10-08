import { IBaseItem } from ".";
export interface IAddOrUpdateResult<T extends IBaseItem> {
    item: T;
    error?: Error;
}
