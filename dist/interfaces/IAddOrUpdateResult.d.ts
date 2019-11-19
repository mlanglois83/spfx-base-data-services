import { IBaseItem } from ".";
/**
 * Interface to describe add or update operation result
 */
export interface IAddOrUpdateResult<T extends IBaseItem> {
    /**
     * Result item
     */
    item: T;
    /**
     * Error if an error occured, undefined else
     */
    error?: Error;
}
