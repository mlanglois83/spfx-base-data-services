/**
 * Interface to describe minimal items manipulated by all data services
 */
export interface IBaseItem {
     /**
     * Item identifier
     */
    id: number | string;
    /**
     * Item unique identifier
     */
    uniqueId?: string;
    /**
     * Item title
     */
    title?: string;
    /**
     * Item version
     */
    version?: number;
    /**
     * Last update error
     */
    error?: Error;
    /**
     * Item deleted
     */
    deleted?: boolean;
}