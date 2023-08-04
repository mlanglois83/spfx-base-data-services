/**
 * Interface to describe minimal items manipulated by all data services
 */
export interface IBaseItem<T extends string | number> {
    defaultKey: T;
    isCreatedOffline?: boolean;
     /**
     * Item identifier
     */
    id: T;
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

    fromObject: (obj) => void;
}