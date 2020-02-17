/**
 * Interface to describe minimal items manipulated by all data services
 */
export interface IBaseItem {
     /**
     * Item identifier
     */
    id: number | string;
    /**
     * Item title
     */
    title: string;
    /**
     * Queries associated with the item
     */
    queries?: Array<number>;
    /**
     * Item version
     */
    version?: number;
}