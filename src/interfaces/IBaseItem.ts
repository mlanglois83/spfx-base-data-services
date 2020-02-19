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
     * Item version
     */
    version?: number;
}