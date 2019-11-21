import { IBaseItem } from "../../interfaces/index";
/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
export declare class TaxonomyTerm implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    __internalLinks: any;
    /**
     * WssIds assiciated with term
     */
    wssids: Array<number>;
    /**
     * Term id (Guid)
     */
    id: string;
    /**
     * Term label
     */
    title: string;
    /**
     * Full path of term
     */
    path: string;
    /**
     * Term custom sort order
     */
    customSortOrder?: string;
    /**
     * Term associated custom properties
     */
    customProperties: any;
    /**
     * Instanciates a term object
     * @param term term object from rest call
     */
    constructor(term: any);
}
