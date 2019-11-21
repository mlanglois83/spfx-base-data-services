import { IBaseItem } from "../../interfaces/index";

/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
export class TaxonomyTerm implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks: any = {};
    /**
     * WssIds assiciated with term
     */
    public wssids: Array<number>;
    /**
     * Term id (Guid)
     */
    public id: string;
    /**
     * Term label
     */
    public title: string;
    /**
     * Full path of term
     */
    public path: string;
    /**
     * Term custom sort order
     */
    public customSortOrder?: string;
    /**
     * Term associated custom properties
     */
    public customProperties: any;
    /**
     * Instanciates a term object
     * @param term term object from rest call
     */
    constructor(term: any) {
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.CustomProperties;
        }
    }
}