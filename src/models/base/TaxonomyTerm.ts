import { IBaseItem } from "../../interfaces/index";
import { stringIsNullOrEmpty } from "@pnp/common";

/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
export class TaxonomyTerm implements IBaseItem {
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
     * Deprecated
     */
    public isDeprecated: boolean;
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
    constructor(term? : any) {
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.CustomProperties;
            this.isDeprecated = term.IsDeprecated
        }
    }

    public get fullPathString(): string {
        let result = "";
        if(!stringIsNullOrEmpty(this.path)) {
            let parts = this.path.split(";");
            result = parts.join(" > ");
        }
        return result;
    }
}