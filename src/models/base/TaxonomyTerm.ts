import { IBaseItem } from "../../interfaces/index";
import { stringIsNullOrEmpty } from "@pnp/common";
import { assign } from "@microsoft/sp-lodash-subset";

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
     * Term description
     */
    public description: string;
    /**
     * Term associated custom properties
     */
    public customProperties: any;
    /**
     * Instanciates a term object
     * @param - term term object from rest call
     */
    constructor(term? : any) {
        if (term != undefined) {
            this.title = term.Name ? term.Name : "";
            this.description = term.Description ? term.Description : "";
            this.id = term.Id ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.CustomProperties;
            this.isDeprecated = term.IsDeprecated;
        }
    }
    public isParentOf(term: TaxonomyTerm): boolean {
        return (
            term && 
            !stringIsNullOrEmpty(this.path) && 
            !stringIsNullOrEmpty(term.path) &&
            this.path.split(";").length + 1 === term.path.split(";").length &&
            term.path.indexOf(this.path + ";") === 0
        );
    }
    public contains(term: TaxonomyTerm): boolean {
        return (
            term && 
            !stringIsNullOrEmpty(this.path) && 
            !stringIsNullOrEmpty(term.path) &&
            term.path.indexOf(this.path + ";") === 0
        );
    }

    
    public fromObject(object: any): void {
        assign(this, object);
    }

}