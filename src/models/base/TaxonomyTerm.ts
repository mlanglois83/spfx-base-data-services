import { stringIsNullOrEmpty } from "@pnp/common/util";
import { BaseItem } from "../base/BaseItem";

/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
export class TaxonomyTerm extends BaseItem {
    /**
     * WssIds assiciated with term
     */
    public wssids: Array<number> = [];
    /**
     * Term id (Guid)
     */
    public id = '00000000-0000-0000-0000-000000000000';
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
    public customProperties: {[key: string]: string};
    
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

}