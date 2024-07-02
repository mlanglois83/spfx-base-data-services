import { stringIsNullOrEmpty } from "@pnp/core";
import { BaseStringItem } from "./BaseStringItem";

/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
export class TaxonomyTerm extends BaseStringItem {
    /**
     * WssIds assiciated with term
     */
    public wssids: Array<number> = [];
    /**
     * Term id (Guid)
     */
    public id;
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
            term.path.startsWith(this.path + ";")
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

    public updatePath(oldPath:string, oldTitle: string, newTitle: string): void {
        this.path = this.path || this.title;
        const oldParts = oldPath.split(';');
        const newPath = oldParts.map((p, i)  => 
            (i === oldParts.length - 1 && p === oldTitle) ?
            newTitle
            :
            p
        ).join(';');
        if(this.path.startsWith(newPath + ';')) {
            this.path = this.path.replace(oldPath + ';', newPath + ';');
        }
        else if(this.path === oldPath) {
            this.path = newPath;
        }
    }

}