import { IBaseItem } from "../../interfaces/index";

export class TaxonomyTerm implements IBaseItem {
    public wssids: Array<number>;
    public id: string;
    public title: string;
    public path: string;
    public customSortOrder?: string;
    public customProperties: any;

    constructor(term: any) {
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.customProperties;
        }
    }
    public convert(): any {
        throw new Error("Not implemented");
    }
}