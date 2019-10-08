
import { IBaseItem } from "../../interfaces/index";
import { TaxonomyTerm } from "./TaxonomyTerm";
/**
 * Base object for sharepoint abstraction objects
 */
export abstract class SPItem implements IBaseItem {

    public id: number = -1;
    public title: string;
    public version?: number;
    public queries?: Array<number>;


    /**
     * Constructs a SPItem object
     * @param item object returned by sp call
     */
    constructor(item?: any) {
        if (item != undefined) {
            this.title = item["Title"] != undefined ? item["Title"] : "";
            this.id = item["ID"] != undefined ? item["ID"] : -1;
            this.version = item["OData__UIVersionString"] ? parseFloat(item["OData__UIVersionString"]): undefined;
        }
    }


    /**
     * Returns a copy of the object compatible with sp calls
     */
    public convert(): any {
        let result = {};
        result["Title"] = this.title;
        result["ID"] = this.id;
        return result;
    }


    protected convertTaxonomyFieldValue(value: TaxonomyTerm): any {
        let result: any = null;
        if (value) {
            result = {
                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                Label: value.title,
                TermGuid: value.id,
                WssId: -1 // fake
            };
        }
        return result;
    }

    public get isValid(): boolean {
        return true;
    }

    public onAddCompleted(addResultData: any): void {
        this.id = addResultData.Id;
        if(addResultData["OData__UIVersionString"]) {
            this.version = parseFloat(addResultData["OData__UIVersionString"]);
        }
    }
    public onUpdateCompleted(updateResult: any): void {
        if(updateResult["OData__UIVersionString"]) {
            this.version = parseFloat(updateResult["OData__UIVersionString"]);
        }
    }

}