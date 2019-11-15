
import { IBaseItem } from "../../interfaces/index";
import { TaxonomyTerm } from "./TaxonomyTerm";
import { User } from "../graph/User";
import { UserService } from "../../services";
import { spField } from "../../decorators";
import { FieldType } from "../../constants";
/**
 * Base object for sharepoint abstraction objects
 */
export abstract class SPItem implements IBaseItem {
    @spField({fieldName: "ID", fieldType: FieldType.Simple, defaultValue: -1 })
    public id: number = -1;
    @spField({fieldName: "Title", fieldType: FieldType.Simple, defaultValue: "" })
    public title: string;
    @spField({fieldName: "OData__UIVersionString", fieldType: FieldType.Simple, defaultValue: undefined })
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
    public async convert(): Promise<any> {
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
    protected async convertSingleUserFieldValue(value: User): Promise<any> {
        let result: any = null;
        if (value) {
            if(!value.spId || value.spId <=0) {
                let userService:UserService = new UserService();
                value = await userService.linkToSpUser(value);

            }
            result = value.spId;
        }
        return result;
    }
    protected async convertMultiUserFieldValue(value: User[]): Promise<any> {
        let result: any = null;
        if (value) {
            result = await Promise.all(value.map((val) => {
                return this.convertSingleUserFieldValue(val);
            }));
        }
        return result;
    }

    public get isValid(): boolean {
        return true;
    }

    /**
     * called after update was made on sp list
     * @param addResultData added item from rest call
     */
    public onAddCompleted(addResultData: any): void {
        this.id = addResultData.Id;
        if(addResultData["OData__UIVersionString"]) {
            this.version = parseFloat(addResultData["OData__UIVersionString"]);
        }
    }
    /**
     * called after update was made on sp list
     * @param updateResult updated item from rest call
     */
    public onUpdateCompleted(updateResult: any): void {
        if(updateResult["OData__UIVersionString"]) {
            this.version = parseFloat(updateResult["OData__UIVersionString"]);
        }
    }
}