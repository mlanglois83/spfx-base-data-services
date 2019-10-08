import { IBaseItem } from "../../interfaces/index";
import { TaxonomyTerm } from "./TaxonomyTerm";
/**
 * Base object for sharepoint abstraction objects
 */
export declare abstract class SPItem implements IBaseItem {
    id: number;
    title: string;
    version?: number;
    queries?: Array<number>;
    /**
     * Constructs a SPItem object
     * @param item object returned by sp call
     */
    constructor(item?: any);
    /**
     * Returns a copy of the object compatible with sp calls
     */
    convert(): any;
    protected convertTaxonomyFieldValue(value: TaxonomyTerm): any;
    readonly isValid: boolean;
    onAddCompleted(addResultData: any): void;
    onUpdateCompleted(updateResult: any): void;
}
