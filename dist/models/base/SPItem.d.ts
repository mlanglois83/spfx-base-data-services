import { IBaseItem } from "../../interfaces/index";
import { TaxonomyTerm } from "./TaxonomyTerm";
import { User } from "../graph/User";
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
    convert(): Promise<any>;
    protected convertTaxonomyFieldValue(value: TaxonomyTerm): any;
    protected convertSingleUserFieldValue(value: User): Promise<any>;
    protected convertMultiUserFieldValue(value: User[]): Promise<any>;
    readonly isValid: boolean;
    /**
     * called after update was made on sp list
     * @param addResultData added item from rest call
     */
    onAddCompleted(addResultData: any): void;
    /**
     * called after update was made on sp list
     * @param updateResult updated item from rest call
     */
    onUpdateCompleted(updateResult: any): void;
    /**
     * called before updating local db
     * update lookup, user, taxonomy ids here from stored objects
     */
    beforeUpdateDb(): void;
}
