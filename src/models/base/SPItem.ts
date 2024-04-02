import { Decorators } from "../../decorators";
import { BaseNumberItem } from "./BaseNumberItem";

const field = Decorators.field;

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class SPItem extends BaseNumberItem {
    /**
     * Item id
     */
    @field({ fieldName: "ID" })
    public id;
    /**
     * Item title
     */
    @field({ fieldName: "Title", defaultValue: "" })
    public title: string;
    /**
     * Version number
     */
    @field({ fieldName: "OData__UIVersionString" })
    public version?: number;
}