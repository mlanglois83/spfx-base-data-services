import { Decorators } from "../../decorators";
import { BaseItem } from "./BaseItem";

const field = Decorators.field;

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class SPItem extends BaseItem {
    /**
     * Item id
     */
    @field({ fieldName: "ID", defaultValue: -1 })
    public id = -1;
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