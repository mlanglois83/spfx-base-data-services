import { Decorators } from "../../decorators";
import { BaseItem } from "./BaseItem";

const field = Decorators.field;

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class SPItem extends BaseItem<number> {
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