import { Decorators } from "../../decorators";
import { BaseItem } from "./BaseItem";

const field = Decorators.field;

/**
 * Base object for rest item abstraction objects
 */
export abstract class RestItem extends BaseItem {
    /**
     * Item id
     */
    @field({ fieldName: "Id", defaultValue: -1 })
    public id = -1;
    /**
     * Version number
     */
    @field({ fieldName: "Version" })
    public version?: number;
}