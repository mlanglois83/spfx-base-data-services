import { Guid } from "@microsoft/sp-core-library";
import { Constants } from "../../constants";
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
    @field({ defaultValue: -1 })
    public id = -1;
    /**
     * Version number
     */
    @field()
    public version?: number;   
    /**
     * Unique id
     */ 
    @field({fieldName: Constants.commonRestFields.uniqueid, defaultValue: Guid.newGuid().toString()})
    public uniqueId: string = Guid.newGuid().toString();  
}