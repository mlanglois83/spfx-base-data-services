import { Constants } from "../../constants";
import { Decorators } from "../../decorators";
import { UtilsService } from "../../services";
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
    @field({fieldName: Constants.commonRestFields.uniqueid, defaultValue: 'AAAAAAAA-AAAA-4AAA-BAAA-AAAAAAAAAAAA'.replace(/[AB]/g, 
    // Callback for String.replace() when generating a guid.
    function (character) {
        const randomNumber = Math.random();
        /* tslint:disable:no-bitwise */
        const num = (randomNumber * 16) | 0;
        // Check for 'A' in template string because the first characters in the
        // third and fourth blocks must be specific characters (according to "version 4" UUID from RFC 4122)
        const masked = character === 'A' ? num : (num & 0x3) | 0x8;
        return masked.toString(16);
    })})
    public uniqueId: string = UtilsService.generateGuid();  

}