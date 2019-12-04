
import { IBaseItem } from "../../interfaces/index";
import { spField } from "../../decorators";
/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class SPItem implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks?: any;
    /**
     * Item id
     */
    @spField({fieldName: "ID", defaultValue: -1 })
    public id: number = -1;
    /**
     * Item title
     */
    @spField({fieldName: "Title", defaultValue: "" })
    public title: string;
    /**
     * Version number
     */
    @spField({fieldName: "OData__UIVersionString"})
    public version?: number;
    /**
     * Queries (only used in services)
     */
    public queries?: Array<number>;
    /**
     * Constructs a SPItem object
     */
    constructor() {        
    }
    /**
     * Defines if item is valid for sending it to list
     */
    public get isValid(): boolean {
        return true;
    }
}