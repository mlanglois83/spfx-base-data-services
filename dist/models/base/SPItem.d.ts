import { IBaseItem } from "../../interfaces/index";
/**
 * Base object for sharepoint item abstraction objects
 */
export declare abstract class SPItem implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    __internalLinks: any;
    /**
     * Item id
     */
    id: number;
    /**
     * Item title
     */
    title: string;
    /**
     * Version number
     */
    version?: number;
    /**
     * Queries (only used in services)
     */
    queries?: Array<number>;
    /**
     * Constructs a SPItem object
     */
    constructor();
    /**
     * Defines if item is valid for sending it to list
     */
    get isValid(): boolean;
}
