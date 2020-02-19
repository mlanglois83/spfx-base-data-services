
import { IBaseItem } from "../../interfaces/index";
import { Decorators } from "../../decorators";

const spField = Decorators.spField;

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class SPItem implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks?: any;

    public __getInternalLinks(propertyName: string): any {
        let result = null;
        if(this.__internalLinks) {
            result = this.__internalLinks[propertyName];
        }
        return result;
    }
    public __setInternalLinks(propertyName: string, value: any): void {
        this.__internalLinks = this.__internalLinks || {};
        this.__internalLinks[propertyName] = value;
    }
    public __deleteInternalLinks(propertyName: string): void {
        if(this.__internalLinks) {
            delete this.__internalLinks[propertyName];
        }
    }
    
    public __clearEmptyInternalLinks(): void {
        if(this.__internalLinks && Object.keys(this.__internalLinks).length === 0) { 
            delete this.__internalLinks;
        }       
    }

    /**
     * Item id
     */
    @spField({fieldName: "ID", defaultValue: -1 })
    public id = -1;
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
     * Defines if item is valid for sending it to list
     */
    public get isValid(): boolean {
        return true;
    }
}