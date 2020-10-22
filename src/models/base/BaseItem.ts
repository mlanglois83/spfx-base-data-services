
import { IBaseItem } from "../../interfaces";
import { findIndex } from "@microsoft/sp-lodash-subset";

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class BaseItem implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks?: any;

    public __getInternalLinks(propertyName: string): any {
        let result = null;
        if (this.__internalLinks) {
            result = this.__internalLinks[propertyName];
        }
        return result;
    }
    public __setInternalLinks(propertyName: string, value: any): void {
        this.__internalLinks = this.__internalLinks || {};
        this.__internalLinks[propertyName] = value;
    }


    public __setReplaceInternalLinks(propertyName: string, oldValue: any, newValue: any): void {
        const links = this.__getInternalLinks(propertyName) || [];

        const lookupidx = findIndex(links, (id) => { return id === oldValue; });
        if (lookupidx > -1) {
            links[lookupidx] = newValue;
        }
    }


    public __deleteInternalLinks(propertyName: string): void {
        if (this.__internalLinks) {
            delete this.__internalLinks[propertyName];
        }
    }

    public __clearEmptyInternalLinks(): void {
        if (this.__internalLinks && Object.keys(this.__internalLinks).length === 0) {
            delete this.__internalLinks;
        }
    }
    /**
     * Item id
     */
    public id: number;
    /**
     * Item title
     */
    public title?: string;
    /**
     * Version number
     */
    public version?: number;
    /**
     * Last update error
     */
    public error?: Error;
    /**
     * Defines if item is valid for sending it to list
     */
    public get isValid(): boolean {
        return true;
    }
}