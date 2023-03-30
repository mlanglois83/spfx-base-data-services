import { assign, findIndex } from "lodash";
import { stringIsNullOrEmpty } from "@pnp/core";
import { FieldType } from "../../constants";
import { IBaseItem, IFieldDescriptor } from "../../interfaces";
import { ServiceFactory } from "../../services/ServiceFactory";

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class BaseItem implements IBaseItem {

    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks?:  any;

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

    public __clearInternalLinks(): void {
        delete this.__internalLinks;
    }

    public __clearEmptyInternalLinks(): void {
        if (this.__internalLinks && Object.keys(this.__internalLinks).length === 0) {
            delete this.__internalLinks;
        }
    }
    /**
     * Item id
     */
    public id: number | string;
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
     * Deleted item
     */
    public deleted?: boolean;

    /**
     * Defines if item is valid for sending it to list
     */
    public get isValid(): boolean {
        return true;
    }

    public cleanBeforeStorage(): void {
        for (const propertyName in this) {
            if (this.hasOwnProperty(propertyName)) {
                if (!this.ItemFields.hasOwnProperty(propertyName) && typeof (this[propertyName]) === "function") {
                    delete this[propertyName];
                }
            }
        }
    }

    public get ItemFields(): {[propertyName: string]: IFieldDescriptor } {
        return ServiceFactory.getModelFields(this.constructor["name"]);        
    }
    
    public fromObject(object: any): void {
        assign(this, object);
        // fields
        for (const propertyName in this.ItemFields) {
            if (this.ItemFields.hasOwnProperty(propertyName) && object[propertyName]) {
                const fieldDescriptor = this.ItemFields[propertyName];
                if(fieldDescriptor.fieldType === FieldType.Date && typeof(object[propertyName]) === "string" && !stringIsNullOrEmpty(object[propertyName])) {
                    this[propertyName] = new Date(object[propertyName]);
                }
                else if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                    switch (fieldDescriptor.fieldType) {
                        case FieldType.Json:                     
                            // get object from serviceFactory
                            const objectConstructor = ServiceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                            const result = new objectConstructor();
                            this[propertyName] = assign(result, this[propertyName]);
                            break;
                        case FieldType.User:
                        case FieldType.Taxonomy:
                        case FieldType.Lookup:
                            // get model from serviceFactory
                            const singleModelValue =  ServiceFactory.getItemByName(fieldDescriptor.modelName);
                            singleModelValue.fromObject(this[propertyName]);
                            this[propertyName] = singleModelValue;
                            break;
                        case FieldType.UserMulti:
                        case FieldType.TaxonomyMulti:
                        case FieldType.LookupMulti:
                            if(Array.isArray(this[propertyName])){
                                // get model from serviceFactory
                                const modelConstructor = ServiceFactory.getItemTypeByName(fieldDescriptor.modelName);
                                for (let index = 0; index < this[propertyName].length; index++) {
                                    const element = this[propertyName][index];
                                    const modelValue = new modelConstructor();
                                    modelValue.fromObject(element);
                                    this[propertyName][index] = modelValue;
                                }
                            }
                            break;
                        default:                        
                            break;
                    }          
                }
                

            }
        }
        
    }
}