
import { IBaseItem } from "../../interfaces";
import { assign, findIndex } from "@microsoft/sp-lodash-subset";
import { FieldType } from "../..";
import { stringIsNullOrEmpty } from "@pnp/pnpjs";
import { ServicesConfiguration } from "../../configuration";

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

    private _itemFields = null;
    public get ItemFields(): any {
        if(this._itemFields) {
            return this._itemFields;
        }
        else {
            this._itemFields = {};
            if (this.constructor["Fields"][this.constructor["name"]]) {
                assign(this._itemFields, this.constructor["Fields"][this.constructor["name"]]);
            }
            let parentType = this.constructor; 
            do {
                parentType = Object.getPrototypeOf(parentType);
                if(this.constructor["Fields"][parentType["name"]]) {
                    for (const key in this.constructor["Fields"][parentType["name"]]) {
                        if (Object.prototype.hasOwnProperty.call(this.constructor["Fields"][parentType["name"]], key)) {
                            if(this._itemFields[key] === undefined || this._itemFields[key] === null) {
                                // keep higher level redefinition
                                this._itemFields[key] = this.constructor["Fields"][parentType["name"]][key];
                            }                            
                        }
                    }
                }
            } while(parentType["name"] !== BaseItem["name"]);
        }
        return this._itemFields;
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
                            const objectConstructor = ServicesConfiguration.configuration.serviceFactory.getObjectTypeByName(fieldDescriptor.modelName);
                            const result = new objectConstructor();
                            this[propertyName] = assign(result, this[propertyName]);
                            break;
                        case FieldType.User:
                        case FieldType.Taxonomy:
                        case FieldType.Lookup:
                            // get model from serviceFactory
                            const singleModelConstructor = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(fieldDescriptor.modelName);
                            const singleModelValue = new singleModelConstructor();
                            singleModelValue.fromObject(this[propertyName]);
                            this[propertyName] = singleModelValue;
                            break;
                        case FieldType.UserMulti:
                        case FieldType.TaxonomyMulti:
                        case FieldType.LookupMulti:
                            if(Array.isArray(this[propertyName])){
                                // get model from serviceFactory
                                const modelConstructor = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(fieldDescriptor.modelName);
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