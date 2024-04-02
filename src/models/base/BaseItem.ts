import { assign, findIndex } from "lodash";
import { stringIsNullOrEmpty } from "@pnp/core";
import { Constants, FieldType } from "../../constants";
import { IBaseItem, IFieldDescriptor } from "../../interfaces";
import { ServiceFactory } from "../../services/ServiceFactory";

/**
 * Base object for sharepoint item abstraction objects
 */
export abstract class BaseItem<T extends string | number> implements IBaseItem<T> {

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
    public get isLocal(): boolean {
        return this.id === this.defaultKey || this.isCreatedOffline; 
    }
    public get isCreatedOffline(): boolean {
        return (
            this.id !== this.defaultKey
            &&
            (
                (typeof(this.id) === "string" && this.id.indexOf(Constants.models.offlineCreatedPrefix) === 0)
                ||
                (typeof(this.id) === "number" && this.id < 0)
            )
        );
    }
    /**
     * default value for id
     */
    public get defaultKey(): T { return undefined; }

    /**
     * typed value for id
     */
    public abstract get typedKey(): T;

    /**
     * Item id
     */
    public id: T;
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
    
    
    constructor() {
        this.id = this.defaultKey;
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
    public equals(other: unknown): boolean {
        let result = other !== null && other !== undefined && other instanceof BaseItem;
        if (result) {
            Object.keys(this.ItemFields).forEach(k => {
                const thisVal = this[k];
                const otherVal = other[k];
                if (Array.isArray(thisVal)) {
                    result &&= thisVal.length === otherVal.length && thisVal.every(v => otherVal.some(ov => this.valuesEquals(v, ov)));
                }
                else {
                    result &&= this.valuesEquals(thisVal, other[k]);
                }
            });
        }
        return result;
    }

    private valuesEquals(a: unknown, b: unknown): boolean {
        if (a && b && typeof (a) === typeof (b)) {
            if (a instanceof BaseItem) {
                return b instanceof BaseItem && a.id === b.id;
            }
            else if (typeof (a) === 'object') {
                return JSON.stringify(a) === JSON.stringify(b);
            }
            else {
                return a === b;
            }
        }
        else {
            return a === b;
        }
    }

}