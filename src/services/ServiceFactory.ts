import { BaseDataService } from "./base/BaseDataService";
import { stringIsNullOrEmpty } from "@pnp/core";
import { ServicesConfiguration } from "../configuration/ServicesConfiguration";
import { BaseItem } from "../models/base/BaseItem";
import { IFieldDescriptor } from "../interfaces";
import { assign } from "lodash";
export class ServiceFactory {    
    
    private static __services: {[modelName: string]: BaseDataService<BaseItem>} = {};
    private static __serviceInitializing: {[modelName: string]: boolean} = {};
    private static __itemFields: {[modelName: string]: {[propertyName: string]: IFieldDescriptor}} = {};

    public static isServiceInitializing(modelName: string): boolean {
        return ServiceFactory.__serviceInitializing[modelName] === true;
    }

    public static isServiceManaged(modelName: string): boolean {
        return Object.keys(ServicesConfiguration.__factory.models).indexOf(modelName) !== -1;
    }
    /**
     * Constructs a service given model name
     * @param  typeName - name of the model for which a service has to be instanciated
     */
     public static getServiceByModelName(modelName: string, ...args: any[]): BaseDataService<BaseItem> {
        if(!ServiceFactory.__services[modelName]) {
            if(!ServicesConfiguration.__factory.services[modelName]) {
                console.log(`modelname: ${modelName}`);
                console.error("Unknown model name");
                throw Error("Unknown model name");
            }
            ServiceFactory.__serviceInitializing[modelName] = true;
            ServiceFactory.__services[modelName] = new ServicesConfiguration.__factory.services[modelName](...args);            
            delete ServiceFactory.__serviceInitializing[modelName];
        }        
        return ServiceFactory.__services[modelName];
    }

    public static getService<T extends BaseItem>(model: (new (item?: any) => T), ...args: any[]): BaseDataService<T> {
        return ServiceFactory.getServiceByModelName(model["name"], ...args) as BaseDataService<T>;
    }

    /**
     * Returns an item contructor given its type name
     * @param typeName - model type name
     */
    public static getItemTypeByName(modelName: string): (new (item?: any) => BaseItem) {
        if(!ServicesConfiguration.__factory.models[modelName]) {
            console.error("Unknown model name");
            throw Error("Unknown model name");
        }
        return ServicesConfiguration.__factory.models[modelName];
    }

    /**
     * Returns an item given its type name
     * @param typeName - model type name
     */
     public static getItemByName(modelName: string): BaseItem {
        const itemType = ServiceFactory.getItemTypeByName(modelName);
        return new itemType();
    }

    public static getModelFields(modelName: string): {[propertyName: string]: IFieldDescriptor} {
        if(!ServiceFactory.__itemFields[modelName]) {
            const itemType = ServiceFactory.getItemTypeByName(modelName);

            ServiceFactory.__itemFields[modelName] = {};
            if (itemType["Fields"] && itemType["Fields"][itemType["name"]]) {
                assign(ServiceFactory.__itemFields[modelName], itemType["Fields"][itemType["name"]]);
            }
            let parentType = itemType;
            do {
                parentType = Object.getPrototypeOf(parentType);
                if (itemType["Fields"] && itemType["Fields"][parentType["name"]]) {
                    for (const key in itemType["Fields"][parentType["name"]]) {
                        if (itemType["Fields"][parentType["name"]].hasOwnProperty(key)) {
                            if (ServiceFactory.__itemFields[modelName][key] === undefined || ServiceFactory.__itemFields[modelName][key] === null) {
                                // keep higher level redefinition
                                ServiceFactory.__itemFields[modelName][key] = itemType["Fields"][parentType["name"]][key];
                            }
                        }
                    }
                }
            } while (parentType["name"] !== BaseItem["name"]);
        }
        return ServiceFactory.__itemFields[modelName];
    }


    /**
     * Returns an object contructor given its type name
     * @param typeName - model type name
     */
    public static getObjectTypeByName(typeName: string): (new () => any) {
        if(stringIsNullOrEmpty(typeName)) {
            throw new Error("Type is required");
        }
        if(!ServicesConfiguration.__factory.objects[typeName]) {
            console.error("Unknown type name");
            throw Error("Unknown type name");
        }
        return ServicesConfiguration.__factory.objects[typeName];
    }


}
