import { BaseDataService } from "./base/BaseDataService";
import { getHashCode, stringIsNullOrEmpty } from "@pnp/core";
import { ServicesConfiguration } from "../configuration/ServicesConfiguration";
import { BaseItem } from "../models/base/BaseItem";
import { IFieldDescriptor } from "../interfaces";
import { assign } from "lodash";
import { Constants } from "../constants";
export class ServiceFactory {    
    
    private static get servicesVarName(): string {
        return Constants.windowVars.servicesVarName + (ServicesConfiguration.configuration.serviceKey ? "-" + ServicesConfiguration.configuration.serviceKey : "");
    }

    private static get windowVar(): { 
        __services: {[modelName: string]: {[key: string]: {instanceId: string, service: BaseDataService<BaseItem<string | number>>}}}, 
        __serviceInitializing: {[modelName: string]: boolean}, 
        __itemFields: {[modelName: string]: {[propertyName: string]: IFieldDescriptor}}
    } {
        if(!window[ServiceFactory.servicesVarName]) {
            window[ServiceFactory.servicesVarName] = {
                __services: {},
                __serviceInitializing: {},
                __itemFields: {},
                __keyTypes: {}
            };
        } 
        return window[ServiceFactory.servicesVarName];
    }

    public static resetServicesCache() {
        ServiceFactory.windowVar.__serviceInitializing = {};
        ServiceFactory.windowVar.__services = {};
    }

    public static removeServicesForCurrentInstance() {
        Object.keys(ServiceFactory.windowVar.__services).forEach(modelName => {
            Object.keys(ServiceFactory.windowVar.__services[modelName]).forEach(serviceHash => {
                if(ServiceFactory.windowVar.__services[modelName][serviceHash].instanceId === ServicesConfiguration.context?.instanceId) {
                    delete ServiceFactory.windowVar.__services[modelName][serviceHash];
                }
            });
            if(Object.keys(ServiceFactory.windowVar.__services[modelName]).length === 0) {
                delete ServiceFactory.windowVar.__services[modelName];
            }
        });    
    }

    public static isServiceInitializing(modelName: string): boolean {
        return ServiceFactory.windowVar.__serviceInitializing[modelName] === true;
    }

    public static isServiceManaged(modelName: string): boolean {
        return Object.keys(ServicesConfiguration.__factory.models).indexOf(modelName) !== -1;
    }
    /**
     * Constructs a service given model name
     * @param  typeName - name of the model for which a service has to be instanciated
     */
    public static getServiceByModelName(modelName: string, ...args: any[]): BaseDataService<BaseItem<string | number>> {
        const hash = getHashCode(ServicesConfiguration.serverRelativeUrl + JSON.stringify((args || "")));
        if(!ServiceFactory.windowVar.__services[modelName] || !ServiceFactory.windowVar.__services[modelName][hash]) {
            if(!ServicesConfiguration.__factory.services[modelName]) {
                console.log(`modelname: ${modelName}`);
                console.error("Unknown model name");
                throw Error("Unknown model name");
            }
            ServiceFactory.windowVar.__serviceInitializing[modelName] = true;
            ServiceFactory.windowVar.__services[modelName] = ServiceFactory.windowVar.__services[modelName] || {};
            ServiceFactory.windowVar.__services[modelName][hash] = { 
                instanceId: ServicesConfiguration.context?.instanceId,
                service: new ServicesConfiguration.__factory.services[modelName](...args)
            } ;            
            delete ServiceFactory.windowVar.__serviceInitializing[modelName];
        }        
        return ServiceFactory.windowVar.__services[modelName][hash].service;
    }

    public static getService<T extends BaseItem<string | number>>(model: (new (item?: any) => T), ...args: any[]): BaseDataService<T> {
        return ServiceFactory.getServiceByModelName(model["name"], ...args) as BaseDataService<T>;
    }

    /**
     * Returns an item contructor given its type name
     * @param typeName - model type name
     */
    public static getItemTypeByName(modelName: string): (new (item?: any) => BaseItem<string | number>) {
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
     public static getItemByName(modelName: string): BaseItem<string | number> {
        const itemType = ServiceFactory.getItemTypeByName(modelName);
        return new itemType();
    }


    public static getModelFields(modelName: string): {[propertyName: string]: IFieldDescriptor} {
        if(!ServiceFactory.windowVar.__itemFields[modelName]) {
            const itemType = ServiceFactory.getItemTypeByName(modelName);

            ServiceFactory.windowVar.__itemFields[modelName] = {};
            if (itemType["Fields"] && itemType["Fields"][itemType["name"]]) {
                assign(ServiceFactory.windowVar.__itemFields[modelName], itemType["Fields"][itemType["name"]]);
            }
            let parentType = itemType;
            do {
                parentType = Object.getPrototypeOf(parentType);
                if (itemType["Fields"] && itemType["Fields"][parentType["name"]]) {
                    for (const key in itemType["Fields"][parentType["name"]]) {
                        if (itemType["Fields"][parentType["name"]].hasOwnProperty(key)) {
                            if (ServiceFactory.windowVar.__itemFields[modelName][key] === undefined || ServiceFactory.windowVar.__itemFields[modelName][key] === null) {
                                // keep higher level redefinition
                                ServiceFactory.windowVar.__itemFields[modelName][key] = itemType["Fields"][parentType["name"]][key];
                            }
                        }
                    }
                }
            } while (parentType["name"] !== BaseItem["name"]);
        }
        return ServiceFactory.windowVar.__itemFields[modelName];
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
