import { assign } from "lodash";
import { ServicesConfiguration } from "../configuration/ServicesConfiguration";
import { TraceLevel } from "../constants";
import { IFieldDescriptor, IRestServiceDescriptor } from "../interfaces";


export namespace Decorators {


    function _getItemTypes(type: any): string[] {
        const typeName: string = type?.constructor?.name;
        let itemTypes: string[];

        if (typeName && typeName !== 'Function') {
            itemTypes = [typeName]; const parentTypes = _getItemTypes(Object.getPrototypeOf(type));
            itemTypes = itemTypes.concat(parentTypes).filter(t => t); return itemTypes;
        }
        return;
    }



    /**
     * Decorator function used for SPItem derived models fields
     * @param declaration - field declaration for binding
     * @deprecated use field instead
     */
    export function spField(declaration?: IFieldDescriptor): (target: any, propertyKey: string) => void {
        return (target: any, propertyKey: string): void => {
            if (!declaration) {
                declaration = {
                    fieldName: propertyKey
                };
            }
            if (!declaration.fieldName) {
                declaration.fieldName = propertyKey;
            }
            // constructs a static dictionnary on SPItem class
            if (!target.constructor.Fields) {
                target.constructor.Fields = {};
            }
            // First key : model name
            if (!target.constructor.Fields[target.constructor["name"]]) {
                target.constructor.Fields[target.constructor["name"]] = {};
            }
            // Second key : model field name
            target.constructor.Fields[target.constructor["name"]][propertyKey] = declaration;


            // Merge field with parent classes
            const types = _getItemTypes(target);
            if (types)
                types.forEach(type => {
                    if (target.constructor.Fields[type])
                        assign(target.constructor.Fields[target.constructor["name"]], target.constructor.Fields[type]);
                });
        };
    }


    export function trace(traceLevel: TraceLevel): (target: any, propertyKey: string) => void {
        return (target: any, propertyKey: string): void => {
            if (typeof (target[propertyKey] === "function")) {
                target.constructor.tracedMembers = target.constructor.tracedMembers || {};
                if (!target.constructor.tracedMembers[propertyKey]) {
                    target.constructor.tracedMembers[propertyKey] = traceLevel;
                }
            }
        };
    }


    /**
     * Decorator function used for models fields
     * @param declaration - field declaration for binding
     */
    export function field(declaration?: IFieldDescriptor): (target: any, propertyKey: string) => void {
        return spField(declaration);
    }

    export function restService(declaration: IRestServiceDescriptor): (target: any) => void {
        return (target: any): void => {
            target.serviceProps = declaration;
        };
    }

    export function dataService(modelName: string): (target: any) => void {
        return (target: any): void => {
            if (!ServicesConfiguration.__factory) {
                ServicesConfiguration.__factory = {
                    models: {},
                    services: {},
                    objects: {}
                };
            }
            if (!ServicesConfiguration.__factory.services) {
                ServicesConfiguration.__factory.services = {};
            }
            // First key : model name
            if (!ServicesConfiguration.__factory.services[modelName]) {
                ServicesConfiguration.__factory.services[modelName] = target;
            }
        };
    }

    export function dataModel(): (target: any) => void {
        return (target: any): void => {
            if (!ServicesConfiguration.__factory) {
                ServicesConfiguration.__factory = {
                    models: {},
                    services: {},
                    objects: {}
                };
            }
            if (!ServicesConfiguration.__factory.models) {
                ServicesConfiguration.__factory.models = {};
            }
            // First key : model name
            if (!ServicesConfiguration.__factory.models[target["name"]]) {
                ServicesConfiguration.__factory.models[target["name"]] = target;
            }
        };
    }
} 