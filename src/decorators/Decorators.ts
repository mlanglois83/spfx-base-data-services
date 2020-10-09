import { IFieldDescriptor, IRestServiceDescriptor } from "../interfaces";


export namespace Decorators {
    /**
     * Decorator function used for SPItem derived models fields
     * @param declaration - field declaration for binding
     * @deprecated use field instead
     */
    export function spField(declaration?: IFieldDescriptor): (target: any, propertyKey: string) => void {
        return (target: any, propertyKey: string): void => {
            if(!declaration){
                declaration = {
                    fieldName: propertyKey
                };
            }
            if(!declaration.fieldName) {
                declaration.fieldName = propertyKey;
            }
            // constructs a static dictionnary on SPItem class
            if(!target.constructor.Fields) {
                target.constructor.Fields = {};
            }
            // First key : model name
            if(!target.constructor.Fields[target.constructor["name"]]) {
                target.constructor.Fields[target.constructor["name"]] = {};
            }
            // Second key : model field name
            target.constructor.Fields[target.constructor["name"]][propertyKey] = declaration;
        };
    }
    /**
     * Decorator function used for models fields
     * @param declaration - field declaration for binding
     */
    export function field (declaration?: IFieldDescriptor): (target: any, propertyKey: string) => void { 
        return spField(declaration);
    }

    export function restService(declaration: IRestServiceDescriptor): (target: any) => void { 
        return (target: any): void => {
            target.serviceProps = declaration;            
        };
    }
} 