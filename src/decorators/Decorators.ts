import { IFieldDescriptor } from "..";


export namespace Decorators {
    /**
     * Decorator function used for SPItem derived models fields
     * @param declaration - field declaration for binding
     */
    export function spField(declaration: IFieldDescriptor): (target: any, propertyKey: string) => void {
        return (target: any, propertyKey: string): void => {
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
} 