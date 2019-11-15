import { IFieldDescriptor } from "..";

export function spField(declaration: IFieldDescriptor) {
    return (target: any, propertyKey: string) => {
        if(!target.constructor.Fields) {
            target.constructor.Fields = {};
        }
        target.constructor.Fields[propertyKey] = declaration;
    };
}