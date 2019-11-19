import { IFieldDescriptor } from "..";
/**
 * Decorator function used for SPItem derived models fields
 * @param declaration field declaration for binding
 */
export declare function spField(declaration: IFieldDescriptor): (target: any, propertyKey: string) => void;
