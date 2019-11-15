import { FieldType } from "..";
export interface IFieldDescriptor {
    fieldName: string;
    fieldType: FieldType;
    defaultValue: any;
    serviceName?: string;
}
