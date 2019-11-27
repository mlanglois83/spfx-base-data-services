import { FieldType } from "..";

export interface IFieldDescriptor {
    /**
     * Internal name of SharePoint field
     */
    fieldName: string;
    /**
     * Field type. If not set Simple is used
     */
    fieldType?: FieldType;
    /**
     * Default value if field is not set. If not set, undefined is used
     */
    defaultValue?: any;
    /**
     * Model name used for linked objects.
     */
    modelName?: string;
    /**
     * Referenced item model type name for taxonomy types only
     */
    refItemName?: string;
}