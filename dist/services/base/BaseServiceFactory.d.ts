import { IBaseItem, IDataService } from "../../interfaces";
export declare class BaseServiceFactory {
    /**
     * Constructs a service given its name
     * @param serviceName Name of the service instance to be instanciated
     */
    create<T extends IBaseItem>(serviceName: string): IDataService<T>;
    /**
     * Returns an item contructor given its type name
     * @param typeName model type name
     */
    getItemTypeByName(typeName: any): (new (item?: any) => IBaseItem);
}
