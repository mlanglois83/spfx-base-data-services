import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
export declare class BaseServiceFactory {
    /**
     * Constructs a service given its name
     * @param serviceName Name of the service instance to be instanciated
     */
    create(serviceName: string): BaseDataService<IBaseItem>;
    /**
     * Returns an item contructor given its type name
     * @param typeName model type name
     */
    getItemTypeByName(typeName: any): (new (item?: any) => IBaseItem);
}
