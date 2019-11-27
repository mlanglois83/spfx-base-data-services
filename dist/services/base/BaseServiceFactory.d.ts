import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
export declare class BaseServiceFactory {
    /**
     * Constructs a service given model name
     * @param  typeName Name of the model for which a service has to be instanciated
     */
    create(typeName: string): BaseDataService<IBaseItem>;
    /**
     * Returns an item contructor given its type name
     * @param typeName model type name
     */
    getItemTypeByName(typeName: any): (new (item?: any) => IBaseItem);
}
