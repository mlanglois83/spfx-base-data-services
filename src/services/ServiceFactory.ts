import { BaseDataService } from "./base/BaseDataService";
import { IBaseItem } from "../interfaces";
import { BaseItem } from "../models";
import { stringIsNullOrEmpty } from "@pnp/common";
export class ServiceFactory {    
    
    private static __objectFactory = {};
    private static __services = {};
    
    /**
     * Constructs a service given model name
     * @param  typeName - name of the model for which a service has to be instanciated
     */
     public static getServiceByModelName(modelName: string): BaseDataService<IBaseItem> {
        if(!this.__services[modelName]) {
            if(!BaseDataService.__factory[modelName]) {
                throw Error("Unknown model name");
            }
            this.__services[modelName] = new BaseDataService.__factory[modelName]();
        }        
        return this.__services[modelName];
    }

    public static getService<T extends IBaseItem>(model: (new (item?: any) => T)): BaseDataService<T> {
        return ServiceFactory.getServiceByModelName(model["name"]) as BaseDataService<T>;
    }

    /**
     * Returns an item contructor given its type name
     * @param typeName - model type name
     */
    public static getItemTypeByName(modelName: string): (new (item?: any) => IBaseItem) {
        if(!BaseItem.__factory[modelName]) {
            throw Error("Unknown model name");
        }
        return BaseItem.__factory[modelName];
    }

    public static addObjectMapping(typeName: string, objectConstructor: (new () => any)): void {
        ServiceFactory.__objectFactory[typeName] = objectConstructor;
    }

    /**
     * Returns an object contructor given its type name
     * @param typeName - model type name
     */
    public static getObjectTypeByName(typeName: string): (new () => any) {
        if(stringIsNullOrEmpty(typeName)) {
            throw new Error("Type is required");
        }
        return ServiceFactory.__objectFactory[typeName];
    }


}
