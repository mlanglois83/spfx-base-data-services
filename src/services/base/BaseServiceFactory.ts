import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
import { SPFile, User, TaxonomyHidden } from "../../models";
import { stringIsNullOrEmpty } from "@pnp/common";
export abstract class BaseServiceFactory {

    /**
     * Constructs a service given model name
     * @param  typeName - name of the model for which a service has to be instanciated
     */
    public abstract create(typeName: string): BaseDataService<IBaseItem>;

    /**
     * Returns an item contructor given its type name
     * @param typeName - model type name
     */
    public getItemTypeByName(typeName: string): (new (item?: any) => IBaseItem) {
        let result = null;
        switch (typeName) {
            case SPFile["name"]:
                result = SPFile;
                break;
            case User["name"]:
                result = User;
                break;            
            case TaxonomyHidden["name"]:
                    result = TaxonomyHidden;
                    break;
            default:
                break;
        }
        return result;
    }

    /**
     * Returns an object contructor given its type name
     * @param typeName - model type name
     */
    public getObjectTypeByName(typeName: string): (new () => any) {
        if(stringIsNullOrEmpty(typeName)) {
            throw new Error("Type is required");
        }
        return null;
    }


}
