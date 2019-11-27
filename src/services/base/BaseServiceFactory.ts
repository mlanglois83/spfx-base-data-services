import { BaseDataService } from "./BaseDataService";
import { IBaseItem, IDataService } from "../../interfaces";
import { SPFile, User, TaxonomyHidden } from "../../models";
import { UserService } from "../graph/UserService";
import { TaxonomyHiddenListService } from "../sp/TaxonomyHiddenListService";
export class BaseServiceFactory {

    /**
     * Constructs a service given model name
     * @param  typeName Name of the model for which a service has to be instanciated
     */
    public create(typeName: string): BaseDataService<IBaseItem> {
        let result = null;        
        /*switch(typeName) {
            case User["name"]:
                result = new UserService();
                break;
            case TaxonomyHidden["name"]:
                result = new TaxonomyHiddenListService();
                break;
            default: break;
        }*/
        return result;
    }

    /**
     * Returns an item contructor given its type name
     * @param typeName model type name
     */
    public getItemTypeByName(typeName): (new (item?: any) => IBaseItem) {
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

}
