import { BaseDataService } from "./BaseDataService";
import { IBaseItem, IDataService } from "../../interfaces";
import { SPFile, User, TaxonomyHidden } from "../../models";
import { UserService } from "../graph/UserService";
import { TaxonomyHiddenListService } from "../sp/TaxonomyHiddenListService";
export class BaseServiceFactory {

    /**
     * Constructs a service given its name
     * @param serviceName Name of the service instance to be instanciated
     */
    public create(serviceName: string): BaseDataService<IBaseItem> {
        let result = null;        
        /*switch(serviceName) {
            case UserService["name"]:
                result = new UserService();
                break;
            case TaxonomyHiddenListService["name"]:
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
