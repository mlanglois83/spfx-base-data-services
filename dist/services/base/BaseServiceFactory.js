import { SPFile, User, TaxonomyHidden } from "../../models";
var BaseServiceFactory = /** @class */ (function () {
    function BaseServiceFactory() {
    }
    /**
     * Constructs a service given its name
     * @param serviceName Name of the service instance to be instanciated
     */
    BaseServiceFactory.prototype.create = function (serviceName) {
        var result = null;
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
    };
    /**
     * Returns an item contructor given its type name
     * @param typeName model type name
     */
    BaseServiceFactory.prototype.getItemTypeByName = function (typeName) {
        var result = null;
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
    };
    return BaseServiceFactory;
}());
export { BaseServiceFactory };
//# sourceMappingURL=BaseServiceFactory.js.map