import { SPFile } from "../../models";
var ServiceFactory = /** @class */ (function () {
    function ServiceFactory() {
    }
    ServiceFactory.create = function (context, serviceName) {
        return null;
    };
    ServiceFactory.getItemTypeByName = function (typeName) {
        var result = null;
        switch (typeName) {
            case SPFile["name"]:
                result = SPFile;
                break;
            default:
                break;
        }
        return result;
    };
    return ServiceFactory;
}());
export { ServiceFactory };
//# sourceMappingURL=ServiceFactory.js.map