import { SPFile } from "../../models";
var BaseServiceFactory = /** @class */ (function () {
    function BaseServiceFactory() {
    }
    BaseServiceFactory.prototype.create = function (serviceName) {
        return null;
    };
    BaseServiceFactory.prototype.getItemTypeByName = function (typeName) {
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
    return BaseServiceFactory;
}());
export { BaseServiceFactory };
//# sourceMappingURL=BaseServiceFactory.js.map