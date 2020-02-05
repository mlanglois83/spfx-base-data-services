var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
import { Decorators } from "../../decorators";
var spField = Decorators.spField;
/**
 * Base object for sharepoint item abstraction objects
 */
var SPItem = /** @class */ (function () {
    /**
     * Constructs a SPItem object
     */
    function SPItem() {
        /**
         * Item id
         */
        this.id = -1;
    }
    Object.defineProperty(SPItem.prototype, "isValid", {
        /**
         * Defines if item is valid for sending it to list
         */
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    __decorate([
        spField({ fieldName: "ID", defaultValue: -1 })
    ], SPItem.prototype, "id", void 0);
    __decorate([
        spField({ fieldName: "Title", defaultValue: "" })
    ], SPItem.prototype, "title", void 0);
    __decorate([
        spField({ fieldName: "OData__UIVersionString" })
    ], SPItem.prototype, "version", void 0);
    return SPItem;
}());
export { SPItem };
//# sourceMappingURL=SPItem.js.map