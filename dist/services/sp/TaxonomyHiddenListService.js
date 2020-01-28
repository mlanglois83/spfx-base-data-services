var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Constants } from "../../constants/index";
import { TaxonomyHidden } from "../../models/";
import { BaseListItemService } from "../base/BaseListItemService";
var cacheDuration = 10;
/**
 * Service allowing to retrieve risks (online only)
 */
var TaxonomyHiddenListService = /** @class */ (function (_super) {
    __extends(TaxonomyHiddenListService, _super);
    function TaxonomyHiddenListService() {
        return _super.call(this, TaxonomyHidden, Constants.taxonomyHiddenList.relativeUrl, Constants.taxonomyHiddenList.tableName, cacheDuration) || this;
    }
    return TaxonomyHiddenListService;
}(BaseListItemService));
export { TaxonomyHiddenListService };
//# sourceMappingURL=TaxonomyHiddenListService.js.map