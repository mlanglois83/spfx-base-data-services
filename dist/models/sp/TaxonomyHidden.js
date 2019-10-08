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
import { SPItem } from "../";
var TaxonomyHidden = /** @class */ (function (_super) {
    __extends(TaxonomyHidden, _super);
    function TaxonomyHidden(item) {
        var _this = _super.call(this, item) || this;
        if (item != undefined) {
            _this.termId = item.IdForTerm;
        }
        return _this;
    }
    TaxonomyHidden.prototype.convert = function () {
        throw new Error("Not implemented");
    };
    return TaxonomyHidden;
}(SPItem));
export { TaxonomyHidden };
//# sourceMappingURL=TaxonomyHidden.js.map