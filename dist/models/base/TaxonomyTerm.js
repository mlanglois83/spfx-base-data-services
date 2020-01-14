import { stringIsNullOrEmpty } from "@pnp/common";
/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
var TaxonomyTerm = /** @class */ (function () {
    /**
     * Instanciates a term object
     * @param term term object from rest call
     */
    function TaxonomyTerm(term) {
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.CustomProperties;
        }
    }
    Object.defineProperty(TaxonomyTerm.prototype, "fullPathString", {
        get: function () {
            var result = "";
            if (!stringIsNullOrEmpty(this.path)) {
                var parts = this.path.split(";");
                result = parts.join(" > ");
            }
            return result;
        },
        enumerable: true,
        configurable: true
    });
    return TaxonomyTerm;
}());
export { TaxonomyTerm };
//# sourceMappingURL=TaxonomyTerm.js.map