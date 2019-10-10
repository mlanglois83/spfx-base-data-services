var TaxonomyTerm = /** @class */ (function () {
    function TaxonomyTerm(term) {
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.customProperties;
        }
    }
    TaxonomyTerm.prototype.convert = function () {
        throw new Error("Not implemented");
    };
    return TaxonomyTerm;
}());
export { TaxonomyTerm };
//# sourceMappingURL=TaxonomyTerm.js.map