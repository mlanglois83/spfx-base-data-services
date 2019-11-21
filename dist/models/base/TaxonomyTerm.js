/**
 * Base object for sharepoint taxonomy term abstraction objects
 */
var TaxonomyTerm = /** @class */ (function () {
    /**
     * Instanciates a term object
     * @param term term object from rest call
     */
    function TaxonomyTerm(term) {
        /**
         * internal field for linked items not stored in db
         */
        this.__internalLinks = {};
        if (term != undefined) {
            this.title = term.Name != undefined ? term.Name : "";
            this.id = term.Id != undefined ? term.Id.replace(/\/Guid\(([^)]+)\)\//g, "$1") : "";
            this.path = term.PathOfTerm != undefined ? term.PathOfTerm : "";
            this.customSortOrder = term.CustomSortOrder;
            this.customProperties = term.CustomProperties;
        }
    }
    return TaxonomyTerm;
}());
export { TaxonomyTerm };
//# sourceMappingURL=TaxonomyTerm.js.map