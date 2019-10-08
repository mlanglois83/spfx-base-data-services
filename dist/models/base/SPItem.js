/**
 * Base object for sharepoint abstraction objects
 */
var SPItem = /** @class */ (function () {
    /**
     * Constructs a SPItem object
     * @param item object returned by sp call
     */
    function SPItem(item) {
        this.id = -1;
        if (item != undefined) {
            this.title = item["Title"] != undefined ? item["Title"] : "";
            this.id = item["ID"] != undefined ? item["ID"] : -1;
            this.version = item["OData__UIVersionString"] ? parseFloat(item["OData__UIVersionString"]) : undefined;
        }
    }
    /**
     * Returns a copy of the object compatible with sp calls
     */
    SPItem.prototype.convert = function () {
        var result = {};
        result["Title"] = this.title;
        result["ID"] = this.id;
        return result;
    };
    SPItem.prototype.convertTaxonomyFieldValue = function (value) {
        var result = null;
        if (value) {
            result = {
                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                Label: value.title,
                TermGuid: value.id,
                WssId: -1 // fake
            };
        }
        return result;
    };
    Object.defineProperty(SPItem.prototype, "isValid", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    SPItem.prototype.onAddCompleted = function (addResultData) {
        this.id = addResultData.Id;
        if (addResultData["OData__UIVersionString"]) {
            this.version = parseFloat(addResultData["OData__UIVersionString"]);
        }
    };
    SPItem.prototype.onUpdateCompleted = function (updateResult) {
        if (updateResult["OData__UIVersionString"]) {
            this.version = parseFloat(updateResult["OData__UIVersionString"]);
        }
    };
    return SPItem;
}());
export { SPItem };
//# sourceMappingURL=SPItem.js.map