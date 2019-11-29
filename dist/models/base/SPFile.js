/**
 * Data model for a SharePoint File
 */
var SPFile = /** @class */ (function () {
    /**
     * Instanciate an SPFile object
     * @param fileItem file item from rest call (can be file or item)
     */
    function SPFile(fileItem) {
        /**
         * internal field for linked items not stored in db
         */
        this.__internalLinks = undefined;
        if (fileItem) {
            this.serverRelativeUrl = (fileItem.FileRef ? fileItem.FileRef : fileItem.ServerRelativeUrl);
            this.name = (fileItem.FileLeafRef ? fileItem.FileLeafRef : fileItem.Name);
        }
    }
    Object.defineProperty(SPFile.prototype, "serverRelativeUrl", {
        /**
         * Get or set file server relative url
         */
        get: function () {
            return this.id;
        },
        /**
         * Get or set file server relative url
         */
        set: function (val) {
            this.id = val;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPFile.prototype, "name", {
        /**
         * Get or set file name
         */
        get: function () {
            return this.title;
        },
        /**
         * Get or set file name
         */
        set: function (val) {
            this.title = val;
        },
        enumerable: true,
        configurable: true
    });
    return SPFile;
}());
export { SPFile };
//# sourceMappingURL=SPFile.js.map