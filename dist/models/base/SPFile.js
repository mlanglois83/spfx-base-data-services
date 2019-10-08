var SPFile = /** @class */ (function () {
    function SPFile(fileItem) {
        if (fileItem) {
            this.serverRelativeUrl = (fileItem.FileRef ? fileItem.FileRef : fileItem.ServerRelativeUrl);
            this.name = (fileItem.FileLeafRef ? fileItem.FileLeafRef : fileItem.Name);
        }
    }
    Object.defineProperty(SPFile.prototype, "serverRelativeUrl", {
        get: function () {
            return this.id;
        },
        set: function (val) {
            this.id = val;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPFile.prototype, "name", {
        get: function () {
            return this.title;
        },
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