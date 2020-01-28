var BaseService = /** @class */ (function () {
    function BaseService() {
    }
    BaseService.prototype.hashCode = function (obj) {
        var hash = 0;
        var str = JSON.stringify(obj);
        if (str.length == 0)
            return hash;
        for (var i = 0; i < str.length; i++) {
            var char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32bit integer
        }
        return hash;
    };
    BaseService.prototype.getDomainUrl = function (web) {
        return web.absoluteUrl.replace(web.serverRelativeUrl, "");
    };
    return BaseService;
}());
export { BaseService };
