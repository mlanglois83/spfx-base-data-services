import { sp } from "@pnp/sp";
import { taxonomy } from "@pnp/sp-taxonomy";
import { Constants } from "../../constants";
import { find } from "@microsoft/sp-lodash-subset";
var _dataServiceConfig = {
    DbName: "spfx-db",
    Version: 1,
    checkOnline: false,
    context: null,
    serviceFactory: null,
    tableNames: [],
    translations: {
        AddLabel: "Add",
        DeleteLabel: "Delete",
        IndexedDBNotDefined: "IDB not defined",
        SynchronisationErrorFormat: "Sync error",
        UpdateLabel: "Update",
        UploadLabel: "Upload",
        versionHigherErrorMessage: "Version conflict",
        typeTranslations: []
    }
};
var BaseService = /** @class */ (function () {
    function BaseService() {
    }
    Object.defineProperty(BaseService, "Configuration", {
        get: function () {
            return _dataServiceConfig;
        },
        set: function (configuration) {
            _dataServiceConfig = configuration;
        },
        enumerable: true,
        configurable: true
    });
    BaseService.Init = function (configuration) {
        BaseService.Configuration = configuration;
        configuration.tableNames = configuration.tableNames || [];
        if (!find(configuration.tableNames, function (s) { return s === Constants.taxonomyHiddenList.tableName; })) {
            configuration.tableNames.push(Constants.taxonomyHiddenList.tableName);
        }
        sp.setup({
            spfxContext: BaseService.Configuration.context,
            sp: {
                headers: {
                    'Cache-Control': 'no-cache'
                }
            }
        });
        taxonomy.setup({
            spfxContext: BaseService.Configuration.context,
            sp: {
                headers: {
                    'Cache-Control': 'no-cache'
                }
            }
        });
    };
    BaseService.prototype.hashCode = function (str) {
        var hash = 0;
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
//# sourceMappingURL=BaseService.js.map