import { sp } from "@pnp/sp";
import { taxonomy } from "@pnp/sp-taxonomy";
import { Constants } from "../constants";
import { find } from "@microsoft/sp-lodash-subset";
import { graph } from "@pnp/graph";
/**
 * Configuration class for spfx base data services
 */
var ServicesConfiguration = /** @class */ (function () {
    function ServicesConfiguration() {
    }
    Object.defineProperty(ServicesConfiguration, "context", {
        /**
         * Web Part Context
         */
        get: function () {
            return ServicesConfiguration.configuration.context;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ServicesConfiguration, "configuration", {
        /**
         * Configuration object
         */
        get: function () {
            return ServicesConfiguration.configurationInternal;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Initializes configuration, must be called before services instanciation
     * @param configuration IConfiguration object
     */
    ServicesConfiguration.Init = function (configuration) {
        ServicesConfiguration.configurationInternal = configuration;
        configuration.tableNames = configuration.tableNames || [];
        if (!find(configuration.tableNames, function (s) { return s === Constants.taxonomyHiddenList.tableName; })) {
            configuration.tableNames.push(Constants.taxonomyHiddenList.tableName);
        }
        // SP calls init with no cache
        sp.setup({
            spfxContext: ServicesConfiguration.context,
            sp: {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Cache-Control': 'no-cache'
                }
            }
        });
        taxonomy.setup({
            spfxContext: ServicesConfiguration.context,
            sp: {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Cache-Control': 'no-cache'
                }
            }
        });
        graph.setup({
            spfxContext: ServicesConfiguration.context
        });
    };
    /**
     * Default configuration
     */
    ServicesConfiguration.configurationInternal = {
        dbName: "spfx-db",
        dbVersion: 1,
        lastConnectionCheckResult: false,
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
    return ServicesConfiguration;
}());
export { ServicesConfiguration };
