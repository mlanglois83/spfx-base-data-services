import { BaseComponentContext } from "@microsoft/sp-component-base";
import { sp } from "@pnp/sp";
import { taxonomy } from "@pnp/sp-taxonomy";
import { IConfiguration } from "../interfaces";
import { Constants } from "../constants";
import { find } from "@microsoft/sp-lodash-subset";
import { graph } from "@pnp/graph";

/**
 * Configuration class for spfx base data services
 */
export class ServicesConfiguration {

    /**
     * Web Part Context
     */
    public static get context(): BaseComponentContext {
        return ServicesConfiguration.configuration.context;
    }

    /**
     * Configuration object
     */
    public static get configuration(): IConfiguration {
        return ServicesConfiguration.configurationInternal;
    }

    /**
     * Default configuration
     */
    private static configurationInternal: IConfiguration= {
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

    /**
     * Initializes configuration, must be called before services instanciation
     * @param configuration IConfiguration object
     */
    public static Init(configuration: IConfiguration): void {
        ServicesConfiguration.configurationInternal = configuration;
        configuration.tableNames = configuration.tableNames || [];
        if(!find(configuration.tableNames, (s) => {return s === Constants.taxonomyHiddenList.tableName})) {
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
            spfxContext: ServicesConfiguration.context,
            graph: {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Cache-Control': 'no-cache'
                }
            }
        });
    }
}