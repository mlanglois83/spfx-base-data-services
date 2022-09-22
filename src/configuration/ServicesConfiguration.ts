import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
import { IConfiguration, IFactoryMapping } from "../interfaces";
import { Constants, TraceLevel } from "../constants";

/**
 * Configuration class for spfx base data services
 */
export class ServicesConfiguration {



    public static __factory: IFactoryMapping = {
        models: {},
        services: {},
        objects: {}
    };

    /**
     * Web Part Context
     */
    public static get context(): any {//BaseComponentContext 
        return ServicesConfiguration.configuration.context;
    }
    /**
     * Web Url
     */
     public static get baseUrl(): string {//BaseComponentContext 
        return ServicesConfiguration.configuration.context ? ServicesConfiguration.context.pageContext.web.absoluteUrl : ServicesConfiguration.configuration.baseUrl;
    }

    public static get serverRelativeUrl(): string {
        return ServicesConfiguration.configuration.context ? ServicesConfiguration.context.pageContext.web.serverRelativeUrl : ServicesConfiguration.configuration.baseUrl.replace(/^https?:\/\/[^/]+(\/.*)$/g, "$1");
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
    private static configurationInternal: IConfiguration = {
        spVersion: "Online",
        dbName: "spfx-db",
        dbVersion: 1,
        lastConnectionCheckResult: false,
        checkOnline: false,
        onlineCheckPage: "",
        context: null,
        currentUserId: -1,
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
     * @param configuration - configuration object
     */
    public static Init(configuration: IConfiguration): void {
        ServicesConfiguration.configurationInternal = configuration;  
        configuration.spVersion = configuration.spVersion || "Online";      
        configuration.traceLevel = configuration.traceLevel || TraceLevel.None;
        configuration.tableNames = Constants.tableNames.concat(configuration.tableNames || []);
        configuration.lastConnectionCheckResult = false;
        configuration.checkOnline = configuration.checkOnline === true;
        configuration.translations = configuration.translations || {
            AddLabel: "Add",
            DeleteLabel: "Delete",
            IndexedDBNotDefined: "IDB not defined",
            SynchronisationErrorFormat: "Sync error",
            UpdateLabel: "Update",
            UploadLabel: "Upload",
            versionHigherErrorMessage: "Version conflict",
            typeTranslations: []
        };
        configuration.currentUserId = configuration.currentUserId > 0 ? configuration.currentUserId : -1;
        
        const allModels = ServicesConfiguration.__factory?.models || {};
        for (const key in allModels) {
            if (allModels.hasOwnProperty(key)) {
                configuration.tableNames.push(key); 
            }
        }
        // SP calls init with no cache
        sp.setup({
            spfxContext: ServicesConfiguration.context,
            sp: {
                baseUrl: ServicesConfiguration.configuration.baseUrl,
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Cache-Control': 'no-cache'
                }
            }
        });
        if(ServicesConfiguration.context) { // no graph without context
            graph.setup({
                spfxContext: ServicesConfiguration.context,
                graph:{
                    headers: {
                        "Accept": "application/json;odata.metadata=minimal",
                        'Cache-Control': 'no-cache'
                    }
                }
            });
        }
    }

    public static addObjectMapping(typeName: string, objectConstructor: (new () => any)): void {
        ServicesConfiguration.__factory.objects[typeName] = objectConstructor;
    }
}