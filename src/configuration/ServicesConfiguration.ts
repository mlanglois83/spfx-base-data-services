import { SPBrowser, SPFI, spfi, SPFx as spSPFx } from "@pnp/sp/presets/all";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph/presets/all";
import { IConfiguration, IFactoryMapping } from "../interfaces";
import { Constants, TraceLevel } from "../constants";
import PnPTelemetry from "@pnp/"

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
     * Sp object for pnp call
     */
    public static get sp(): SPFI {
        if(ServicesConfiguration.context) {
            return spfi().using(spSPFx(ServicesConfiguration.context));
        }
        else {
            return spfi().using(SPBrowser({ baseUrl: ServicesConfiguration.baseUrl }));
        }
    }
    /**
     * graph object for pnp call
     */
    public static get graph(): GraphFI {
        return graphfi().using(graphSPFx(ServicesConfiguration.context));
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
        // disable telemtry
        const telemetry = PnPTelemetry.getInstance();
        telemetry.optOut();
    }

    public static addObjectMapping(typeName: string, objectConstructor: (new () => any)): void {
        ServicesConfiguration.__factory.objects[typeName] = objectConstructor;
    }
}