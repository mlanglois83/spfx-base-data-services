import { SPBrowser, SPFI, spfi, SPFx as spSPFx } from "@pnp/sp/presets/all";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph/presets/all";
import { IConfiguration, IFactoryMapping } from "../interfaces";
import { Constants, TraceLevel } from "../constants";

/**
 * Configuration class for spfx base data services
 */
/**
 * The `ServicesConfiguration` class provides static methods and properties to manage the configuration
 * and initialization of services, models, and objects within the application.
 * 
 * @remarks
 * This class includes methods to initialize the configuration, retrieve context and SP/Graph objects,
 * and manage object mappings. It also provides default configuration settings.
 * 
 * @example
 * To initialize the configuration:
 * ```typescript
 * ServicesConfiguration.Init({
 *     spVersion: "Online",
 *     dbName: "my-db",
 *     dbVersion: 1,
 *     context: myContext,
 *     baseUrl: "https://my.sharepoint.com",
 *     // other configuration settings
 * });
 * ```
 * 
 * @public
 */
export class ServicesConfiguration {



    public static __factory: IFactoryMapping = {
        models: {},
        services: {},
        objects: {}
    };

    /**
     * Retrieves the context from the ServicesConfiguration.
     * If the context is defined as a function, it invokes the function to get the context.
     * Otherwise, it directly returns the context.
     *
     * @returns {any} The context, which can be of any type for compatibility purpose.
     */
    public static get context(): any {//BaseComponentContext 
        if(typeof(ServicesConfiguration.configuration.context) === "function") {
            return ServicesConfiguration.configuration.context();
        }
        return ServicesConfiguration.configuration.context;
    }
    
    /**
     * Retrieves an instance of the SharePoint Fluent API (SPFI) configured with the appropriate context.
     * 
     * @returns {SPFI} An instance of SPFI configured with either the SPFx context or a browser-based context.
     * 
     * If `ServicesConfiguration.context` is defined, the SPFI instance will be configured using the SPFx context.
     * Otherwise, it will be configured using a browser-based context with the base URL from `ServicesConfiguration.baseUrl`.
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
     * Gets an instance of GraphFI configured with the SPFx context.
     *
     * @returns {GraphFI} An instance of GraphFI using the SPFx context.
     */
    public static get graph(): GraphFI {
        return graphfi().using(graphSPFx(ServicesConfiguration.context));
    }
    /**
     * Web Url
     */
     public static get baseUrl(): string {//BaseComponentContext 
        return ServicesConfiguration.context ? ServicesConfiguration.context.pageContext.web.absoluteUrl : ServicesConfiguration.configuration.baseUrl;
    }

    /**
     * Gets the server relative URL.
     * 
     * If the `context` is available, it returns the `serverRelativeUrl` from the `pageContext.web`.
     * Otherwise, it extracts and returns the relative URL from the `baseUrl` in the configuration.
     * 
     * @returns {string} The server relative URL.
     */
    public static get serverRelativeUrl(): string {
        return ServicesConfiguration.context ? ServicesConfiguration.context.pageContext.web.serverRelativeUrl : ServicesConfiguration.configuration.baseUrl.replace(/^https?:\/\/[^/]+(\/.*)$/g, "$1");
    }

    
    /**
     * Retrieves the internal configuration for the services.
     * 
     * @returns {IConfiguration} The current services configuration.
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
        useLocalStorage: false,
        onlineCheckPage: "",
        context: null,
        currentUserId: -1,
        serviceKey: "global",
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
     * Initializes the ServicesConfiguration with the provided configuration.
     * 
     * @param {IConfiguration} configuration - The configuration object to initialize with.
     * 
     * @remarks
     * This method sets default values for various configuration properties if they are not already provided.
     * It also ensures that the `tableNames` array is populated with both default and provided table names.
     * 
     * The following properties are set with default values if not provided:
     * - `spVersion`: Defaults to "Online".
     * - `traceLevel`: Defaults to `TraceLevel.None`.
     * - `tableNames`: Defaults to an empty array.
     * - `lastConnectionCheckResult`: Defaults to `false`.
     * - `checkOnline`: Defaults to `false`.
     * - `useLocalStorage`: Defaults to `false`.
     * - `serviceKey`: Defaults to "global".
     * - `translations`: Defaults to an object with predefined labels and messages.
     * - `currentUserId`: Defaults to `-1` if not greater than `0`.
     * 
     * Additionally, it appends all model keys from models with decorator to the `tableNames` array.
     */
    public static Init(configuration: IConfiguration): void {
        ServicesConfiguration.configurationInternal = configuration;  
        configuration.spVersion = configuration.spVersion || "Online";      
        configuration.traceLevel = configuration.traceLevel || TraceLevel.None;
        configuration.tableNames = Constants.tableNames.concat(configuration.tableNames || []);
        configuration.lastConnectionCheckResult = false;
        configuration.checkOnline = configuration.checkOnline === true;
        configuration.useLocalStorage = configuration.useLocalStorage === true;
        configuration.serviceKey = configuration.serviceKey || "global";
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
        
    }

    /**
     * Adds a mapping between a type name and its corresponding object constructor.
     *
     * @param typeName - The name of the type to be mapped.
     * @param objectConstructor - The constructor function for the object associated with the type name.
     */
    public static addObjectMapping(typeName: string, objectConstructor: (new () => any)): void {
        ServicesConfiguration.__factory.objects[typeName] = objectConstructor;
    }
}