import { TraceLevel } from "../constants";
import { ITranslationLabels } from "./";

/**
 * Configuration format for spfx base data services
 */
export interface IConfiguration {
    /**
     * Name of indexed db
     */
    dbName: string;
    /**
     * Data Base version used to manage structure updates
     */
    dbVersion: number;
    /**
     * SharePoint version for rest calls
     */
    spVersion?: "Online" | "SP2013" | "SP2016" | "SP2019";
    /**
     * Result of the last connection test call
     */
    lastConnectionCheckResult?: boolean;
    /**
     * true : services can retrieve data when offline, false : every time a network call is performed
     */
    checkOnline?: boolean;
    /**
    * empty : indicate a specific url to test online/offline instead of site root (creating 302 to default page). Fill this value avoid making to many request
    */
    onlineCheckPage?: string;
    /**
     * SPFX component context
     */
    context?: any; // To avoid spfx dependency --> BaseComponentContext
    /**
     * SP Web url in case ther is no context
     */
    baseUrl?: string; // for SP farms
    /**
     * Data table names used to update structure (1 by data service)
     */
    tableNames?: Array<string>;
    /**
     * Translations used by synchronization service when an operation or an error is reported
     */
    translations?: ITranslationLabels;
    /**
     * Current user id
     */
    currentUserId?: number;
    /**
     * Id of Azure AD app registered to get authentication token
     */
    aadAppId?: string;
    /**
     * Add traces to services calls
     */
    traceLevel?: TraceLevel;
    /**
     * Limit simultaneous db calls (0 or undefined --> no limit)
     */
    maxSimultaneousDbAccess?: number;
    /**
     * Limit simultaneous network queries in services (0 or undefined --> no limit)
     * @todo
     */
    maxSimultaneousQueries?: number;
    /**
     * Key used to share services accross instances
     */
    serviceKey?: string;
}