import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ITranslationLabels } from "./";
import { BaseServiceFactory } from "../services";

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
     * Result of the last connection test call
     */
    lastConnectionCheckResult:boolean;
    /**
     * true : services can retrieve data when offline, false : every time a network call is performed
     */
    checkOnline: boolean;
    /**
     * SPFX component context
     */
    context: BaseComponentContext;
    /**
     * Data table names used to update structure (1 by data service)
     */
    tableNames: Array<string>;
    /**
     * Translations used by synchronization service when an operation or an error is reported
     */
    translations: ITranslationLabels;
    /**
     * Service factory able to instanciate services allowing the synchronization service to work
     */
    serviceFactory: BaseServiceFactory;
}