import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IConfiguration } from "../interfaces";
/**
 * Configuration class for spfx base data services
 */
export declare class ServicesConfiguration {
    /**
     * Web Part Context
     */
    static get context(): BaseComponentContext;
    /**
     * Configuration object
     */
    static get configuration(): IConfiguration;
    /**
     * Default configuration
     */
    private static configurationInternal;
    /**
     * Initializes configuration, must be called before services instanciation
     * @param configuration IConfiguration object
     */
    static Init(configuration: IConfiguration): void;
}
