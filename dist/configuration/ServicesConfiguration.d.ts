import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IConfiguration } from "../interfaces";
export declare class ServicesConfiguration {
    static readonly context: BaseComponentContext;
    static readonly configuration: IConfiguration;
    private static configurationInternal;
    static Init(configuration: IConfiguration): void;
}
