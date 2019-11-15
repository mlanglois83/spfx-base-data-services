import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IConfiguration } from "../interfaces";
export declare class ServicesConfiguration {
    static get context(): BaseComponentContext;
    static get configuration(): IConfiguration;
    private static configurationInternal;
    static Init(configuration: IConfiguration): void;
}
