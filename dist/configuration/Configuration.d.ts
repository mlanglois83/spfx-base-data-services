import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IConfiguration } from "../interfaces";
export default class ServicesConfiguration {
    static readonly context: BaseComponentContext;
    private static configuration;
    static Init(configuration: IConfiguration): void;
}
