import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ITranslationLabels } from "./";
import { BaseServiceFactory } from "../services";
export interface IConfiguration {
    DbName: string;
    Version: number;
    checkOnline: boolean;
    context: BaseComponentContext;
    versionHigherErrorMessage: string;
    tableNames: Array<string>;
    translations: ITranslationLabels;
    serviceFactory: BaseServiceFactory;
}
