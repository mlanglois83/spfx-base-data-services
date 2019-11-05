import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ITranslationLabels } from "./";
import { BaseServiceFactory } from "../services";

export interface IConfiguration {
    dbName: string;
    dbVersion: number;
    lastConnectionCheckResult:boolean;
    checkOnline: boolean;
    context: BaseComponentContext;
    tableNames: Array<string>;
    translations: ITranslationLabels;
    serviceFactory: BaseServiceFactory;
}