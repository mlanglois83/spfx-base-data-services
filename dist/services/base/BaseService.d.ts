import { SPWeb } from "@microsoft/sp-page-context";
import { IConfiguration } from "../../interfaces";
export declare abstract class BaseService {
    protected static Configuration: IConfiguration;
    static Init(configuration: IConfiguration): void;
    protected hashCode(str: String): number;
    constructor();
    getDomainUrl(web: SPWeb): string;
}
