import { SPWeb } from "@microsoft/sp-page-context";
export declare abstract class BaseService {
    protected hashCode(str: String): number;
    constructor();
    getDomainUrl(web: SPWeb): string;
}
