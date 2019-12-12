import { SPWeb } from "@microsoft/sp-page-context";
export declare abstract class BaseService {
    protected hashCode(obj: any): number;
    constructor();
    getDomainUrl(web: SPWeb): string;
}
