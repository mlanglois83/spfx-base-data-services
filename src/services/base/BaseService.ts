import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { TraceLevel } from "../../constants";
import { LoggingService } from "../LoggingService";


export abstract class BaseService {   
    

    public get serviceName(): string {
        return this.constructor["name"];
    }

    protected get logFormat(): string {
        return LoggingService.defaultLogFormat;
    }
    constructor() {        
        if(this.debug) {
            LoggingService.addLoggingToTaggedMembers(this, this.logFormat);
        }
    }

    protected hashCode(obj: any): number {
        let hash = 0;
        const str = JSON.stringify((obj || ""));
        if (str.length == 0) return hash;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32bit integer
        }
        return hash;
    }

    public getDomainUrl(): string {
        return ServicesConfiguration.context.pageContext.web.absoluteUrl.replace(ServicesConfiguration.context.pageContext.web.serverRelativeUrl, "");
    }
    protected get debug(): boolean {
        return ServicesConfiguration.configuration.traceLevel !== TraceLevel.None;
    }

    /**************************************************************** Promise Concurency ******************************************************************************/

    /**
     * Stored promises to avoid multiple calls
     */
     protected static promises = {};

    protected getExistingPromise(key = "all"): Promise<any> {
        const pkey = this.serviceName + "-" + key;
        if (BaseService.promises[pkey]) {
            return BaseService.promises[pkey];
        }
        else return null;
    }

    protected storePromise(promise: Promise<any>, key = "all"): void {
        const pkey = this.serviceName + "-" + key;
        BaseService.promises[pkey] = promise;
        promise.then(() => {
            this.removePromise(key);
        }).catch(() => {
            this.removePromise(key);
        });
    }

    protected removePromise(key = "all"): void {
        const pkey = this.serviceName + "-" + key;
        delete BaseService.promises[pkey];
    }
    /*****************************************************************************************************************************************************************/

}

