import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { TraceLevel } from "../../constants";
import { LoggingService } from "../LoggingService";
import { UtilsService } from "../UtilsService";


export abstract class BaseService {   
    
    protected __thisArgs: any[];
    public get serviceName(): string {
        return this.constructor["name"];
    }

    protected get logFormat(): string {
        return LoggingService.defaultLogFormat;
    }
    constructor(...args: any[]) {     
        this.__thisArgs = args;   
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

    protected async callAsyncWithPromiseManagement<T>(
        promiseGenerator: () => Promise<T>,
        key = "all"
      ): Promise<T> {
        const pkey = this.serviceName + "-" + this.hashCode(this.__thisArgs) + "-" + key;
        return UtilsService.callAsyncWithPromiseManagement(pkey, promiseGenerator);
    }
    
    /*****************************************************************************************************************************************************************/

}

