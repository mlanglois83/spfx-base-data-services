import { SPWeb } from "@microsoft/sp-page-context";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { TraceLevel } from "../../constants";
import { LoggingService } from "../LoggingService";


export abstract class BaseService {
    protected get logFormat(): string {
        return LoggingService.defaultLogFormat;
    }
    constructor() {        
        if(ServicesConfiguration.configuration.traceLevel !== TraceLevel.None) {
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

    public getDomainUrl(web: SPWeb): string {
        return web.absoluteUrl.replace(web.serverRelativeUrl, "");
    }
}

