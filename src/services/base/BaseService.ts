import { SPWeb } from "@microsoft/sp-page-context";


export abstract class BaseService {
    
    protected hashCode(str: String): number {
        var hash = 0;
        if (str.length == 0) return hash;
        for (let i = 0; i < str.length; i++) {
            let char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32bit integer
        }
        return hash;
    }

    constructor() {
    }

    public getDomainUrl(web: SPWeb): string {
        return web.absoluteUrl.replace(web.serverRelativeUrl, "");
    }
}

