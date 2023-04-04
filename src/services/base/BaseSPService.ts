import { stringIsNullOrEmpty } from "@pnp/core";
import { SPBrowser, spfi, SPFI } from "@pnp/sp";
import { ServicesConfiguration } from "../../configuration";
import { BaseItem } from "../../models";
import { UtilsService } from "../UtilsService";
import { BaseDataService } from "./BaseDataService";

export abstract class BaseSPService<T extends BaseItem> extends BaseDataService<T> {
    protected baseUrl: string;

    public get sp(): SPFI {
        if(stringIsNullOrEmpty(this.baseUrl)) {
            return ServicesConfiguration.sp;
        }
        else {
            return spfi().using(SPBrowser({ baseUrl: this.baseUrl }));
        }
    }

    public get baseRelativeUrl(): string {
        if(stringIsNullOrEmpty(this.baseUrl)) {
            return ServicesConfiguration.serverRelativeUrl;
        }
        else {
            return UtilsService.getRelativeUrl(this.baseUrl);
        }
        
    }

    /**
     * 
     * @param type - type of items
     * @param context - context of the current wp
     */
    constructor(type: (new (item?: any) => T), cacheDuration = -1, baseUrl = undefined) {
        super(type, cacheDuration);  
        this.baseUrl = baseUrl;      
    }
}