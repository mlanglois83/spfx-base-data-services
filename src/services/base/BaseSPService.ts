import { stringIsNullOrEmpty } from "@pnp/core";
import { SPBrowser, spfi, SPFI } from "@pnp/sp";
import { ServicesConfiguration } from "../../configuration";
import { IBaseSPServiceOptions } from "../../interfaces";
import { BaseItem } from "../../models";
import { UtilsService } from "../UtilsService";
import { BaseDataService } from "./BaseDataService";

export abstract class BaseSPService<T extends BaseItem<string | number>> extends BaseDataService<T> {
    protected serviceOptions: IBaseSPServiceOptions;

    protected get baseUrl(): string {
        return this.serviceOptions?.baseUrl;
    }

    public get sp(): SPFI {
        if(stringIsNullOrEmpty(this.baseUrl)) {
            return ServicesConfiguration.sp;
        }
        else {
            return spfi().using(SPBrowser({ baseUrl: this.baseUrl }));
        }
    }

    protected get cacheKeyUrl(): string {
        if(stringIsNullOrEmpty(this.baseUrl)) {
            return super.cacheKeyUrl;
        }
        else {
            return UtilsService.getRelativeUrl(this.baseUrl);
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
    constructor(itemType: (new (item?: any) => T), options?: IBaseSPServiceOptions, ...args: any[]) {
        super(itemType, options, ...args);  
    }
}