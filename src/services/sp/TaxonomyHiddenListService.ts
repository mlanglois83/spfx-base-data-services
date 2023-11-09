import { Constants } from "../../constants/index";
import { Decorators } from "../../decorators";
import { IBaseListItemServiceOptions } from "../../interfaces";
import { TaxonomyHidden } from "../../models/";
import { BaseListItemService } from "../base/BaseListItemService";


const dataService = Decorators.dataService;
const cacheDuration = 60;

/**
 * Service allowing to retrieve risks (online only)
 */
@dataService("TaxonomyHidden")
export class TaxonomyHiddenListService extends BaseListItemService<TaxonomyHidden> {
    constructor(options?: Partial<IBaseListItemServiceOptions>, ...args: any[]) {
        super(TaxonomyHidden, Constants.taxonomyHiddenList.relativeUrl, options, ...args);
        this.serviceOptions.cacheDuration = this.serviceOptions.cacheDuration === undefined ? cacheDuration : this.serviceOptions.cacheDuration;
    }
}