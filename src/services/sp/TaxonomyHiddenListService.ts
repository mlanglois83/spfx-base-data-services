import { Constants } from "../../constants/index";
import { TaxonomyHidden } from "../../models/";
import { BaseListItemService } from "../base/BaseListItemService";



const cacheDuration = 1440;

/**
 * Service allowing to retrieve risks (online only)
 */
export class TaxonomyHiddenListService extends BaseListItemService<TaxonomyHidden> {


    constructor() {
        super(TaxonomyHidden, Constants.taxonomyHiddenList.relativeUrl, Constants.taxonomyHiddenList.tableName, cacheDuration);

    }


    /**
    * Cache has to be relaod ?
    *
    * @readonly
    * @protected
    * @type {boolean}
    * @memberof BaseDataService
    */
    protected async needRefreshCache(key = "all"): Promise<boolean> {

        let result: boolean = this.cacheDuration === -1;
        //if cache defined
        if (!result) {

            const cachedDataDate = this.getCachedData(key);
            if (cachedDataDate) {
                //add cache duration
                cachedDataDate.setMinutes(cachedDataDate.getMinutes() + this.cacheDuration);

                const now = new Date();

                //cache has expired
                result = cachedDataDate < now;
            } else {
                result = true;
            }

        }

        return result;
    }

}