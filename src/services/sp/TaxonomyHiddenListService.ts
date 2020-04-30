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


}