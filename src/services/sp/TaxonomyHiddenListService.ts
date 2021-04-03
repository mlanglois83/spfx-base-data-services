import { Constants } from "../../constants/index";
import { Decorators } from "../../decorators";
import { TaxonomyHidden } from "../../models/";
import { BaseListItemService } from "../base/BaseListItemService";


const dataService = Decorators.dataService;
const cacheDuration = 60;

/**
 * Service allowing to retrieve risks (online only)
 */
@dataService("TaxonomyHidden")
export class TaxonomyHiddenListService extends BaseListItemService<TaxonomyHidden> {
    constructor() {
        super(TaxonomyHidden, Constants.taxonomyHiddenList.relativeUrl, cacheDuration, false);
    }
}