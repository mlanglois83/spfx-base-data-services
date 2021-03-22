import { SPItem } from "../";
import { Decorators } from "../../decorators";

const field = Decorators.field;
const dataModel = Decorators.dataModel;
/**
 * Taxonomy hidden list data model
 */
@dataModel()
export class TaxonomyHidden extends SPItem {
    /**
     * Term id (guid)
     */
    @field({fieldName: "IdForTerm", defaultValue: -1 })
    public termId: string;
    /**
     * Instanciate a new TaxonomyHidden object
     */
    constructor() {
        super();
    }
}