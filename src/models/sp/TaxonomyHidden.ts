import { SPItem } from "../";
import { spField } from "../..";

/**
 * Taxonomy hidden list data model
 */
export class TaxonomyHidden extends SPItem {
    /**
     * Term id (guid)
     */
    @spField({fieldName: "IdForTerm", defaultValue: -1 })
    public termId: string;
    /**
     * Instanciate a new TaxonomyHidden object
     */
    constructor() {
        super();
    }
}