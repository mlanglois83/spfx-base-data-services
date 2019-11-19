/**
 * SP Field types used for models decorators
 */
export declare enum FieldType {
    /**
     * Common field types as text and integer
     * Model field type must be string or number
     */
    Simple = 0,
    /**
     * Date field
     * Model field type must be Date
     */
    Date = 1,
    /**
     * Single lookup type, please provide an item model type for linking
     * Model field type must be integer
     */
    Lookup = 2,
    /**
     * Multi lookup type, please provide an item model type for linking
     * Model field type must be array of integers
     */
    LookupMulti = 3,
    /**
     * Single taxonomy type, please provide an service name for linking
     * Model field type must inherit from TaxonomyTerm
     */
    Taxonomy = 4,
    /**
     * Multi taxonomy type, please provide an service name for linking
     * Model field type must be an array of TaxonomyTerm child
     */
    TaxonomyMulti = 5,
    /**
     * User type resolving a O365 user
     * Model field must be User
     */
    O365User = 6,
    /**
     * Multi User type resolving a O365 user
     * Model field must be an array of User
     */
    O365UserMulti = 7,
    /**
     * Text field parsed to json
     */
    Json = 8
}
