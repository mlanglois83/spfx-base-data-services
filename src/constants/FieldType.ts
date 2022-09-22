/**
 * SP Field types used for models decorators
 */
export enum FieldType {
    /**
     * Common field types as text and integer
     * Model field type must be string or number
     */
    Simple,
    /**
     * UrlField
     * Model field type must be IUrl
     */
    Url,
    /**
     * Date field
     * Model field type must be Date
     */
    Date,
    /**
     * Single lookup type, please provide an item model type for linking
     * Model field type must be integer or typed with linked model type if model name is defined
     */
    Lookup,
    /**
     * Multi lookup type, please provide an item model type for linking
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    LookupMulti,
    /**
     * Single taxonomy type, please provide a model name name for linking 
     * Model field type must inherit from TaxonomyTerm
     */
    Taxonomy,
    /**
     * Multi taxonomy type, please provide an service name for linking 
     * Model field type must be an array of TaxonomyTerm child
     */
    TaxonomyMulti,
    /**
     * User type resolving a O365 user
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    User, 
    /**
     * Multi User type resolving a O365 user
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    UserMulti,
    /**
     * Text field parsed to json
     */
    Json
}