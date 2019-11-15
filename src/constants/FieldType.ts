export enum FieldType {
    Simple, // no format check 
    Lookup, // get unique id 
    LookupMulti, // get id array 
    Taxonomy, // get taxonomyTerm
    TaxonomyMulti, // get array off taxonomy terms
    O365User // get o365 user (via userservice)
}