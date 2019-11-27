/**
 * SP Field types used for models decorators
 */
export var FieldType;
(function (FieldType) {
    /**
     * Common field types as text and integer
     * Model field type must be string or number
     */
    FieldType[FieldType["Simple"] = 0] = "Simple";
    /**
     * Date field
     * Model field type must be Date
     */
    FieldType[FieldType["Date"] = 1] = "Date";
    /**
     * Single lookup type, please provide an item model type for linking
     * Model field type must be integer or typed with linked model type if model name is defined
     */
    FieldType[FieldType["Lookup"] = 2] = "Lookup";
    /**
     * Multi lookup type, please provide an item model type for linking
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    FieldType[FieldType["LookupMulti"] = 3] = "LookupMulti";
    /**
     * Single taxonomy type, please provide a model name name for linking
     * Model field type must inherit from TaxonomyTerm
     */
    FieldType[FieldType["Taxonomy"] = 4] = "Taxonomy";
    /**
     * Multi taxonomy type, please provide an service name for linking
     * Model field type must be an array of TaxonomyTerm child
     */
    FieldType[FieldType["TaxonomyMulti"] = 5] = "TaxonomyMulti";
    /**
     * User type resolving a O365 user
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    FieldType[FieldType["User"] = 6] = "User";
    /**
     * Multi User type resolving a O365 user
     * Model field type must be array of integers or an array of linked model type if model name is defined
     */
    FieldType[FieldType["UserMulti"] = 7] = "UserMulti";
    /**
     * Text field parsed to json
     */
    FieldType[FieldType["Json"] = 8] = "Json";
})(FieldType || (FieldType = {}));
//# sourceMappingURL=FieldType.js.map