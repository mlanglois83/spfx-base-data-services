export var FieldType;
(function (FieldType) {
    FieldType[FieldType["Simple"] = 0] = "Simple";
    FieldType[FieldType["Lookup"] = 1] = "Lookup";
    FieldType[FieldType["LookupMulti"] = 2] = "LookupMulti";
    FieldType[FieldType["Taxonomy"] = 3] = "Taxonomy";
    FieldType[FieldType["TaxonomyMulti"] = 4] = "TaxonomyMulti";
    FieldType[FieldType["O365User"] = 5] = "O365User"; // get o365 user (via userservice)
})(FieldType || (FieldType = {}));
//# sourceMappingURL=FieldType.js.map