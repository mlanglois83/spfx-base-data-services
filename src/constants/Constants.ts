export const Constants = { 
    /**
     * Common list item fields
     */
    commonFields: {
        created: "Created",
        modified: "Modified",
        author: "Author",
        editor: "Editor",
        attachments: "AttachmentFiles",
        version: "OData__UIVersionString"
    },
    /**
     *  Error codes
     */  
    Errors: {
        /**
         * Newer item found on server
         */
        ItemVersionConfict: "VERSION_CONFLICT"
    },
    /**
     * Default cache keys
     */
    cacheKeys: {
        /**
         * Termsset sort order, {0} --> site relative url, {1} --> Termset name or id
         */
        termsetCustomOrder: "spfxdataservice-ts-custom-order-{0}-{1}",
        /**
         * Cache key for data service
         * {0} --> web server relative url
         * {1} --> service name
         * {2} --> operation key
         */
        latestDataLoadFormat: "spfxdataservice-latestDataLoad-{0}-{1}-{2}"
    },
    /**
     * Constants for SP Taxaonomy hidden list
     */
    taxonomyHiddenList: {
        /**
         * Table name in indexed db
         */
        tableName: "TaxonomyHiddenList",
        /**
         * list web relative url
         */
        relativeUrl: "/lists/taxonomyhiddenlist"
    },
    /**
     * Standard table names
     */
    tableNames: [
        "Transaction",
        "TransactionFiles",
        "TaxonomyHiddenList", 
        "Users"]
};