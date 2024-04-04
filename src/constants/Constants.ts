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
        version: "OData__UIVersionString",
        id: "ID"
    },
    /**
     * Common rest item fields
     */
    commonRestFields: {
        created: "created",
        modified: "modified",
        author: "creator",
        editor: "modifiedby",
        version: "version",
        id: "id",
        uniqueid: "uniqueid"
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
         * Termsset sort order, 
         * {0} --> app key
         * {1} --> site relative url, 
         * {2} --> Termset name or id
         */
        termsetCustomOrder: "spfxdataservice-ts-custom-order-{0}-{1}-{2}",
        /**
         * Termsset site collection group id, 
         * {0} --> app key
         * {1} --> site relative url, 
         * {2} --> Termset name or id
         */
        termsetSiteCollectionGroupId: "spfxdataservice-ts-sitecollection-group-id-{0}-{1}-{2}",
        /**
         * Termsstore language, 
         * {0} --> app key
         * {1} --> site relative url
         */
        termStoreDefaultLanguageTag: "spfxdataservice-termstore-default-language-{0}-{1}",
        /**
         * Termsset id, 
         * {0} --> app key
         * {1} --> site relative url, 
         * {2} --> Termset name or id
         */
        termsetId: "spfxdataservice-ts-id-{0}-{1}-{2}",
        /**
         * Cache key for data service
         * {0} --> app key
         * {1} --> web server relative url
         * {2} --> service name
         * {3} --> operation key
         */
        latestDataLoadFormat: "spfxdataservice-latestDataLoad-{0}-{1}-{2}-{3}-{4}",
        /**
         * Cache key for data service
         * {0} --> app key
         * {1} --> web server relative url
         * {2} --> service name
         * {3} --> service args hashcode
         * {4} --> key
         */
        localStorageTableFormat: "spfxdataservice-table-{0}-{1}-{2}",
        /**
         * DbName
         * {0} --> app key
         * {1} --> web server relative url
         * {2} --> dbName
         */
        dbNameFormat: "{0}-{1}-{2}"
    },
    /**
     * Constants for SP Taxaonomy hidden list
     */
    taxonomyHiddenList: {
        /**
         * list web relative url
         */
        relativeUrl: "/lists/taxonomyhiddenlist"
    },
    /**
     * Standard table names
     */
    tableNames: [
        "OfflineTransaction",
        "OfflineTransactionFiles",
        "ListAttachments",
        "RestMapping"
    ],
    windowVars: {
        promiseVarName: "spfxBaseDataServicesPromises",
        servicesVarName: "spfxBaseDataServicesServices"
    },
    models: {
        offlineCreatedPrefix : "==OFFLINE=="
    }
};