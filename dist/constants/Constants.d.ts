export declare const Constants: {
    /**
     * Common list item fields
     */
    commonFields: {
        created: string;
        modified: string;
        author: string;
        editor: string;
        attachments: string;
        version: string;
    };
    /**
     *  Error codes
     */
    Errors: {
        /**
         * Newer item found on server
         */
        ItemVersionConfict: string;
    };
    /**
     * Default cache keys
     */
    cacheKeys: {
        /**
         * Termsset sort order, {0} --> site relative url, {1} --> Termset name or id
         */
        termsetCustomOrder: string;
        /**
         * Cache key for data service
         * {0} --> web server relative url
         * {1} --> service name
         * {2} --> operation key
         */
        latestDataLoadFormat: string;
    };
    /**
     * Constants for SP Taxaonomy hidden list
     */
    taxonomyHiddenList: {
        /**
         * Table name in indexed db
         */
        tableName: string;
        /**
         * list web relative url
         */
        relativeUrl: string;
    };
    /**
     * Standard table names
     */
    tableNames: string[];
};
