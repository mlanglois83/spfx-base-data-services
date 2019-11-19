/**
 * Labels for synchronization messages
 */
export interface ITranslationLabels {
    /**
     * Message for file uploaded
     */
    UploadLabel: string;
    /**
     * Message for item added
     */
    AddLabel: string;
    /**
     * Message for item updated
     */
    UpdateLabel: string;
    /**
     * Message for item deleted
     */
    DeleteLabel: string;
    /**
     * Indexed db not defined error message
     */
    IndexedDBNotDefined: string;
    /**
     * Version conflict error message
     */
    versionHigherErrorMessage: string;
    /**
     * Synchronization error message with tokens
     * {0} --> item type label
     * {1} --> operation label
     * {2} --> item title
     * {3} --> item id
     * {4} --> message
     */
    SynchronisationErrorFormat: string;
    /**
     * Dictionnary of type labels
     * key: Model type name
     * Value : model label
     */
    typeTranslations: any;
}
