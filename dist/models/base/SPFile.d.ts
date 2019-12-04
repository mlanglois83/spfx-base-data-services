import { IBaseItem } from "../..";
/**
 * Data model for a SharePoint File
 */
export declare class SPFile implements IBaseItem {
    /**
     * File content (binary data)
     */
    content?: ArrayBuffer;
    /**
     * File mime type
     */
    mimeType: string;
    /**
     * File Id (server relative url)
     */
    id: string;
    /**
     * File title (name)
     */
    title: string;
    /**
     * Get or set file server relative url
     */
    get serverRelativeUrl(): string;
    /**
     * Get or set file server relative url
     */
    set serverRelativeUrl(val: string);
    /**
     * Get or set file name
     */
    get name(): string;
    /**
     * Get or set file name
     */
    set name(val: string);
    /**
     * Instanciate an SPFile object
     * @param fileItem file item from rest call (can be file or item)
     */
    constructor(fileItem?: any);
}
