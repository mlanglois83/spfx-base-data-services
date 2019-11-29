import { IBaseItem } from "../..";

/**
 * Data model for a SharePoint File
 */
export class SPFile implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    public __internalLinks: any = undefined;
    /**
     * File content (binary data)
     */
    public content?: ArrayBuffer;
    /**
     * File mime type
     */
    public mimeType: string; 
    /**
     * File Id (server relative url)
     */
    public id: string;
    /**
     * File title (name)
     */
    public title: string;

    /**
     * Get or set file server relative url
     */
    public get serverRelativeUrl(): string {
        return this.id;
    }
    /**
     * Get or set file server relative url
     */
    public set serverRelativeUrl(val: string) {
        this.id = val;
    }
    /**
     * Get or set file name
     */
    public get name(): string {
        return this.title;
    }
    /**
     * Get or set file name
     */
    public set name(val: string) {
        this.title = val;
    }

    /**
     * Instanciate an SPFile object
     * @param fileItem file item from rest call (can be file or item)
     */
    constructor(fileItem?:any){
        if(fileItem) {
            this.serverRelativeUrl = (fileItem.FileRef ? fileItem.FileRef : fileItem.ServerRelativeUrl);
            this.name = (fileItem.FileLeafRef ? fileItem.FileLeafRef : fileItem.Name);
        }
    }


}