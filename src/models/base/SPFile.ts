import { BaseFile } from "./BaseFile";
/**
 * Data model for a SharePoint File
 */
export class SPFile extends BaseFile<string> {

    public get defaultKey(): string { return undefined}
    /**
     * File Id (server relative url)
     */
    public id: string;
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
     * @param fileItem - file item from rest call (can be file or item or attachment)
     */
    constructor(fileItem?: any){
        super();
        if(fileItem) {
            this.serverRelativeUrl = (fileItem.FileRef ? fileItem.FileRef : fileItem.ServerRelativeUrl);
            this.name = (fileItem.FileLeafRef ? fileItem.FileLeafRef : (fileItem.Name ? fileItem.Name : fileItem.FileName));
        }
    }


}