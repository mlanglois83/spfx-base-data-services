import { IBaseFile } from "../../interfaces";
import { BaseItem } from "./BaseItem";

/**
 * Data model for a SharePoint File
 */
export class BaseFile<T extends string | number> extends BaseItem<T> implements IBaseFile<T> {
    
    public mimeType: string;  
    private  _content: ArrayBuffer;
    public get content(): ArrayBuffer {
        return this._content;
    }
    public set content(value: ArrayBuffer) {
        this._content = value;
    }
    private _serverRelativeUrl: string;
    public get serverRelativeUrl(): string {
        return this._serverRelativeUrl;
    }
    public set serverRelativeUrl(value: string) {
        this._serverRelativeUrl = value;
    }

}