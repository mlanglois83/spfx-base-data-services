import { BaseFile } from "./BaseFile";
import { Decorators } from "../../decorators";
import { Constants } from "../../constants";
import { Guid } from "@microsoft/sp-core-library";

const field = Decorators.field;
export abstract class RestFile extends BaseFile {
    /**
     * File id
     */
    @field({ defaultValue: -1 })
    public id = -1;
    /**
     * File name
     */
    @field({fieldName: "name"})
    public title: string;
    /**
     * File content (base 64)
     */
    @field({fieldName: "content"})
    public base64Content: string;
    /**
     * File url
     */
    @field({fieldName: "uri"})
    public url: string;
    /**
     * Version number
     */
    @field()
    public version?: number;   
    /**
     * Unique id
     */
    @field({fieldName: Constants.commonRestFields.uniqueid, defaultValue: Guid.newGuid().toString()})
    public uniqueId: string = Guid.newGuid().toString();  
    
    public get content(): ArrayBuffer {
        const binary_string = atob(this.base64Content);
        const len = binary_string.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binary_string.charCodeAt(i);
        }
        return bytes.buffer;
    }
    public set content(value: ArrayBuffer) {
        let binary = "";
        const bytes = new Uint8Array( value );
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode( bytes[i] );
        }
        this.base64Content = btoa(binary);
    } 
    public get serverRelativeUrl(): string {        
        return this.url;
    }
    public set serverRelativeUrl(value: string) {
        this.url = value;
    }
}