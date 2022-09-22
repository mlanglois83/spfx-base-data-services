import { BaseFile } from "./BaseFile";
import { Decorators } from "../../decorators";
import { Constants } from "../../constants";
import { stringIsNullOrEmpty } from "@pnp/common/util";
import { UtilsService } from "../../services";

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
    @field({fieldName: Constants.commonRestFields.uniqueid, defaultValue: 'AAAAAAAA-AAAA-4AAA-BAAA-AAAAAAAAAAAA'.replace(/[AB]/g, 
    // Callback for String.replace() when generating a guid.
    function (character) {
        const randomNumber = Math.random();
        /* tslint:disable:no-bitwise */
        const num = (randomNumber * 16) | 0;
        // Check for 'A' in template string because the first characters in the
        // third and fourth blocks must be specific characters (according to "version 4" UUID from RFC 4122)
        const masked = character === 'A' ? num : (num & 0x3) | 0x8;
        return masked.toString(16);
    })})
    public uniqueId: string = UtilsService.generateGuid();  
    
    public get content(): ArrayBuffer {
        if(!stringIsNullOrEmpty(this.base64Content)) {
            const binary_string = atob(this.base64Content);
            const len = binary_string.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                bytes[i] = binary_string.charCodeAt(i);
            }
            return bytes.buffer;
        }
        return null;
    }
    public set content(value: ArrayBuffer) {
        if(value && value.byteLength > 0) {
            let binary = "";
            const bytes = new Uint8Array( value );
            const len = bytes.byteLength;
            for (let i = 0; i < len; i++) {
                binary += String.fromCharCode( bytes[i] );
            }
            this.base64Content = btoa(binary);
        }
        else {
            this.base64Content = null;
        }
    } 
    public get serverRelativeUrl(): string {        
        return this.url;
    }
    public set serverRelativeUrl(value: string) {
        this.url = value;
    }
}