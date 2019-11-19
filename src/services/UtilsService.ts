import { BaseService } from "./base/BaseService";
import { ServicesConfiguration } from "../";

/**
 * Utility class
 */
export class UtilsService extends BaseService {



    constructor() {
        super();
    }

    /**
     * check is user has connexion
     */
    public static async CheckOnline(): Promise<boolean> {
        let result = false;


        try {
            const response = await fetch(ServicesConfiguration.context.pageContext.web.absoluteUrl, { method: 'HEAD', mode: 'no-cors' }); // head method not cached
            result = (response && (response.ok || response.type === 'opaque'));
        }
        catch (ex) {
            result = false;
        }
        ServicesConfiguration.configuration.lastConnectionCheckResult = result;
        return result;

    }

    /**
     * Converts blob object to array buffer
     * @param blob source blob
     */
    public static blobToArrayBuffer(blob): Promise<ArrayBuffer> {
        return new Promise<ArrayBuffer>((resolve, reject) => {
            const reader = new FileReader();
            reader.addEventListener('loadend', (e) => {
                resolve(<ArrayBuffer>reader.result);
            });
            reader.addEventListener('error', reject);
            reader.readAsArrayBuffer(blob);
        });
    }

    /**
     * Converts array buffer to blob
     * @param buffer source array buffer
     * @param type file mime type
     */
    public static arrayBufferToBlob(buffer: ArrayBuffer, type: string) {
        return new Blob([buffer], { type: type });
    }

    /**
     * Return base 64 url from file content
     * @param fileData file content
     */
    public static getOfflineFileUrl(fileData: Blob): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const reader = new FileReader;
            reader.onerror = reject;
            reader.onload = () => {
                let val = reader.result.toString();
                resolve(val);
            };
            reader.readAsDataURL(fileData);
        });
    }
    /**
     * Return parent folder url from url
     * @param url child url 
     */
    public static getParentFolderUrl(url: string): string {
        let urlParts = url.split('/');
        urlParts.pop();
        return urlParts.join("/");
    }

    /**
     * Concatenatee array buffers
     * @param arrays array buffers
     */
    public static concatArrayBuffers(...arrays: ArrayBuffer[]): ArrayBuffer {
        let length = 0;
        let buffer = null;
        arrays.forEach((a) => {
            length += a.byteLength;
        });
        let joined = new Uint8Array(length);
        let offset = 0;
        arrays.forEach((a) => {
            joined.set(new Uint8Array(a), offset);
            offset += a.byteLength;
        });
        return joined.buffer;
    }

    

    /**
     * Escapes a string for use in a regex
     * @param value string to escape
     */
    public static escapeRegExp(value: string) {
        return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }

    /**
     * transform an array to the corresponding caml in clause values (surrounded with <Values></Values> tag)
     * @param values array of value to transform to in values
     * @param fieldType sp field type
     */
    public static getCamlInValues(values: Array<number | string>, fieldType: string): string {
        return values && values.length > 0 ? "<Values>" + values.map((value) => { return `<Value Type="${fieldType}">${value}</Value>`; }).join('') + "</Values>" : `<Values><Value Type="${fieldType}">-1</Value></Values>`;
    }
}