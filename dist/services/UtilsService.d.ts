import { BaseService } from "./base/BaseService";
/**
 * Utility class
 */
export declare class UtilsService extends BaseService {
    constructor();
    /**
     * check is user has connexion
     */
    static CheckOnline(): Promise<boolean>;
    /**
     * Converts blob object to array buffer
     * @param blob source blob
     */
    static blobToArrayBuffer(blob: any): Promise<ArrayBuffer>;
    /**
     * Converts array buffer to blob
     * @param buffer source array buffer
     * @param type file mime type
     */
    static arrayBufferToBlob(buffer: ArrayBuffer, type: string): Blob;
    /**
     * Return base 64 url from file content
     * @param fileData file content
     */
    static getOfflineFileUrl(fileData: Blob): Promise<string>;
    /**
     * Return parent folder url from url
     * @param url child url
     */
    static getParentFolderUrl(url: string): string;
    /**
     * Concatenatee array buffers
     * @param arrays array buffers
     */
    static concatArrayBuffers(...arrays: ArrayBuffer[]): ArrayBuffer;
    /**
     * Escapes a string for use in a regex
     * @param value string to escape
     */
    static escapeRegExp(value: string): string;
    /**
     * transform an array to the corresponding caml in clause values (surrounded with <Values></Values> tag)
     * @param values array of value to transform to in values
     * @param fieldType sp field type
     */
    static getCamlInValues(values: Array<number | string>, fieldType: string): string;
}
