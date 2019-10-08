import { BaseService } from "./base/BaseService";
import { TaxonomyTerm } from "../models/index";
export declare class UtilsService extends BaseService {
    constructor();
    /**
     * check is user has connexion
     */
    CheckOnline(): Promise<boolean>;
    static blobToArrayBuffer(blob: any): Promise<ArrayBuffer>;
    static arrayBufferToBlob(buffer: ArrayBuffer, type: string): Blob;
    static getOfflineFileUrl(fileData: Blob): Promise<string>;
    static getParentFolderUrl(url: string): string;
    static concatArrayBuffers(...arrays: ArrayBuffer[]): ArrayBuffer;
    static isUrlInCache(url: string, cacheKey: string): Promise<boolean>;
    static getTaxonomyTermByWssId<T extends TaxonomyTerm>(wssid: number, terms: Array<T>): T;
}
