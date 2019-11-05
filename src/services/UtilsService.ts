import { BaseService } from "./base/BaseService";
import { TaxonomyTerm } from "../models/index";
import { find } from "@microsoft/sp-lodash-subset";
import { ServicesConfiguration } from "../";

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
    public static arrayBufferToBlob(buffer: ArrayBuffer, type: string) {
        return new Blob([buffer], { type: type });
    }
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
    public static getParentFolderUrl(url: string): string {
        let urlParts = url.split('/');
        urlParts.pop();
        return urlParts.join("/");
    }

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

    public static getTaxonomyTermByWssId<T extends TaxonomyTerm>(wssid: number, terms: Array<T>): T {
        return find(terms, (term) => {
            return (term.wssids && term.wssids.indexOf(wssid) > -1);
        });
    }

}