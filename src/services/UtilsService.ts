import { BaseService } from "./base/BaseService";
import { ServicesConfiguration } from "../configuration/ServicesConfiguration";
import { cloneDeep, find } from "lodash";
//import { Batch } from "@pnp/odata";
import { TaxonomyTerm } from "../models/base/TaxonomyTerm";
import { stringIsNullOrEmpty } from "@pnp/core";
import { Constants, TraceLevel } from "../constants";
/**
 * Utility class
 */
export class UtilsService extends BaseService {



    constructor() {
        super();
    }

    
    private static checkOnlinePromise: Promise<boolean> = undefined;
    /**
     * check is user has connexion
     */
    public static async CheckOnline(): Promise<boolean> {
        if(!UtilsService.checkOnlinePromise) {            
            UtilsService.checkOnlinePromise = new Promise<boolean>(async (resolve) => {                
                let result = false;
                try {
                    const response = await fetch(ServicesConfiguration.baseUrl + (ServicesConfiguration.configuration.onlineCheckPage || ""), { method: 'HEAD', mode: 'no-cors' }); // head method not cached
                    result = (response && (response.ok || response.type === 'opaque'));
                }
                catch (ex) {
                    result = false;
                }
                ServicesConfiguration.configuration.lastConnectionCheckResult = result;
                UtilsService.checkOnlinePromise = undefined;
                resolve(result);
            });            
        }        
        return UtilsService.checkOnlinePromise;
    }

    /**
     * Converts blob object to array buffer
     * @param blob - source blob
     */
    public static blobToArrayBuffer(blob): Promise<ArrayBuffer> {
        return new Promise<ArrayBuffer>((resolve, reject) => {
            const reader = new FileReader();
            reader.addEventListener('loadend', () => {
                resolve(reader.result as ArrayBuffer);
            });
            reader.addEventListener('error', reject);
            reader.readAsArrayBuffer(blob);
        });
    }

    /**
     * Converts array buffer to blob
     * @param buffer - source array buffer
     * @param type - file mime type
     */
    public static arrayBufferToBlob(buffer: ArrayBuffer, type: string): Blob {
        return new Blob([buffer], { type: type });
    }

    /**
     * Return base 64 url from file content
     * @param fileData - file content
     */
    public static getOfflineFileUrl(fileData: Blob): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const reader = new FileReader;
            reader.onerror = reject;
            reader.onload = (): void => {
                const val = reader.result.toString();
                resolve(val);
            };
            reader.readAsDataURL(fileData);
        });
    }
    /**
     * Return parent folder url from url
     * @param url - child url 
     */
    public static getParentFolderUrl(url: string): string {
        const urlParts = url.split('/');
        urlParts.pop();
        return urlParts.join("/");
    }

    /**
     * Return relative url
     * @param url - url 
     */
     public static getRelativeUrl(url: string): string {
         if(!stringIsNullOrEmpty(url)) {
             url = url.replace(/((https?:)?\/\/[^\/]+)?(.*)/,"$3");
             url = stringIsNullOrEmpty(url) ? "/" : url; 
         }
         return url;
    }

    /**
     * Concatenate array buffers
     * @param arrays - array buffers to concatenate
     */
    public static concatArrayBuffers(...arrays: ArrayBuffer[]): ArrayBuffer {
        let length = 0;
        arrays.forEach((a) => {
            length += a.byteLength;
        });
        const joined = new Uint8Array(length);
        let offset = 0;
        arrays.forEach((a) => {
            joined.set(new Uint8Array(a), offset);
            offset += a.byteLength;
        });
        return joined.buffer;
    }



    /**
     * Escapes a string for use in a regex
     * @param value - string to escape
     */
    public static escapeRegExp(value: string): string {
        return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
    }

    /**
     * transform an array to the corresponding caml in clause values (surrounded with <Values></Values> tag)
     * @param values - array of value to transform to in values
     * @param fieldType - sp field type
     * @deprecated Use getCamlInQuery (for limit management)
     */
    public static getCamlInValues(values: Array<number | string>, fieldType: string): string {
        return values && values.length > 0 ? "<Values>" + values.map((value) => { return `<Value Type="${fieldType}">${value}</Value>`; }).join('') + "</Values>" : `<Values><Value Type="${fieldType}">-1</Value></Values>`;
    }

    /**  
     * Construction récursive d'une requête CAML
     * @param operande - operator name (Or, And)
     * @param listClauses - where clauses list
     */
    public static buildCAMLQueryRecursive(operande: string, listClauses: Array<string>): string {
        if (!listClauses || listClauses.length === 0)
            return "";
        if (listClauses.length === 1)
            return listClauses[0];
        const clause = listClauses[0];
        return `<${operande}>${clause}${UtilsService.buildCAMLQueryRecursive(operande, listClauses.slice(1))}</${operande}>`;
    }

    /**
     * transform an array to the corresponding caml in clause (surrounded with <Values></Values> tag)
     * @param fieldName - internal name of field
     * @param fieldType - sp field type
     * @param values - array of value to transform to in values
     * @param isLookup - true if query is based on lookup id (default false)
     */
    public static getCamlInQuery(fieldName: string, fieldType: string, values: Array<number | string>, isLookup = false): string {
        const copy = cloneDeep(values);
        if (copy && copy.length > 0) {
            const orClauses = [];
            while (copy.length) {
                const subValues = copy.splice(0, 500);
                orClauses.push(`<In><FieldRef LookupId="${isLookup ? "TRUE" : "FALSE"}" Name="${fieldName}"></FieldRef><Values>${subValues.map((value) => { return `<Value Type="${fieldType}">${value}</Value>`; }).join('')}</Values></In>`);
            }
            return UtilsService.buildCAMLQueryRecursive("Or", orClauses);
        }
        else {
            return `<In><FieldRef LookupId="${isLookup ? "TRUE" : "FALSE"}" Name="${fieldName}"></FieldRef><Values><Value Type="${fieldType}">-1</Value></Values></In>`;
        }
    }
    public static divideArray<T>(source: Array<T>, segments: number): Array<Array<T>> {
        if (segments < 2) {
            return [source];
        }

        const len = source.length;
        const out: Array<Array<T>> = [];

        let i = 0;
        let size: number;

        if (len % segments === 0) {
            size = Math.floor(len / segments);
            while (i < len) {
                out.push(source.slice(i, (i += size)));
            }
        } else {
            while (i < len) {
                size = Math.ceil((len - i) / segments);
                out.push(source.slice(i, i + size));
                i += size;
                segments -= 1;
            }
        }
        return out;
    }
    public static async runPromisesInStacks<T>(promises: Promise<T>[], stackCount: number): Promise<T[]> {
        const result: T[] = [];
        const segments = UtilsService.divideArray(promises, stackCount);
        const results = await Promise.all(segments.map((s) => {
            return UtilsService.chainPromises(s);
        }));
        results.forEach((r) => {
            result.push(...r);
        });
        return result;
    }

    public static async chainPromises<T>(promises: Promise<T>[]): Promise<T[]> {
        const result: T[] = [];
        while (promises.length > 0) {
            const currentPromise = promises.shift();
            const currentResult = await currentPromise;
            result.push(currentResult);
        }
        return result;
    }

    public static async executePromisesInStacks<T>(promiseGenerators: (() => Promise<T>)[], stackCount: number): Promise<T[]> {
        const result: T[] = [];
        const segments = UtilsService.divideArray(promiseGenerators, stackCount);
        const results = await Promise.all(segments.map((s) => {
            return UtilsService.chainPromiseGenerators(s);
        }));
        results.forEach((r) => {
            result.push(...r);
        });
        return result;
    }

    public static async chainPromiseGenerators<T>(promiseGenerators: (() => Promise<T>)[]): Promise<T[]> {
        const result: T[] = [];
        while (promiseGenerators.length > 0) {
            const currentPromise = promiseGenerators.shift()();
            const currentResult = await currentPromise;
            result.push(currentResult);
        }
        return result;
    }

    public static async runBatchesInStacks(batches: Array<() => Promise<void>>, stackCount: number): Promise<void> {
        const segments = UtilsService.divideArray(batches, stackCount);
        await Promise.all(segments.map((s) => {
            return UtilsService.chainBatches(s);
        }));
    }

    public static async chainBatches(batches: Array<() => Promise<void>>): Promise<void> {
        while (batches.length > 0) {
            const currentBatch = batches.shift();
            await currentBatch();
        }
    }

    public static getTermPathPartString(term: TaxonomyTerm, allTerms: Array<TaxonomyTerm>, level: number): string {
        const parts = term.path.split(";");
        if (parts.length - level < 1) {
            return "";
        }
        else if (parts.length - level === 1){
            return term.title;
        }
        else {
            const subParts = parts.slice(0, level + 1);
            const currentPath = subParts.join(";");
            const refTerm = find(allTerms, { path: currentPath });
            if(refTerm) {
                return refTerm.title;
            }
            else {
                return parts[level];
            }
        }
    }
    
    
    public static getTermFullPathString(term: TaxonomyTerm, allTerms: Array<TaxonomyTerm>, baseLevel = 0): string {
        const parts = term.path.split(";");
        if (parts.length - baseLevel <= 1) {
            return term.title;
        }
        else {
            const resultParts = [];
            const iterator = [];
            for (let index = 0; index < parts.length - 1; index++) {
                const part = parts[index];
                iterator.push(part);
                if (index >= baseLevel) {
                    const currentPath = iterator.join(";");
                    const refTerm = find(allTerms, { path: currentPath });
                    if (!refTerm) {
                        return term.path.split(";").slice(baseLevel).join(" > ");
                    }
                    resultParts.push(refTerm.title);
                }
            }
            resultParts.push(term.title);
            return resultParts.join(" > ");
        }
    }

    public static async isUrlInCache(url: string, cacheKey: string): Promise<boolean> {
        const cache = await caches.open(cacheKey);
        const response = await cache.match(url);
        return response !== undefined && response !== null;
    }

    public static formatText(s: string, ...values: unknown[]): string {        
        if (s === null) {
            // tslint:disable-line:no-null-keyword
            throw new Error(`The value for "${s}" must not be null`);
        }
        if (s === undefined) {
            throw new Error(`The value for "${s}" must not be undefined`);
        }
        return s.replace(/\{\d+\}/g, (match: string): string => {
            // The matchID is the numeric value contained in the brackets. {01} gets converted to 1.
            const matchId = parseInt(match.replace(/[\{\}]/g, ''), 10);
            // The matchValue is the string contained in the values' matchId index.
            // If matchId = 1 and values = ['the', 'bear', 'ate'], then matchValue = 'bear'
            let matchValue = `${values[matchId]}`;
            // If you index outside of the values array, return the original matchId in brackets
            if (matchId >= values.length || matchId < 0) {
                matchValue = match;
            }
            // Logic to convert null and undefined into readable strings
            /* tslint:disable-next-line:no-null-keyword */
            if (matchValue === null) {
                matchValue = 'null';
            }
            else if (matchValue === undefined) {
                matchValue = 'undefined';
            }
            return matchValue;
        });
    }

    public static generateGuid(): string {
        return 'AAAAAAAA-AAAA-4AAA-BAAA-AAAAAAAAAAAA'.replace(/[AB]/g, 
        // Callback for String.replace() when generating a guid.
        function (character) {
            const randomNumber = Math.random();
            /* tslint:disable:no-bitwise */
            const num = (randomNumber * 16) | 0;
            // Check for 'A' in template string because the first characters in the
            // third and fourth blocks must be specific characters (according to "version 4" UUID from RFC 4122)
            const masked = character === 'A' ? num : (num & 0x3) | 0x8;
            return masked.toString(16);
        });
    }

    public static getServerRelativeUrl(url: string): string {
        return url?.replace(/^https?:\/\/[^/]+(\/.*)$/g, "$1");
    }
    
    public static getAbsoluteUrl(url: string): string {
        if(url && !url?.match(/^(https?:)?\/\//g)) {
            url = url.replace(/^\/+/g, "");
            url = location.protocol + "//" + location.host +"/" + url;
        }
        return url;
    }

    public static callAsyncWithPromiseManagement<T>(
        key: string,
        promiseGenerator: () => Promise<T>
      ): Promise<T> {
        let promise = UtilsService.getExistingPromise<T>(key);
        if (promise) {
            if(ServicesConfiguration.configuration.traceLevel !== TraceLevel.None) {
                console.log(key + ' : load allready called before, sharing promise');
            }
        } else {
            promise = promiseGenerator();
            UtilsService.storePromise(promise, key);
        }
        return promise;
    }

    /* eslint-disable */
  /**************************************************************** Promise Concurency ******************************************************************************/
  /**
   * Stored promises to avoid multiple calls
   */
  protected static getExistingPromise<T>(key: string): Promise<T> {
    if (window[Constants.promiseVarName] && window[Constants.promiseVarName][key]) {
      return window[Constants.promiseVarName][key];
    } else return null;
  }

  protected static storePromise<T>(promise: Promise<T>, key: string): void {
    if (!window[Constants.promiseVarName]) {
      window[Constants.promiseVarName] = {};
    }
    window[Constants.promiseVarName][key] = promise;
    promise
      .then(() => {
        UtilsService.removePromise(key);
      })
      .catch(() => {
        UtilsService.removePromise(key);
      });
  }

  protected static removePromise(key: string): void {
    if (window[Constants.promiseVarName] && window[Constants.promiseVarName][key]) {
      delete window[Constants.promiseVarName][key];
    }
  }
  /*****************************************************************************************************************************************************************/
  /* eslint-enable */

}