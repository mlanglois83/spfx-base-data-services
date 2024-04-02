import { BaseFile } from "./BaseFile";

/**
 * Data model for a SharePoint File
 */
export class BaseNumberFile extends BaseFile<number> {    
    public get typedKey(): number{ return 0; }   
}