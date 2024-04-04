import { RestItem } from "./RestItem";

/**
 * Base object for rest item abstraction objects
 */
 export abstract class RestStringItem extends RestItem<string> { 
     public get typedKey(): string {
         return "";
     }    
}