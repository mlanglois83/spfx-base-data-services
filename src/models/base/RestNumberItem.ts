import { RestItem } from "./RestItem";

/**
 * Base object for rest item abstraction objects
 */
 export abstract class RestNumberItem extends RestItem<number> { 
     public get typedKey(): number {
         return 0;
     }    
}