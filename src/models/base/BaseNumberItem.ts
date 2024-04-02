import { BaseItem } from "./BaseItem";

export abstract class BaseNumberItem extends BaseItem<number> {
    public get typedKey(): number {
        return 0;
    }    
}