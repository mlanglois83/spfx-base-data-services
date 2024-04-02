import { BaseItem } from "./BaseItem";

export abstract class BaseStringItem extends BaseItem<string> {
    public get typedKey(): string {
        return "";
    }    
}