import { assign } from "lodash";
import { IBaseItem } from "../../interfaces";

export class RestResultMapping<T extends string | number> implements IBaseItem<string> {
    public get typedKey(): string { return "" }
    public get defaultKey(): string { return undefined; }
    public id: string;
    public itemIds: T[] = [];

    public fromObject(object: any): void {
        assign(this, object);
    }

}