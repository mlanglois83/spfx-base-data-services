import { assign } from "@microsoft/sp-lodash-subset";
import { IBaseItem } from "../../interfaces";

export class RestResultMapping implements IBaseItem {
    public id: string;
    public itemIds: number[] = [];

    public fromObject(object: any): void {
        assign(this, object);
    }

}