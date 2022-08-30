import { assign } from "lodash";
import { IBaseItem } from "../../interfaces";

export class RestResultMapping implements IBaseItem {
    public id: string;
    public itemIds: number[] = [];

    public fromObject(object: any): void {
        assign(this, object);
    }

}