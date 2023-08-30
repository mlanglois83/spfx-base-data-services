import { IBaseItem } from "./IBaseItem";

export interface IBaseFile<T extends string | number> extends IBaseItem<T> {
    serverRelativeUrl?: string;
    content?: ArrayBuffer;
    mimeType: string;
}