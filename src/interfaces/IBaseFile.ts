import { IBaseItem } from "./IBaseItem";

export interface IBaseFile extends IBaseItem {
    content?: ArrayBuffer;
    mimeType: string;
}