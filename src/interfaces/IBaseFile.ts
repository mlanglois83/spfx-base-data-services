import { IBaseItem } from "./IBaseItem";

export interface IBaseFile extends IBaseItem {
    serverRelativeUrl?: string;
    content?: ArrayBuffer;
    mimeType: string;
}