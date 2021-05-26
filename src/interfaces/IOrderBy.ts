import { IBaseItem } from "./IBaseItem";

export interface IOrderBy<T extends IBaseItem, K extends keyof T> {
    type: "orderby";
    propertyName: K;
    ascending?: boolean;
}