import { IBaseItem } from "../../interfaces/index";
export declare class User implements IBaseItem {
    id: string;
    title: string;
    mail: string;
    spId?: number;
    userPrincipalName: string;
    queries?: Array<number>;
    get displayName(): string;
    set displayName(val: string);
    /***** graph object ******/
    constructor(graphUser?: any);
    convert(): Promise<any>;
}
