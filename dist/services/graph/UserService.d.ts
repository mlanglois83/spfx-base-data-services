import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
export declare class UserService<T extends User> extends BaseDataService<T> {
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration?: number);
    protected get_Internal(query: any): Promise<Array<T>>;
    protected addOrUpdateItem_Internal(item: T): Promise<T>;
    protected deleteItem_Internal(item: T): Promise<void>;
    /**
     * Retrieve all users
     */
    protected getAll_Internal(): Promise<Array<T>>;
    getById_Internal(id: string): Promise<T>;
    linkToSpUser(user: T): Promise<T>;
    prot: any;
    getByDisplayName(displayName: string): Promise<Array<T>>;
    static getPictureUrl(user: User, size?: PictureSize): string;
}
