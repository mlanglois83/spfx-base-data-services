import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
export declare class UserService extends BaseDataService<User> {
    private _spUsers;
    private spUsers;
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    constructor(cacheDuration?: number);
    protected get_Internal(query: any): Promise<Array<User>>;
    protected addOrUpdateItem_Internal(item: User): Promise<User>;
    protected addOrUpdateItems_Internal(items: Array<User>): Promise<Array<User>>;
    protected deleteItem_Internal(item: User): Promise<void>;
    /**
     * Retrieve all users (sp)
     */
    protected getAll_Internal(): Promise<Array<User>>;
    getItemById_Internal(id: string): Promise<User>;
    getItemsById_Internal(ids: Array<string>): Promise<Array<User>>;
    linkToSpUser(user: User): Promise<User>;
    getByDisplayName(displayName: string): Promise<Array<User>>;
    getBySpId(spId: number): Promise<User>;
    static getPictureUrl(user: User, size?: PictureSize): string;
}
