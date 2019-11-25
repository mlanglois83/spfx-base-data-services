import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
export declare class UserService extends BaseDataService<User> {
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    constructor(cacheDuration?: number);
    protected get_Internal(query: any): Promise<Array<User>>;
    protected addOrUpdateItem_Internal(item: User): Promise<User>;
    protected deleteItem_Internal(item: User): Promise<void>;
    /**
     * Retrieve all users from site
     */
    protected getAll_Internal(): Promise<Array<User>>;
    getById_Internal(id: string): Promise<User>;
    linkToSpUser(user: User): Promise<User>;
    getByDisplayName(displayName: string): Promise<Array<User>>;
    getBySpId(spId: number): Promise<User>;
    static getPictureUrl(user: User, size?: PictureSize): string;
}
