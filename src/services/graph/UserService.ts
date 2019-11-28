import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { Constants } from "../../constants";
import { UtilsService } from "../UtilsService";
import { find } from "@microsoft/sp-lodash-subset";

const standardUserCacheDuration: number = 10;
export class UserService extends BaseDataService<User> {
    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param termsetname termset name
     */
    constructor(cacheDuration: number = standardUserCacheDuration) {
        super(User, "Users", cacheDuration);
    }

    protected async get_Internal(query: any): Promise<Array<User>> {
        query = query.trim();
        let reverseFilter = query;
        let parts = query.split(" ");
        if (parts.length > 1) {
            reverseFilter = parts[1].trim() + " " + parts[0].trim();
        }
        let users = await graph.users
        .filter(
            `startswith(displayName,'${query}') or 
            startswith(displayName,'${reverseFilter}') or 
            startswith(givenName,'${query}') or 
            startswith(surname,'${query}') or 
            startswith(mail,'${query}') or 
            startswith(userPrincipalName,'${query}')`
        )
        .get();
        return users.map((u) => { 
            return new this.itemType(u);
        });
    }


    protected async addOrUpdateItem_Internal(item: User): Promise<User> {
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: User): Promise<void> {
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users from site
     */
    protected async getAll_Internal(): Promise<Array<User>> {
        let results = [];
        let spUsers  = await sp.web.siteUsers.get();
        let batch = graph.createBatch();
        spUsers.forEach((spu) => {
            if(spu.UserPrincipalName) {
                graph.users.filter(`userPrincipalName eq '${encodeURIComponent(spu.UserPrincipalName)}'`).inBatch(batch).get().then((graphUser) => {
                    if(graphUser && graphUser.length > 0) {
                        let result = new User(graphUser[0]);
                        result.spId = spu.Id;
                        results.push(result);
                    }
                });
            }
        });
        await batch.execute();
        return results;       
    }

    public async getById_Internal(id: string): Promise<User> {
        let graphUser = await graph.users.getById(id).get();
        return new this.itemType(graphUser);
    }

    public async linkToSpUser(user: User): Promise<User> {
        if(user.spId === undefined) {
            let result = await sp.web.ensureUser(user.userPrincipalName);
            user.spId = result.data.Id;
            this.dbService.addOrUpdateItem(user);
        }
        return user;        
    }


    public async getByDisplayName(displayName: string): Promise<Array<User>> {
        let users = await this.get(displayName);
        if(users.length == 0) {
            users = await this.getAll();

            displayName = displayName.trim();
            let reverseFilter = displayName;
            let parts = displayName.split(" ");
            if (parts.length > 1) {
                reverseFilter = parts[1].trim() + " " + parts[0].trim();
            }

            users = users.filter((user) => {
                return user.displayName.indexOf(displayName) === 0 ||
                user.displayName.indexOf(reverseFilter) === 0 ||
                user.mail.indexOf(displayName) === 0 ||
                user.userPrincipalName.indexOf(displayName) === 0;
            });
        }
        return users;
    }

    public async getBySpId(spId: number): Promise<User> {
        let allUsers = await this.getAll();
        return find(allUsers, (user) => {return user.spId === spId;});
    }

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    }
}
