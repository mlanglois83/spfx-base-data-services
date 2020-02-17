import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { find } from "@microsoft/sp-lodash-subset";

const standardUserCacheDuration = 10;
export class UserService extends BaseDataService<User> {

    private _spUsers: Array<any> = null;
    private async spUsers(): Promise<Array<any>>{
        if(this._spUsers === null) {
            this._spUsers = await sp.web.siteUsers.select("UserPrincipalName", "Id").get();
        }
        return this._spUsers;
    }


    /**
     * Instanciates a user service
     * @param cacheDuration Cache duration in minutes (default : 10)
     */
    constructor(cacheDuration: number = standardUserCacheDuration) {
        super(User, "Users", cacheDuration);
    }

    protected async get_Internal(query: any): Promise<Array<User>> {
        query = query.trim();
        let reverseFilter = query;
        const parts = query.split(" ");
        if (parts.length > 1) {
        reverseFilter = parts[1].trim() + " " + parts[0].trim();
        }
        const [users, spUsers] = await Promise.all([graph.users
        .filter(
            `startswith(displayName,'${query}') or 
            startswith(displayName,'${reverseFilter}') or 
            startswith(givenName,'${query}') or 
            startswith(surname,'${query}') or 
            startswith(mail,'${query}') or 
            startswith(userPrincipalName,'${query}')`
        )
        .get(), this.spUsers]);
        
        return users.map((u) => {
            const spuser = find(spUsers, (spu: any) => {return spu.UserPrincipalName === u.userPrincipalName;});
            const result =  new User(u);
            if(spuser) {
                result.spId = spuser.Id;
            }
            return result;
        });
    }


    protected async addOrUpdateItem_Internal(item: User): Promise<User> {
        console.log("[" + this.serviceName + ".addOrUpdateItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }
    
    protected async addOrUpdateItems_Internal(items: Array<User>): Promise<Array<User>> {
        console.log("[" + this.serviceName + ".addOrUpdateItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: User): Promise<void> {
        console.log("[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users (sp)
     */
    protected async getAll_Internal(): Promise<Array<User>> {
        const results = [];
        const spUsers  = await this.spUsers();
        const batch = graph.createBatch();
        spUsers.forEach((spu) => {
            if(spu.UserPrincipalName) {
                graph.users.select("id","userPrincipalName","mail","displayName").filter(`userPrincipalName eq '${encodeURIComponent(spu.UserPrincipalName)}'`).inBatch(batch).get().then((graphUser) => {
                    if(graphUser && graphUser.length > 0) {
                        const result = new User(graphUser[0]);
                        result.spId = spu.Id;
                        results.push(result);
                    }
                });
            }
        });
        await batch.execute();
        return results;       
    }

    public async getItemById_Internal(id: string): Promise<User> {
        let result = null;
        const [graphUser, spUsers] = await Promise.all([graph.users.getById(id).select("id","userPrincipalName","mail","displayName").get(), this.spUsers]);
        if(graphUser) {
            const spuser = find(spUsers, (spu: any)=> {
                return spu.UserPrincipalName === graphUser.userPrincipalName;
            });
             result= new User(graphUser);
             if(spuser) {
                 result.spId = spuser.Id;
             }
        }
        return result;
    }

    public async getItemsById_Internal(ids: Array<string>): Promise<Array<User>> {
        const results: Array<User> = [];
        const spUsers = await this.spUsers();
        const batch = graph.createBatch();
        ids.forEach(id => {
            graph.users.getById(id).select("id","userPrincipalName","mail","displayName").inBatch(batch).get().then((graphUser) => {
                const spuser = find(spUsers, (spu: any)=> {
                    return spu.UserPrincipalName === graphUser.userPrincipalName;
                });
                const result= new User(graphUser);
                if(spuser) {
                    result.spId = spuser.Id;
                }
                results.push(result);
            });
        });
        await batch.execute();
        return results;    
    }

    public async linkToSpUser(user: User): Promise<User> {
        if(user.spId === undefined) {
            const result = await sp.web.ensureUser(user.userPrincipalName);
            user.spId = result.data.Id;
            this.dbService.addOrUpdateItem(user);
        }
        return user;        
    }


    public async getByDisplayName(displayName: string): Promise<Array<User>> {
        let users = await this.get(displayName);
        if(users.length === 0) {
            users = await this.getAll();

            displayName = displayName.trim();
            let reverseFilter = displayName;
            const parts = displayName.split(" ");
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
        const allUsers = await this.getAll();
        return find(allUsers, (user) => {return user.spId === spId;});
    }

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    }
}
