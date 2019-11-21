import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { find } from "@microsoft/sp-lodash-subset";

const standardUserCacheDuration: number = 10;
export class UserService extends BaseDataService<User> {

    private _spUsers: Array<any> = null;
    private async spUsers(): Promise<Array<any>>{
        if(this._spUsers === null) {
            this._spUsers = await sp.web.siteUsers.select("UserPrincipalName", "Id").get();
        }
        return this._spUsers;
    }


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
        let [users, spUsers] = await Promise.all([graph.users
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
            let spuser = find(spUsers, (spu: any) => {return spu.UserPrincipalName === u.userPrincipalName;});
            let result =  new User(u);
            if(spuser) {
                result.spId = spuser.Id;
            }
            return result;
        });
    }


    protected async addOrUpdateItem_Internal(item: User): Promise<User> {
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: User): Promise<void> {
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users (sp)
     */
    protected async getAll_Internal(): Promise<Array<User>> {
        let spUsers = await this.spUsers();
        let results: Array<User> = [];
        let batch = graph.createBatch();
        spUsers.forEach((spu) => {
            graph.users.select("id","userPrincipalName","mail","displayName").filter(`userPrincipalName eq '${spu.UserPrincipalName}'`).inBatch(batch).get().then((graphUsers)=> {
                if(graphUsers && graphUsers.length > 0) {
                    let graphUser = graphUsers[0];
                    let spuser = find(spUsers, (spu)=> {
                        return spu.UserPrincipalName === graphUser.userPrincipalName;
                    });
                    let result= new User(graphUser);
                    if(spuser) {
                        result.spId = spuser.Id;
                    }
                    results.push(result);
                }
            })
        });
        await batch.execute();
        return results;         
    }

    public async getItemById_Internal(id: string): Promise<User> {
        let result = null;
        let [graphUser, spUsers] = await Promise.all([graph.users.getById(id).select("id","userPrincipalName","mail","displayName").get(), this.spUsers]);
        if(graphUser) {
            let spuser = find(spUsers, (spu: any)=> {
                return spu.UserPrincipalName === graphUser.userPrincipalName;
            });
             let result= new User(graphUser);
             if(spuser) {
                 result.spId = spuser.Id;
             }
        }
        return result;
    }

    public async getItemsById_Internal(ids: Array<string>): Promise<Array<User>> {
        let [graphUsers, spUsers] = await Promise.all([graph.users.filter(ids.map((id) => { return `id eq '${id}'`}).join(' or ')).select("id","userPrincipalName","mail","displayName").get(), this.spUsers]);
        return graphUsers.map((u) => { 
            let spuser = find(spUsers, (spu: any)=> {
                return spu.UserPrincipalName === u.userPrincipalName;
            });
            let result= new User(u);
            if(spuser) {
                result.spId = spuser.Id;
            }
            return result;
        });          
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
        if(users.length === 0) {
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
