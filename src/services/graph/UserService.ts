import { BaseDataService } from "..";
import { User, PictureSize, IQuery } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { find, cloneDeep } from "@microsoft/sp-lodash-subset";
import { TestOperator } from "../../constants";
import { IPredicate } from "../../interfaces";

const standardUserCacheDuration = 10;
export class UserService extends BaseDataService<User> {
    /**
     * Instanciates a user service
     * @param cacheDuration - cache duration in minutes (default : 10)
     */
    constructor(cacheDuration: number = standardUserCacheDuration) {
        super(User, "Users", cacheDuration);
    }

    protected async get_Internal(query: IQuery): Promise<Array<User>> {
        let queryStr = (query.test as IPredicate).value;
        queryStr = queryStr.trim();
        let reverseFilter = queryStr;
        const parts = queryStr.split(" ");
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
        .get(), sp.web.siteUsers.select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").get()]);
        
        return users.map((u) => {
            const spuser = find(spUsers, (spu: any) => {return spu.UserPrincipalName === u.userPrincipalName;});
            const result =  new User(u);
            if(spuser) {
                result.id = spuser.Id;
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
        const spUsers  = await sp.web.siteUsers.select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").get();
        return spUsers.map(spu => new User(spu));
    }

    public async getItemById_Internal(id: number): Promise<User> {
        const spu = await sp.web.siteUsers.getById(id).select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").get();
        if(spu)
            return new User(spu);
        return null;
    }

    public async getItemsById_Internal(ids: Array<number>): Promise<Array<User>> {
        const results: Array<User> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while(copy.length > 0) {
            const sub = copy.splice(0,100);
            const batch = sp.createBatch();
            sub.forEach((id) => {
                sp.web.siteUsers.getById(id).select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").inBatch(batch).get().then((spu) => {                
                    const result= new User(spu);
                    results.push(result);
                });
            });
            batches.push(batch);
        }        
        while(batches.length > 0) {
            const sub = batches.splice(0,5);
            await Promise.all(sub.map(b => b.execute()));
        }
        return results;    
    }

    public async linkToSpUser(user: User): Promise<User> {
        if(user.id === -1) {
            const result = await sp.web.ensureUser(user.userPrincipalName);
            user.id = result.data.Id;
            const dbresult = await this.dbService.addOrUpdateItem(user);
            user = dbresult;
        }
        return user;        
    }


    public async getByDisplayName(displayName: string): Promise<Array<User>> {
        let users = await this.get({test: {type: "predicate", propertyName: "displayName", operator: TestOperator.BeginsWith, value: displayName}});
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

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    }
}
