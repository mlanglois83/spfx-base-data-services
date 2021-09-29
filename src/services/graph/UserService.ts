import { Text } from "@microsoft/sp-core-library";
import { cloneDeep, find } from "@microsoft/sp-lodash-subset";
import { stringIsNullOrEmpty } from "@pnp/common/util";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { PictureSize, TestOperator } from "../../constants";
import { Decorators } from "../../decorators";
import { IPredicate, IQuery } from "../../interfaces";
import { User } from "../../models";
import { BaseDataService } from "../base/BaseDataService";
import { UtilsService } from "../UtilsService";


const standardUserCacheDuration = 10;
const dataService = Decorators.dataService;

@dataService("User")
export class UserService extends BaseDataService<User> {
    /**
     * Instanciates a user service
     * @param cacheDuration - cache duration in minutes (default : 10)
     */
    constructor(cacheDuration: number = standardUserCacheDuration) {
        super(User, cacheDuration);
    }

    public async currentUser(extendedProperties: Array<string>): Promise<User> {
        let result: User = null;
        const me = await graph.me.select("displayName", "givenName", "jobTitle", "mail", "mobilePhone", "officeLocation", "preferredLanguage", "surname", "userPrincipalName", "id", ...extendedProperties).get();
        if (me) {
            result = new User(me);
        }
        return result;
    }

    public async currentSPUser(): Promise<User> {
        let result: User = null;
        const me = await sp.web.currentUser.select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").get();
        if (me) {
            result = new User(me);
        }
        return result;
    }

    protected async get_Query(query: IQuery<User>): Promise<Array<any>> {
        let queryStr = (query.test as IPredicate<User, keyof User>).value;
        queryStr = queryStr.trim();
        let reverseFilter = queryStr;
        const parts = queryStr.split(" ");
        if (parts.length > 1) {
            reverseFilter = parts[1].trim() + " " + parts[0].trim();
        }

        const [users, spUsers] = await Promise.all([graph.users
        .filter(
            `startswith(displayName,'${queryStr}') or ` + 
            `startswith(displayName,'${reverseFilter}') or ` +
            `startswith(givenName,'${queryStr}') or ` +
            `startswith(surname,'${queryStr}') or ` +
            `startswith(mail,'${queryStr}') or ` +
            `startswith(userPrincipalName,'${queryStr}')`
        )
        .get(), sp.web.siteUsers.select("Id","UserPrincipalName","Email","Title","IsSiteAdmin").get()]);
        
        return users.map((u) => {
            const spuser = find(spUsers, (spu: any) => { return spu.UserPrincipalName?.toLowerCase() === u.userPrincipalName?.toLowerCase(); });
            if (spuser) {
                u['id'] = spuser.Id;
            }
            return u;
        });
    }


    protected async addOrUpdateItem_Internal(item: User): Promise<User> {
        console.log("[" + this.serviceName + ".addOrUpdateItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItems_Internal(items: Array<User>/*, onItemUpdated?: (oldItem: User, newItem: User) => void*/): Promise<Array<User>> {
        console.log("[" + this.serviceName + ".addOrUpdateItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: User): Promise<User> {
        console.log("[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async deleteItems_Internal(items: Array<User>): Promise<Array<User>> {
        console.log("[" + this.serviceName + ".deleteItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users (sp)
     */
    protected async getAll_Query(): Promise<Array<any>> {
        const spUsers = await sp.web.siteUsers.select("Id", "UserPrincipalName", "Email", "Title", "IsSiteAdmin").get();
        return spUsers.filter(u => !stringIsNullOrEmpty(u.UserPrincipalName));
    }

    public async getItemById_Query(id: number): Promise<any> {
        return sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "Title", "IsSiteAdmin").get();
    }

    public async getItemsById_Query(ids: Array<number>): Promise<Array<any>> {
        const results: Array<any> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 100);
            const batch = sp.createBatch();
            sub.forEach((id) => {
                sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "Title", "IsSiteAdmin").inBatch(batch).get().then((spu) => {
                    if (spu) {
                        results.push(spu);
                    }
                    else {
                        console.log(`[${this.serviceName}] - user with id ${id} not found`);
                    }
                });
            });
            batches.push(batch);
        }
        await UtilsService.runBatchesInStacks(batches, 3);
        return results;
    }

    public async linkToSpUser(user: User): Promise<User> {    
        // user is not registered (or created offline)    
        if (user.id < 0) {
            const allItems = await this.getAll();
            const existing = find(allItems, u => u.userPrincipalName?.toLowerCase() === user.userPrincipalName?.toLowerCase() && u.id > 0);
            // remove existing
            if(existing) {
                user = existing;
            }
            else {
                // remove existing without id
                const sameUsers = allItems.filter(u => u.userPrincipalName?.toLowerCase() === user.userPrincipalName?.toLowerCase());
                if(sameUsers.length > 0) {
                    await this.dbService.deleteItems(sameUsers);
                }
                // register user
                const result = await sp.web.ensureUser(user.userPrincipalName);
                const userItem = await result.user.select("Id", "UserPrincipalName", "Email", "Title", "IsSiteAdmin").get();
                user = new User(userItem);
                // cache 
                const dbresult = await this.dbService.addOrUpdateItem(user);
                user = dbresult;
            }            
        }
        return user;
    }


    public async getByDisplayName(displayName: string): Promise<Array<User>> {
        if(stringIsNullOrEmpty(displayName)) {
            return [];
        }
        let users = await this.get({ test: { type: "predicate", propertyName: "displayName", operator: TestOperator.BeginsWith, value: displayName } });
        if (users.length === 0) {
            users = await this.getAll();

            displayName = displayName.trim();
            let reverseFilter = displayName;
            const parts = displayName.split(" ");
            if (parts.length > 1) {
                reverseFilter = parts[1].trim() + " " + parts[0].trim();
            }
            users = users.filter((user) => {
                return user.displayName?.toLowerCase().indexOf(displayName.toLowerCase()) === 0 ||
                    user.displayName?.toLowerCase().indexOf(reverseFilter.toLowerCase()) === 0 ||
                    user.mail?.toLowerCase().indexOf(displayName.toLowerCase()) === 0 ||
                    user.userPrincipalName?.toLowerCase().indexOf(displayName.toLowerCase()) === 0;
            });
        }
        return users;
    }

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    }
}
