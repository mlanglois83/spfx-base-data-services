import { stringIsNullOrEmpty } from "@pnp/common/util";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import { IPrincipalInfo, PrincipalSource, PrincipalType, sp } from "@pnp/sp";
import "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";
import { cloneDeep, find, isArray } from "lodash";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { PictureSize, TestOperator } from "../../constants";
import { IPredicate, IQuery } from "../../interfaces";
import { User } from "../../models";
import { BaseDataService } from "../base/BaseDataService";
import { UtilsService } from "../UtilsService";


const standardUserCacheDuration = 10;

export abstract class BaseUserService<T extends User> extends BaseDataService<T> {

    protected groupsToo = false;

    public static get userField(): keyof Pick<User, "userPrincipalName" | "loginName"> {
        return ServicesConfiguration.configuration.spVersion === "Online" ? "userPrincipalName" : "loginName";
    }

    protected get spUserField(): keyof Pick<ISiteUserInfo, "UserPrincipalName" | "LoginName"> {
        return ServicesConfiguration.configuration.spVersion === "Online" ? "UserPrincipalName" : "LoginName";
    }

    /**
     * Instanciates a user service
     * @param cacheDuration - cache duration in minutes (default : 10)
     */
    constructor(type: new (item?: any) => T, cacheDuration: number = standardUserCacheDuration, groupsToo = false) {
        super(type, cacheDuration);
        this.groupsToo = groupsToo;
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
        const me = await sp.web.currentUser.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin").get();
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
        if(ServicesConfiguration.context) {
            const [users, spUsers, cached] = await Promise.all([graph.users
                .filter(
                    `startswith(displayName,'${queryStr}') or ` +
                    `startswith(displayName,'${reverseFilter}') or ` +
                    `startswith(givenName,'${queryStr}') or ` +
                    `startswith(surname,'${queryStr}') or ` +
                    `startswith(mail,'${queryStr}') or ` +
                    `startswith(userPrincipalName,'${queryStr}')`
                )
                .get(), 
                sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin").get(),
                this.dbService.getAll()
            ]);

            return users.map((u: any) => {
                const spuser = find(spUsers, (spu: ISiteUserInfo) => { return spu[this.spUserField]?.toLowerCase() === u[this.spUserField]?.toLowerCase(); });
                const cachedUser = find(cached, (spu) => { return spu[BaseUserService.userField]?.toLowerCase() === u[this.spUserField]?.toLowerCase(); });
                if (spuser) {
                    u.id = spuser.Id;
                }
                else if(cachedUser) {
                    u.id = cachedUser.id;
                }
                return u;
            });
        }
        else {
            const [searchResults, spUsers, cached] = await Promise.all([
                sp.utility.searchPrincipals(queryStr, (PrincipalType.User | (this.groupsToo ? PrincipalType.SecurityGroup : PrincipalType.None)) , PrincipalSource.All,"", 15),
                sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin").get(),
                this.dbService.getAll()
            ]);
            let searchConv = searchResults;
            if(!isArray(searchConv)) // parsing error
            {
                searchConv = ((searchConv as any).SearchPrincipalsUsingContextWeb?.results as Array<IPrincipalInfo>) || [];
            }
            return searchConv.map((sr): Partial<ISiteUserInfo> => {
                const spuser = find(spUsers, (spu: any) => { return spu[this.spUserField]?.toLowerCase() === sr.LoginName.toLowerCase(); });
                const cachedUser = find(cached, (spu) => { return spu.loginName?.toLowerCase() === sr.LoginName.toLowerCase(); });
                return {
                    Id: spuser ? spuser.Id : (cachedUser ? cachedUser.id : sr.PrincipalId),
                    LoginName: sr.LoginName,
                    Title: sr.DisplayName,
                    Email: sr.Email,
                    IsSiteAdmin: spuser ? spuser.IsSiteAdmin : false
                };
            }) || [];
        }
    }


    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        console.log("[" + this.serviceName + ".addOrUpdateItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItems_Internal(items: Array<T>/*, onItemUpdated?: (oldItem: User, newItem: User) => void*/): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".addOrUpdateItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<T> {
        console.log("[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".deleteItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async recycleItem_Internal(item: T): Promise<T> {
        console.log("[" + this.serviceName + ".recycleItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async recycleItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".recycleItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users (sp)
     */
    protected async getAll_Query(): Promise<Array<any>> {
        const spUsers = await sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin").get();
        if (this.groupsToo)
            return spUsers;
        else
            return spUsers.filter(u => ServicesConfiguration.configuration.spVersion === "Online" ? !stringIsNullOrEmpty(u.UserPrincipalName) : u.LoginName.indexOf("i:0#") === 0);
    }

    public async getItemById_Query(id: number): Promise<any> {
        return sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin").get();
    }

    public async getItemsById_Query(ids: Array<number>): Promise<Array<any>> {
        // TODO ON PREM
        const results: Array<any> = [];
        if(ServicesConfiguration.configuration.spVersion !== "SP2013") {
            const batches = [];
            const copy = cloneDeep(ids);
            while (copy.length > 0) {
                const sub = copy.splice(0, 100);
                const batch = sp.createBatch();
                sub.forEach((id) => {
                    sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin").inBatch(batch).get().then((spu) => {
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
        }
        else {
            const promises = ids.map(id => ((): Promise<ISiteUserInfo> => sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin").get()));
            const responses = await UtilsService.executePromisesInStacks(promises, 3);
            responses.forEach((spu, idx) => {
                if (spu) {
                    results.push(spu);
                }
                else {
                    console.log(`[${this.serviceName}] - user with id ${ids[idx]} not found`);
                }
            });
        }
        return results;
        
    }

    public async linkToSpUser(user: T): Promise<T> {
        // user is not registered (or created offline)    
        if (user.id < 0) {
            const allItems = await this.getAll();
            const existing = find(allItems, u => u[BaseUserService.userField]?.toLowerCase() === user[BaseUserService.userField]?.toLowerCase() && u.id > 0);
            // remove existing
            if (existing) {
                user = existing;
            }
            else {
                // remove existing without id
                const sameUsers = allItems.filter(u => u[BaseUserService.userField]?.toLowerCase() === user[BaseUserService.userField]?.toLowerCase());
                if (sameUsers.length > 0) {
                    await this.dbService.deleteItems(sameUsers);
                }
                // register user
                const result = await sp.web.ensureUser(user[BaseUserService.userField]);
                const userItem = await result.user.select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin").get();
                user = new this.itemType(userItem);
                // cache 
                const dbresult = await this.dbService.addOrUpdateItem(user);
                user = dbresult;
            }
        }
        return user;
    }


    public async getByDisplayName(displayName: string): Promise<Array<User>> {
        if (stringIsNullOrEmpty(displayName)) {
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
                    user.userPrincipalName?.toLowerCase().indexOf(displayName.toLowerCase()) === 0 ||
                    user.cleanLoginName?.toLowerCase().indexOf(displayName.toLowerCase()) === 0 ||
                    user.cleanLoginNameNoDomain?.toLowerCase().indexOf(displayName.toLowerCase()) === 0;
            });
        }
        return users;
    }

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user[BaseUserService.userField] ? UtilsService.formatText("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.baseUrl, encodeURIComponent(user[BaseUserService.userField]), size) : "";
    }
}
