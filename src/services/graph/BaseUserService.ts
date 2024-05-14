import { stringIsNullOrEmpty } from "@pnp/core";
import "@pnp/graph/users";
import { IPrincipalInfo, PrincipalSource, PrincipalType } from "@pnp/sp";
import "@pnp/sp/site-users";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";
import { cloneDeep, find, isArray } from "lodash";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { PictureSize, TestOperator } from "../../constants";
import { IBaseUserServiceOptions, IPredicate, IQuery } from "../../interfaces";
import { User } from "../../models";
import { UtilsService } from "../UtilsService";
import { BaseSPService } from "../base/BaseSPService";


const standardUserCacheDuration = 10;

export abstract class BaseUserService<T extends User> extends BaseSPService<T> {

    protected serviceOptions: IBaseUserServiceOptions;

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
    constructor(itemType: (new (item?: any) => T), options?: IBaseUserServiceOptions, ...args: any[]) {
        super(itemType, options, ...args);        
        this.serviceOptions.cacheDuration = this.serviceOptions.cacheDuration === undefined ? standardUserCacheDuration : this.serviceOptions.cacheDuration;
    }

    public async currentUser(extendedProperties: Array<string>): Promise<User> {
        let result: User = null;
        const me = await ServicesConfiguration.graph.me.select("displayName", "givenName", "jobTitle", "mail", "mobilePhone", "officeLocation", "preferredLanguage", "surname", "userPrincipalName", "id", ...extendedProperties)();
        if (me) {
            result = new User(me);
        }
        return result;
    }

    public async currentSPUser(): Promise<User> {
        let result: User = null;
        const me = await this.sp.web.currentUser.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin")();
        if (me) {
            result = new User(me);
        }
        return result;
    }

    protected async get_Query(query: IQuery<User>): Promise<Array<any>> {
        const property = (query.test as IPredicate<User, keyof User>).propertyName;
        const operator = (query.test as IPredicate<User, keyof User>).operator;
        if(property === 'displayName' && operator === TestOperator.BeginsWith) {
            return this.searchByDisplayName_Query((query.test as IPredicate<User, keyof User>).value);
        }
        if(operator === TestOperator.In) {
            let userProp: string;
            switch(property) {
                case "displayName":
                    userProp = 'Title';
                    break;
                case "userPrincipalName":
                    userProp = 'UserPrincipalName';
                    break;
                case "id":
                    userProp = 'Id';
                    break;
                default:
                    break
            }
            return this.getByPropertyMulti_query(userProp, (query.test as IPredicate<User, keyof User>).value);
        }
        return [];
        
    }

    protected async searchByDisplayName_Query(displayName): Promise<any[]> {
        let queryStr = displayName;
        queryStr = queryStr.trim();
        let reverseFilter = queryStr;
        const parts = queryStr.split(" ");
        if (parts.length > 1) {
            reverseFilter = parts[1].trim() + " " + parts[0].trim();
        }
        if(ServicesConfiguration.context) {
            const [users, spUsers, cached] = await Promise.all([ServicesConfiguration.graph.users
                .filter(
                    `startswith(displayName,'${queryStr}') or ` +
                    `startswith(displayName,'${reverseFilter}') or ` +
                    `startswith(givenName,'${queryStr}') or ` +
                    `startswith(surname,'${queryStr}') or ` +
                    `startswith(mail,'${queryStr}') or ` +
                    `startswith(userPrincipalName,'${queryStr}')`
                )
                (), 
                this.sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin")(),
                this.hasCache ? this.cacheService.getAll() : Promise.resolve([])
            ]);

            return users.map((u: any) => {
                const spuser = find(spUsers, (spu: ISiteUserInfo) => { return spu[this.spUserField]?.toLowerCase() === u[BaseUserService.userField]?.toLowerCase(); });
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
                this.sp.utility.searchPrincipals(queryStr, (PrincipalType.User | (this.serviceOptions.includeGroups ? PrincipalType.SecurityGroup : PrincipalType.None)) , PrincipalSource.All,"", 15),
                this.sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin")(),
                this.hasCache ? this.cacheService.getAll() : Promise.resolve([])
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

    protected getFilterValueString<TVal extends string | number>(value: TVal) {
        if(typeof(value) === "string") {
            return `'${value}'`
        }
        else return value.toString();
    }

    protected async getByPropertyMulti_query<TVal extends string | number>(propertyName: string, values: TVal[]) {
        const results: Array<any> = [];
        const copy = cloneDeep(values);
        const batchesIndex = [];
        while (copy.length > 0) {
            const sub = copy.splice(0, 5);
            batchesIndex.push(sub);
        }
        if(ServicesConfiguration.configuration.spVersion !== "SP2013") {
            const batches = [];            
            const copyIndex = cloneDeep(batchesIndex);
            while (copyIndex.length > 0) {
                const sub = copyIndex.splice(0, 100);
                const [batchedSP, execute] = this.sp.batched();
                sub.forEach((dns: string[]) => {
                    batchedSP.web.siteUsers.filter(dns.map(dn => `${propertyName} eq ${this.getFilterValueString(dn)}`).join(' or ')).select("Id", "UserPrincipalName", "Email", "Title", "LoginName" , "IsSiteAdmin")().then((spusers) => {
                        if (spusers) {
                            results.push(...spusers);
                        }
                    });
                });
                batches.push(execute);
            }
            await UtilsService.runBatchesInStacks(batches, 3);
        }
        else {
            const promises = batchesIndex.map(dns => ((): Promise<ISiteUserInfo[]> => this.sp.web.siteUsers.filter(dns.map(dn => `${propertyName} eq ${this.getFilterValueString(dn)}`).join(' or ')).select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin")()));
            const responses = await UtilsService.executePromisesInStacks(promises, 3);
            responses.forEach((spus) => {
                if (spus) {
                    results.push(...spus);
                }
            });
        }
        return results;
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
        const spUsers = await this.sp.web.siteUsers.select("Id", "UserPrincipalName", "LoginName", "Email", "Title", "IsSiteAdmin")();
        if (this.serviceOptions.includeGroups)
            return spUsers;
        else
            return spUsers.filter(u => ServicesConfiguration.configuration.spVersion === "Online" ? !stringIsNullOrEmpty(u.UserPrincipalName) : u.LoginName.indexOf("i:0#") === 0);
    }

    public async getItemById_Query(id: number): Promise<any> {
        return this.sp.web.siteUsers.getById(id).select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin")();
    }

    public async getItemsById_Query(ids: Array<number>): Promise<Array<any>> {
        return this.get_Query({ 
            test: { 
                type: "predicate", 
                propertyName: "id", 
                operator: TestOperator.In, 
                value: ids 
            }
        });        
    }


    public async linkToSpUser(user: T): Promise<T> {
        // user is not registered (or created offline)    
        if (user.isLocal) {
            const allItems = await this.getAll();
            const existing = find(allItems, u => u[BaseUserService.userField]?.toLowerCase() === user[BaseUserService.userField]?.toLowerCase() && !u.isLocal);
            // remove existing
            if (existing) {
                user = existing;
            }
            else {
                // remove existing without id
                const sameUsers = allItems.filter(u => u[BaseUserService.userField]?.toLowerCase() === user[BaseUserService.userField]?.toLowerCase());
                if (sameUsers.length > 0 && this.hasCache) {
                    await this.cacheService.deleteItems(sameUsers);
                }
                // register user
                const result = await this.sp.web.ensureUser(user[BaseUserService.userField]);
                const userItem = await result.user.select("Id", "UserPrincipalName", "Email", "LoginName", "Title", "IsSiteAdmin")();
                user = new this.itemType(userItem);
                // cache 
                if(this.hasCache) {
                    user = await this.cacheService.addOrUpdateItem(user);
                }
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

    public async getByEmails(emails: Array<string>): Promise<Array<T>> {
        return this.get({ 
            test: { 
                type: "predicate", 
                propertyName: "userPrincipalName", 
                operator: TestOperator.In, 
                value: emails 
            }
        });        
    }
    public async getByDisplayNames(displayNames: Array<string>): Promise<Array<T>> {
        return this.get({ 
            test: { 
                type: "predicate", 
                propertyName: "displayName", 
                operator: TestOperator.In, 
                value: displayNames 
            }
        });        
    }

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user[BaseUserService.userField] ? UtilsService.formatText("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.baseUrl, encodeURIComponent(user[BaseUserService.userField]), size) : "";
    }
}
