import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";

const standardUserCacheDuration: number = 10;
export class UserService<T extends User> extends BaseDataService<T> {
    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param termsetname termset name
     */
    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration: number = standardUserCacheDuration) {
        super(type, tableName, cacheDuration);
    }

    protected async get_Internal(query: any): Promise<Array<T>> {
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


    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<void> {
        throw new Error("Not implemented");
    }

    /**
     * Retrieve all users
     */
    protected async getAll_Internal(): Promise<Array<T>> {
       let users = await graph.users.get();
       return users.map((u) => { 
        return new this.itemType(u);
       });                    
    }

    public async getById_Internal(id: string): Promise<T> {
        let graphUser = await graph.users.getById(id).get();
        return new this.itemType(graphUser);
    }

    public async linkToSpUser(user: T): Promise<T> {
        if(user.spId === undefined) {
            let result = await sp.web.ensureUser(user.userPrincipalName);
            user.spId = result.data.Id
            this.dbService.addOrUpdateItem(user);
        }
        return user;        
    }

    prot

    public async getByDisplayName(displayName: string): Promise<Array<T>> {
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

    public static getPictureUrl(user: User, size: PictureSize = PictureSize.Large): string {
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    }
  }
}