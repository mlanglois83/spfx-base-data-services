import "@pnp/graph/users";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import { Decorators } from "../../decorators";
import { User } from "../../models";
import { BaseUserService } from "./BaseUserService";


const dataService = Decorators.dataService;
const standardUserCacheDuration = 10;

@dataService("User")
export class UserService extends BaseUserService<User> {
    /**
     * Instanciates a user service
     * @param cacheDuration - cache duration in minutes (default : 10)
     */
    constructor(cacheDuration: number = standardUserCacheDuration, baseUrl?: string, ...args: any[]) {
        super(User, cacheDuration, false, baseUrl, ...args);
    }
    
}
