import "@pnp/graph/users";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import { Decorators } from "../../decorators";
import { User } from "../../models";
import { BaseUserService } from "./BaseUserService";
import { IBaseUserServiceOptions } from "../../interfaces";


const dataService = Decorators.dataService;

@dataService("User")
export class UserService extends BaseUserService<User> {
    /**
     * Instanciates a user service
     * @param cacheDuration - cache duration in minutes (default : 10)
     */
    constructor(options?: IBaseUserServiceOptions, ...args: any[]) {
        super(User, options, ...args);
    }
    
}
