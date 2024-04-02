import { Decorators } from "../../decorators";
import { BaseNumberItem } from "../base/BaseNumberItem";
const dataModel = Decorators.dataModel;
/**
 * Abstraction class for O365 user associated with a SP User
 */
@dataModel()
export class User extends BaseNumberItem {
    /**
     * id of the user
     */
    public id: number;
    /**
     * Graph id of the user
     */
    public o365id: string;
    /**
     * User display name
     */
    public title: string;
    /**
     * User email
     */
    public mail: string;
    /**
     * User principal name (login)
     */
    public userPrincipalName: string;

    public loginName: string;
    /*
    * User is site admin
    */
    public isSiteAdmin = false;

    public firstName: string;
    public lastName: string;

    public extendedProperties: Map<string, any>;
    /**
     * Get or Set User display name
     */
    public get displayName(): string {
        return this.title;
    }
    /**
     * Get or Set User display name
     */
    public set displayName(val: string) {
        this.title = val;
    }

    /**
     * Get User login name without claims
     */
    public get cleanLoginName(): string {
        return this.loginName?.replace(/(.*\|)?([^|]+)/, "$2");
    }

    /**
     * Get User login name without claims and domain
     */
     public get cleanLoginNameNoDomain(): string {
        return this.loginName?.replace(/(.*\|)?([^|\\]+)\\?([^|\\]+)/, "$3");
    }
    
    

    /**
     * Instancate an user object
     * @param userObj - user object returned by graph api or sp
     */
    constructor(userObj?: any) {
        super();
        if (userObj) {
            this.title = userObj.displayName ? userObj.displayName : (userObj.Title ? userObj.Title : "");
            this.id = userObj.Id ? userObj.Id : -1;
            this.o365id = userObj.id ? userObj.id : "";
            this.mail = userObj.mail ? userObj.mail : (userObj.Email ? userObj.Email : "");
            this.userPrincipalName = userObj.userPrincipalName ? userObj.userPrincipalName : (userObj.UserPrincipalName ? userObj.UserPrincipalName : "");
            this.isSiteAdmin = userObj.IsSiteAdmin === true;
            this.loginName = userObj.loginName ? userObj.loginName : (userObj.LoginName ? userObj.LoginName : "");

            this.firstName = userObj.givenName;
            this.lastName = userObj.surname;

            this.extendedProperties = new Map<string, any>();

            for (const key of Object.keys(userObj)) {
                this.extendedProperties.set(key, userObj[key]);
            }
        }
    }

}