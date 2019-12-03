import { IBaseItem } from "../../interfaces/index";
/**
 * Abstraction class for O365 user associated with a SP User
 */
export class User implements IBaseItem {
    /**
     * O365 id of the user
     */
    public id: string;
    /**
     * User display name
     */
    public title: string;
    /**
     * User email
     */
    public mail: string;
    /**
     * Associated SP User ID
     */
    public spId?: number;
    /**
     * User principal name (login)
     */
    public userPrincipalName: string;
    /**
     * Queries used to retrieve user (only used in data services)
     */
    public queries?: Array<number>;
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
     * Instancate an user object
     * @param graphUser User object returned by graph api
     */
    constructor(graphUser?: any) {
        if (graphUser) {
            this.title = graphUser.displayName ? graphUser.displayName : "";
            this.id = graphUser.id ? graphUser.id  : "";
            this.mail = graphUser.mail ? graphUser.mail : "";
            this.userPrincipalName = graphUser.userPrincipalName ? graphUser.userPrincipalName : "";
        }
    }

}