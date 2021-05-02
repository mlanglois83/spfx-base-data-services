import { IBaseItem } from "../../interfaces";
/**
 * Abstraction class for O365 user associated with a SP User
 */
export class User implements IBaseItem {
    /**
     * id of the user
     */
    public id = -1;
    /**
     * id of the user
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
    /*
    * User is site admin
    */
    public isSiteAdmin = false;
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
     * @param userObj - user object returned by graph api or sp
     */
    constructor(userObj?: any) {
        if (userObj) {
            this.title = userObj.displayName ? userObj.displayName : (userObj.Title ? userObj.Title : "");
            this.id = userObj.Id ? userObj.Id  : -1;
            this.o365id = userObj.id ? userObj.id : '';
            this.mail = userObj.mail ? userObj.mail : (userObj.Email ? userObj.Email : "");
            this.userPrincipalName = userObj.userPrincipalName ? userObj.userPrincipalName : (userObj.UserPrincipalName ? userObj.UserPrincipalName : "");            
            this.isSiteAdmin = userObj.IsSiteAdmin === true;
        }
    }

}