import { IBaseItem } from "../../interfaces/index";
/**
 * Abstraction class for O365 user associated with a SP User
 */
export declare class User implements IBaseItem {
    /**
     * O365 id of the user
     */
    id: string;
    /**
     * User display name
     */
    title: string;
    /**
     * User email
     */
    mail: string;
    /**
     * Associated SP User ID
     */
    spId?: number;
    /**
     * User principal name (login)
     */
    userPrincipalName: string;
    /**
     * Queries used to retrieve user (only used in data services)
     */
    queries?: Array<number>;
    /**
     * Get or Set User display name
     */
    /**
    * Get or Set User display name
    */
    displayName: string;
    /**
     * Instancate an user object
     * @param graphUser User object returned by graph api
     */
    constructor(graphUser?: any);
}
