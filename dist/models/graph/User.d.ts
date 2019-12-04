import { IBaseItem } from "../../interfaces/index";
/**
 * Abstraction class for O365 user associated with a SP User
 */
export declare class User implements IBaseItem {
    /**
     * internal field for linked items not stored in db
     */
    __internalLinks: any;
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
    get displayName(): string;
    /**
     * Get or Set User display name
     */
    set displayName(val: string);
    /**
     * Instancate an user object
     * @param graphUser User object returned by graph api
     */
    constructor(graphUser?: any);
}
