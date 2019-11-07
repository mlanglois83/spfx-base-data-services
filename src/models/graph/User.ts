import { IBaseItem } from "../../interfaces/index";
export class User implements IBaseItem {

    public id: string;
    public title: string;
    public mail: string;
    public spId?: number;
    public userPrincipalName: string;
    public queries?: Array<number>;

    public get displayName(): string {
        return this.title;
    }
    public set displayName(val: string) {
        this.title = val;
    }

    /***** graph object ******/
    /*"businessPhones": [],
    "displayName": "Conf Room Adams",
    "givenName": null,
    "jobTitle": null,
    "mail": "Adams@M365x214355.onmicrosoft.com",
    "mobilePhone": null,
    "officeLocation": null,
    "preferredLanguage": null,
    "surname": null,
    "userPrincipalName": "Adams@M365x214355.onmicrosoft.com",
    "id": "6e7b768e-07e2-4810-8459-485f84f8f204"*/

    constructor(graphUser?: any) {
        if (graphUser != undefined) {
            this.title = graphUser.displayName != undefined ? graphUser.displayName : "";
            this.id = graphUser.id != undefined ? graphUser.id  : "";
            this.mail = graphUser.mail != undefined ? graphUser.mail : "";
            this.userPrincipalName = graphUser.userPrincipalName != undefined ? graphUser.userPrincipalName : "";
        }
    }
    public convert(): Promise<any> {
        throw new Error("Not implemented");
    }

}