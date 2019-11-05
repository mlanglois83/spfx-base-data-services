var User = /** @class */ (function () {
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
    function User(graphUser) {
        if (graphUser != undefined) {
            this.title = graphUser.displayName != undefined ? graphUser.displayName : "";
            this.id = graphUser.id != undefined ? graphUser.id : "";
            this.mail = graphUser.mail != undefined ? graphUser.mail : "";
            this.userPrincipalName = graphUser.userPrincipalName != undefined ? graphUser.userPrincipalName : "";
        }
    }
    Object.defineProperty(User.prototype, "displayName", {
        get: function () {
            return this.title;
        },
        set: function (val) {
            this.title = val;
        },
        enumerable: true,
        configurable: true
    });
    User.prototype.convert = function () {
        throw new Error("Not implemented");
    };
    return User;
}());
export { User };
//# sourceMappingURL=User.js.map