/**
 * Abstraction class for O365 user associated with a SP User
 */
var User = /** @class */ (function () {
    /**
     * Instancate an user object
     * @param graphUser User object returned by graph api
     */
    function User(graphUser) {
        /**
         * internal field for linked items not stored in db
         */
        this.__internalLinks = {};
        if (graphUser != undefined) {
            this.title = graphUser.displayName != undefined ? graphUser.displayName : "";
            this.id = graphUser.id != undefined ? graphUser.id : "";
            this.mail = graphUser.mail != undefined ? graphUser.mail : "";
            this.userPrincipalName = graphUser.userPrincipalName != undefined ? graphUser.userPrincipalName : "";
        }
    }
    Object.defineProperty(User.prototype, "displayName", {
        /**
         * Get or Set User display name
         */
        get: function () {
            return this.title;
        },
        /**
         * Get or Set User display name
         */
        set: function (val) {
            this.title = val;
        },
        enumerable: true,
        configurable: true
    });
    return User;
}());
export { User };
//# sourceMappingURL=User.js.map