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
        this.__internalLinks = undefined;
        if (graphUser) {
            this.title = graphUser.displayName ? graphUser.displayName : "";
            this.id = graphUser.id ? graphUser.id : "";
            this.mail = graphUser.mail ? graphUser.mail : "";
            this.userPrincipalName = graphUser.userPrincipalName ? graphUser.userPrincipalName : "";
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