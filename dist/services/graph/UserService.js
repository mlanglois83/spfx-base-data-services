var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { BaseDataService } from "..";
import { User, PictureSize } from "../..";
import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { find } from "@microsoft/sp-lodash-subset";
var standardUserCacheDuration = 10;
var UserService = /** @class */ (function (_super) {
    __extends(UserService, _super);
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    function UserService(cacheDuration) {
        if (cacheDuration === void 0) { cacheDuration = standardUserCacheDuration; }
        var _this = _super.call(this, User, "Users", cacheDuration) || this;
        _this._spUsers = null;
        return _this;
    }
    UserService.prototype.spUsers = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(this._spUsers === null)) return [3 /*break*/, 2];
                        _a = this;
                        return [4 /*yield*/, sp.web.siteUsers.select("UserPrincipalName", "Id").get()];
                    case 1:
                        _a._spUsers = _b.sent();
                        _b.label = 2;
                    case 2: return [2 /*return*/, this._spUsers];
                }
            });
        });
    };
    UserService.prototype.get_Internal = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var reverseFilter, parts, _a, users, spUsers;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        query = query.trim();
                        reverseFilter = query;
                        parts = query.split(" ");
                        if (parts.length > 1) {
                            reverseFilter = parts[1].trim() + " " + parts[0].trim();
                        }
                        return [4 /*yield*/, Promise.all([graph.users
                                    .filter("startswith(displayName,'" + query + "') or \n            startswith(displayName,'" + reverseFilter + "') or \n            startswith(givenName,'" + query + "') or \n            startswith(surname,'" + query + "') or \n            startswith(mail,'" + query + "') or \n            startswith(userPrincipalName,'" + query + "')")
                                    .get(), this.spUsers])];
                    case 1:
                        _a = _b.sent(), users = _a[0], spUsers = _a[1];
                        return [2 /*return*/, users.map(function (u) {
                                var spuser = find(spUsers, function (spu) { return spu.UserPrincipalName === u.userPrincipalName; });
                                var result = new User(u);
                                if (spuser) {
                                    result.spId = spuser.Id;
                                }
                                return result;
                            })];
                }
            });
        });
    };
    UserService.prototype.addOrUpdateItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error("Not implemented");
            });
        });
    };
    UserService.prototype.addOrUpdateItems_Internal = function (items) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error("Not implemented");
            });
        });
    };
    UserService.prototype.deleteItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error("Not implemented");
            });
        });
    };
    /**
     * Retrieve all users (sp)
     */
    UserService.prototype.getAll_Internal = function () {
        return __awaiter(this, void 0, void 0, function () {
            var results, spUsers, batch;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = [];
                        return [4 /*yield*/, this.spUsers()];
                    case 1:
                        spUsers = _a.sent();
                        batch = graph.createBatch();
                        spUsers.forEach(function (spu) {
                            if (spu.UserPrincipalName) {
                                graph.users.select("id", "userPrincipalName", "mail", "displayName").filter("userPrincipalName eq '" + encodeURIComponent(spu.UserPrincipalName) + "'").inBatch(batch).get().then(function (graphUser) {
                                    if (graphUser && graphUser.length > 0) {
                                        var result = new User(graphUser[0]);
                                        result.spId = spu.Id;
                                        results.push(result);
                                    }
                                });
                            }
                        });
                        return [4 /*yield*/, batch.execute()];
                    case 2:
                        _a.sent();
                        return [2 /*return*/, results];
                }
            });
        });
    };
    UserService.prototype.getItemById_Internal = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, _a, graphUser, spUsers, spuser, result_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        result = null;
                        return [4 /*yield*/, Promise.all([graph.users.getById(id).select("id", "userPrincipalName", "mail", "displayName").get(), this.spUsers])];
                    case 1:
                        _a = _b.sent(), graphUser = _a[0], spUsers = _a[1];
                        if (graphUser) {
                            spuser = find(spUsers, function (spu) {
                                return spu.UserPrincipalName === graphUser.userPrincipalName;
                            });
                            result_1 = new User(graphUser);
                            if (spuser) {
                                result_1.spId = spuser.Id;
                            }
                        }
                        return [2 /*return*/, result];
                }
            });
        });
    };
    UserService.prototype.getItemsById_Internal = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var results, spUsers, batch;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = [];
                        return [4 /*yield*/, this.spUsers()];
                    case 1:
                        spUsers = _a.sent();
                        batch = graph.createBatch();
                        ids.forEach(function (id) {
                            graph.users.getById(id).select("id", "userPrincipalName", "mail", "displayName").inBatch(batch).get().then(function (graphUser) {
                                var spuser = find(spUsers, function (spu) {
                                    return spu.UserPrincipalName === graphUser.userPrincipalName;
                                });
                                var result = new User(graphUser);
                                if (spuser) {
                                    result.spId = spuser.Id;
                                }
                                results.push(result);
                            });
                        });
                        return [4 /*yield*/, batch.execute()];
                    case 2:
                        _a.sent();
                        return [2 /*return*/, results];
                }
            });
        });
    };
    UserService.prototype.linkToSpUser = function (user) {
        return __awaiter(this, void 0, void 0, function () {
            var result;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(user.spId === undefined)) return [3 /*break*/, 2];
                        return [4 /*yield*/, sp.web.ensureUser(user.userPrincipalName)];
                    case 1:
                        result = _a.sent();
                        user.spId = result.data.Id;
                        this.dbService.addOrUpdateItem(user);
                        _a.label = 2;
                    case 2: return [2 /*return*/, user];
                }
            });
        });
    };
    UserService.prototype.getByDisplayName = function (displayName) {
        return __awaiter(this, void 0, void 0, function () {
            var users, reverseFilter_1, parts;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.get(displayName)];
                    case 1:
                        users = _a.sent();
                        if (!(users.length === 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.getAll()];
                    case 2:
                        users = _a.sent();
                        displayName = displayName.trim();
                        reverseFilter_1 = displayName;
                        parts = displayName.split(" ");
                        if (parts.length > 1) {
                            reverseFilter_1 = parts[1].trim() + " " + parts[0].trim();
                        }
                        users = users.filter(function (user) {
                            return user.displayName.indexOf(displayName) === 0 ||
                                user.displayName.indexOf(reverseFilter_1) === 0 ||
                                user.mail.indexOf(displayName) === 0 ||
                                user.userPrincipalName.indexOf(displayName) === 0;
                        });
                        _a.label = 3;
                    case 3: return [2 /*return*/, users];
                }
            });
        });
    };
    UserService.prototype.getBySpId = function (spId) {
        return __awaiter(this, void 0, void 0, function () {
            var allUsers;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getAll()];
                    case 1:
                        allUsers = _a.sent();
                        return [2 /*return*/, find(allUsers, function (user) { return user.spId === spId; })];
                }
            });
        });
    };
    UserService.getPictureUrl = function (user, size) {
        if (size === void 0) { size = PictureSize.Large; }
        return user.mail ? Text.format("{0}/_layouts/15/userphoto.aspx?accountname={1}&size={2}", ServicesConfiguration.context.pageContext.web.absoluteUrl, user.mail, size) : "";
    };
    return UserService;
}(BaseDataService));
export { UserService };
