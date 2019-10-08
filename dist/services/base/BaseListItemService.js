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
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { Constants } from "../../constants/index";
import { BaseDataService } from "./BaseDataService";
import { BaseService } from "./BaseService";
/**
 *
 * Base service for sp list items operations
 */
var BaseListItemService = /** @class */ (function (_super) {
    __extends(BaseListItemService, _super);
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param listRelativeUrl list web relative url
     */
    function BaseListItemService(type, listRelativeUrl, tableName, cacheDuration) {
        var _this = _super.call(this, type, tableName, cacheDuration) || this;
        _this.listRelativeUrl = BaseService.Configuration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
        _this.itemType = type;
        return _this;
    }
    Object.defineProperty(BaseListItemService.prototype, "listItemType", {
        get: function () {
            return this.itemType;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseListItemService.prototype, "list", {
        /**
         * Associeted list (pnpjs)
         */
        get: function () {
            return sp.web.getList(this.listRelativeUrl);
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Cache has to be relaod ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    BaseListItemService.prototype.needRefreshCache = function (key) {
        if (key === void 0) { key = "all"; }
        return __awaiter(this, void 0, void 0, function () {
            var result, isconnected, cachedDataDate, response, tempList, lastModifiedDate, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.needRefreshCache.call(this, key)];
                    case 1:
                        result = _a.sent();
                        if (!!result) return [3 /*break*/, 8];
                        return [4 /*yield*/, this.utilService.CheckOnline()];
                    case 2:
                        isconnected = _a.sent();
                        if (!isconnected) return [3 /*break*/, 8];
                        return [4 /*yield*/, _super.prototype.getCachedData.call(this, key)];
                    case 3:
                        cachedDataDate = _a.sent();
                        if (!cachedDataDate) return [3 /*break*/, 8];
                        _a.label = 4;
                    case 4:
                        _a.trys.push([4, 7, , 8]);
                        return [4 /*yield*/, BaseService.Configuration.context.spHttpClient.get(BaseService.Configuration.context.pageContext.web.absoluteUrl + "/_api/web/getList('" + this.listRelativeUrl + "')", SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata.metadata=minimal',
                                    'Cache-Control': 'no-cache'
                                }
                            })];
                    case 5:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 6:
                        tempList = _a.sent();
                        lastModifiedDate = new Date(tempList.LastItemUserModifiedDate ? tempList.LastItemUserModifiedDate : tempList.d.LastItemUserModifiedDate);
                        result = lastModifiedDate > cachedDataDate;
                        return [3 /*break*/, 8];
                    case 7:
                        error_1 = _a.sent();
                        console.error(error_1);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     *
     * TODO avoid getting all fields
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    BaseListItemService.prototype.get_Internal = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var results, items;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = new Array();
                        return [4 /*yield*/, this.list.getItemsByCAMLQuery({
                                ViewXml: '<View Scope="RecursiveAll"><Query>' + query + '</Query></View>'
                            }, 'FieldValuesAsText')];
                    case 1:
                        items = _a.sent();
                        return [2 /*return*/, items.map(function (r) { return new _this.itemType(r); })];
                }
            });
        });
    };
    /**
     *
     * @param id
     */
    BaseListItemService.prototype.getById_Internal = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, temp;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        return [4 /*yield*/, this.list.items.getById(id).get()];
                    case 1:
                        temp = _a.sent();
                        if (temp) {
                            result = new this.itemType(temp);
                        }
                        return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Retrieve all items
     *
     * TODO avoid getting all fields
     */
    BaseListItemService.prototype.getAll_Internal = function () {
        return __awaiter(this, void 0, void 0, function () {
            var items;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.list.items.getAll()];
                    case 1:
                        items = _a.sent();
                        return [2 /*return*/, items.map(function (r) { return new _this.itemType(r); })];
                }
            });
        });
    };
    BaseListItemService.prototype.addOrUpdateItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var result, addResult, existing, error, version, updateResult;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = cloneDeep(item);
                        if (!(item.id < 0)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.list.items.add(item.convert())];
                    case 1:
                        addResult = _a.sent();
                        if (result.onAddCompleted) {
                            result.onAddCompleted(addResult.data);
                        }
                        return [3 /*break*/, 10];
                    case 2:
                        if (!item.version) return [3 /*break*/, 8];
                        return [4 /*yield*/, this.list.items.getById(item.id).select("OData__UIVersionString").get()];
                    case 3:
                        existing = _a.sent();
                        if (!(parseFloat(existing["OData__UIVersionString"]) > item.version)) return [3 /*break*/, 4];
                        error = new Error(BaseService.Configuration.translations.versionHigherErrorMessage);
                        error.name = Constants.Errors.ItemVersionConfict;
                        throw error;
                    case 4: return [4 /*yield*/, this.list.items.getById(item.id).update(item.convert())];
                    case 5:
                        _a.sent();
                        return [4 /*yield*/, this.list.items.getById(item.id).get()];
                    case 6:
                        version = _a.sent();
                        if (result.onUpdateCompleted) {
                            result.onUpdateCompleted(version);
                        }
                        _a.label = 7;
                    case 7: return [3 /*break*/, 10];
                    case 8: return [4 /*yield*/, this.list.items.getById(item.id).update(item.convert())];
                    case 9:
                        updateResult = _a.sent();
                        if (result.onUpdateCompleted) {
                            result.onUpdateCompleted(updateResult.data);
                        }
                        _a.label = 10;
                    case 10: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseListItemService.prototype.deleteItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.list.items.getById(item.id).delete()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    return BaseListItemService;
}(BaseDataService));
export { BaseListItemService };
//# sourceMappingURL=BaseListItemService.js.map