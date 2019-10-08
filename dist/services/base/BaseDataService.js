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
import { assign } from "@microsoft/sp-lodash-subset";
import { OfflineTransaction } from "../../models/index";
import { UtilsService } from "../index";
import { TransactionService } from "../synchronization/TransactionService";
import { BaseDbService } from "./BaseDbService";
import { BaseService } from "./BaseService";
import { Text } from "@microsoft/sp-core-library";
import { TransactionType, Constants } from "../../constants";
/**
 * Base class for data service allowing automatic management of online/offline mode with links to db and sp
 */
var BaseDataService = /** @class */ (function (_super) {
    __extends(BaseDataService, _super);
    /**
     *
     * @param type type of items
     * @param context context of the current wp
     * @param tableName name of indexedDb table
     */
    function BaseDataService(type, tableName, cacheDuration) {
        if (cacheDuration === void 0) { cacheDuration = -1; }
        var _this = _super.call(this) || this;
        _this.cacheDuration = -1;
        _this.itemType = type;
        _this.cacheDuration = cacheDuration;
        _this.dbService = new BaseDbService(type, tableName);
        _this.transactionService = new TransactionService();
        _this.utilService = new UtilsService();
        return _this;
    }
    Object.defineProperty(BaseDataService.prototype, "serviceName", {
        get: function () {
            return this.constructor["name"];
        },
        enumerable: true,
        configurable: true
    });
    BaseDataService.prototype.getCacheKey = function (key) {
        if (key === void 0) { key = "all"; }
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, BaseService.Configuration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
    };
    BaseDataService.prototype.getExistingPromise = function (key) {
        if (key === void 0) { key = "all"; }
        var pkey = this.serviceName + "-" + key;
        if (BaseDataService.promises[pkey]) {
            return BaseDataService.promises[pkey];
        }
        else
            return null;
    };
    BaseDataService.prototype.storePromise = function (promise, key) {
        if (key === void 0) { key = "all"; }
        var pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = promise;
    };
    BaseDataService.prototype.removePromise = function (key) {
        if (key === void 0) { key = "all"; }
        var pkey = this.serviceName + "-" + key;
        BaseDataService.promises[pkey] = undefined;
    };
    /***
     *
     */
    BaseDataService.prototype.getCachedData = function (key) {
        if (key === void 0) { key = "all"; }
        return __awaiter(this, void 0, void 0, function () {
            var cacheKey, lastDataLoadString, lastDataLoad;
            return __generator(this, function (_a) {
                cacheKey = this.getCacheKey(key);
                lastDataLoadString = window.sessionStorage.getItem(cacheKey);
                lastDataLoad = null;
                if (lastDataLoadString) {
                    lastDataLoad = new Date(JSON.parse(window.sessionStorage.getItem(cacheKey)));
                }
                return [2 /*return*/, lastDataLoad];
            });
        });
    };
    /**
     * Cache has to be relaod ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseDataService
     */
    BaseDataService.prototype.needRefreshCache = function (key) {
        if (key === void 0) { key = "all"; }
        return __awaiter(this, void 0, void 0, function () {
            var result, cachedDataDate, now;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = this.cacheDuration == -1;
                        if (!!result) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.getCachedData(key)];
                    case 1:
                        cachedDataDate = _a.sent();
                        if (cachedDataDate) {
                            //add cache duration
                            cachedDataDate.setMinutes(cachedDataDate.getMinutes() + this.cacheDuration);
                            now = new Date();
                            //cache has expired
                            result = cachedDataDate < now;
                        }
                        else {
                            result = true;
                        }
                        _a.label = 2;
                    case 2: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseDataService.prototype.UpdateCacheData = function (key) {
        if (key === void 0) { key = "all"; }
        var result = this.cacheDuration == -1;
        //if cache defined
        if (!result) {
            var cacheKey = this.getCacheKey(key);
            window.sessionStorage.setItem(cacheKey, JSON.stringify(new Date()));
        }
    };
    /*
     * Retrieve all elements from datasource depending on connection is enabled
     * If service is not configured as offline, an exception is thrown;
     */
    BaseDataService.prototype.getAll = function () {
        return __awaiter(this, void 0, void 0, function () {
            var promise;
            var _this = this;
            return __generator(this, function (_a) {
                promise = this.getExistingPromise();
                if (promise) {
                    console.log(this.serviceName + " getAll : load allready called before, sharing promise");
                }
                else {
                    promise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var result, reloadData, error_1;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 9, , 10]);
                                    result = new Array();
                                    return [4 /*yield*/, this.needRefreshCache()];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!reloadData) return [3 /*break*/, 3];
                                    return [4 /*yield*/, this.utilService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 6];
                                    return [4 /*yield*/, this.getAll_Internal()];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, this.dbService.replaceAll(result)];
                                case 5:
                                    _a.sent();
                                    this.UpdateCacheData();
                                    return [3 /*break*/, 8];
                                case 6: return [4 /*yield*/, this.dbService.getAll()];
                                case 7:
                                    result = _a.sent();
                                    _a.label = 8;
                                case 8:
                                    this.removePromise();
                                    resolve(result);
                                    return [3 /*break*/, 10];
                                case 9:
                                    error_1 = _a.sent();
                                    this.removePromise();
                                    reject(error_1);
                                    return [3 /*break*/, 10];
                                case 10: return [2 /*return*/];
                            }
                        });
                    }); });
                    this.storePromise(promise);
                }
                return [2 /*return*/, promise];
            });
        });
    };
    BaseDataService.prototype.get = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var keyCached, promise;
            var _this = this;
            return __generator(this, function (_a) {
                keyCached = _super.prototype.hashCode.call(this, query).toString();
                promise = this.getExistingPromise(keyCached);
                if (promise) {
                    console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
                }
                else {
                    promise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var result, reloadData, error_2;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 9, , 10]);
                                    result = new Array();
                                    return [4 /*yield*/, this.needRefreshCache(keyCached)];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!reloadData) return [3 /*break*/, 3];
                                    return [4 /*yield*/, this.utilService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 6];
                                    return [4 /*yield*/, this.get_Internal(query)];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, this.dbService.addOrUpdateItems(result, query)];
                                case 5:
                                    _a.sent();
                                    this.UpdateCacheData(keyCached);
                                    return [3 /*break*/, 8];
                                case 6: return [4 /*yield*/, this.dbService.get(query)];
                                case 7:
                                    result = _a.sent();
                                    _a.label = 8;
                                case 8:
                                    this.removePromise(keyCached);
                                    resolve(result);
                                    return [3 /*break*/, 10];
                                case 9:
                                    error_2 = _a.sent();
                                    this.removePromise(keyCached);
                                    reject(error_2);
                                    return [3 /*break*/, 10];
                                case 10: return [2 /*return*/];
                            }
                        });
                    }); });
                    this.storePromise(promise, keyCached);
                }
                return [2 /*return*/, promise];
            });
        });
    };
    BaseDataService.prototype.getById = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var keyCached, promise;
            var _this = this;
            return __generator(this, function (_a) {
                keyCached = "getById_" + id.toString();
                promise = this.getExistingPromise(keyCached);
                if (promise) {
                    console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
                }
                else {
                    promise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var result, reloadData, temp, error_3;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 9, , 10]);
                                    result = void 0;
                                    return [4 /*yield*/, this.needRefreshCache(keyCached)];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!reloadData) return [3 /*break*/, 3];
                                    return [4 /*yield*/, this.utilService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 6];
                                    return [4 /*yield*/, this.getById_Internal(id)];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, this.dbService.addOrUpdateItems([result], keyCached)];
                                case 5:
                                    _a.sent();
                                    this.UpdateCacheData(_super.prototype.hashCode.call(this, keyCached).toString());
                                    return [3 /*break*/, 8];
                                case 6: return [4 /*yield*/, this.dbService.get(keyCached)];
                                case 7:
                                    temp = _a.sent();
                                    if (temp && temp.length > 0) {
                                        result = temp[0];
                                    }
                                    _a.label = 8;
                                case 8:
                                    this.removePromise(keyCached);
                                    resolve(result);
                                    return [3 /*break*/, 10];
                                case 9:
                                    error_3 = _a.sent();
                                    this.removePromise(keyCached);
                                    reject(error_3);
                                    return [3 /*break*/, 10];
                                case 10: return [2 /*return*/];
                            }
                        });
                    }); });
                    this.storePromise(promise, keyCached);
                }
                return [2 /*return*/, promise];
            });
        });
    };
    BaseDataService.prototype.addOrUpdateItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var result, itemResult, isconnected, error_4, ot;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        itemResult = null;
                        return [4 /*yield*/, this.utilService.CheckOnline()];
                    case 1:
                        isconnected = _a.sent();
                        if (!isconnected) return [3 /*break*/, 11];
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 5, , 10]);
                        return [4 /*yield*/, this.addOrUpdateItem_Internal(item)];
                    case 3:
                        itemResult = _a.sent();
                        return [4 /*yield*/, this.dbService.addOrUpdateItem(itemResult)];
                    case 4:
                        _a.sent();
                        result = {
                            item: itemResult
                        };
                        return [3 /*break*/, 10];
                    case 5:
                        error_4 = _a.sent();
                        if (!(error_4.name === Constants.Errors.ItemVersionConfict)) return [3 /*break*/, 8];
                        return [4 /*yield*/, this.getById_Internal(item.id)];
                    case 6:
                        itemResult = _a.sent();
                        return [4 /*yield*/, this.dbService.addOrUpdateItems([itemResult])];
                    case 7:
                        _a.sent();
                        result = {
                            item: itemResult,
                            error: error_4
                        };
                        return [3 /*break*/, 9];
                    case 8:
                        result = {
                            item: item,
                            error: error_4
                        };
                        _a.label = 9;
                    case 9: return [3 /*break*/, 10];
                    case 10: return [3 /*break*/, 14];
                    case 11: return [4 /*yield*/, this.dbService.addOrUpdateItem(item)];
                    case 12:
                        result = _a.sent();
                        ot = new OfflineTransaction();
                        ot.itemData = assign({}, result.item);
                        ot.itemType = result.item.constructor["name"];
                        ot.serviceName = this.serviceName;
                        ot.title = TransactionType.AddOrUpdate;
                        return [4 /*yield*/, this.transactionService.addOrUpdateItem(ot)];
                    case 13:
                        _a.sent();
                        _a.label = 14;
                    case 14: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseDataService.prototype.deleteItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var isconnected, ot;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.utilService.CheckOnline()];
                    case 1:
                        isconnected = _a.sent();
                        if (!isconnected) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.deleteItem_Internal(item)];
                    case 2:
                        _a.sent();
                        return [4 /*yield*/, this.dbService.deleteItem(item)];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 7];
                    case 4: return [4 /*yield*/, this.dbService.deleteItem(item)];
                    case 5:
                        _a.sent();
                        ot = new OfflineTransaction();
                        ot.itemData = assign({}, item);
                        ot.itemType = item.constructor["name"];
                        ot.serviceName = this.serviceName;
                        ot.title = TransactionType.Delete;
                        return [4 /*yield*/, this.transactionService.addOrUpdateItem(ot)];
                    case 6:
                        _a.sent();
                        _a.label = 7;
                    case 7: return [2 /*return*/, null];
                }
            });
        });
    };
    /**
     * Stored promises to avoid multiple calls
     */
    BaseDataService.promises = {};
    return BaseDataService;
}(BaseService));
export { BaseDataService };
//# sourceMappingURL=BaseDataService.js.map