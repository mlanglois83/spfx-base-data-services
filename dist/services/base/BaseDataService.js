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
import { ServicesConfiguration } from "../..";
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
        _this.itemModelType = type;
        _this.cacheDuration = cacheDuration;
        _this.dbService = new BaseDbService(type, tableName);
        _this.transactionService = new TransactionService();
        return _this;
    }
    Object.defineProperty(BaseDataService.prototype, "ItemFields", {
        get: function () {
            return {};
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseDataService.prototype, "serviceName", {
        get: function () {
            return this.constructor["name"];
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseDataService.prototype, "itemType", {
        get: function () {
            return this.itemModelType;
        },
        enumerable: true,
        configurable: true
    });
    BaseDataService.prototype.Init = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/];
            });
        });
    };
    BaseDataService.prototype.getCacheKey = function (key) {
        if (key === void 0) { key = "all"; }
        return Text.format(Constants.cacheKeys.latestDataLoadFormat, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName, key);
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
                        var result, reloadData, convresult, tmp, error_1;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 11, , 12]);
                                    result = new Array();
                                    return [4 /*yield*/, this.needRefreshCache()];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!(reloadData && ServicesConfiguration.configuration.checkOnline)) return [3 /*break*/, 3];
                                    return [4 /*yield*/, UtilsService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 7];
                                    return [4 /*yield*/, this.getAll_Internal()];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, Promise.all(result.map(function (res) {
                                            return _this.convertItemToDbFormat(res);
                                        }))];
                                case 5:
                                    convresult = _a.sent();
                                    return [4 /*yield*/, this.dbService.replaceAll(convresult)];
                                case 6:
                                    _a.sent();
                                    this.UpdateCacheData();
                                    return [3 /*break*/, 10];
                                case 7: return [4 /*yield*/, this.dbService.getAll()];
                                case 8:
                                    tmp = _a.sent();
                                    return [4 /*yield*/, Promise.all(tmp.map(function (res) {
                                            return _this.mapItem(res);
                                        }))];
                                case 9:
                                    result = _a.sent();
                                    _a.label = 10;
                                case 10:
                                    this.removePromise();
                                    resolve(result);
                                    return [3 /*break*/, 12];
                                case 11:
                                    error_1 = _a.sent();
                                    this.removePromise();
                                    reject(error_1);
                                    return [3 /*break*/, 12];
                                case 12: return [2 /*return*/];
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
                        var result, reloadData, convresult, tmp, error_2;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 11, , 12]);
                                    result = new Array();
                                    return [4 /*yield*/, this.needRefreshCache(keyCached)];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!(reloadData && ServicesConfiguration.configuration.checkOnline)) return [3 /*break*/, 3];
                                    return [4 /*yield*/, UtilsService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 7];
                                    return [4 /*yield*/, this.get_Internal(query)];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, Promise.all(result.map(function (res) {
                                            return _this.convertItemToDbFormat(res);
                                        }))];
                                case 5:
                                    convresult = _a.sent();
                                    return [4 /*yield*/, this.dbService.addOrUpdateItems(convresult, query)];
                                case 6:
                                    _a.sent();
                                    this.UpdateCacheData(keyCached);
                                    return [3 /*break*/, 10];
                                case 7: return [4 /*yield*/, this.dbService.get(query)];
                                case 8:
                                    tmp = _a.sent();
                                    return [4 /*yield*/, Promise.all(tmp.map(function (res) {
                                            return _this.mapItem(res);
                                        }))];
                                case 9:
                                    result = _a.sent();
                                    _a.label = 10;
                                case 10:
                                    this.removePromise(keyCached);
                                    resolve(result);
                                    return [3 /*break*/, 12];
                                case 11:
                                    error_2 = _a.sent();
                                    this.removePromise(keyCached);
                                    reject(error_2);
                                    return [3 /*break*/, 12];
                                case 12: return [2 /*return*/];
                            }
                        });
                    }); });
                    this.storePromise(promise, keyCached);
                }
                return [2 /*return*/, promise];
            });
        });
    };
    BaseDataService.prototype.getItemById = function (id) {
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
                        var result, reloadData, converted, temp, error_3;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 11, , 12]);
                                    result = void 0;
                                    return [4 /*yield*/, this.needRefreshCache(keyCached)];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!(reloadData && ServicesConfiguration.configuration.checkOnline)) return [3 /*break*/, 3];
                                    return [4 /*yield*/, UtilsService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 7];
                                    return [4 /*yield*/, this.getItemById_Internal(id)];
                                case 4:
                                    result = _a.sent();
                                    return [4 /*yield*/, this.convertItemToDbFormat(result)];
                                case 5:
                                    converted = _a.sent();
                                    return [4 /*yield*/, this.dbService.addOrUpdateItem(converted)];
                                case 6:
                                    _a.sent();
                                    this.UpdateCacheData(_super.prototype.hashCode.call(this, keyCached).toString());
                                    return [3 /*break*/, 10];
                                case 7: return [4 /*yield*/, this.dbService.getItemById(id)];
                                case 8:
                                    temp = _a.sent();
                                    if (!temp) return [3 /*break*/, 10];
                                    return [4 /*yield*/, this.mapItem(temp)];
                                case 9:
                                    result = _a.sent();
                                    _a.label = 10;
                                case 10:
                                    this.removePromise(keyCached);
                                    resolve(result);
                                    return [3 /*break*/, 12];
                                case 11:
                                    error_3 = _a.sent();
                                    this.removePromise(keyCached);
                                    reject(error_3);
                                    return [3 /*break*/, 12];
                                case 12: return [2 /*return*/];
                            }
                        });
                    }); });
                    this.storePromise(promise, keyCached);
                }
                return [2 /*return*/, promise];
            });
        });
    };
    BaseDataService.prototype.getItemsById = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var keyCached, promise;
            var _this = this;
            return __generator(this, function (_a) {
                keyCached = "getByIds_" + ids.join();
                promise = this.getExistingPromise(keyCached);
                if (promise) {
                    console.log(this.serviceName + " " + keyCached + " : load allready called before, sharing promise");
                }
                else {
                    promise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var results, reloadData, convresults, tmp, error_4;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 11, , 12]);
                                    results = void 0;
                                    return [4 /*yield*/, this.needRefreshCache(keyCached)];
                                case 1:
                                    reloadData = _a.sent();
                                    if (!(reloadData && ServicesConfiguration.configuration.checkOnline)) return [3 /*break*/, 3];
                                    return [4 /*yield*/, UtilsService.CheckOnline()];
                                case 2:
                                    reloadData = _a.sent();
                                    _a.label = 3;
                                case 3:
                                    if (!reloadData) return [3 /*break*/, 7];
                                    return [4 /*yield*/, this.getItemsById_Internal(ids)];
                                case 4:
                                    results = _a.sent();
                                    return [4 /*yield*/, Promise.all(results.map(function (res) { return __awaiter(_this, void 0, void 0, function () {
                                            return __generator(this, function (_a) {
                                                return [2 /*return*/, this.convertItemToDbFormat(res)];
                                            });
                                        }); }))];
                                case 5:
                                    convresults = _a.sent();
                                    return [4 /*yield*/, this.dbService.addOrUpdateItems(convresults)];
                                case 6:
                                    _a.sent();
                                    this.UpdateCacheData(_super.prototype.hashCode.call(this, keyCached).toString());
                                    return [3 /*break*/, 10];
                                case 7: return [4 /*yield*/, this.dbService.getItemsById(ids)];
                                case 8:
                                    tmp = _a.sent();
                                    return [4 /*yield*/, Promise.all(tmp.map(function (res) {
                                            return _this.mapItem(res);
                                        }))];
                                case 9:
                                    results = _a.sent();
                                    _a.label = 10;
                                case 10:
                                    this.removePromise(keyCached);
                                    resolve(results);
                                    return [3 /*break*/, 12];
                                case 11:
                                    error_4 = _a.sent();
                                    this.removePromise(keyCached);
                                    reject(error_4);
                                    return [3 /*break*/, 12];
                                case 12: return [2 /*return*/];
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
            var result, itemResult, isconnected, converted, error_5, converted, dbItem, ot;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        itemResult = null;
                        isconnected = true;
                        if (!ServicesConfiguration.configuration.checkOnline) return [3 /*break*/, 2];
                        return [4 /*yield*/, UtilsService.CheckOnline()];
                    case 1:
                        isconnected = _a.sent();
                        _a.label = 2;
                    case 2:
                        if (!isconnected) return [3 /*break*/, 14];
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 7, , 13]);
                        return [4 /*yield*/, this.addOrUpdateItem_Internal(item)];
                    case 4:
                        itemResult = _a.sent();
                        return [4 /*yield*/, this.convertItemToDbFormat(itemResult)];
                    case 5:
                        converted = _a.sent();
                        return [4 /*yield*/, this.dbService.addOrUpdateItem(converted)];
                    case 6:
                        _a.sent();
                        result = {
                            item: itemResult
                        };
                        return [3 /*break*/, 13];
                    case 7:
                        error_5 = _a.sent();
                        console.error(error_5);
                        if (!(error_5.name === Constants.Errors.ItemVersionConfict)) return [3 /*break*/, 11];
                        return [4 /*yield*/, this.getItemById_Internal(item.id)];
                    case 8:
                        itemResult = _a.sent();
                        return [4 /*yield*/, this.convertItemToDbFormat(itemResult)];
                    case 9:
                        converted = _a.sent();
                        return [4 /*yield*/, this.dbService.addOrUpdateItem(converted)];
                    case 10:
                        _a.sent();
                        result = {
                            item: itemResult,
                            error: error_5
                        };
                        return [3 /*break*/, 12];
                    case 11:
                        result = {
                            item: item,
                            error: error_5
                        };
                        _a.label = 12;
                    case 12: return [3 /*break*/, 13];
                    case 13: return [3 /*break*/, 18];
                    case 14: return [4 /*yield*/, this.convertItemToDbFormat(item)];
                    case 15:
                        dbItem = _a.sent();
                        return [4 /*yield*/, this.dbService.addOrUpdateItem(dbItem)];
                    case 16:
                        result = _a.sent();
                        result.item = item;
                        ot = new OfflineTransaction();
                        ot.itemData = assign({}, dbItem);
                        ot.itemType = result.item.constructor["name"];
                        ot.title = TransactionType.AddOrUpdate;
                        return [4 /*yield*/, this.transactionService.addOrUpdateItem(ot)];
                    case 17:
                        _a.sent();
                        _a.label = 18;
                    case 18: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseDataService.prototype.deleteItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var isconnected, ot, converted;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        isconnected = true;
                        if (!ServicesConfiguration.configuration.checkOnline) return [3 /*break*/, 2];
                        return [4 /*yield*/, UtilsService.CheckOnline()];
                    case 1:
                        isconnected = _a.sent();
                        _a.label = 2;
                    case 2:
                        if (!isconnected) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.deleteItem_Internal(item)];
                    case 3:
                        _a.sent();
                        return [4 /*yield*/, this.dbService.deleteItem(item)];
                    case 4:
                        _a.sent();
                        return [3 /*break*/, 9];
                    case 5: return [4 /*yield*/, this.dbService.deleteItem(item)];
                    case 6:
                        _a.sent();
                        ot = new OfflineTransaction();
                        return [4 /*yield*/, this.convertItemToDbFormat(item)];
                    case 7:
                        converted = _a.sent();
                        ot.itemData = assign({}, converted);
                        ot.itemType = item.constructor["name"];
                        ot.title = TransactionType.Delete;
                        return [4 /*yield*/, this.transactionService.addOrUpdateItem(ot)];
                    case 8:
                        _a.sent();
                        _a.label = 9;
                    case 9: return [2 /*return*/, null];
                }
            });
        });
    };
    BaseDataService.prototype.convertItemToDbFormat = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                delete item.__internalLinks;
                return [2 /*return*/, item];
            });
        });
    };
    BaseDataService.prototype.mapItem = function (item) {
        return Promise.resolve(item);
    };
    BaseDataService.prototype.updateLinkedTransactions = function (oldId, newId, nextTransactions) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, nextTransactions];
            });
        });
    };
    BaseDataService.prototype.__getFromCache = function (id) {
        return this.dbService.getItemById(id);
    };
    BaseDataService.prototype.__getAllFromCache = function () {
        return this.dbService.getAll();
    };
    BaseDataService.prototype.__updateCache = function () {
        var items = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            items[_i] = arguments[_i];
        }
        return this.dbService.addOrUpdateItems(items);
    };
    /**
     * Stored promises to avoid multiple calls
     */
    BaseDataService.promises = {};
    return BaseDataService;
}(BaseService));
export { BaseDataService };
//# sourceMappingURL=BaseDataService.js.map