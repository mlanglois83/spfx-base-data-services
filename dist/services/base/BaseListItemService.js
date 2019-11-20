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
import { ServicesConfiguration } from "../..";
import { SPHttpClient } from '@microsoft/sp-http';
import { cloneDeep, find, assign } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { Constants, FieldType } from "../../constants/index";
import { BaseDataService } from "./BaseDataService";
import { UtilsService } from "..";
import { SPItem } from "../../models";
import { UserService } from "../graph/UserService";
import { isArray } from "@pnp/common";
/**
 *
 * Base service for sp list items operations
 */
var BaseListItemService = /** @class */ (function (_super) {
    __extends(BaseListItemService, _super);
    /***************************** Constructor **************************************/
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param listRelativeUrl list web relative url
     */
    function BaseListItemService(type, listRelativeUrl, tableName, cacheDuration) {
        var _this = _super.call(this, type, tableName, cacheDuration) || this;
        _this.initValues = {};
        /***************************** External sources init and access **************************************/
        _this.initialized = false;
        _this.initPromise = null;
        _this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
        _this.itemType = type;
        return _this;
    }
    Object.defineProperty(BaseListItemService.prototype, "ItemFields", {
        get: function () {
            var result = {};
            assign(result, this.itemType["Fields"][SPItem["name"]]);
            if (this.itemType["Fields"][this.itemType["name"]]) {
                assign(result, this.itemType["Fields"][this.itemType["name"]]);
            }
            return result;
        },
        enumerable: true,
        configurable: true
    });
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
    Object.defineProperty(BaseListItemService.prototype, "isInitialized", {
        get: function () {
            return this.initialized;
        },
        enumerable: true,
        configurable: true
    });
    BaseListItemService.prototype.Init = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                if (!this.initPromise) {
                    this.initPromise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var fields, services, key, fieldDescription, error_1;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    if (!this.initialized) return [3 /*break*/, 1];
                                    resolve();
                                    return [3 /*break*/, 6];
                                case 1:
                                    _a.trys.push([1, 5, , 6]);
                                    if (!this.init_internal) return [3 /*break*/, 3];
                                    return [4 /*yield*/, this.init_internal()];
                                case 2:
                                    _a.sent();
                                    _a.label = 3;
                                case 3:
                                    fields = this.ItemFields;
                                    services = [];
                                    for (key in fields) {
                                        if (fields.hasOwnProperty(key)) {
                                            fieldDescription = fields[key];
                                            if (fieldDescription.serviceName && services.indexOf(fieldDescription.serviceName) === -1) {
                                                services.push(fieldDescription.serviceName);
                                            }
                                            else if ((fieldDescription.fieldType === FieldType.O365User || fieldDescription.fieldType === FieldType.O365UserMulti) &&
                                                services.indexOf(UserService["name"]) === -1) {
                                                services.push(UserService["name"]);
                                            }
                                        }
                                    }
                                    return [4 /*yield*/, Promise.all(services.map(function (serviceName) { return __awaiter(_this, void 0, void 0, function () {
                                            var service, values;
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        if (!!this.initValues[serviceName]) return [3 /*break*/, 2];
                                                        service = ServicesConfiguration.configuration.serviceFactory.create(serviceName);
                                                        return [4 /*yield*/, service.getAll()];
                                                    case 1:
                                                        values = _a.sent();
                                                        this.initValues[serviceName] = values;
                                                        _a.label = 2;
                                                    case 2: return [2 /*return*/];
                                                }
                                            });
                                        }); }))];
                                case 4:
                                    _a.sent();
                                    this.initialized = true;
                                    this.initPromise = null;
                                    resolve();
                                    return [3 /*break*/, 6];
                                case 5:
                                    error_1 = _a.sent();
                                    this.initPromise = null;
                                    reject(error_1);
                                    return [3 /*break*/, 6];
                                case 6: return [2 /*return*/];
                            }
                        });
                    }); });
                }
                return [2 /*return*/, this.initPromise];
            });
        });
    };
    BaseListItemService.prototype.getServiceInitValues = function (serviceName) {
        return this.initValues[serviceName];
    };
    /****************************** get item methods ***********************************/
    BaseListItemService.prototype.getItemFromRest = function (spitem) {
        var _this = this;
        var item = new this.listItemType();
        Object.keys(this.ItemFields).map(function (propertyName) {
            var fieldDescription = _this.ItemFields[propertyName];
            item[propertyName] = _this.getFieldValue(spitem, fieldDescription);
        });
        return item;
    };
    BaseListItemService.prototype.getFieldValue = function (spitem, fieldDescriptor) {
        var value = fieldDescriptor.defaultValue;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if (fieldDescriptor.fieldName === "OData__UIVersionString") {
                    value = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                }
                else {
                    value = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Date:
                value = spitem[fieldDescriptor.fieldName] ? new Date(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
            case FieldType.LookupMulti:
                value = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : fieldDescriptor.defaultValue;
                break;
            case FieldType.O365User:
                var id_1 = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if (id_1 !== -1) {
                    var users = this.getServiceInitValues(UserService["name"]);
                    value = find(users, function (user) { return user.spId === id_1; });
                }
                break;
            case FieldType.O365UserMulti:
                var ids = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : [];
                if (ids.length > 0) {
                    var users_1 = this.getServiceInitValues(UserService["name"]);
                    value = ids.map(function (userid) { return find(users_1, function (user) { return user.spId === userid; }); });
                }
                break;
            case FieldType.Taxonomy:
                var wssid = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if (id_1 !== -1) {
                    var terms_1 = this.getServiceInitValues(fieldDescriptor.serviceName);
                    value = this.getTaxonomyTermByWssId(wssid, terms_1);
                }
                break;
            case FieldType.TaxonomyMulti:
                var terms = spitem[fieldDescriptor.fieldName];
                if (terms) {
                    var allterms_1 = this.getServiceInitValues(fieldDescriptor.serviceName);
                    value = terms.map(function (term) {
                        return term.getTaxonomyTermByWssId(term.WssId, allterms_1);
                    });
                }
                break;
            case FieldType.Json:
                value = spitem[fieldDescriptor.fieldName] ? JSON.parse(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
        return value;
    };
    /****************************** Send item methods ***********************************/
    BaseListItemService.prototype.getSPRestItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var spitem;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        spitem = {};
                        return [4 /*yield*/, Promise.all(Object.keys(this.ItemFields).map(function (propertyName) { return __awaiter(_this, void 0, void 0, function () {
                                var fieldDescription, value;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            fieldDescription = this.ItemFields[propertyName];
                                            if (!(propertyName != "Version")) return [3 /*break*/, 2];
                                            return [4 /*yield*/, this.convertFieldValueToRest(item[propertyName], fieldDescription)];
                                        case 1:
                                            value = _a.sent();
                                            assign(spitem[fieldDescription.fieldName], value);
                                            _a.label = 2;
                                        case 2: return [2 /*return*/];
                                    }
                                });
                            }); }))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, spitem];
                }
            });
        });
    };
    BaseListItemService.prototype.convertFieldValueToRest = function (itemValue, fieldDescriptor) {
        return __awaiter(this, void 0, void 0, function () {
            var value, _a, _b, _c, _d, _e;
            var _this = this;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        value = {};
                        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
                        _a = fieldDescriptor.fieldType;
                        switch (_a) {
                            case FieldType.Simple: return [3 /*break*/, 1];
                            case FieldType.Date: return [3 /*break*/, 1];
                            case FieldType.Lookup: return [3 /*break*/, 2];
                            case FieldType.LookupMulti: return [3 /*break*/, 3];
                            case FieldType.O365User: return [3 /*break*/, 4];
                            case FieldType.O365UserMulti: return [3 /*break*/, 6];
                            case FieldType.Taxonomy: return [3 /*break*/, 10];
                            case FieldType.TaxonomyMulti: return [3 /*break*/, 11];
                            case FieldType.Json: return [3 /*break*/, 12];
                        }
                        return [3 /*break*/, 13];
                    case 1:
                        value[fieldDescriptor.fieldName] = itemValue;
                        return [3 /*break*/, 13];
                    case 2:
                        value[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                        _f.label = 3;
                    case 3:
                        value[fieldDescriptor.fieldName + "Id"] = itemValue && isArray(itemValue) && itemValue.length > 0 ? itemValue : [];
                        return [3 /*break*/, 13];
                    case 4:
                        _b = value;
                        _c = fieldDescriptor.fieldName + "Id";
                        return [4 /*yield*/, this.convertSingleUserFieldValue(itemValue)];
                    case 5:
                        _b[_c] = _f.sent();
                        return [3 /*break*/, 13];
                    case 6:
                        if (!(itemValue && isArray(itemValue) && itemValue.length > 0)) return [3 /*break*/, 8];
                        _d = value;
                        _e = fieldDescriptor.fieldName + "Id";
                        return [4 /*yield*/, Promise.all(itemValue.map(function (user) {
                                return _this.convertSingleUserFieldValue(user);
                            }))];
                    case 7:
                        _d[_e] = _f.sent();
                        return [3 /*break*/, 9];
                    case 8:
                        value[fieldDescriptor.fieldName + "Id"] = [];
                        _f.label = 9;
                    case 9: return [3 /*break*/, 13];
                    case 10:
                        value[fieldDescriptor.fieldName] = this.convertTaxonomyFieldValue(itemValue);
                        return [3 /*break*/, 13];
                    case 11:
                        if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                            value[fieldDescriptor.fieldName] = itemValue.map(function (term) {
                                return _this.convertTaxonomyFieldValue(term);
                            });
                        }
                        else {
                            value[fieldDescriptor.fieldName] = [];
                        }
                        return [3 /*break*/, 13];
                    case 12:
                        value[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                        return [3 /*break*/, 13];
                    case 13: return [2 /*return*/, value];
                }
            });
        });
    };
    /********************** SP Fields conversion helpers *****************************/
    BaseListItemService.prototype.convertTaxonomyFieldValue = function (value) {
        var result = null;
        if (value) {
            result = {
                __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
                Label: value.title,
                TermGuid: value.id,
                WssId: -1 // fake
            };
        }
        return result;
    };
    BaseListItemService.prototype.convertSingleUserFieldValue = function (value) {
        return __awaiter(this, void 0, void 0, function () {
            var result, userService;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        if (!value) return [3 /*break*/, 3];
                        if (!(!value.spId || value.spId <= 0)) return [3 /*break*/, 2];
                        userService = new UserService();
                        return [4 /*yield*/, userService.linkToSpUser(value)];
                    case 1:
                        value = _a.sent();
                        _a.label = 2;
                    case 2:
                        result = value.spId;
                        _a.label = 3;
                    case 3: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     *
     * @param wssid
     * @param terms
     */
    BaseListItemService.prototype.getTaxonomyTermByWssId = function (wssid, terms) {
        return find(terms, function (term) {
            return (term.wssids && term.wssids.indexOf(wssid) > -1);
        });
    };
    /******************************************* Cache Management *************************************************/
    /**
     * Cache has to be reloaded ?
     *
     * @readonly
     * @protected
     * @type {boolean}
     * @memberof BaseListItemService
     */
    BaseListItemService.prototype.needRefreshCache = function (key) {
        if (key === void 0) { key = "all"; }
        return __awaiter(this, void 0, void 0, function () {
            var result, isconnected, cachedDataDate, response, tempList, lastModifiedDate, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.needRefreshCache.call(this, key)];
                    case 1:
                        result = _a.sent();
                        if (!!result) return [3 /*break*/, 8];
                        return [4 /*yield*/, UtilsService.CheckOnline()];
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
                        return [4 /*yield*/, ServicesConfiguration.context.spHttpClient.get(ServicesConfiguration.context.pageContext.web.absoluteUrl + "/_api/web/getList('" + this.listRelativeUrl + "')", SPHttpClient.configurations.v1, {
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
                        error_2 = _a.sent();
                        console.error(error_2);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/, result];
                }
            });
        });
    };
    /***************** SP Calls associated to service standard operations ********************/
    /**
     * Get items by query
     * @protected
     * @param {*} query
     * @returns {Promise<Array<T>>}
     * @memberof BaseListItemService
     */
    BaseListItemService.prototype.get_Internal = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var results, selectFields, items;
            var _a;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        results = new Array();
                        selectFields = this.getOdataFieldNames();
                        return [4 /*yield*/, (_a = this.list).select.apply(_a, selectFields).getItemsByCAMLQuery({
                                ViewXml: "<View Scope=\"RecursiveAll\"><Query>" + query + "</Query></View>"
                            })];
                    case 1:
                        items = _b.sent();
                        if (!(items && items.length > 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.Init()];
                    case 2:
                        _b.sent();
                        results = items.map(function (r) {
                            return _this.getItemFromRest(r);
                        });
                        _b.label = 3;
                    case 3: return [2 /*return*/, results];
                }
            });
        });
    };
    /**
     * Get an item by id
     * @param id item id
     */
    BaseListItemService.prototype.getById_Internal = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, selectFields, temp;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        result = null;
                        selectFields = this.getOdataFieldNames();
                        return [4 /*yield*/, (_a = this.list.items.getById(id)).select.apply(_a, selectFields).get()];
                    case 1:
                        temp = _b.sent();
                        if (!temp) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.Init()];
                    case 2:
                        _b.sent();
                        result = this.getItemFromRest(temp);
                        return [2 /*return*/, result];
                    case 3: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Retrieve all items
     *
     */
    BaseListItemService.prototype.getAll_Internal = function () {
        return __awaiter(this, void 0, void 0, function () {
            var results, selectFields, items;
            var _a;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        results = [];
                        selectFields = this.getOdataFieldNames();
                        return [4 /*yield*/, (_a = this.list.items).select.apply(_a, selectFields).getAll()];
                    case 1:
                        items = _b.sent();
                        if (!(items && items.length > 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.Init()];
                    case 2:
                        _b.sent();
                        results = items.map(function (r) {
                            return _this.getItemFromRest(r);
                        });
                        _b.label = 3;
                    case 3: return [2 /*return*/, results];
                }
            });
        });
    };
    /**
     * Add or update an item
     * @param item SPItem derived object to be converted
     */
    BaseListItemService.prototype.addOrUpdateItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var result, converted, addResult, existing, error, converted, updateResult, version, converted, updateResult, version;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = cloneDeep(item);
                        if (!(item.id < 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.getSPRestItem(item)];
                    case 1:
                        converted = _a.sent();
                        return [4 /*yield*/, this.list.items.add(converted)];
                    case 2:
                        addResult = _a.sent();
                        if (addResult.data["OData__UIVersionString"]) {
                            result.version = parseFloat(addResult.data["OData__UIVersionString"]);
                        }
                        return [3 /*break*/, 14];
                    case 3:
                        if (!item.version) return [3 /*break*/, 10];
                        return [4 /*yield*/, this.list.items.getById(item.id).select("OData__UIVersionString").get()];
                    case 4:
                        existing = _a.sent();
                        if (!(parseFloat(existing["OData__UIVersionString"]) > item.version)) return [3 /*break*/, 5];
                        error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                        error.name = Constants.Errors.ItemVersionConfict;
                        throw error;
                    case 5: return [4 /*yield*/, this.getSPRestItem(item)];
                    case 6:
                        converted = _a.sent();
                        return [4 /*yield*/, this.list.items.getById(item.id).update(converted)];
                    case 7:
                        updateResult = _a.sent();
                        return [4 /*yield*/, updateResult.item.select("OData__UIVersionString").get()];
                    case 8:
                        version = _a.sent();
                        if (version["OData__UIVersionString"]) {
                            result.version = parseFloat(version["OData__UIVersionString"]);
                        }
                        _a.label = 9;
                    case 9: return [3 /*break*/, 14];
                    case 10: return [4 /*yield*/, this.getSPRestItem(item)];
                    case 11:
                        converted = _a.sent();
                        return [4 /*yield*/, this.list.items.getById(item.id).update(converted)];
                    case 12:
                        updateResult = _a.sent();
                        return [4 /*yield*/, updateResult.item.select("OData__UIVersionString").get()];
                    case 13:
                        version = _a.sent();
                        if (version["OData__UIVersionString"]) {
                            result.version = parseFloat(version["OData__UIVersionString"]);
                        }
                        _a.label = 14;
                    case 14: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Delete an item
     * @param item SPItem derived class to be deletes
     */
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
    /************************** Query filters ***************************/
    /**
     * Retrive all fields to include in odata setect parameter
     */
    BaseListItemService.prototype.getOdataFieldNames = function () {
        var fields = this.ItemFields;
        var fieldNames = Object.keys(fields).filter(function (propertyName) {
            return fields.hasOwnProperty(propertyName);
        }).map(function (prop) {
            var result = fields[prop].fieldName;
            switch (fields[prop].fieldType) {
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                case FieldType.O365User:
                case FieldType.O365UserMulti:
                    result += "Id";
                default:
                    break;
            }
            return result;
        });
        return fieldNames;
    };
    /**
     * Retrive all fields to include in odata setect parameter
     */
    BaseListItemService.prototype.getCamlViewFields = function () {
        var fields = this.ItemFields;
        var fieldNames = Object.keys(fields).filter(function (propertyName) {
            return fields.hasOwnProperty(propertyName);
        }).map(function (prop) {
            var result = fields[prop].fieldName;
            switch (fields[prop].fieldType) {
                case FieldType.Lookup:
                case FieldType.LookupMulti:
                case FieldType.O365User:
                case FieldType.O365UserMulti:
                    result += "Id";
                default:
                    break;
            }
            return "<FieldRef Name=\"" + result + "\"></FieldRef>";
        });
        var fieldRefs = fieldNames.map(function (fieldName) {
            return "<FieldRef Name=\"" + fieldName + "\"></FieldRef>";
        });
        return "<ViewFields>" + fieldRefs.join('') + "</ViewFields>";
    };
    return BaseListItemService;
}(BaseDataService));
export { BaseListItemService };
//# sourceMappingURL=BaseListItemService.js.map