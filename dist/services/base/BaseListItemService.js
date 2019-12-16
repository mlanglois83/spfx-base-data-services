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
import { cloneDeep, find, assign, findIndex } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { Constants, FieldType } from "../../constants/index";
import { BaseDataService } from "./BaseDataService";
import { UtilsService } from "..";
import { SPItem, User, TaxonomyTerm } from "../../models";
import { UserService } from "../graph/UserService";
import { isArray, stringIsNullOrEmpty } from "@pnp/common";
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
        _this.taxoMultiFieldNames = {};
        /***************************** External sources init and access **************************************/
        _this.initialized = false;
        _this.initPromise = null;
        /********** init for taxo multi ************/
        _this.fieldsInitialized = false;
        _this.initFieldsPromise = null;
        _this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
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
    BaseListItemService.prototype.init_internal = function () {
        return __awaiter(this, void 0, void 0, function () { return __generator(this, function (_a) {
            return [2 /*return*/];
        }); });
    };
    ;
    BaseListItemService.prototype.Init = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                if (!this.initPromise) {
                    this.initPromise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var fields, models, key, fieldDescription, error_1;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    if (!this.initialized) return [3 /*break*/, 1];
                                    resolve();
                                    return [3 /*break*/, 7];
                                case 1:
                                    this.initValues = {};
                                    _a.label = 2;
                                case 2:
                                    _a.trys.push([2, 6, , 7]);
                                    if (!this.init_internal) return [3 /*break*/, 4];
                                    return [4 /*yield*/, this.init_internal()];
                                case 3:
                                    _a.sent();
                                    _a.label = 4;
                                case 4:
                                    fields = this.ItemFields;
                                    models = [];
                                    for (key in fields) {
                                        if (fields.hasOwnProperty(key)) {
                                            fieldDescription = fields[key];
                                            if (fieldDescription.modelName && models.indexOf(fieldDescription.modelName) === -1) {
                                                models.push(fieldDescription.modelName);
                                            }
                                        }
                                    }
                                    return [4 /*yield*/, Promise.all(models.map(function (modelName) { return __awaiter(_this, void 0, void 0, function () {
                                            var service, values;
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        if (!!this.initValues[modelName]) return [3 /*break*/, 2];
                                                        service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                                        return [4 /*yield*/, service.getAll()];
                                                    case 1:
                                                        values = _a.sent();
                                                        this.initValues[modelName] = values;
                                                        _a.label = 2;
                                                    case 2: return [2 /*return*/];
                                                }
                                            });
                                        }); }))];
                                case 5:
                                    _a.sent();
                                    this.initialized = true;
                                    this.initPromise = null;
                                    resolve();
                                    return [3 /*break*/, 7];
                                case 6:
                                    error_1 = _a.sent();
                                    this.initPromise = null;
                                    reject(error_1);
                                    return [3 /*break*/, 7];
                                case 7: return [2 /*return*/];
                            }
                        });
                    }); });
                }
                return [2 /*return*/, this.initPromise];
            });
        });
    };
    BaseListItemService.prototype.initFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                if (!this.initFieldsPromise) {
                    this.initFieldsPromise = new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
                        var fields, taxofields, key, fieldDescription, error_2;
                        var _this = this;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    if (!this.fieldsInitialized) return [3 /*break*/, 1];
                                    resolve();
                                    return [3 /*break*/, 5];
                                case 1:
                                    this.taxoMultiFieldNames = {};
                                    _a.label = 2;
                                case 2:
                                    _a.trys.push([2, 4, , 5]);
                                    fields = this.ItemFields;
                                    taxofields = [];
                                    for (key in fields) {
                                        if (fields.hasOwnProperty(key)) {
                                            fieldDescription = fields[key];
                                            if (fieldDescription.fieldType === FieldType.TaxonomyMulti) {
                                                taxofields.push(fieldDescription.fieldName);
                                            }
                                        }
                                    }
                                    return [4 /*yield*/, Promise.all(taxofields.map(function (tf) { return __awaiter(_this, void 0, void 0, function () {
                                            var hiddenField;
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0: return [4 /*yield*/, this.list.fields.getByTitle(tf + "_0").select("InternalName").get()];
                                                    case 1:
                                                        hiddenField = _a.sent();
                                                        this.taxoMultiFieldNames[tf] = hiddenField.InternalName;
                                                        return [2 /*return*/];
                                                }
                                            });
                                        }); }))];
                                case 3:
                                    _a.sent();
                                    this.fieldsInitialized = true;
                                    this.initFieldsPromise = null;
                                    resolve();
                                    return [3 /*break*/, 5];
                                case 4:
                                    error_2 = _a.sent();
                                    this.initFieldsPromise = null;
                                    reject(error_2);
                                    return [3 /*break*/, 5];
                                case 5: return [2 /*return*/];
                            }
                        });
                    }); });
                }
                return [2 /*return*/, this.initFieldsPromise];
            });
        });
    };
    BaseListItemService.prototype.getServiceInitValues = function (modelName) {
        return this.initValues[modelName];
    };
    /****************************** get item methods ***********************************/
    BaseListItemService.prototype.getItemFromRest = function (spitem) {
        var _this = this;
        var item = new this.itemType();
        Object.keys(this.ItemFields).map(function (propertyName) {
            var fieldDescription = _this.ItemFields[propertyName];
            _this.setFieldValue(spitem, item, propertyName, fieldDescription);
        });
        return item;
    };
    BaseListItemService.prototype.setFieldValue = function (spitem, destItem, propertyName, fieldDescriptor) {
        var _this = this;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
                if (fieldDescriptor.fieldName === Constants.commonFields.version) {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? parseFloat(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                }
                else {
                    destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName] : fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Date:
                destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? new Date(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                var lookupId_1 = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if (lookupId_1 !== -1) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        var destElements = this.getServiceInitValues(fieldDescriptor.modelName);
                        var existing = find(destElements, function (destElement) {
                            return destElement.id === lookupId_1;
                        });
                        destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        destItem[propertyName] = lookupId_1;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.LookupMulti:
                var lookupIds = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results : spitem[fieldDescriptor.fieldName + "Id"]) : [];
                if (lookupIds.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        var val_1 = [];
                        var targetItems_1 = this.getServiceInitValues(fieldDescriptor.modelName);
                        lookupIds.forEach(function (id) {
                            var existing = find(targetItems_1, function (item) {
                                return item.id === id;
                            });
                            if (existing) {
                                val_1.push(existing);
                            }
                        });
                        destItem[propertyName] = val_1;
                    }
                    else {
                        destItem[propertyName] = lookupIds;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.User:
                var id_1 = spitem[fieldDescriptor.fieldName + "Id"] ? spitem[fieldDescriptor.fieldName + "Id"] : -1;
                if (id_1 !== -1) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        var users = this.getServiceInitValues(fieldDescriptor.modelName);
                        var existing = find(users, function (user) {
                            return user.spId === id_1;
                        });
                        destItem[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                    }
                    else {
                        destItem[propertyName] = id_1;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.UserMulti:
                var ids = spitem[fieldDescriptor.fieldName + "Id"] ? (spitem[fieldDescriptor.fieldName + "Id"].results ? spitem[fieldDescriptor.fieldName + "Id"].results : spitem[fieldDescriptor.fieldName + "Id"]) : [];
                if (ids.length > 0) {
                    if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                        // get values from init values
                        var val_2 = [];
                        var users_1 = this.getServiceInitValues(fieldDescriptor.modelName);
                        ids.forEach(function (id) {
                            var existing = find(users_1, function (user) {
                                return user.spId === id;
                            });
                            if (existing) {
                                val_2.push(existing);
                            }
                        });
                        destItem[propertyName] = val_2;
                    }
                    else {
                        destItem[propertyName] = ids;
                    }
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Taxonomy:
                var wssid = spitem[fieldDescriptor.fieldName] ? spitem[fieldDescriptor.fieldName].WssId : -1;
                if (id_1 !== -1) {
                    var terms_1 = this.getServiceInitValues(fieldDescriptor.modelName);
                    destItem[propertyName] = this.getTaxonomyTermByWssId(wssid, terms_1);
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.TaxonomyMulti:
                var terms = spitem[fieldDescriptor.fieldName] ? (spitem[fieldDescriptor.fieldName].results ? spitem[fieldDescriptor.fieldName].results : spitem[fieldDescriptor.fieldName]) : [];
                if (terms.length > 0) {
                    var allterms_1 = this.getServiceInitValues(fieldDescriptor.modelName);
                    destItem[propertyName] = terms.map(function (term) {
                        return _this.getTaxonomyTermByWssId(term.WssId, allterms_1);
                    });
                }
                else {
                    destItem[propertyName] = fieldDescriptor.defaultValue;
                }
                break;
            case FieldType.Json:
                destItem[propertyName] = spitem[fieldDescriptor.fieldName] ? JSON.parse(spitem[fieldDescriptor.fieldName]) : fieldDescriptor.defaultValue;
                break;
        }
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
                                var fieldDescription;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            fieldDescription = this.ItemFields[propertyName];
                                            return [4 /*yield*/, this.setRestFieldValue(item, spitem, propertyName, fieldDescription)];
                                        case 1:
                                            _a.sent();
                                            return [2 /*return*/];
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
    BaseListItemService.prototype.setRestFieldValue = function (item, destItem, propertyName, fieldDescriptor) {
        return __awaiter(this, void 0, void 0, function () {
            var itemValue, _a, firstLookupVal, idArray, _b, _c, firstUserVal, userIds, hiddenFieldName;
            var _this = this;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        itemValue = item[propertyName];
                        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
                        _a = fieldDescriptor.fieldType;
                        switch (_a) {
                            case FieldType.Simple: return [3 /*break*/, 1];
                            case FieldType.Date: return [3 /*break*/, 1];
                            case FieldType.Lookup: return [3 /*break*/, 2];
                            case FieldType.LookupMulti: return [3 /*break*/, 3];
                            case FieldType.User: return [3 /*break*/, 4];
                            case FieldType.UserMulti: return [3 /*break*/, 10];
                            case FieldType.Taxonomy: return [3 /*break*/, 16];
                            case FieldType.TaxonomyMulti: return [3 /*break*/, 17];
                            case FieldType.Json: return [3 /*break*/, 18];
                        }
                        return [3 /*break*/, 19];
                    case 1:
                        if (fieldDescriptor.fieldName !== Constants.commonFields.author &&
                            fieldDescriptor.fieldName !== Constants.commonFields.created &&
                            fieldDescriptor.fieldName !== Constants.commonFields.editor &&
                            fieldDescriptor.fieldName !== Constants.commonFields.modified &&
                            fieldDescriptor.fieldName !== Constants.commonFields.version) {
                            destItem[fieldDescriptor.fieldName] = itemValue;
                        }
                        return [3 /*break*/, 19];
                    case 2:
                        if (itemValue) {
                            if (typeof (itemValue) === "number") {
                                destItem[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                            }
                            else {
                                destItem[fieldDescriptor.fieldName + "Id"] = itemValue.id > 0 ? itemValue.id : null;
                            }
                        }
                        else {
                            destItem[fieldDescriptor.fieldName + "Id"] = null;
                        }
                        return [3 /*break*/, 19];
                    case 3:
                        if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                            firstLookupVal = itemValue[0];
                            if (typeof (firstLookupVal) === "number") {
                                destItem[fieldDescriptor.fieldName + "Id"] = { results: itemValue };
                            }
                            else {
                                idArray = destItem[fieldDescriptor.fieldName + "Id"] = { results: itemValue.map(function (lookupMultiElt) { return lookupMultiElt.id; }) };
                            }
                        }
                        else {
                            destItem[fieldDescriptor.fieldName + "Id"] = { results: [] };
                        }
                        return [3 /*break*/, 19];
                    case 4:
                        if (!itemValue) return [3 /*break*/, 8];
                        if (!(typeof (itemValue) === "number")) return [3 /*break*/, 5];
                        destItem[fieldDescriptor.fieldName + "Id"] = itemValue > 0 ? itemValue : null;
                        return [3 /*break*/, 7];
                    case 5:
                        _b = destItem;
                        _c = fieldDescriptor.fieldName + "Id";
                        return [4 /*yield*/, this.convertSingleUserFieldValue(itemValue)];
                    case 6:
                        _b[_c] = _d.sent();
                        _d.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        destItem[fieldDescriptor.fieldName + "Id"] = null;
                        _d.label = 9;
                    case 9: return [3 /*break*/, 19];
                    case 10:
                        if (!(itemValue && isArray(itemValue) && itemValue.length > 0)) return [3 /*break*/, 14];
                        firstUserVal = itemValue[0];
                        if (!(typeof (firstUserVal) === "number")) return [3 /*break*/, 11];
                        destItem[fieldDescriptor.fieldName + "Id"] = { results: itemValue };
                        return [3 /*break*/, 13];
                    case 11: return [4 /*yield*/, Promise.all(itemValue.map(function (user) {
                            return _this.convertSingleUserFieldValue(user);
                        }))];
                    case 12:
                        userIds = _d.sent();
                        destItem[fieldDescriptor.fieldName + "Id"] = { results: userIds };
                        _d.label = 13;
                    case 13: return [3 /*break*/, 15];
                    case 14:
                        destItem[fieldDescriptor.fieldName + "Id"] = { results: [] };
                        _d.label = 15;
                    case 15: return [3 /*break*/, 19];
                    case 16:
                        destItem[fieldDescriptor.fieldName] = this.convertTaxonomyFieldValue(itemValue);
                        return [3 /*break*/, 19];
                    case 17:
                        hiddenFieldName = this.taxoMultiFieldNames[fieldDescriptor.fieldName];
                        if (itemValue && isArray(itemValue) && itemValue.length > 0) {
                            destItem[hiddenFieldName] = this.convertTaxonomyMultiFieldValue(itemValue);
                        }
                        else {
                            destItem[hiddenFieldName] = null;
                        }
                        return [3 /*break*/, 19];
                    case 18:
                        destItem[fieldDescriptor.fieldName] = itemValue ? JSON.stringify(itemValue) : null;
                        return [3 /*break*/, 19];
                    case 19: return [2 /*return*/];
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
    BaseListItemService.prototype.convertTaxonomyMultiFieldValue = function (value) {
        var result = null;
        if (value) {
            result = value.map(function (term) { return "-1;#" + term.title + "|" + term.id + ";#"; }).join("");
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
            var result, isconnected, cachedDataDate, response, tempList, lastModifiedDate, error_3;
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
                        error_3 = _a.sent();
                        console.error(error_3);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/, result];
                }
            });
        });
    };
    /**********************************Service specific calls  *******************************/
    /**
     * Get items by caml query
     * @param query caml query (<Where></Where>)
     * @param orderBy array of <FieldRef Name='Field1' Ascending='TRUE'/>
     * @param limit  number of lines
     * @param lastId last id for paged queries
     */
    BaseListItemService.prototype.getByCamlQuery = function (query, orderBy, limit, lastId) {
        var queryXml = this.getQuery(query, orderBy, limit);
        var camlQuery = {
            ViewXml: queryXml
        };
        if (lastId !== undefined) {
            camlQuery.ListItemCollectionPosition = {
                "PagingInfo": "Paged=TRUE&p_ID=" + lastId
            };
        }
        return this.get(camlQuery);
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
                        return [4 /*yield*/, (_a = this.list).select.apply(_a, selectFields).getItemsByCAMLQuery(query)];
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
    BaseListItemService.prototype.getItemById_Internal = function (id) {
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
     * Get a list of items by id
     * @param id item id
     */
    BaseListItemService.prototype.getItemsById_Internal = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var results, selectFields, batch;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = [];
                        selectFields = this.getOdataFieldNames();
                        batch = sp.createBatch();
                        ids.forEach(function (id) {
                            var _a;
                            (_a = _this.list.items.getById(id)).select.apply(_a, selectFields).inBatch(batch).get().then(function (item) {
                                results.push(_this.getItemFromRest(item));
                            });
                        });
                        return [4 /*yield*/, batch.execute()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, results];
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
            var result, selectFields, converted, addResult, existing, error, converted, updateResult, version, converted, updateResult, version;
            var _a, _b, _c, _d;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        result = cloneDeep(item);
                        return [4 /*yield*/, this.initFields()];
                    case 1:
                        _e.sent();
                        selectFields = this.getOdataCommonFieldNames();
                        if (!(item.id < 0)) return [3 /*break*/, 8];
                        return [4 /*yield*/, this.getSPRestItem(item)];
                    case 2:
                        converted = _e.sent();
                        return [4 /*yield*/, (_a = this.list.items).select.apply(_a, selectFields).add(converted)];
                    case 3:
                        addResult = _e.sent();
                        return [4 /*yield*/, this.populateCommonFields(result, addResult.data)];
                    case 4:
                        _e.sent();
                        return [4 /*yield*/, this.updateWssIds(result, addResult.data)];
                    case 5:
                        _e.sent();
                        if (!(item.id < -1)) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.updateLinksInDb(Number(item.id), Number(result.id))];
                    case 6:
                        _e.sent();
                        _e.label = 7;
                    case 7: return [3 /*break*/, 23];
                    case 8:
                        if (!item.version) return [3 /*break*/, 17];
                        return [4 /*yield*/, this.list.items.getById(item.id).select(Constants.commonFields.version).get()];
                    case 9:
                        existing = _e.sent();
                        if (!(parseFloat(existing[Constants.commonFields.version]) > item.version)) return [3 /*break*/, 10];
                        error = new Error(ServicesConfiguration.configuration.translations.versionHigherErrorMessage);
                        error.name = Constants.Errors.ItemVersionConfict;
                        throw error;
                    case 10: return [4 /*yield*/, this.getSPRestItem(item)];
                    case 11:
                        converted = _e.sent();
                        return [4 /*yield*/, (_b = this.list.items.getById(item.id)).select.apply(_b, selectFields).update(converted)];
                    case 12:
                        updateResult = _e.sent();
                        return [4 /*yield*/, (_c = updateResult.item).select.apply(_c, selectFields).get()];
                    case 13:
                        version = _e.sent();
                        return [4 /*yield*/, this.populateCommonFields(result, version)];
                    case 14:
                        _e.sent();
                        return [4 /*yield*/, this.updateWssIds(result, version)];
                    case 15:
                        _e.sent();
                        _e.label = 16;
                    case 16: return [3 /*break*/, 23];
                    case 17: return [4 /*yield*/, this.getSPRestItem(item)];
                    case 18:
                        converted = _e.sent();
                        return [4 /*yield*/, this.list.items.getById(item.id).update(converted)];
                    case 19:
                        updateResult = _e.sent();
                        return [4 /*yield*/, (_d = updateResult.item).select.apply(_d, selectFields).get()];
                    case 20:
                        version = _e.sent();
                        return [4 /*yield*/, this.populateCommonFields(result, version)];
                    case 21:
                        _e.sent();
                        return [4 /*yield*/, this.updateWssIds(result, version)];
                    case 22:
                        _e.sent();
                        _e.label = 23;
                    case 23: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Delete an item
     * @param item SPItem derived class to be deleted
     */
    BaseListItemService.prototype.deleteItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.list.items.getById(item.id).recycle()];
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
                case FieldType.User:
                case FieldType.UserMulti:
                    result += "Id";
                default:
                    break;
            }
            return result;
        });
        return fieldNames;
    };
    BaseListItemService.prototype.getOdataCommonFieldNames = function () {
        var fields = this.ItemFields;
        var fieldNames = [Constants.commonFields.version];
        Object.keys(fields).filter(function (propertyName) {
            return fields.hasOwnProperty(propertyName);
        }).forEach(function (prop) {
            var fieldName = fields[prop].fieldName;
            if (fieldName === Constants.commonFields.author ||
                fieldName === Constants.commonFields.created ||
                fieldName === Constants.commonFields.editor ||
                fieldName === Constants.commonFields.modified) {
                var result = fields[prop].fieldName;
                switch (fields[prop].fieldType) {
                    case FieldType.Lookup:
                    case FieldType.LookupMulti:
                    case FieldType.User:
                    case FieldType.UserMulti:
                        result += "Id";
                    default:
                        break;
                }
                fieldNames.push(result);
            }
        });
        return fieldNames;
    };
    BaseListItemService.prototype.populateCommonFields = function (item, restItem) {
        return __awaiter(this, void 0, void 0, function () {
            var fields;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (item.id < 0) {
                            // update id
                            item.id = restItem.Id;
                        }
                        if (restItem[Constants.commonFields.version]) {
                            item.version = parseFloat(restItem[Constants.commonFields.version]);
                        }
                        fields = this.ItemFields;
                        return [4 /*yield*/, Promise.all(Object.keys(fields).filter(function (propertyName) {
                                var result = false;
                                if (fields.hasOwnProperty(propertyName)) {
                                    var fieldName = fields[propertyName].fieldName;
                                    return (fieldName === Constants.commonFields.author ||
                                        fieldName === Constants.commonFields.created ||
                                        fieldName === Constants.commonFields.editor ||
                                        fieldName === Constants.commonFields.modified);
                                }
                            }).map(function (prop) { return __awaiter(_this, void 0, void 0, function () {
                                var fieldName, _a, id_2, user, users, userService;
                                return __generator(this, function (_b) {
                                    switch (_b.label) {
                                        case 0:
                                            fieldName = fields[prop].fieldName;
                                            _a = fields[prop].fieldType;
                                            switch (_a) {
                                                case FieldType.Date: return [3 /*break*/, 1];
                                                case FieldType.User: return [3 /*break*/, 2];
                                            }
                                            return [3 /*break*/, 6];
                                        case 1:
                                            item[prop] = new Date(restItem[fieldName]);
                                            return [3 /*break*/, 7];
                                        case 2:
                                            id_2 = restItem[fieldName + "Id"];
                                            user = null;
                                            if (!this.initialized) return [3 /*break*/, 3];
                                            users = this.getServiceInitValues(User["name"]);
                                            user = find(users, function (u) { return u.spId === id_2; });
                                            return [3 /*break*/, 5];
                                        case 3:
                                            userService = new UserService();
                                            return [4 /*yield*/, userService.getBySpId(id_2)];
                                        case 4:
                                            user = _b.sent();
                                            _b.label = 5;
                                        case 5:
                                            item[prop] = user;
                                            return [3 /*break*/, 7];
                                        case 6:
                                            item[prop] = restItem[fieldName];
                                            return [3 /*break*/, 7];
                                        case 7: return [2 /*return*/];
                                    }
                                });
                            }); }))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * convert full item to db format (with links only)
     * @param item full provisionned item
     */
    BaseListItemService.prototype.convertItemToDbFormat = function (item) {
        var result = cloneDeep(item);
        delete result.__internalLinks;
        var _loop_1 = function (propertyName) {
            if (this_1.ItemFields.hasOwnProperty(propertyName)) {
                var fieldDescriptor = this_1.ItemFields[propertyName];
                switch (fieldDescriptor.fieldType) {
                    case FieldType.Lookup:
                    case FieldType.User:
                    case FieldType.Taxonomy:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            result.__internalLinks = result.__internalLinks || {};
                            result.__internalLinks[propertyName] = item[propertyName] ? item[propertyName].id : undefined;
                            delete result[propertyName];
                        }
                        break;
                    case FieldType.LookupMulti:
                    case FieldType.UserMulti:
                    case FieldType.TaxonomyMulti:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            var ids_1 = [];
                            if (item[propertyName]) {
                                item[propertyName].forEach(function (element) {
                                    if (element.id) {
                                        if ((typeof (element.id) === "number" && element.id > 0) || (typeof (element.id) === "string" && !stringIsNullOrEmpty(element.id))) {
                                            ids_1.push(element.id);
                                        }
                                    }
                                });
                            }
                            result.__internalLinks = result.__internalLinks || {};
                            result.__internalLinks[propertyName] = ids_1.length > 0 ? ids_1 : [];
                            delete result[propertyName];
                        }
                        break;
                    default:
                        break;
                }
            }
        };
        var this_1 = this;
        for (var propertyName in this.ItemFields) {
            _loop_1(propertyName);
        }
        return result;
    };
    /**
     * populate item from db storage
     * @param item db item with links in __internalLinks fields
     */
    BaseListItemService.prototype.mapItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var result, _loop_2, this_2, propertyName;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = cloneDeep(item);
                        return [4 /*yield*/, this.Init()];
                    case 1:
                        _a.sent();
                        _loop_2 = function (propertyName) {
                            if (this_2.ItemFields.hasOwnProperty(propertyName)) {
                                var fieldDescriptor = this_2.ItemFields[propertyName];
                                switch (fieldDescriptor.fieldType) {
                                    case FieldType.Lookup:
                                    case FieldType.User:
                                    case FieldType.Taxonomy:
                                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                            // get values from init values
                                            var id_3 = item.__internalLinks[propertyName] ? item.__internalLinks[propertyName] : null;
                                            if (id_3 !== null) {
                                                var destElements = this_2.getServiceInitValues(fieldDescriptor.modelName);
                                                var existing = find(destElements, function (destElement) {
                                                    return destElement.id === id_3;
                                                });
                                                result[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                                            }
                                            else {
                                                result[propertyName] = fieldDescriptor.defaultValue;
                                            }
                                        }
                                        break;
                                    case FieldType.LookupMulti:
                                    case FieldType.UserMulti:
                                    case FieldType.TaxonomyMulti:
                                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                                            // get values from init values
                                            var ids = item.__internalLinks[propertyName] ? item.__internalLinks[propertyName] : [];
                                            if (ids.length > 0) {
                                                var val_3 = [];
                                                var targetItems_2 = this_2.getServiceInitValues(fieldDescriptor.modelName);
                                                ids.forEach(function (id) {
                                                    var existing = find(targetItems_2, function (item) {
                                                        return item.id === id;
                                                    });
                                                    if (existing) {
                                                        val_3.push(existing);
                                                    }
                                                });
                                                result[propertyName] = val_3;
                                            }
                                            else {
                                                result[propertyName] = fieldDescriptor.defaultValue;
                                            }
                                        }
                                        break;
                                    default:
                                        result[propertyName] = item[propertyName];
                                        break;
                                }
                            }
                        };
                        this_2 = this;
                        for (propertyName in this.ItemFields) {
                            _loop_2(propertyName);
                        }
                        delete result.__internalLinks;
                        return [2 /*return*/, result];
                }
            });
        });
    };
    BaseListItemService.prototype.updateLinkedTransactions = function (oldId, newId, nextTransactions) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                // Update items pointing to this in transactions
                nextTransactions.forEach(function (transaction) {
                    var currentObject = null;
                    var needUpdate = false;
                    var service = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
                    var fields = service.ItemFields;
                    // search for lookup fields
                    for (var propertyName in fields) {
                        if (fields.hasOwnProperty(propertyName)) {
                            var fieldDescription = fields[propertyName];
                            if (fieldDescription.refItemName === _this.itemType["name"]) {
                                // get object if not done yet
                                if (!currentObject) {
                                    var destType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                                    var currentObject_1 = new destType();
                                    assign(currentObject_1, transaction.itemData);
                                }
                                if (fieldDescription.fieldType === FieldType.Lookup) {
                                    if (fieldDescription.modelName) {
                                        // search in __internalLinks
                                        if (currentObject.__internalLinks && currentObject.__internalLinks[propertyName] === oldId) {
                                            currentObject.__internalLinks[propertyName] = newId;
                                            needUpdate = true;
                                        }
                                    }
                                    else if (currentObject[propertyName] === oldId) {
                                        // change field
                                        currentObject[propertyName] = newId;
                                        needUpdate = true;
                                    }
                                }
                                else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                    if (fieldDescription.modelName) {
                                        // serch in __internalLinks
                                        if (currentObject.__internalLinks && currentObject.__internalLinks[propertyName] && isArray(currentObject.__internalLinks[propertyName])) {
                                            // find item
                                            var lookupidx = findIndex(currentObject.__internalLinks[propertyName], function (id) { return id === oldId; });
                                            // change id
                                            if (lookupidx > -1) {
                                                currentObject.__internalLinks[propertyName] = newId;
                                                needUpdate = true;
                                            }
                                        }
                                    }
                                    else if (currentObject[propertyName] && isArray(currentObject[propertyName])) {
                                        // find index
                                        var lookupidx = findIndex(currentObject[propertyName], function (id) { return id === oldId; });
                                        // change field
                                        // change id
                                        if (lookupidx > -1) {
                                            currentObject[propertyName] = newId;
                                            needUpdate = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (needUpdate) {
                        transaction.itemData = assign({}, currentObject);
                        _this.transactionService.addOrUpdateItem(transaction);
                    }
                });
                return [2 /*return*/, nextTransactions];
            });
        });
    };
    BaseListItemService.prototype.updateLinksInDb = function (oldId, newId) {
        return __awaiter(this, void 0, void 0, function () {
            var allFields, _loop_3, _a, _b, _i, modelName;
            var _this = this;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        allFields = assign({}, this.itemType["Fields"]);
                        delete allFields[SPItem["name"]];
                        delete allFields[this.itemType["name"]];
                        _loop_3 = function (modelName) {
                            var modelFields_1, lookupProperties_1, service, allitems, updated_1;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        if (!allFields.hasOwnProperty(modelName)) return [3 /*break*/, 3];
                                        modelFields_1 = allFields[modelName];
                                        lookupProperties_1 = Object.keys(modelFields_1).filter(function (prop) {
                                            return modelFields_1[prop].refItemName &&
                                                modelFields_1[prop].refItemName === _this.itemType["name"];
                                        });
                                        if (!(lookupProperties_1.length > 0)) return [3 /*break*/, 3];
                                        service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                        return [4 /*yield*/, service.__getAllFromCache()];
                                    case 1:
                                        allitems = _a.sent();
                                        updated_1 = [];
                                        allitems.forEach(function (element) {
                                            var needUpdate = false;
                                            lookupProperties_1.forEach(function (propertyName) {
                                                var fieldDescription = modelFields_1[propertyName];
                                                if (fieldDescription.fieldType === FieldType.Lookup) {
                                                    if (fieldDescription.modelName) {
                                                        // serch in __internalLinks
                                                        if (element.__internalLinks && element.__internalLinks[propertyName] === oldId) {
                                                            element.__internalLinks[propertyName] = newId;
                                                            needUpdate = true;
                                                        }
                                                    }
                                                    else if (element[propertyName] === oldId) {
                                                        // change field
                                                        element[propertyName] = newId;
                                                        needUpdate = true;
                                                    }
                                                }
                                                else if (fieldDescription.fieldType === FieldType.LookupMulti) {
                                                    if (fieldDescription.modelName) {
                                                        // serch in __internalLinks
                                                        if (element.__internalLinks && element.__internalLinks[propertyName] && isArray(element.__internalLinks[propertyName])) {
                                                            // find item
                                                            var lookupidx = findIndex(element.__internalLinks[propertyName], function (id) { return id === oldId; });
                                                            // change id
                                                            if (lookupidx > -1) {
                                                                element.__internalLinks[propertyName] = newId;
                                                                needUpdate = true;
                                                            }
                                                        }
                                                    }
                                                    else if (element[propertyName] && isArray(element[propertyName])) {
                                                        // find index
                                                        var lookupidx = findIndex(element[propertyName], function (id) { return id === oldId; });
                                                        // change field
                                                        // change id
                                                        if (lookupidx > -1) {
                                                            element[propertyName] = newId;
                                                            needUpdate = true;
                                                        }
                                                    }
                                                }
                                            });
                                            if (needUpdate) {
                                                updated_1.push(element);
                                            }
                                        });
                                        if (!(updated_1.length > 0)) return [3 /*break*/, 3];
                                        return [4 /*yield*/, service.__updateCache.apply(service, updated_1)];
                                    case 2:
                                        _a.sent();
                                        _a.label = 3;
                                    case 3: return [2 /*return*/];
                                }
                            });
                        };
                        _a = [];
                        for (_b in allFields)
                            _a.push(_b);
                        _i = 0;
                        _c.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 4];
                        modelName = _a[_i];
                        return [5 /*yield**/, _loop_3(modelName)];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    BaseListItemService.prototype.updateWssIds = function (item, spItem) {
        return __awaiter(this, void 0, void 0, function () {
            var fields, _loop_4, this_3, _a, _b, _i, propertyName;
            var _this = this;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        fields = this.ItemFields;
                        _loop_4 = function (propertyName) {
                            var fieldDescription_1, needUpdate, wssid, id_4, service, term, idx, updated_2, terms, service_1;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        if (!fields.hasOwnProperty(propertyName)) return [3 /*break*/, 8];
                                        fieldDescription_1 = fields[propertyName];
                                        if (!(fieldDescription_1.fieldType === FieldType.Taxonomy)) return [3 /*break*/, 4];
                                        needUpdate = false;
                                        wssid = spItem[fieldDescription_1.fieldName] ? spItem[fieldDescription_1.fieldName].WssId : -1;
                                        if (!(wssid !== -1)) return [3 /*break*/, 3];
                                        id_4 = item[propertyName].id;
                                        service = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription_1.modelName);
                                        return [4 /*yield*/, service.__getFromCache(id_4)];
                                    case 1:
                                        term = _a.sent();
                                        if (term instanceof TaxonomyTerm) {
                                            term.wssids = term.wssids || [];
                                            if (term.wssids.indexOf(wssid) === -1) {
                                                term.wssids.push(wssid);
                                                needUpdate = true;
                                            }
                                        }
                                        if (!needUpdate) return [3 /*break*/, 3];
                                        return [4 /*yield*/, service.__updateCache(term)];
                                    case 2:
                                        _a.sent();
                                        // update initValues
                                        if (this_3.initialized) {
                                            idx = findIndex(this_3.initValues[fieldDescription_1.modelName], function (t) { return t.id === id_4; });
                                            if (idx !== -1) {
                                                this_3.initValues[fieldDescription_1.modelName][idx] = term;
                                            }
                                        }
                                        _a.label = 3;
                                    case 3: return [3 /*break*/, 8];
                                    case 4:
                                        if (!(fieldDescription_1.fieldType === FieldType.TaxonomyMulti)) return [3 /*break*/, 8];
                                        updated_2 = [];
                                        terms = spItem[fieldDescription_1.fieldName] ? spItem[fieldDescription_1.fieldName].results : [];
                                        service_1 = ServicesConfiguration.configuration.serviceFactory.create(fieldDescription_1.modelName);
                                        if (!(terms && terms.length > 0)) return [3 /*break*/, 6];
                                        return [4 /*yield*/, Promise.all(terms.map(function (termitem) { return __awaiter(_this, void 0, void 0, function () {
                                                var wssid, id, term;
                                                return __generator(this, function (_a) {
                                                    switch (_a.label) {
                                                        case 0:
                                                            wssid = termitem.WssId;
                                                            id = termitem.TermGuid;
                                                            term = find(updated_2, function (u) { return u.id === id; });
                                                            if (!!term) return [3 /*break*/, 2];
                                                            return [4 /*yield*/, service_1.__getFromCache(id)];
                                                        case 1:
                                                            term = _a.sent();
                                                            _a.label = 2;
                                                        case 2:
                                                            if (term instanceof TaxonomyTerm) {
                                                                term.wssids = term.wssids || [];
                                                                if (term.wssids.indexOf(wssid) === -1) {
                                                                    term.wssids.push(wssid);
                                                                    if (!find(updated_2, function (u) { return u.id === id; })) {
                                                                        updated_2.push(term);
                                                                    }
                                                                }
                                                            }
                                                            return [2 /*return*/];
                                                    }
                                                });
                                            }); }))];
                                    case 5:
                                        _a.sent();
                                        _a.label = 6;
                                    case 6:
                                        if (!(updated_2.length > 0)) return [3 /*break*/, 8];
                                        return [4 /*yield*/, service_1.__updateCache.apply(service_1, updated_2)];
                                    case 7:
                                        _a.sent();
                                        // update initValues
                                        if (this_3.initialized) {
                                            updated_2.forEach(function (u) {
                                                var idx = findIndex(_this.initValues[fieldDescription_1.modelName], function (t) { return t.id === u.id; });
                                                if (idx !== -1) {
                                                    _this.initValues[fieldDescription_1.modelName][idx] = u;
                                                }
                                            });
                                        }
                                        _a.label = 8;
                                    case 8: return [2 /*return*/];
                                }
                            });
                        };
                        this_3 = this;
                        _a = [];
                        for (_b in fields)
                            _a.push(_b);
                        _i = 0;
                        _c.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 4];
                        propertyName = _a[_i];
                        return [5 /*yield**/, _loop_4(propertyName)];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     *
     * @param query caml query (<Where></Where>)
     * @param orderBy array of <FieldRef Name='Field1' Ascending='TRUE'/>
     * @param limit  number of lines
     */
    BaseListItemService.prototype.getQuery = function (query, orderBy, limit) {
        return "<View Scope=\"RecursiveAll\">\n            <Query>\n                " + query + "\n                " + (orderBy ? "<OrderBy>" + orderBy.join('') + "</OrderBy>" : "") + "\n            </Query>            \n            " + (limit !== undefined ? "<RowLimit>" + limit + "</RowLimit>" : "") + "\n        </View>";
    };
    return BaseListItemService;
}(BaseDataService));
export { BaseListItemService };
//# sourceMappingURL=BaseListItemService.js.map