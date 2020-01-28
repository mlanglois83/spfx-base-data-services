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
import { taxonomy } from "@pnp/sp-taxonomy";
import { UtilsService } from "../";
import { Constants } from "../../constants/index";
import { BaseDataService } from "./BaseDataService";
import { TaxonomyHiddenListService } from "../";
import { find } from "@microsoft/sp-lodash-subset";
import { Text } from "@microsoft/sp-core-library";
import { stringIsNullOrEmpty } from "@pnp/common";
import { ServicesConfiguration } from "../..";
var standardTermSetCacheDuration = 10;
/**
 * Base service for sp termset operations
 */
var BaseTermsetService = /** @class */ (function (_super) {
    __extends(BaseTermsetService, _super);
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param termsetname termset name
     */
    function BaseTermsetService(type, termsetnameorid, tableName, isGlobal, cacheDuration) {
        if (isGlobal === void 0) { isGlobal = true; }
        if (cacheDuration === void 0) { cacheDuration = standardTermSetCacheDuration; }
        var _this = _super.call(this, type, tableName, cacheDuration) || this;
        _this.utilsService = new UtilsService();
        _this.taxonomyHiddenListService = new TaxonomyHiddenListService();
        _this.termsetnameorid = termsetnameorid;
        _this.isGlobal = isGlobal;
        return _this;
    }
    Object.defineProperty(BaseTermsetService.prototype, "termset", {
        /**
         * Associeted termset (pnpjs)
         */
        get: function () {
            if (this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)) {
                if (this.isGlobal) {
                    return taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(this.termsetnameorid);
                }
                else {
                    return taxonomy.getDefaultSiteCollectionTermStore().getSiteCollectionGroup().termSets.getById(this.termsetnameorid);
                }
            }
            else {
                if (this.isGlobal) {
                    return taxonomy.getDefaultSiteCollectionTermStore().getTermSetsByName(this.termsetnameorid, 1033).getByName(this.termsetnameorid);
                }
                else {
                    return taxonomy.getDefaultSiteCollectionTermStore().getSiteCollectionGroup().termSets.getByName(this.termsetnameorid);
                }
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(BaseTermsetService.prototype, "customSortOrder", {
        get: function () {
            return localStorage.getItem(Text.format(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName));
        },
        set: function (value) {
            localStorage.setItem(Text.format(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName), value ? value : "");
        },
        enumerable: true,
        configurable: true
    });
    BaseTermsetService.prototype.getWssIds = function (termId) {
        return __awaiter(this, void 0, void 0, function () {
            var taxonomyHiddenItems;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.taxonomyHiddenListService.getAll()];
                    case 1:
                        taxonomyHiddenItems = _a.sent();
                        return [2 /*return*/, taxonomyHiddenItems.filter(function (taxItem) {
                                return taxItem.termId === termId;
                            }).map(function (filteredItem) {
                                return filteredItem.id;
                            })];
                }
            });
        });
    };
    /**
     * Retrieve all terms
     */
    BaseTermsetService.prototype.getAll_Internal = function () {
        return __awaiter(this, void 0, void 0, function () {
            var spterms, ts, taxonomyHiddenItems;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.termset.terms.get()];
                    case 1:
                        spterms = _a.sent();
                        return [4 /*yield*/, this.termset.get()];
                    case 2:
                        ts = _a.sent();
                        this.customSortOrder = ts.CustomSortOrder;
                        return [4 /*yield*/, this.taxonomyHiddenListService.getAll()];
                    case 3:
                        taxonomyHiddenItems = _a.sent();
                        return [2 /*return*/, spterms.map(function (term) {
                                var result = new _this.itemType(term);
                                result.wssids = [];
                                for (var _i = 0, taxonomyHiddenItems_1 = taxonomyHiddenItems; _i < taxonomyHiddenItems_1.length; _i++) {
                                    var taxonomyHiddenItem = taxonomyHiddenItems_1[_i];
                                    if (taxonomyHiddenItem.termId == result.id) {
                                        result.wssids.push(taxonomyHiddenItem.id);
                                    }
                                }
                                return result;
                            })];
                }
            });
        });
    };
    BaseTermsetService.prototype.getItemById_Internal = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, spterm;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        return [4 /*yield*/, this.termset.terms.getById(id)];
                    case 1:
                        spterm = _a.sent();
                        if (spterm) {
                            result = new this.itemType(spterm);
                        }
                        return [2 /*return*/, result];
                }
            });
        });
    };
    BaseTermsetService.prototype.getItemsById_Internal = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var results, batch;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = [];
                        batch = taxonomy.createBatch();
                        ids.forEach(function (id) {
                            _this.termset.terms.getById(id).inBatch(batch).get().then(function (term) {
                                results.push(new _this.itemType(term));
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
    BaseTermsetService.prototype.get_Internal = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error('Not Implemented');
            });
        });
    };
    BaseTermsetService.prototype.addOrUpdateItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error("Not implemented");
            });
        });
    };
    BaseTermsetService.prototype.deleteItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error("Not implemented");
            });
        });
    };
    BaseTermsetService.prototype.getOrderedChildTerms = function (term, allTerms) {
        var _this = this;
        //items.sort((a: T,b: T) => {return a.path.localeCompare(b.path);});
        var result = [];
        var childterms = allTerms.filter(function (t) { return t.path.indexOf(term.path) == 0; });
        var level = term.path.split(";").length;
        var directChilds = childterms.filter(function (ct) { return ct.path.split(";").length === level + 1; });
        if (!stringIsNullOrEmpty(term.customSortOrder)) {
            var terms_1 = new Array();
            var orderIds = term.customSortOrder.split(":");
            orderIds.forEach(function (id) {
                var t = find(directChilds, function (spterm) {
                    return spterm.id === id;
                });
                terms_1.push(t);
            });
            directChilds = terms_1;
        }
        directChilds.forEach(function (dc) {
            result.push(dc);
            var dcchildren = _this.getOrderedChildTerms(dc, childterms);
            if (dcchildren.length > 0) {
                result.push.apply(result, dcchildren);
            }
        });
        return result;
    };
    BaseTermsetService.prototype.getAll = function () {
        return __awaiter(this, void 0, void 0, function () {
            var items, result, rootTerms, terms_2, orderIds;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.getAll.call(this)];
                    case 1:
                        items = _a.sent();
                        result = [];
                        rootTerms = items.filter(function (item) { return item.path.indexOf(";") === -1; });
                        if (!stringIsNullOrEmpty(this.customSortOrder)) {
                            terms_2 = new Array();
                            orderIds = this.customSortOrder.split(":");
                            orderIds.forEach(function (id) {
                                var term = find(rootTerms, function (spterm) {
                                    return spterm.id === id;
                                });
                                terms_2.push(term);
                            });
                            rootTerms = terms_2;
                        }
                        rootTerms.forEach(function (rt) {
                            result.push(rt);
                            var rtchildren = _this.getOrderedChildTerms(rt, items);
                            if (rtchildren.length > 0) {
                                result.push.apply(result, rtchildren);
                            }
                        });
                        return [2 /*return*/, result];
                }
            });
        });
    };
    return BaseTermsetService;
}(BaseDataService));
export { BaseTermsetService };
//# sourceMappingURL=BaseTermsetService.js.map