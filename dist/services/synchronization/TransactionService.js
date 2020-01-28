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
import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction, SPFile } from "../../models/index";
import { assign } from "@microsoft/sp-lodash-subset";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
var TransactionService = /** @class */ (function (_super) {
    __extends(TransactionService, _super);
    function TransactionService() {
        var _this = _super.call(this, OfflineTransaction, "Transaction") || this;
        _this.transactionFileService = new BaseDbService(SPFile, "TransactionFiles");
        return _this;
    }
    /**
     * Add or update an item in DB and returns updated item
     * @param item Item to add or update
     */
    TransactionService.prototype.addOrUpdateItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var result, existing, file, baseUrl;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        if (!this.isFile(item.itemType)) return [3 /*break*/, 6];
                        return [4 /*yield*/, this.getItemById(item.id)];
                    case 1:
                        existing = _a.sent();
                        if (!existing) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.deleteItem(existing)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        file = assign(new SPFile(), item.itemData);
                        baseUrl = file.serverRelativeUrl;
                        item.itemData = new Date().getTime() + "_" + file.serverRelativeUrl;
                        file.serverRelativeUrl = item.itemData;
                        return [4 /*yield*/, this.transactionFileService.addOrUpdateItem(file)];
                    case 4:
                        _a.sent();
                        return [4 /*yield*/, _super.prototype.addOrUpdateItem.call(this, item)];
                    case 5:
                        result = _a.sent();
                        // reassign values for result
                        file.serverRelativeUrl = baseUrl;
                        result.item.itemData = assign({}, file);
                        return [3 /*break*/, 8];
                    case 6: return [4 /*yield*/, _super.prototype.addOrUpdateItem.call(this, item)];
                    case 7:
                        result = _a.sent();
                        _a.label = 8;
                    case 8: return [2 /*return*/, result];
                }
            });
        });
    };
    TransactionService.prototype.deleteItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var transaction, file;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.isFile(item.itemType)) return [3 /*break*/, 3];
                        return [4 /*yield*/, _super.prototype.getItemById.call(this, item.id)];
                    case 1:
                        transaction = _a.sent();
                        file = new SPFile();
                        file.serverRelativeUrl = transaction.itemData;
                        return [4 /*yield*/, this.transactionFileService.deleteItem(file)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3: return [4 /*yield*/, _super.prototype.deleteItem.call(this, item)];
                    case 4:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * add items in table (ids updated)
     * @param newItems
     */
    TransactionService.prototype.addOrUpdateItems = function (newItems) {
        return __awaiter(this, void 0, void 0, function () {
            var updateResults;
            var _this = this;
            return __generator(this, function (_a) {
                updateResults = Promise.all(newItems.map(function (item) { return __awaiter(_this, void 0, void 0, function () {
                    var result;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0: return [4 /*yield*/, this.addOrUpdateItem(item)];
                            case 1:
                                result = _a.sent();
                                return [2 /*return*/, result.item];
                        }
                    });
                }); }));
                return [2 /*return*/, updateResults];
            });
        });
    };
    /**
     * Retrieve all items from db table
     */
    TransactionService.prototype.getAll = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.getAll.call(this)];
                    case 1:
                        result = _a.sent();
                        return [4 /*yield*/, Promise.all(result.map(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                var file;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            if (!this.isFile(item.itemType)) return [3 /*break*/, 2];
                                            return [4 /*yield*/, this.transactionFileService.getItemById(item.itemData)];
                                        case 1:
                                            file = _a.sent();
                                            if (file) {
                                                file.serverRelativeUrl = file.serverRelativeUrl.replace(/^\d+_(.*)$/g, "$1");
                                                item.itemData = assign({}, file);
                                            }
                                            _a.label = 2;
                                        case 2: return [2 /*return*/, item];
                                    }
                                });
                            }); }))];
                    case 2:
                        result = _a.sent();
                        return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Get a transaction given its id
     * @param id transaction id
     */
    TransactionService.prototype.getItemById = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, file;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.getItemById.call(this, id)];
                    case 1:
                        result = _a.sent();
                        if (!(result && result.itemType === SPFile["name"])) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.transactionFileService.getItemById(result.itemData)];
                    case 2:
                        file = _a.sent();
                        if (file) {
                            file.serverRelativeUrl = file.serverRelativeUrl.replace(/^\d+_(.*)$/g, "$1");
                            result.itemData = assign({}, file);
                        }
                        _a.label = 3;
                    case 3: return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Clear table
     */
    TransactionService.prototype.clear = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.transactionFileService.clear()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, _super.prototype.clear.call(this)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    TransactionService.prototype.isFile = function (itemTypeName) {
        var itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(itemTypeName);
        var instance = new itemType();
        return (instance instanceof SPFile);
    };
    return TransactionService;
}(BaseDbService));
export { TransactionService };
//# sourceMappingURL=TransactionService.js.map