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
var __spreadArrays = (this && this.__spreadArrays) || function () {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};
import { BaseService } from "../base/BaseService";
import { SPFile } from "../../models/index";
import { TransactionType, Constants } from "../../constants/index";
import { assign } from "@microsoft/sp-lodash-subset";
import { TransactionService } from "./TransactionService";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
var SynchronizationService = /** @class */ (function (_super) {
    __extends(SynchronizationService, _super);
    function SynchronizationService() {
        var _this = _super.call(this) || this;
        _this.transactionService = new TransactionService();
        return _this;
    }
    SynchronizationService.prototype.run = function () {
        return __awaiter(this, void 0, void 0, function () {
            var errors, transactions, _loop_1, this_1, index;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        errors = [];
                        return [4 /*yield*/, this.transactionService.getAll()];
                    case 1:
                        transactions = _a.sent();
                        _loop_1 = function (index) {
                            var transaction, itemType, dataService, item, _a, oldId_1, isAdd, dbItem, updatedItem_1, nextTransactions, error_1;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        transaction = transactions[index];
                                        itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                                        dataService = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
                                        // init service for tardive links
                                        return [4 /*yield*/, dataService.Init()];
                                    case 1:
                                        // init service for tardive links
                                        _b.sent();
                                        item = assign(new itemType(), transaction.itemData);
                                        _a = transaction.title;
                                        switch (_a) {
                                            case TransactionType.AddOrUpdate: return [3 /*break*/, 2];
                                            case TransactionType.Delete: return [3 /*break*/, 15];
                                        }
                                        return [3 /*break*/, 20];
                                    case 2:
                                        oldId_1 = item.id;
                                        isAdd = typeof (oldId_1) === "number" && oldId_1 < 0;
                                        return [4 /*yield*/, dataService.mapItem(item)];
                                    case 3:
                                        dbItem = _b.sent();
                                        return [4 /*yield*/, dataService.addOrUpdateItem(dbItem)];
                                    case 4:
                                        updatedItem_1 = _b.sent();
                                        if (!(isAdd && !updatedItem_1.error)) return [3 /*break*/, 9];
                                        nextTransactions = [];
                                        if (!(index < transactions.length - 1)) return [3 /*break*/, 6];
                                        return [4 /*yield*/, Promise.all(transactions.slice(index + 1).map(function (updatedTr) { return __awaiter(_this, void 0, void 0, function () {
                                                return __generator(this, function (_a) {
                                                    switch (_a.label) {
                                                        case 0:
                                                            if (!(updatedTr.itemType === transaction.itemType &&
                                                                updatedTr.itemData.id === oldId_1)) return [3 /*break*/, 2];
                                                            updatedTr.itemData.id = updatedItem_1.item.id;
                                                            updatedTr.itemData.version = updatedItem_1.item.version;
                                                            return [4 /*yield*/, this.transactionService.addOrUpdateItem(updatedTr)];
                                                        case 1:
                                                            _a.sent();
                                                            _a.label = 2;
                                                        case 2: return [2 /*return*/, updatedTr];
                                                    }
                                                });
                                            }); }))];
                                    case 5:
                                        nextTransactions = _b.sent();
                                        _b.label = 6;
                                    case 6:
                                        if (!dataService.updateLinkedTransactions) return [3 /*break*/, 8];
                                        return [4 /*yield*/, dataService.updateLinkedTransactions(oldId_1, updatedItem_1.item.id, nextTransactions)];
                                    case 7:
                                        nextTransactions = _b.sent();
                                        _b.label = 8;
                                    case 8:
                                        if (index < transactions.length - 1) {
                                            transactions.splice.apply(transactions, __spreadArrays([index + 1, transactions.length - index - 1], nextTransactions));
                                        }
                                        _b.label = 9;
                                    case 9:
                                        if (!updatedItem_1.error) return [3 /*break*/, 12];
                                        errors.push(this_1.formatError(transaction, updatedItem_1.error.message));
                                        if (!(updatedItem_1.error.name === Constants.Errors.ItemVersionConfict)) return [3 /*break*/, 11];
                                        return [4 /*yield*/, this_1.transactionService.deleteItem(transaction)];
                                    case 10:
                                        _b.sent();
                                        _b.label = 11;
                                    case 11: return [3 /*break*/, 14];
                                    case 12: return [4 /*yield*/, this_1.transactionService.deleteItem(transaction)];
                                    case 13:
                                        _b.sent();
                                        _b.label = 14;
                                    case 14: return [3 /*break*/, 20];
                                    case 15:
                                        _b.trys.push([15, 18, , 19]);
                                        return [4 /*yield*/, dataService.deleteItem(item)];
                                    case 16:
                                        _b.sent();
                                        return [4 /*yield*/, this_1.transactionService.deleteItem(transaction)];
                                    case 17:
                                        _b.sent();
                                        return [3 /*break*/, 19];
                                    case 18:
                                        error_1 = _b.sent();
                                        errors.push(this_1.formatError(transaction, error_1.message));
                                        return [3 /*break*/, 19];
                                    case 19: return [3 /*break*/, 20];
                                    case 20: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        index = 0;
                        _a.label = 2;
                    case 2:
                        if (!(index < transactions.length)) return [3 /*break*/, 5];
                        return [5 /*yield**/, _loop_1(index)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4:
                        index++;
                        return [3 /*break*/, 2];
                    case 5: 
                    //return errors list
                    return [2 /*return*/, errors];
                }
            });
        });
    };
    SynchronizationService.prototype.formatError = function (transaction, message) {
        var operationLabel;
        var itemTypeLabel;
        var itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
        var item = assign(new itemType(), transaction.itemData);
        switch (transaction.title) {
            case TransactionType.AddOrUpdate:
                if (item instanceof SPFile) {
                    operationLabel = ServicesConfiguration.configuration.translations.UploadLabel;
                }
                else if (item.id < 0) {
                    operationLabel = ServicesConfiguration.configuration.translations.AddLabel;
                }
                else {
                    operationLabel = ServicesConfiguration.configuration.translations.UpdateLabel;
                }
                break;
            case TransactionType.Delete:
                operationLabel = ServicesConfiguration.configuration.translations.DeleteLabel;
                break;
            default: break;
        }
        itemTypeLabel = ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] ? ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] : transaction.itemType;
        return Text.format(ServicesConfiguration.configuration.translations.SynchronisationErrorFormat, itemTypeLabel, operationLabel, item.title, item.id, message);
    };
    return SynchronizationService;
}(BaseService));
export { SynchronizationService };
