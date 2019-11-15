var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
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
import { UserService } from "../../services";
import { spField } from "../../decorators";
import { FieldType } from "../../constants";
/**
 * Base object for sharepoint abstraction objects
 */
var SPItem = /** @class */ (function () {
    /**
     * Constructs a SPItem object
     * @param item object returned by sp call
     */
    function SPItem(item) {
        this.id = -1;
        if (item != undefined) {
            this.title = item["Title"] != undefined ? item["Title"] : "";
            this.id = item["ID"] != undefined ? item["ID"] : -1;
            this.version = item["OData__UIVersionString"] ? parseFloat(item["OData__UIVersionString"]) : undefined;
        }
    }
    /**
     * Returns a copy of the object compatible with sp calls
     */
    SPItem.prototype.convert = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result;
            return __generator(this, function (_a) {
                result = {};
                result["Title"] = this.title;
                result["ID"] = this.id;
                return [2 /*return*/, result];
            });
        });
    };
    SPItem.prototype.convertTaxonomyFieldValue = function (value) {
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
    SPItem.prototype.convertSingleUserFieldValue = function (value) {
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
    SPItem.prototype.convertMultiUserFieldValue = function (value) {
        return __awaiter(this, void 0, void 0, function () {
            var result;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        if (!value) return [3 /*break*/, 2];
                        return [4 /*yield*/, Promise.all(value.map(function (val) {
                                return _this.convertSingleUserFieldValue(val);
                            }))];
                    case 1:
                        result = _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/, result];
                }
            });
        });
    };
    Object.defineProperty(SPItem.prototype, "isValid", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * called after update was made on sp list
     * @param addResultData added item from rest call
     */
    SPItem.prototype.onAddCompleted = function (addResultData) {
        this.id = addResultData.Id;
        if (addResultData["OData__UIVersionString"]) {
            this.version = parseFloat(addResultData["OData__UIVersionString"]);
        }
    };
    /**
     * called after update was made on sp list
     * @param updateResult updated item from rest call
     */
    SPItem.prototype.onUpdateCompleted = function (updateResult) {
        if (updateResult["OData__UIVersionString"]) {
            this.version = parseFloat(updateResult["OData__UIVersionString"]);
        }
    };
    __decorate([
        spField({ fieldName: "ID", fieldType: FieldType.Simple, defaultValue: -1 })
    ], SPItem.prototype, "id", void 0);
    __decorate([
        spField({ fieldName: "Title", fieldType: FieldType.Simple, defaultValue: "" })
    ], SPItem.prototype, "title", void 0);
    __decorate([
        spField({ fieldName: "OData__UIVersionString", fieldType: FieldType.Simple, defaultValue: undefined })
    ], SPItem.prototype, "version", void 0);
    return SPItem;
}());
export { SPItem };
//# sourceMappingURL=SPItem.js.map