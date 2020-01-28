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
import { sp } from "@pnp/sp";
import * as mime from "mime-types";
import { UtilsService } from "../";
import { SPFile } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { ServicesConfiguration } from "../..";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
/**
 * Base service for sp files operations
 */
var BaseFileService = /** @class */ (function (_super) {
    __extends(BaseFileService, _super);
    /**
     *
     * @param type items type
     * @param context current sp component context
     * @param listRelativeUrl list web relative url
     */
    function BaseFileService(type, listRelativeUrl, tableName) {
        var _this = _super.call(this, type, tableName) || this;
        _this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
        return _this;
    }
    Object.defineProperty(BaseFileService.prototype, "list", {
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
     * Retrieve all items
     */
    BaseFileService.prototype.getAll_Internal = function () {
        return __awaiter(this, void 0, void 0, function () {
            var files;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.list.items.filter('FSObjType eq 0').select('FileRef', 'FileLeafRef').get()];
                    case 1:
                        files = _a.sent();
                        return [4 /*yield*/, Promise.all(files.map(function (file) {
                                return _this.createFileObject(file);
                            }))];
                    case 2: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    BaseFileService.prototype.get_Internal = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                throw new Error('Not Implemented');
            });
        });
    };
    BaseFileService.prototype.getItemById_Internal = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, file;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        return [4 /*yield*/, sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get()];
                    case 1:
                        file = _a.sent();
                        if (file) {
                            result = this.createFileObject(file);
                        }
                        return [2 /*return*/, result];
                }
            });
        });
    };
    BaseFileService.prototype.getItemsById_Internal = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var results, batch;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = [];
                        batch = sp.createBatch();
                        ids.forEach(function (id) {
                            sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get().then(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                var fo;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.createFileObject(item)];
                                        case 1:
                                            fo = _a.sent();
                                            results.push(fo);
                                            return [2 /*return*/];
                                    }
                                });
                            }); });
                        });
                        return [4 /*yield*/, batch.execute()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, results];
                }
            });
        });
    };
    BaseFileService.prototype.createFileObject = function (file) {
        return __awaiter(this, void 0, void 0, function () {
            var resultFile;
            return __generator(this, function (_a) {
                resultFile = new this.itemType(file);
                if (resultFile instanceof SPFile) {
                    resultFile.mimeType = mime.lookup(resultFile.name) || 'application/octet-stream';
                    //resultFile.content = await sp.web.getFileByServerRelativeUrl(resultFile.serverRelativeUrl).getBuffer();
                }
                return [2 /*return*/, resultFile];
            });
        });
    };
    BaseFileService.prototype.getFilesInFolder = function (folderListRelativeUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var result, folderUrl, folderExists, files;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = new Array();
                        folderUrl = this.listRelativeUrl + folderListRelativeUrl;
                        return [4 /*yield*/, this.folderExists(folderListRelativeUrl)];
                    case 1:
                        folderExists = _a.sent();
                        if (!folderExists) return [3 /*break*/, 5];
                        return [4 /*yield*/, sp.web.getFolderByServerRelativeUrl(folderUrl).files.get()];
                    case 2:
                        files = _a.sent();
                        return [4 /*yield*/, Promise.all(files.map(function (file) {
                                return _this.createFileObject(file);
                            }))];
                    case 3: return [4 /*yield*/, _a.sent()];
                    case 4:
                        result = _a.sent();
                        _a.label = 5;
                    case 5: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseFileService.prototype.folderExists = function (folderUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var result, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = false;
                        if (folderUrl.indexOf(this.listRelativeUrl) === -1) {
                            folderUrl = this.listRelativeUrl + folderUrl;
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, sp.web.getFolderByServerRelativeUrl(folderUrl).get()];
                    case 2:
                        _a.sent();
                        result = true;
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseFileService.prototype.addOrUpdateItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var folderUrl, folder, exists;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(item instanceof SPFile && item.content)) return [3 /*break*/, 7];
                        folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
                        folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
                        return [4 /*yield*/, this.folderExists(folderUrl)];
                    case 1:
                        exists = _a.sent();
                        if (!!exists) return [3 /*break*/, 3];
                        return [4 /*yield*/, sp.web.folders.add(folderUrl)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        if (!(item.content.byteLength <= 10485760)) return [3 /*break*/, 5];
                        // small upload
                        return [4 /*yield*/, folder.files.add(item.name, item.content, true)];
                    case 4:
                        // small upload
                        _a.sent();
                        return [3 /*break*/, 7];
                    case 5: 
                    // large upload
                    return [4 /*yield*/, folder.files.addChunked(item.name, UtilsService.arrayBufferToBlob(item.content, item.mimeType), function (data) {
                            console.log("block:" + data.blockNumber + "/" + data.totalBlocks);
                        }, true)];
                    case 6:
                        // large upload
                        _a.sent();
                        _a.label = 7;
                    case 7: return [2 /*return*/, item];
                }
            });
        });
    };
    BaseFileService.prototype.deleteItem_Internal = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var folderUrl, folder, files;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(item instanceof SPFile)) return [3 /*break*/, 4];
                        return [4 /*yield*/, sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl).recycle()];
                    case 1:
                        _a.sent();
                        folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
                        folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
                        return [4 /*yield*/, folder.files.get()];
                    case 2:
                        files = _a.sent();
                        if (!(!files || files.length === 0)) return [3 /*break*/, 4];
                        return [4 /*yield*/, folder.recycle()];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    BaseFileService.prototype.changeFolderInDb = function (oldFolderListRelativeUrl, newFolderListRelativeUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var oldFolderRelativeUrl, newFolderRelativeUrl, allFiles, files, newFiles;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        oldFolderRelativeUrl = this.listRelativeUrl + oldFolderListRelativeUrl;
                        newFolderRelativeUrl = this.listRelativeUrl + newFolderListRelativeUrl;
                        return [4 /*yield*/, this.dbService.getAll()];
                    case 1:
                        allFiles = _a.sent();
                        files = allFiles.filter(function (f) {
                            return UtilsService.getParentFolderUrl(f.id.toString()).toLowerCase() === oldFolderRelativeUrl.toLowerCase();
                        });
                        newFiles = cloneDeep(files);
                        return [4 /*yield*/, Promise.all(files.map(function (f) {
                                return _this.dbService.deleteItem(f);
                            }))];
                    case 2:
                        _a.sent();
                        newFiles.forEach(function (file) {
                            file.id = newFolderRelativeUrl + "/" + file.title;
                        });
                        return [4 /*yield*/, this.dbService.addOrUpdateItems(newFiles)];
                    case 3:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    return BaseFileService;
}(BaseDataService));
export { BaseFileService };
//# sourceMappingURL=BaseFileService.js.map