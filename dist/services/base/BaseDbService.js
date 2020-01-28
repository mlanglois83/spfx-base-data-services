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
import { Text } from "@microsoft/sp-core-library";
import { assign } from "@microsoft/sp-lodash-subset";
import { openDb } from "idb";
import { SPFile } from "../../models/index";
import { UtilsService } from "../index";
import { BaseService } from "./BaseService";
import { Constants } from "../../constants";
import { ServicesConfiguration } from "../..";
/**
 * Base classe for indexedDB interraction using SP repository
 */
var BaseDbService = /** @class */ (function (_super) {
    __extends(BaseDbService, _super);
    /**
     *
     * @param tableName : Name of the db table the service interracts with
     */
    function BaseDbService(type, tableName) {
        var _this = _super.call(this) || this;
        _this.tableName = tableName;
        _this.db = null;
        _this.itemType = type;
        return _this;
    }
    BaseDbService.prototype.getChunksRegexp = function (fileUrl) {
        var escapedUrl = UtilsService.escapeRegExp(fileUrl);
        return new RegExp("^" + escapedUrl + "_chunk_\\d+$", "g");
    };
    BaseDbService.prototype.getAllKeysInternal = function (store) {
        return __awaiter(this, void 0, void 0, function () {
            var result, cursor;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = [];
                        if (!store.getAllKeys) return [3 /*break*/, 2];
                        return [4 /*yield*/, store.getAllKeys()];
                    case 1:
                        result = _a.sent();
                        return [3 /*break*/, 6];
                    case 2: return [4 /*yield*/, store.openCursor()];
                    case 3:
                        cursor = _a.sent();
                        _a.label = 4;
                    case 4:
                        if (!cursor) return [3 /*break*/, 6];
                        result.push(cursor.primaryKey);
                        return [4 /*yield*/, cursor.continue()];
                    case 5:
                        cursor = _a.sent();
                        return [3 /*break*/, 4];
                    case 6: return [2 /*return*/, result];
                }
            });
        });
    };
    BaseDbService.prototype.getNextAvailableKey = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result, tx, store, keys, minKey;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        tx = this.db.transaction(this.tableName, 'readonly');
                        store = tx.objectStore(this.tableName);
                        return [4 /*yield*/, this.getAllKeysInternal(store)];
                    case 2:
                        keys = _a.sent();
                        if (keys.length > 0) {
                            minKey = Math.min.apply(Math, keys);
                            result = Math.min(-2, minKey - 1);
                        }
                        else {
                            result = -2;
                        }
                        return [4 /*yield*/, tx.complete];
                    case 3:
                        _a.sent();
                        return [2 /*return*/, result];
                }
            });
        });
    };
    /**
     * Opens indexed db, update structure if needed
     */
    BaseDbService.prototype.OpenDb = function () {
        return __awaiter(this, void 0, void 0, function () {
            var dbName, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(this.db == null)) return [3 /*break*/, 2];
                        if (!('indexedDB' in window)) {
                            throw new Error(ServicesConfiguration.configuration.translations.IndexedDBNotDefined);
                        }
                        dbName = Text.format(ServicesConfiguration.configuration.dbName, ServicesConfiguration.context.pageContext.web.serverRelativeUrl);
                        _a = this;
                        return [4 /*yield*/, openDb(dbName, ServicesConfiguration.configuration.dbVersion, function (UpgradeDB) {
                                var tableNames = Constants.tableNames.concat(ServicesConfiguration.configuration.tableNames);
                                // add new tables
                                for (var _i = 0, tableNames_1 = tableNames; _i < tableNames_1.length; _i++) {
                                    var tableName = tableNames_1[_i];
                                    if (!UpgradeDB.objectStoreNames.contains(tableName)) {
                                        UpgradeDB.createObjectStore(tableName, { keyPath: 'id', autoIncrement: tableName == "Transaction" });
                                    }
                                }
                                // TODO : remove old tables
                            })];
                    case 1:
                        _a.db = _b.sent();
                        _b.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Add or update an item in DB and returns updated item
     * @param item Item to add or update
     */
    BaseDbService.prototype.addOrUpdateItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var nextid, tx, store, keys, chunkRegex_1, chunkkeys, idx, size, firstidx, lastidx, chunk, chunkitem, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.getNextAvailableKey()];
                    case 2:
                        nextid = _a.sent();
                        tx = this.db.transaction(this.tableName, 'readwrite');
                        store = tx.objectStore(this.tableName);
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 13, , 14]);
                        if (typeof (item.id) === "number" && !store.autoIncrement && item.id === -1) {
                            item.id = nextid;
                        }
                        if (!(item instanceof SPFile && item.content && item.content.byteLength >= 10485760)) return [3 /*break*/, 9];
                        return [4 /*yield*/, this.getAllKeysInternal(store)];
                    case 4:
                        keys = _a.sent();
                        chunkRegex_1 = this.getChunksRegexp(item.serverRelativeUrl);
                        chunkkeys = keys.filter(function (k) {
                            var match = k.match(chunkRegex_1);
                            return match && match.length > 0;
                        });
                        return [4 /*yield*/, Promise.all(chunkkeys.map(function (k) {
                                return store.delete(k);
                            }))];
                    case 5:
                        _a.sent();
                        idx = 0;
                        size = 0;
                        _a.label = 6;
                    case 6:
                        if (!(size < item.content.byteLength)) return [3 /*break*/, 8];
                        firstidx = idx * 10485760;
                        lastidx = Math.min(item.content.byteLength, firstidx + 10485760);
                        chunk = item.content.slice(firstidx, lastidx);
                        chunkitem = new SPFile();
                        chunkitem.serverRelativeUrl = item.serverRelativeUrl + (idx === 0 ? "" : "_chunk_" + idx);
                        chunkitem.name = item.name;
                        chunkitem.mimeType = item.mimeType;
                        chunkitem.content = chunk;
                        return [4 /*yield*/, store.put(assign({}, chunkitem))];
                    case 7:
                        _a.sent();
                        idx++;
                        size += chunk.byteLength;
                        return [3 /*break*/, 6];
                    case 8: return [3 /*break*/, 11];
                    case 9: return [4 /*yield*/, store.put(assign({}, item))];
                    case 10:
                        _a.sent(); // store simple object with data only 
                        _a.label = 11;
                    case 11: return [4 /*yield*/, tx.complete];
                    case 12:
                        _a.sent();
                        return [2 /*return*/, {
                                item: item
                            }];
                    case 13:
                        error_1 = _a.sent();
                        console.error(error_1.message + " - " + error_1.Name);
                        try {
                            tx.abort();
                        }
                        catch (_b) {
                            // error allready thrown
                        }
                        return [2 /*return*/, {
                                item: item,
                                error: error_1
                            }];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    BaseDbService.prototype.deleteItem = function (item) {
        return __awaiter(this, void 0, void 0, function () {
            var tx, store, deleteKeys, keys, chunkRegex_2, chunkkeys, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        tx = this.db.transaction(this.tableName, 'readwrite');
                        store = tx.objectStore(this.tableName);
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 7, , 8]);
                        deleteKeys = [item.id];
                        if (!(item instanceof SPFile)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getAllKeysInternal(store)];
                    case 3:
                        keys = _a.sent();
                        chunkRegex_2 = this.getChunksRegexp(item.serverRelativeUrl);
                        chunkkeys = keys.filter(function (k) {
                            var match = k.match(chunkRegex_2);
                            return match && match.length > 0;
                        });
                        deleteKeys.push.apply(deleteKeys, chunkkeys);
                        _a.label = 4;
                    case 4: return [4 /*yield*/, Promise.all(deleteKeys.map(function (k) {
                            return store.delete(k);
                        }))];
                    case 5:
                        _a.sent();
                        return [4 /*yield*/, tx.complete];
                    case 6:
                        _a.sent();
                        return [3 /*break*/, 8];
                    case 7:
                        error_2 = _a.sent();
                        console.error(error_2.message + " - " + error_2.Name);
                        try {
                            tx.abort();
                        }
                        catch (_b) {
                            // error allready thrown
                        }
                        throw error_2;
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    BaseDbService.prototype.get = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var results, hash, items, _i, items_1, item;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        results = new Array();
                        hash = _super.prototype.hashCode.call(this, query);
                        return [4 /*yield*/, this.getAll()];
                    case 1:
                        items = _a.sent();
                        for (_i = 0, items_1 = items; _i < items_1.length; _i++) {
                            item = items_1[_i];
                            if (item.queries && item.queries.indexOf(hash) >= 0) {
                                results.push(item);
                            }
                        }
                        return [2 /*return*/, results];
                }
            });
        });
    };
    /**
     * add items in table (ids updated)
     * @param newItems
     */
    BaseDbService.prototype.addOrUpdateItems = function (newItems, query) {
        return __awaiter(this, void 0, void 0, function () {
            var nextid, tx, store, error_3;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.getNextAvailableKey()];
                    case 2:
                        nextid = _a.sent();
                        tx = this.db.transaction(this.tableName, 'readwrite');
                        store = tx.objectStore(this.tableName);
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 6, , 7]);
                        return [4 /*yield*/, Promise.all(newItems.map(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                var keys, chunkRegex_3, chunkkeys, idx, size, firstidx, lastidx, chunk, chunkitem, hash, temp;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            if (typeof (item.id) === "number" && !store.autoIncrement && item.id === -1) {
                                                item.id = nextid--;
                                            }
                                            if (!(item instanceof SPFile && item.content && item.content.byteLength >= 10485760)) return [3 /*break*/, 6];
                                            return [4 /*yield*/, this.getAllKeysInternal(store)];
                                        case 1:
                                            keys = _a.sent();
                                            chunkRegex_3 = this.getChunksRegexp(item.serverRelativeUrl);
                                            chunkkeys = keys.filter(function (k) {
                                                var match = k.match(chunkRegex_3);
                                                return match && match.length > 0;
                                            });
                                            return [4 /*yield*/, Promise.all(chunkkeys.map(function (k) {
                                                    return store.delete(k);
                                                }))];
                                        case 2:
                                            _a.sent();
                                            idx = 0;
                                            size = 0;
                                            _a.label = 3;
                                        case 3:
                                            if (!(size < item.content.byteLength)) return [3 /*break*/, 5];
                                            firstidx = idx * 10485760;
                                            lastidx = Math.min(item.content.byteLength, firstidx + 10485760);
                                            chunk = item.content.slice(firstidx, lastidx);
                                            chunkitem = new SPFile();
                                            chunkitem.serverRelativeUrl = item.serverRelativeUrl + (idx === 0 ? "" : "_chunk_" + idx);
                                            chunkitem.name = item.name;
                                            chunkitem.mimeType = item.mimeType;
                                            chunkitem.content = chunk;
                                            return [4 /*yield*/, store.put(assign({}, chunkitem))];
                                        case 4:
                                            _a.sent();
                                            idx++;
                                            size += chunk.byteLength;
                                            return [3 /*break*/, 3];
                                        case 5: return [3 /*break*/, 10];
                                        case 6:
                                            if (!query) return [3 /*break*/, 8];
                                            item.queries = new Array();
                                            hash = this.hashCode(query);
                                            return [4 /*yield*/, store.get(item.id)];
                                        case 7:
                                            temp = _a.sent();
                                            //if exist    
                                            if (temp) {
                                                //if item never store from query, init array
                                                if (!temp.queries) {
                                                    temp.queries = new Array();
                                                }
                                                //if query never launched
                                                //add query to item db
                                                if (temp.queries.indexOf(hash) < 0) {
                                                    temp.queries.push(hash);
                                                }
                                                item.queries = temp.queries;
                                            }
                                            else {
                                                item.queries.push(hash);
                                            }
                                            _a.label = 8;
                                        case 8: return [4 /*yield*/, store.put(assign({}, item))];
                                        case 9:
                                            _a.sent(); // store simple object with data only 
                                            _a.label = 10;
                                        case 10: return [2 /*return*/];
                                    }
                                });
                            }); }))];
                    case 4:
                        _a.sent();
                        return [4 /*yield*/, tx.complete];
                    case 5:
                        _a.sent();
                        return [2 /*return*/, newItems];
                    case 6:
                        error_3 = _a.sent();
                        console.error(error_3.message + " - " + error_3.Name);
                        try {
                            tx.abort();
                        }
                        catch (_b) {
                            // error allready thrown
                        }
                        throw error_3;
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Retrieve all items from db table
     */
    BaseDbService.prototype.getAll = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result, transaction, store, rows_1, error_4;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = new Array();
                        return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        transaction = this.db.transaction(this.tableName, 'readonly');
                        store = transaction.objectStore(this.tableName);
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 5, , 6]);
                        return [4 /*yield*/, store.getAll()];
                    case 3:
                        rows_1 = _a.sent();
                        rows_1.forEach(function (r) {
                            var item = new _this.itemType();
                            var resultItem = assign(item, r);
                            if (item instanceof SPFile) {
                                // item is a part of another file
                                var chunkparts = (/^.*_chunk_\d+$/g).test(item.serverRelativeUrl);
                                if (!chunkparts) {
                                    // verify if there are other parts
                                    var chunkRegex_4 = _this.getChunksRegexp(item.serverRelativeUrl);
                                    var chunks = rows_1.filter(function (chunkedrow) {
                                        var match = chunkedrow.id.match(chunkRegex_4);
                                        return match && match.length > 0;
                                    });
                                    if (chunks.length > 0) {
                                        chunks.sort(function (a, b) {
                                            return parseInt(a.id.replace(/^.*_chunk_(\d+)$/g, "$1")) - parseInt(b.id.replace(/^.*_chunk_(\d+)$/g, "$1"));
                                        });
                                        resultItem.content = UtilsService.concatArrayBuffers.apply(UtilsService, __spreadArrays([resultItem.content], chunks.map(function (c) { return c.content; })));
                                    }
                                    result.push(resultItem);
                                }
                            }
                            else {
                                result.push(resultItem);
                            }
                        });
                        return [4 /*yield*/, transaction.complete];
                    case 4:
                        _a.sent();
                        return [2 /*return*/, result];
                    case 5:
                        error_4 = _a.sent();
                        console.error(error_4.message + " - " + error_4.Name);
                        try {
                            transaction.abort();
                        }
                        catch (_b) {
                            // error allready thrown
                        }
                        throw error_4;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Clear table and insert new items
     * @param newItems Items to insert in place of existing
     */
    BaseDbService.prototype.replaceAll = function (newItems) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clear()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.addOrUpdateItems(newItems)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Clear table
     */
    BaseDbService.prototype.clear = function () {
        return __awaiter(this, void 0, void 0, function () {
            var tx, store, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        tx = this.db.transaction(this.tableName, 'readwrite');
                        store = tx.objectStore(this.tableName);
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 5, , 6]);
                        return [4 /*yield*/, store.clear()];
                    case 3:
                        _a.sent();
                        return [4 /*yield*/, tx.complete];
                    case 4:
                        _a.sent();
                        return [3 /*break*/, 6];
                    case 5:
                        error_5 = _a.sent();
                        console.error(error_5.message + " - " + error_5.Name);
                        try {
                            tx.abort();
                        }
                        catch (_b) {
                            // error allready thrown
                        }
                        throw error_5;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    BaseDbService.prototype.getItemById = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var result, tx, store, obj, chunkparts, allRows, chunkRegex_5, chunks, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = null;
                        return [4 /*yield*/, this.OpenDb()];
                    case 1:
                        _a.sent();
                        tx = this.db.transaction(this.tableName, 'readonly');
                        store = tx.objectStore(this.tableName);
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 8, , 9]);
                        return [4 /*yield*/, store.get(id)];
                    case 3:
                        obj = _a.sent();
                        if (!obj) return [3 /*break*/, 6];
                        result = assign(new this.itemType(), obj);
                        if (!(result instanceof SPFile)) return [3 /*break*/, 6];
                        chunkparts = (/^.*_chunk_\d+$/g).test(result.serverRelativeUrl);
                        if (!!chunkparts) return [3 /*break*/, 5];
                        return [4 /*yield*/, store.getAll()];
                    case 4:
                        allRows = _a.sent();
                        chunkRegex_5 = this.getChunksRegexp(result.serverRelativeUrl);
                        chunks = allRows.filter(function (chunkedrow) {
                            var match = chunkedrow.id.match(chunkRegex_5);
                            return match && match.length > 0;
                        });
                        if (chunks.length > 0) {
                            chunks.sort(function (a, b) {
                                return parseInt(a.id.replace(/^.*_chunk_(\d+)$/g, "$1")) - parseInt(b.id.replace(/^.*_chunk_(\d+)$/g, "$1"));
                            });
                            result.content = UtilsService.concatArrayBuffers.apply(UtilsService, __spreadArrays([result.content], chunks.map(function (c) { return c.content; })));
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        // no chunked parts here
                        result = null;
                        _a.label = 6;
                    case 6: return [4 /*yield*/, tx.complete];
                    case 7:
                        _a.sent();
                        return [2 /*return*/, result];
                    case 8:
                        error_6 = _a.sent();
                        // key not found
                        return [2 /*return*/, null];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    BaseDbService.prototype.getItemsById = function (ids) {
        return __awaiter(this, void 0, void 0, function () {
            var results;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Promise.all(ids.map(function (id) {
                            return _this.getItemById(id);
                        }))];
                    case 1:
                        results = _a.sent();
                        return [2 /*return*/, results];
                }
            });
        });
    };
    return BaseDbService;
}(BaseService));
export { BaseDbService };
