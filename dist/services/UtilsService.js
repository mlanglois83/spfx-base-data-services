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
import { BaseService } from "./base/BaseService";
import { find } from "@microsoft/sp-lodash-subset";
var UtilsService = /** @class */ (function (_super) {
    __extends(UtilsService, _super);
    function UtilsService() {
        return _super.call(this) || this;
    }
    /**
     * check is user has connexion
     */
    UtilsService.CheckOnline = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result, response, ex_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = false;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, fetch(BaseService.Configuration.context.pageContext.web.absoluteUrl, { method: 'HEAD', mode: 'no-cors' })];
                    case 2:
                        response = _a.sent();
                        result = (response && (response.ok || response.type === 'opaque'));
                        return [3 /*break*/, 4];
                    case 3:
                        ex_1 = _a.sent();
                        result = false;
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, result];
                }
            });
        });
    };
    UtilsService.blobToArrayBuffer = function (blob) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.addEventListener('loadend', function (e) {
                resolve(reader.result);
            });
            reader.addEventListener('error', reject);
            reader.readAsArrayBuffer(blob);
        });
    };
    UtilsService.arrayBufferToBlob = function (buffer, type) {
        return new Blob([buffer], { type: type });
    };
    UtilsService.getOfflineFileUrl = function (fileData) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader;
            reader.onerror = reject;
            reader.onload = function () {
                var val = reader.result.toString();
                resolve(val);
            };
            reader.readAsDataURL(fileData);
        });
    };
    UtilsService.getParentFolderUrl = function (url) {
        var urlParts = url.split('/');
        urlParts.pop();
        return urlParts.join("/");
    };
    UtilsService.concatArrayBuffers = function () {
        var arrays = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            arrays[_i] = arguments[_i];
        }
        var length = 0;
        var buffer = null;
        arrays.forEach(function (a) {
            length += a.byteLength;
        });
        var joined = new Uint8Array(length);
        var offset = 0;
        arrays.forEach(function (a) {
            joined.set(new Uint8Array(a), offset);
            offset += a.byteLength;
        });
        return joined.buffer;
    };
    UtilsService.getTaxonomyTermByWssId = function (wssid, terms) {
        return find(terms, function (term) {
            return (term.wssids && term.wssids.indexOf(wssid) > -1);
        });
    };
    return UtilsService;
}(BaseService));
export { UtilsService };
//# sourceMappingURL=UtilsService.js.map