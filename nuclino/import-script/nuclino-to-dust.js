"use strict";
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
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
var axios_1 = require("axios");
var dotenv = require("dotenv");
var bottleneck_1 = require("bottleneck");
var mime_1 = require("mime");
dotenv.config();
// Configuration
var _a = process.env, NUCLINO_API_KEY = _a.NUCLINO_API_KEY, DUST_API_KEY = _a.DUST_API_KEY, DUST_WORKSPACE_ID = _a.DUST_WORKSPACE_ID;
var missingEnvVars = [
    ['NUCLINO_API_KEY', NUCLINO_API_KEY],
    ['DUST_API_KEY', DUST_API_KEY],
    ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
].filter(function (_a) {
    var name = _a[0], value = _a[1];
    return !value;
}).map(function (_a) {
    var name = _a[0];
    return name;
});
if (missingEnvVars.length > 0) {
    throw new Error("Please provide values for the following environment variables in the .env file: ".concat(missingEnvVars.join(', ')));
}
var NuclinoWorkspace = /** @class */ (function () {
    function NuclinoWorkspace(id, teamId, name, createdAt, createdUserId) {
        this.object = 'workspace';
        this.id = id;
        this.teamId = teamId;
        this.name = name;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
    }
    NuclinoWorkspace.fromRawData = function (data) {
        var instance = new NuclinoWorkspace(data.id, data.teamId, data.name, data.createdAt, data.createdUserId);
        instance.fields = data.fields.map(function (field) { return ({
            object: field.object,
            id: field.id,
            type: field.type,
            name: field.name,
        }); });
        instance.childIds = __spreadArray([], data.childIds, true);
        return instance;
    };
    return NuclinoWorkspace;
}());
var NuclinoUser = /** @class */ (function () {
    function NuclinoUser(id, firstName, lastName, email) {
        this.object = "user";
        this.id = id;
        this.firstName = firstName;
        this.lastName = lastName;
        this.email = email;
    }
    NuclinoUser.fromRawData = function (data) {
        return new NuclinoUser(data.id, data.firstName = 'deactivated', data.lastName = '', data.email);
    };
    return NuclinoUser;
}());
var NuclinoCollection = /** @class */ (function () {
    function NuclinoCollection(id, workspaceId, title, createdAt, createdUserId, lastUpdatedAt, lastUpdatedUserId) {
        this.object = 'collection';
        this.id = id;
        this.workspaceId = workspaceId;
        this.title = title;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.lastUpdatedAt = new Date(lastUpdatedAt);
        this.lastUpdatedUserId = lastUpdatedUserId;
    }
    NuclinoCollection.fromRawData = function (data) {
        var instance = new NuclinoCollection(data.id, data.workspaceId, data.title, data.createdAt, data.createdUserId, data.lastUpdatedAt, data.lastUpdatedUserId);
        instance.childIds = __spreadArray([], data.childIds, true);
        return instance;
    };
    return NuclinoCollection;
}());
var NuclinoArticle = /** @class */ (function () {
    function NuclinoArticle(id, workspaceId, title, url, createdAt, createdUserId, lastUpdatedAt, lastUpdatedUserId, content) {
        this.object = 'item';
        this.fields = [];
        this.meta = { itemIds: [], fileIds: [] };
        this.id = id;
        this.workspaceId = workspaceId;
        this.title = title;
        this.url = url;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.lastUpdatedAt = new Date(lastUpdatedAt);
        this.lastUpdatedUserId = lastUpdatedUserId;
        this.content = content;
    }
    NuclinoArticle.fromRawData = function (data) {
        var instance = new NuclinoArticle(data.id, data.workspaceId, data.title, data.url, data.createdAt, data.createdUserId, data.lastUpdatedAt, data.lastUpdatedUserId, data.content);
        /*instance.fields = data.fields.map((field: any) => ({
            object: field.object,
            id: field.id,
            type: field.type,
            name: field.name,
        }));
        */
        instance.meta = data.contentMeta;
        return instance;
    };
    return NuclinoArticle;
}());
var NuclinoFile = /** @class */ (function () {
    function NuclinoFile(id, itemId, fileName, createdAt, createdUserId, download) {
        this.id = id;
        this.itemId = itemId;
        this.fileName = fileName;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.download = {
            url: download.url,
            expiresAt: new Date(download.expiresAt)
        };
    }
    NuclinoFile.fromRawData = function (data) {
        return new NuclinoFile(data.id, data.itemId, data.fileName, data.createdAt, data.createdUserId, data.download);
    };
    NuclinoFile.prototype.getRawData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var body;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, axios_1.default.get(this.download.url)];
                    case 1:
                        body = _a.sent();
                        return [2 /*return*/, body.data.content];
                }
            });
        });
    };
    return NuclinoFile;
}());
var Nuclino = /** @class */ (function () {
    function Nuclino(rateLimiter, apiConnection) {
        this.rateLimiter = rateLimiter;
        this.apiConnection = apiConnection;
        this.users = new Map();
        this.rateLimiter = rateLimiter;
        this.apiConnection = apiConnection;
    }
    Nuclino.prototype.getUserInfo = function (userId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, user, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        if (this.users.has(userId)) {
                            return [2 /*return*/, this.users.get(userId)];
                        }
                        return [4 /*yield*/, this.doThrottledGetRequest("/users/".concat(userId))];
                    case 1:
                        response = _a.sent();
                        user = NuclinoUser.fromRawData(response.data.data);
                        this.users.set(userId, user);
                        return [2 /*return*/, user];
                    case 2:
                        error_1 = _a.sent();
                        console.error("Error fetching  user ".concat(userId, ":"), error_1.message);
                        throw error_1;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    Nuclino.prototype.doThrottledGetRequest = function (url, data) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2 /*return*/, this.rateLimiter.schedule(function () { return _this.apiConnection.get(url, data); })];
            });
        });
    };
    Nuclino.prototype.getWorkspace = function (name) {
        return __awaiter(this, void 0, void 0, function () {
            var response, workspaceRawata, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.doThrottledGetRequest("/workspaces")];
                    case 1:
                        response = _a.sent();
                        workspaceRawata = response.data.data.results.find(function (workspace) { return workspace.name === name; });
                        return [2 /*return*/, NuclinoWorkspace.fromRawData(workspaceRawata)];
                    case 2:
                        error_2 = _a.sent();
                        console.error("Error fetching  workspace ".concat(name, ":"), error_2.message);
                        throw error_2;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    Nuclino.prototype.getItem = function (itemId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.doThrottledGetRequest("/items/".concat(itemId))];
                    case 1:
                        response = _a.sent();
                        switch (response.data.data.object) {
                            case 'collection': {
                                return [2 /*return*/, NuclinoCollection.fromRawData(response.data.data)];
                            }
                            case 'item': {
                                return [2 /*return*/, NuclinoArticle.fromRawData(response.data.data)];
                            }
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        console.error("Error fetching content for item ".concat(itemId, ":"), error_3.message);
                        throw error_3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    Nuclino.prototype.getFile = function (fileId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.doThrottledGetRequest("/files/".concat(fileId))];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, NuclinoFile.fromRawData(response.data.data)];
                    case 2:
                        error_4 = _a.sent();
                        console.error("Error fetching file ".concat(fileId, ":"), error_4.message);
                        throw error_4;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    Nuclino.prototype.getAllItems = function (workspaceId, after) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.doThrottledGetRequest("/items", {
                                params: { workspaceId: workspaceId, limit: 100, after: after },
                            })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response];
                    case 2:
                        error_5 = _a.sent();
                        console.error('Error fetching items from Nuclino:', error_5);
                        throw error_5;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return Nuclino;
}());
/**
 * And this is for all the dust.
 */
var Dust = /** @class */ (function () {
    function Dust(rateLimiter, apiConnection) {
        this.rateLimiter = rateLimiter;
        this.apiConnection = apiConnection;
        this.rateLimiter = rateLimiter;
        this.apiConnection = apiConnection;
    }
    Dust.prototype.upsertArticleToDustDatasource = function (article, author, lastUpdater, breadCrumb, destination) {
        return __awaiter(this, void 0, void 0, function () {
            var documentId, fullPath, content, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        documentId = "article-".concat(article.id);
                        fullPath = "".concat(breadCrumb);
                        content = "\nPath: ".concat(fullPath, "\nUrl: ").concat(article.url, "\nTitle: ").concat(article.title, "\nAuthor: ").concat(author.firstName, " ").concat(author.lastName, "\nLast updated by: ").concat(lastUpdater.firstName, " ").concat(lastUpdater.lastName, "\nCreated At: ").concat(article.createdAt, "\nUpdated At: ").concat(article.lastUpdatedAt, "\nContent:\n").concat(article.content, "\n  ").trim();
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.doThrottledPostRequest("/vaults/".concat(destination.spaceId, "/data_sources/").concat(destination.sourceId, "/documents/").concat(documentId), {
                                title: article.title,
                                mime_type: 'text/markdown',
                                text: content,
                                source_url: article.url,
                            })];
                    case 2:
                        _a.sent();
                        console.log("Upserted article ".concat(article.id, " to Dust datasource"));
                        return [3 /*break*/, 4];
                    case 3:
                        error_6 = _a.sent();
                        console.error("Error upserting article ".concat(article.id, " to Dust datasource:"), error_6);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    Dust.prototype.upsertFileToDustDatasource = function (file, author, breadCrumb, destination) {
        return __awaiter(this, void 0, void 0, function () {
            var error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.doThrottledPostRequest("/vaults/".concat(destination.spaceId, "/data_sources/").concat(destination.sourceId, "/documents/").concat(file.id, "-").concat(file.fileName), {
                                title: file.fileName,
                                text: file.getRawData(),
                                mime_type: mime_1.default.getType(file.fileName),
                            })];
                    case 1:
                        _a.sent();
                        console.log("Upserted article ".concat(file.id, " to Dust datasource"));
                        return [3 /*break*/, 3];
                    case 2:
                        error_7 = _a.sent();
                        console.error("Error upserting article ".concat(file.id, " to Dust datasource:"), error_7);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    Dust.prototype.doThrottledPostRequest = function (url, data, config) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.rateLimiter.schedule(function () { return _this.apiConnection.post(url, data, config); })];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    return Dust;
}());
/**
 * Bootstrap
 */
// Create a Bottleneck limiter for Dust API
var dustLimiter = new bottleneck_1.default({
    minTime: 500, // 500ms between requests. limit of 120 upserts / minute per workspace https://docs.dust.tt/reference/rate-limits
    maxConcurrent: 1, // Only 1 request at a time
});
var nuclinoLimiter = new bottleneck_1.default({
    minTime: 500, // 400 between requests. per https://help.nuclino.com/b147124e-rate-limiting
    maxConcurrent: 1, // Only 1 request at a time
});
var nuclinoApi = axios_1.default.create({
    baseURL: 'https://api.nuclino.com/v0',
    headers: {
        Authorization: "".concat(NUCLINO_API_KEY),
        Accept: 'application/json',
    },
});
var dustApi = axios_1.default.create({
    baseURL: "https://dust.tt/api/v1/w/".concat(DUST_WORKSPACE_ID, "/"),
    headers: {
        Authorization: "Bearer ".concat(DUST_API_KEY),
        'Content-Type': 'application/json',
    },
});
var nuclino = new Nuclino(nuclinoLimiter, nuclinoApi);
var dust = new Dust(dustLimiter, dustApi);
/**
 * Time to run some code
 */
function recursiveSync(nuclinoItem, breadCrumb, dust, nuclino, destination) {
    return __awaiter(this, void 0, void 0, function () {
        var _a, author, lastUpdater, _i, _b, item, itemData;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    _a = nuclinoItem.object;
                    switch (_a) {
                        case 'item': return [3 /*break*/, 1];
                        case 'workspace': return [3 /*break*/, 6];
                        case 'collection': return [3 /*break*/, 6];
                    }
                    return [3 /*break*/, 10];
                case 1:
                    if (!(nuclinoItem instanceof NuclinoArticle)) return [3 /*break*/, 5];
                    return [4 /*yield*/, nuclino.getUserInfo(nuclinoItem.createdUserId)];
                case 2:
                    author = _c.sent();
                    return [4 /*yield*/, nuclino.getUserInfo(nuclinoItem.lastUpdatedUserId)];
                case 3:
                    lastUpdater = _c.sent();
                    return [4 /*yield*/, dust.upsertArticleToDustDatasource(nuclinoItem, author, lastUpdater, breadCrumb, destination)];
                case 4:
                    _c.sent();
                    _c.label = 5;
                case 5: return [2 /*return*/];
                case 6:
                    _i = 0, _b = nuclinoItem.childIds;
                    _c.label = 7;
                case 7:
                    if (!(_i < _b.length)) return [3 /*break*/, 10];
                    item = _b[_i];
                    return [4 /*yield*/, nuclino.getItem(item)];
                case 8:
                    itemData = _c.sent();
                    recursiveSync(itemData, "".concat(breadCrumb, "/").concat(itemData.title), dust, nuclino, destination);
                    _c.label = 9;
                case 9:
                    _i++;
                    return [3 /*break*/, 7];
                case 10: return [2 /*return*/];
            }
        });
    });
}
function migrateContent(workspaceName, destination) {
    return __awaiter(this, void 0, void 0, function () {
        var workspace, error_8;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 3, , 4]);
                    console.log('Starting migration...' + workspaceName);
                    return [4 /*yield*/, nuclino.getWorkspace(workspaceName)];
                case 1:
                    workspace = _a.sent();
                    return [4 /*yield*/, recursiveSync(workspace, workspaceName, dust, nuclino, destination)];
                case 2:
                    _a.sent();
                    console.log('Migration completed successfully!');
                    return [3 /*break*/, 4];
                case 3:
                    error_8 = _a.sent();
                    console.error('Migration failed:', error_8.message);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    });
}
// Run the migration
if (process.argv.length < 5) {
    console.error('Usage: npm run import <nuclinoWorkspaceName> <DustSpaceId> <DustDataSourceId>');
    exit(1);
}
console.log(process.argv);
migrateContent(process.argv[2], { spaceId: process.argv[3], sourceId: process.argv[4] });
