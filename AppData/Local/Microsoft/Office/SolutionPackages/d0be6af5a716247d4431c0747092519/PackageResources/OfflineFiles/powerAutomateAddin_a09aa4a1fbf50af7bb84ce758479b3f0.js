var TypeScriptModule =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/powerAutomateAddin.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/powerAutomateAddin.ts":
/*!***********************************!*\
  !*** ./src/powerAutomateAddin.ts ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/// <reference types="@types/office-js" />
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
Object.defineProperty(exports, "__esModule", { value: true });
// Needs to be changed if login.html is updated.
// The change should be replacing it with the build number pertaining to the build containing the new login.html bits.
// Only make the change once the login.html change has been deployed to at least fast food.
var BUILD_NUMBER = '1.0.2301.17006';
// Reply urls need to match exactly with the urls registered on the first-party registration page.
// The build url changes with every build, thus the reply url we automatically send would change as well.
// Urls registered on the first-party registration page are registered manually.
// This is a workaround so that we do not have to manually register a new url with every new build.
// Since pages on the CDN live indefinitely, we just send an old page containing an up to date login.html page.
var WORKAROUND_REPLY_URL = window.location.origin.startsWith("https://localhost") ?
    'https://localhost:3000/login.html' :
    "https://fa000000113.resources.office.net/f7024bdc-7caf-4ca8-807d-2908f09640d6/" + BUILD_NUMBER + "/en-us_web/login.html";
var formattedPostfix = '_Formatted';
var graphEndpoint = "https://graph.microsoft.com";
var loadFlowWidget = function (endpoint, exelHostDriveid, excelFileId, autoFillParams, flowToken) {
    var sdk = new MsFlowSdk({
        hostName: endpoint,
        locale: 'en-us',
        hostId: 'ExcelSDX',
        enableWidgetV2: true,
    });
    var widgetRenderParams = {
        container: 'flow-div',
        flowsSettings: {
            allowImplicitConsent: true,
            hideTabs: true,
            isMini: true,
            flowsFilter: "operations/any(operation: operation/excel.fileId eq '" + exelHostDriveid + "/" + excelFileId + "')",
            widgetFlowListDisplaySettings: {
                actionMenuOverFlowItems: true,
                actionMenuClassName: 'fl-ActionMenu-ExcelWidget',
                triggerOperationKey: 'SHARED_EXCELONLINEBUSINESS-ONROWSELECTED',
                triggerOperationName: 'OnRowSelected',
                triggerOperationGroupName: 'shared_excelonlinebusiness',
                hideTemplateTitleDietDesigner: true,
                hideTemplateTypeDietDesigner: true,
            },
        },
        templatesSettings: {
            allowCustomFlowName: true,
            metadataSortProperty: 'ExcelTablesPriority',
            templateCategory: 'microsoftexcel_sdx_nativem2',
            useFlowCreatorSurfaceFromTemplateGallery: true,
            enableDietDesigner: true,
            showHiddenTemplates: true,
            enableTemplatesPageShell: true,
            defaultParams: autoFillParams,
        },
        enableOnBehalfOfTokens: true,
        widgetStyleSettings: {
            themeName: getWidgetTheme(),
        }
    };
    var widgetInstance = sdk.renderWidget('flows', widgetRenderParams);
    widgetInstance.listen('GET_ACCESS_TOKEN', function (_requestParams, widgetDoneCallback) {
        widgetDoneCallback(null, { token: flowToken });
    });
    widgetInstance.listen('WIDGET_READY', function () {
        document.getElementById('spinner').style.display = 'none';
        document.getElementById('spinner-container').style.display = 'none';
        document.getElementById('flow-div').style.display = 'block';
    });
    widgetInstance.listen('GET_IMPLICIT_DATA', function (requestParam, widgetDoneCallback) {
        handleImplicitData(requestParam.data, widgetDoneCallback);
    });
};
var getWidgetTheme = function () {
    var lightTheme = 'excel_sdx';
    var grayTheme = 'excel_sdx_gray';
    var darkTheme = 'excel_sdx_dark';
    try {
        // The Office.OfficeTheme API is not supported on these platforms, so fallback to
        // to light theme
        switch (Office.context.platform) {
            case Office.PlatformType.Mac:
            case Office.PlatformType.OfficeOnline:
                return lightTheme;
            default:
            // Intentionally empty to fallback to below logic.
        }
        var officeTheme = Office.context.officeTheme;
        var bodyBackgroundColor = officeTheme
            ? officeTheme.bodyBackgroundColor.toUpperCase()
            : '';
        switch (bodyBackgroundColor) {
            case '#E6E6E6': //OfficeTheme Colorful:
            case '#FFFFFF': //OfficeTheme White
                return lightTheme;
            case '#666666': //OfficeTheme DarkGray
                return grayTheme;
            case '#262626': //OfficeTheme Black
                return darkTheme;
            // If the office theme API does not exist or we receive an unrecognized color,
            // use the light theme.
            default:
                return lightTheme;
        }
    }
    catch (_a) {
        // In case of any other exception, use light theme
        return lightTheme;
    }
};
function getExcelSource(graphEndpoint, graphHeaders, driveType, siteId) {
    return __awaiter(this, void 0, void 0, function () {
        var resp, respJson;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!((driveType === null || driveType === void 0 ? void 0 : driveType.length) > 0 && (siteId === null || siteId === void 0 ? void 0 : siteId.length) > 0)) return [3 /*break*/, 4];
                    if (!(driveType === "documentLibrary")) return [3 /*break*/, 3];
                    return [4 /*yield*/, fetch(graphEndpoint + "/v1.0/sites/" + siteId, {
                            headers: graphHeaders
                        })];
                case 1:
                    resp = _a.sent();
                    if (!resp.ok) {
                        return [2 /*return*/, null];
                    }
                    return [4 /*yield*/, resp.json()];
                case 2:
                    respJson = _a.sent();
                    if (!(respJson === null || respJson === void 0 ? void 0 : respJson.id)) {
                        return [2 /*return*/, null];
                    }
                    return [2 /*return*/, "sites/" + (respJson === null || respJson === void 0 ? void 0 : respJson.id)];
                case 3:
                    if (driveType === "business") {
                        return [2 /*return*/, "me"];
                    }
                    _a.label = 4;
                case 4: return [2 /*return*/, null];
            }
        });
    });
}
;
var getMakeEndpointForToken = function (token) {
    // // Tenant whitelist
    var tenantsForPreview = ["72f988bf-86f1-41af-91ab-2d7cd011db47", "23dc67c7-c437-4c8c-b17d-a8eb749e2ad7"];
    var tenantsForTest = ["0aa46ef4-f834-4456-bcaf-c823c9cdfffd", "48619c59-dbe9-4218-b8fd-e9b746e93508"];
    // Parse token and get the tid claim
    var parsedToken = JSON.parse(atob(token.split('.')[1]));
    var tenantId = parsedToken.tid;
    // If tid claim matches either preview or test, return appropriate endpoint
    // otherwise default to PROD.
    if (tenantsForPreview.includes(tenantId)) {
        return "https://make.preview.powerautomate.com/";
    }
    else if (tenantsForTest.includes(tenantId)) {
        return "https://make.test.powerautomate.com/";
    }
    else {
        return "https://make.powerautomate.com/";
    }
};
var handleImplicitData = function (data, widgetDoneCallback) {
    var columnKeys = Object.keys(data.implicitData || data);
    var rowStartIndex;
    var dataRowStartIndex;
    var rowCount;
    var tableHeaders = [];
    var rawColumnKeys = [];
    var formattedColumnKeys = [];
    Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        var rangeTables = range.getTables();
        rangeTables.load();
        var table = rangeTables.getFirst();
        table.load();
        range.load('address');
        range.load('rowIndex');
        range.load('rowCount');
        var tableHeaderRowRange = table.getHeaderRowRange();
        tableHeaderRowRange.load('values');
        tableHeaderRowRange.load('rowIndex');
        tableHeaderRowRange.load('columnIndex');
        return context.sync().then(function () {
            tableHeaders = tableHeaderRowRange.values[0];
            rowStartIndex = range.rowIndex;
            // Table row index is relative to table position
            dataRowStartIndex = rowStartIndex - tableHeaderRowRange.rowIndex - 1;
            rowCount = range.rowCount;
            rawColumnKeys = columnKeys.filter(function (columnKey) { return tableHeaders.includes(columnKey); });
            formattedColumnKeys = columnKeys.filter(function (columnKey) { return rawColumnKeys.indexOf(columnKey) < 0 && columnKey.endsWith(formattedPostfix); });
        });
    })
        .then(function () {
        // No need to do exception handling here when invalid data is selected
        // We won't allow you to run the flow with invalid table selection
        Excel.run(function (ctx) {
            var range = ctx.workbook.getSelectedRange();
            var rangeTables = range.getTables();
            rangeTables.load();
            var table = rangeTables.getFirst();
            return sendUserData(ctx, tableHeaders, formattedColumnKeys, rawColumnKeys, widgetDoneCallback, table, rowStartIndex, dataRowStartIndex, rowCount);
        });
    })
        .catch(function (error) {
        // If there was an issue getting data (for example, the customer has not selected a row in the table),
        // return an empty object to the widget, this will cause the flow to run without any data but will
        // not block any widget usage.
        widgetDoneCallback(null, null);
        console.error(error);
    });
};
var sendUserData = function (ctx, tableHeaders, formattedColumnKeys, rawColumnKeys, widgetDoneCallback, table, rowStartIndex, dataRowStartIndex, rowCount) {
    var rows = table.rows;
    var rowRanges = [];
    // This is for displaying selected rows in the confirmation page
    var selectedRows = [];
    var boundImplicitData = [];
    for (var i = 0; i < rowCount; i++) {
        var row = rows.getItemAt(dataRowStartIndex + i);
        selectedRows.push(rowStartIndex + i + 1);
        var rowRange = row.getRange();
        rowRange.load('text');
        rowRange.load('values');
        rowRanges.push(rowRange);
    }
    return ctx
        .sync()
        .then(function () {
        var formattedPostfixRegex = new RegExp(formattedPostfix + '+$');
        boundImplicitData.selectedRows = selectedRows;
        var columnMap = [];
        rawColumnKeys.forEach(function (columnKey) {
            columnMap[columnKey] = tableHeaders.indexOf(columnKey);
        });
        formattedColumnKeys.forEach(function (columnKey) {
            columnMap[columnKey] = tableHeaders.indexOf(columnKey.replace(formattedPostfixRegex, ''));
        });
        rowRanges.forEach(function (currentRow, rowIndex) {
            // Row data has to be an object to be correctly stringified during flow run
            var rowData = {};
            rawColumnKeys.forEach(function (rawColumnKey) {
                rowData[rawColumnKey] =
                    columnMap[rawColumnKey] !== -1 ? rowRanges[rowIndex].values[0][columnMap[rawColumnKey]] : '';
            });
            formattedColumnKeys.forEach(function (formattedColumnKey) {
                rowData[formattedColumnKey] =
                    columnMap[formattedColumnKey] !== -1
                        ? rowRanges[rowIndex].text[0][columnMap[formattedColumnKey]].trim()
                        : '';
            });
            boundImplicitData.push(rowData);
        });
        widgetDoneCallback(null, { implicitData: boundImplicitData });
    })
        .catch(function (error) {
        console.error('SendUserData failed. Error: ' + error);
        widgetDoneCallback(null, { implicitData: boundImplicitData });
    });
};
Office.onReady(function () {
    document.getElementById('spinner').style.display = 'block';
    document.getElementById('spinner-container').style.display = 'block';
    document.getElementById('flow-div').style.display = 'none';
    OfficeFirstPartyAuth.load(WORKAROUND_REPLY_URL, ["https://service.flow.microsoft.com/"]).then(function () {
        var graphTokenPromise = OfficeFirstPartyAuth.getAccessToken({ resource: graphEndpoint }, false);
        var tokenParams = {
            'resource': 'https://service.flow.microsoft.com/',
            'authChallenge': '',
            'policy': ''
        };
        var behaviorParam = {
            'popup': false,
            'forceRefresh': false
        };
        var flowTokenPromise = OfficeFirstPartyAuth.getAccessToken(tokenParams, behaviorParam);
        Promise.all([graphTokenPromise, flowTokenPromise]).then(function (_a) {
            var graphTokenResult = _a[0], flowTokenResult = _a[1];
            Office.context.document.getFilePropertiesAsync(function (fileResult) {
                var _a, _b;
                // Step 1: Get the shared file id to call graph.
                if ((_b = (_a = fileResult === null || fileResult === void 0 ? void 0 : fileResult.value) === null || _a === void 0 ? void 0 : _a.url) === null || _b === void 0 ? void 0 : _b.startsWith('https')) {
                    //getGraphSharesApiPathFromUrl
                    // Encode the url to onedrive item-id
                    // https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/shares_get
                    var workbookUrlInBase64 = btoa(encodeURI(fileResult.value.url));
                    var encodedUrl = 'u!' + workbookUrlInBase64.replace('/(^=)/g', '').replace('/', '_').replace('+', '-');
                    var sharedApiPath = "/shares/" + encodedUrl + "/driveItem";
                    // Step 2: Get the file id and drive id.
                    var graphToken = graphTokenResult.accessToken;
                    var graphAuthHeader = "Bearer " + graphToken;
                    var graphApplicationHeader = "WACPowerAutomateSDX, WANPowerAutomateSDX";
                    var graphScenarioHeader = "GetExcelFileMetadata";
                    var graphHeaders_1 = {
                        'Authorization': graphAuthHeader,
                        'Accept': 'application/json',
                        'Application': graphApplicationHeader,
                        'Scenario': graphScenarioHeader
                    };
                    // Parameters for flow list and template
                    var autoFillParams_1 = {};
                    var fileId_1 = "nothing";
                    var hostDriveId_1 = "nothing";
                    fetch(graphEndpoint + "/v1.0" + sharedApiPath, {
                        headers: graphHeaders_1
                    })
                        .then(function (response) { return response.json(); })
                        .then(function (result) { return __awaiter(void 0, void 0, void 0, function () {
                        var tag, path, pathIndex, relativePath, fileName, driveType, siteId, excelSource, makeEndpoint;
                        var _a, _b, _c, _d, _e, _f;
                        return __generator(this, function (_g) {
                            switch (_g.label) {
                                case 0:
                                    if (!(((_a = result === null || result === void 0 ? void 0 : result.id) === null || _a === void 0 ? void 0 : _a.length) > 0 &&
                                        ((_c = (_b = result === null || result === void 0 ? void 0 : result.parentReference) === null || _b === void 0 ? void 0 : _b.driveId) === null || _c === void 0 ? void 0 : _c.length) > 0)) return [3 /*break*/, 2];
                                    tag = "/root:/";
                                    path = (_d = result.parentReference) === null || _d === void 0 ? void 0 : _d.path;
                                    pathIndex = path === null || path === void 0 ? void 0 : path.indexOf(tag);
                                    relativePath = pathIndex > 0 ? "" + path.substr(pathIndex + tag.length - 1) : '';
                                    fileId_1 = result.id;
                                    hostDriveId_1 = result.parentReference.driveId;
                                    fileName = relativePath + "/" + result.name;
                                    driveType = (_e = result === null || result === void 0 ? void 0 : result.parentReference) === null || _e === void 0 ? void 0 : _e.driveType;
                                    siteId = (_f = result === null || result === void 0 ? void 0 : result.parentReference) === null || _f === void 0 ? void 0 : _f.siteId;
                                    return [4 /*yield*/, getExcelSource(graphEndpoint, graphHeaders_1, driveType, siteId)];
                                case 1:
                                    excelSource = _g.sent();
                                    if (excelSource !== null) {
                                        autoFillParams_1 = {
                                            'parameters.officescripts.drive': hostDriveId_1,
                                            'parameters.officescripts.fileId': fileId_1,
                                            'parameters.officescripts.fileName': fileName,
                                            'parameters.officescripts.source': excelSource,
                                        };
                                    }
                                    _g.label = 2;
                                case 2:
                                    makeEndpoint = getMakeEndpointForToken(flowTokenResult.accessToken);
                                    loadFlowWidget(makeEndpoint, hostDriveId_1, fileId_1, autoFillParams_1, flowTokenResult.accessToken);
                                    return [2 /*return*/];
                            }
                        });
                    }); })
                        .catch(function (e) {
                        // TODO: add logging for exception
                    });
                }
            });
        });
    });
});


/***/ })

/******/ });
//# sourceMappingURL=powerAutomateAddin.js.map