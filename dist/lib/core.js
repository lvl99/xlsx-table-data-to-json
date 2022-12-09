"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.convertXLSXToJSON = exports.convertXLSXToTableData = exports.convertSheetToTableData = void 0;
var tslib_1 = require("tslib");
var xlsx = tslib_1.__importStar(require("xlsx"));
var node_fetch_1 = tslib_1.__importDefault(require("node-fetch"));
var dayjs_1 = tslib_1.__importDefault(require("dayjs"));
var utc_1 = tslib_1.__importDefault(require("dayjs/plugin/utc"));
var get_1 = tslib_1.__importDefault(require("lodash/get"));
var table_data_to_json_1 = tslib_1.__importDefault(require("table-data-to-json"));
var constants_1 = require("./constants");
dayjs_1.default.extend(utc_1.default);
/**
 * Convert XLSX sheet from XLSX workbook to array table data.
 */
function convertSheetToTableData(_a) {
    var workbook = _a.workbook, sheet = _a.sheet;
    var data = [];
    Object.entries(sheet).forEach(function (_a) {
        var _b, _c;
        var cellRef = _a[0], cell = _a[1];
        if (cellRef.indexOf("!") !== 0) {
            var coords = xlsx.utils.decode_cell(cellRef);
            if (!data[coords.r])
                data[coords.r] = [];
            var cellValue = cell.v;
            switch (cell.t) {
                case "b":
                    cellValue = ["true", "1", "y"].includes("".concat(cellValue).toLowerCase());
                    break;
                case "d":
                    if (cell.v instanceof Date) {
                        cellValue = dayjs_1.default.utc(cell.v).toISOString();
                    }
                    else {
                        cellValue = xlsx.SSF.parse_date_code(cell.v, {
                            date1904: (_c = (_b = workbook.Workbook) === null || _b === void 0 ? void 0 : _b.WBProps) === null || _c === void 0 ? void 0 : _c.date1904,
                        });
                    }
                    break;
                case "z":
                    cellValue = null;
                    break;
                case "n":
                    if (cell.z &&
                        Object.keys(constants_1.XLSX_DATE_FORMATS_TO_DAYJS_FORMATS).includes(cell.z)) {
                        var parseDateFormat = constants_1.XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z] instanceof Array
                            ? constants_1.XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z][0]
                            : constants_1.XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z];
                        cellValue = dayjs_1.default.utc(cell.w, parseDateFormat).toISOString();
                    }
                    else {
                        cellValue = parseFloat(cellValue);
                    }
                    break;
                case "s":
                default:
                    cellValue = "".concat(cellValue);
            }
            data[coords.r][coords.c] = cellValue;
        }
    });
    return data;
}
exports.convertSheetToTableData = convertSheetToTableData;
function convertXLSXToTableData(_a) {
    var file = _a.file, url = _a.url, data = _a.data, sheetIndex = _a.sheetIndex, sheetName = _a.sheetName, parsingOptions = _a.parsingOptions;
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var parseConfig, workbook, contents, exp_1, sheet, tableData;
        return tslib_1.__generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    parseConfig = tslib_1.__assign(tslib_1.__assign(tslib_1.__assign({}, constants_1.DEFAULT_XLSX_PARSING_OPTIONS), { sheets: sheetIndex || 0 }), parsingOptions);
                    if (!(data && data.Sheets)) return [3 /*break*/, 1];
                    workbook = data;
                    return [3 /*break*/, 8];
                case 1:
                    if (!url) return [3 /*break*/, 7];
                    _b.label = 2;
                case 2:
                    _b.trys.push([2, 5, , 6]);
                    return [4 /*yield*/, (0, node_fetch_1.default)(url)];
                case 3: return [4 /*yield*/, (_b.sent()).arrayBuffer()];
                case 4:
                    contents = _b.sent();
                    workbook = xlsx.read(contents, parseConfig);
                    return [3 /*break*/, 6];
                case 5:
                    exp_1 = _b.sent();
                    throw exp_1;
                case 6: return [3 /*break*/, 8];
                case 7:
                    if (file) {
                        // Read local file
                        if (file instanceof Buffer) {
                            workbook = xlsx.read(file, parseConfig);
                        }
                        else {
                            workbook = xlsx.readFile(file, parseConfig);
                        }
                    }
                    else {
                        throw Error("Please provide XLSX data, url or file (either file contents buffer or local path string)");
                    }
                    _b.label = 8;
                case 8:
                    sheet = (0, get_1.default)(workbook.Sheets, sheetName || workbook.SheetNames[sheetIndex || 0]);
                    tableData = convertSheetToTableData({
                        workbook: workbook,
                        sheet: sheet,
                    });
                    return [2 /*return*/, {
                            workbook: workbook,
                            sheet: sheet,
                            tableData: tableData,
                        }];
            }
        });
    });
}
exports.convertXLSXToTableData = convertXLSXToTableData;
/**
 * Convert an XLSX local file path, remote HTTP URL, Buffer,
 * or xlsx.WorkBook object (or other such file format supported
 * by SheetJS) to JSON.
 *
 * Will convert the data to tableData first, then perform any
 * modifications on headers based on merged cells in the sheet,
 * then convert to JSON.
 */
function convertXLSXToJSON(_a) {
    var file = _a.file, url = _a.url, data = _a.data, sheetIndex = _a.sheetIndex, sheetName = _a.sheetName, parsingOptions = _a.parsingOptions, preset = _a.preset, headers = _a.headers;
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var _b, sheet, tableData;
        return tslib_1.__generator(this, function (_c) {
            switch (_c.label) {
                case 0: return [4 /*yield*/, convertXLSXToTableData({
                        file: file,
                        url: url,
                        data: data,
                        sheetIndex: sheetIndex,
                        sheetName: sheetName,
                        parsingOptions: parsingOptions,
                    })];
                case 1:
                    _b = _c.sent(), sheet = _b.sheet, tableData = _b.tableData;
                    return [2 /*return*/, (0, table_data_to_json_1.default)(tableData, {
                            preset: preset,
                            headers: headers,
                            modifyHeaders: function (headers) {
                                // Calculate header cell widths/heights for any
                                // merged cells.
                                if (sheet["!merges"] && sheet["!merges"].length) {
                                    headers.forEach(function (row) {
                                        row.forEach(function (cell) {
                                            // @ts-ignore
                                            sheet["!merges"].forEach(function (mergedCell) {
                                                if (mergedCell.s.c === cell.c && mergedCell.s.r === cell.r) {
                                                    cell.width = 1 + (mergedCell.e.c - mergedCell.s.c);
                                                    cell.height = 1 + (mergedCell.e.r - mergedCell.s.r);
                                                }
                                            });
                                        });
                                    });
                                }
                            },
                        })];
            }
        });
    });
}
exports.convertXLSXToJSON = convertXLSXToJSON;
exports.default = convertXLSXToJSON;
//# sourceMappingURL=core.js.map