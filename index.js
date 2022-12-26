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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createExcelFromJson = exports.writeExcel = exports.writeExcelFromJson = exports.excelToJson = exports.excelToJsonv2 = exports.sheetToJson = exports.base64ExcelToJson = exports.readExcel = exports.createWorkbook = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const types_1 = require("util/types");
const xlsx_1 = __importDefault(require("xlsx"));
const createWorkbook = () => new exceljs_1.default.Workbook();
exports.createWorkbook = createWorkbook;
const readExcel = (data) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = (0, exports.createWorkbook)();
    if (Buffer.isBuffer(data)) {
        yield workbook.xlsx.load(data);
    }
    else if (typeof data == "string") {
        yield workbook.xlsx.readFile(data);
    }
    else {
        yield workbook.xlsx.read(data);
    }
    return workbook;
});
exports.readExcel = readExcel;
const base64ExcelToJson = (base64, indexOrName, readMap = {}, cb = (m, row, rowNumber) => __awaiter(void 0, void 0, void 0, function* () { return m; })) => __awaiter(void 0, void 0, void 0, function* () {
    return (0, exports.excelToJson)(Buffer.from(base64, "base64"), indexOrName, readMap, cb);
});
exports.base64ExcelToJson = base64ExcelToJson;
const sheetToJson = (sheet, readMap = {}, cb = (m, row, rowNumber) => __awaiter(void 0, void 0, void 0, function* () { return m; })) => __awaiter(void 0, void 0, void 0, function* () {
    let items = [];
    sheet.eachRow((row, rowNumber) => {
        let item = {};
        for (const [key, indexOrKey] of Object.entries(readMap)) {
            item[key] = row.getCell(indexOrKey).value;
        }
        items.push(cb(item, row, rowNumber));
    });
    if ((0, types_1.isAsyncFunction)(cb) || cb.toString().includes("__awaiter")) {
        items = yield Promise.all(items);
    }
    return items;
});
exports.sheetToJson = sheetToJson;
const mapKeys = (m, keyMap) => {
    const item = {};
    for (const [k, v] of Object.entries(m)) {
        const key = keyMap[k];
        if (key) {
            item[key] = v;
        }
        else {
            item[k] = v;
        }
    }
    return item;
};
const excelToJsonv2 = (data, type = "string", indexOrName, range = 0, keyMap = {}, cb = (m, rowNumber) => __awaiter(void 0, void 0, void 0, function* () { return m; })) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = xlsx_1.default.readFile(data, { type });
    const sheet = workbook.Sheets[typeof indexOrName == "string"
        ? indexOrName
        : workbook.SheetNames[indexOrName]];
    if (!sheet) {
        throw new Error("Sheet not found!");
    }
    const items = xlsx_1.default.utils.sheet_to_json(sheet, { range });
    let newItems = [];
    for (let i = 0; i < items.length; i++) {
        newItems.push(cb(mapKeys(items[i], keyMap), i + 1));
    }
    if ((0, types_1.isAsyncFunction)(cb) || cb.toString().includes("__awaiter")) {
        newItems = yield Promise.all(newItems);
    }
    return newItems;
});
exports.excelToJsonv2 = excelToJsonv2;
const excelToJson = (data, indexOrName, readMap = {}, cb = (m, row, rowNumber) => __awaiter(void 0, void 0, void 0, function* () { return m; })) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = yield (0, exports.readExcel)(data);
    const sheet = workbook.getWorksheet(indexOrName);
    if (!sheet) {
        throw new Error("Sheet not found!");
    }
    return (0, exports.sheetToJson)(sheet, readMap, cb);
});
exports.excelToJson = excelToJson;
const writeExcelFromJson = (pages, exportType = "file", fileNameOrStream = "") => __awaiter(void 0, void 0, void 0, function* () {
    const { workbook } = (0, exports.createExcelFromJson)(pages);
    return (0, exports.writeExcel)(workbook, exportType, fileNameOrStream);
});
exports.writeExcelFromJson = writeExcelFromJson;
const writeExcel = (workbook, exportType = "file", fileNameOrStream = "") => __awaiter(void 0, void 0, void 0, function* () {
    switch (exportType) {
        case "file":
            yield workbook.xlsx.writeFile(fileNameOrStream);
            break;
        case "stream":
            yield workbook.xlsx.write(fileNameOrStream);
            break;
        case "buffer":
            return yield workbook.xlsx.writeBuffer();
    }
    return null;
});
exports.writeExcel = writeExcel;
const createExcelFromJson = (pages) => {
    if (pages.length) {
        const workbook = (0, exports.createWorkbook)();
        const sheets = [];
        for (const { data, sheetName } of pages) {
            if (data.length) {
                const headers = Object.keys(data[0]);
                const sheet = workbook.addWorksheet(sheetName);
                sheet.addRows([headers, ...data.map((d) => Object.values(d))]);
                sheets.push(sheet);
            }
        }
        return { workbook, sheets };
    }
};
exports.createExcelFromJson = createExcelFromJson;
