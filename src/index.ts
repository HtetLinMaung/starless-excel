import ExcelJS, { Row, Workbook, Worksheet } from "exceljs";
import fs from "fs";
import { isAsyncFunction } from "util/types";
import xlsx from "xlsx";

export const createWorkbook = () => new ExcelJS.Workbook();

export const readExcel = async (data: string | Buffer | fs.ReadStream) => {
  const workbook = createWorkbook();
  if (Buffer.isBuffer(data)) {
    await workbook.xlsx.load(data);
  } else if (typeof data == "string") {
    await workbook.xlsx.readFile(data);
  } else {
    await workbook.xlsx.read(data);
  }

  return workbook;
};

export const base64ExcelToJson = async (
  base64: string,
  indexOrName: string | number,
  readMap: any = {},
  cb = async (m: any, row: Row, rowNumber: number) => m
) => {
  return excelToJson(Buffer.from(base64, "base64"), indexOrName, readMap, cb);
};

export const sheetToJson = async (
  sheet: Worksheet,
  readMap: any = {},
  cb = async (m: any, row: Row, rowNumber: number) => m
) => {
  let items: any[] = [];
  sheet.eachRow((row, rowNumber) => {
    let item: any = {};
    for (const [key, indexOrKey] of Object.entries(readMap)) {
      item[key] = row.getCell(indexOrKey as string | number).value;
    }

    items.push(cb(item, row, rowNumber));
  });

  if (isAsyncFunction(cb) || cb.toString().includes("__awaiter")) {
    items = await Promise.all(items);
  }
  return items;
};

const mapKeys = (m, keyMap) => {
  const item: any = {};
  for (const [k, v] of Object.entries(m)) {
    const key = keyMap[k];
    if (key) {
      item[key] = v;
    } else {
      item[k] = v;
    }
  }
  return item;
};

export const excelToJsonv2 = async (
  data: any,
  type: "string" | "file" | "base64" | "binary" | "buffer" | "array" = "string",
  indexOrName: string | number,
  range: number = 0,
  keyMap: any = {},
  cb = async (m: any, rowNumber: number) => m
) => {
  const workbook = xlsx.readFile(data, { type });
  const sheet =
    workbook.Sheets[
      typeof indexOrName == "string"
        ? indexOrName
        : workbook.SheetNames[indexOrName]
    ];
  if (!sheet) {
    throw new Error("Sheet not found!");
  }
  const items = xlsx.utils.sheet_to_json(sheet, { range });
  let newItems = [];
  for (let i = 0; i < items.length; i++) {
    newItems.push(cb(mapKeys(items[i], keyMap), i + 1));
  }
  if (isAsyncFunction(cb) || cb.toString().includes("__awaiter")) {
    newItems = await Promise.all(newItems);
  }
  return newItems;
};

export const excelToJson = async (
  data: string | Buffer | fs.ReadStream,
  indexOrName: string | number,
  readMap: any = {},
  cb = async (m: any, row: Row, rowNumber: number) => m
) => {
  const workbook = await readExcel(data);
  const sheet = workbook.getWorksheet(indexOrName);
  if (!sheet) {
    throw new Error("Sheet not found!");
  }
  return sheetToJson(sheet, readMap, cb);
};

export interface ExcelPage {
  sheetName: string;
  data: any[];
}

export const writeExcelFromJson = async (
  pages: ExcelPage[],
  exportType: string = "file",
  fileNameOrStream: any = ""
) => {
  const { workbook } = createExcelFromJson(pages);
  return writeExcel(workbook, exportType, fileNameOrStream);
};

export const writeExcel = async (
  workbook: Workbook,
  exportType: string = "file",
  fileNameOrStream: any = ""
) => {
  switch (exportType) {
    case "file":
      await workbook.xlsx.writeFile(fileNameOrStream);
      break;
    case "stream":
      await workbook.xlsx.write(fileNameOrStream);
      break;
    case "buffer":
      return await workbook.xlsx.writeBuffer();
  }
  return null;
};

export const createExcelFromJson = (pages: ExcelPage[]) => {
  if (pages.length) {
    const workbook = createWorkbook();
    const sheets: Worksheet[] = [];
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
