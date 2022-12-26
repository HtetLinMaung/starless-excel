/// <reference types="node" />
import ExcelJS, { Row, Workbook, Worksheet } from "exceljs";
import fs from "fs";
export declare const createWorkbook: () => ExcelJS.Workbook;
export declare const readExcel: (data: string | Buffer | fs.ReadStream) => Promise<ExcelJS.Workbook>;
export declare const base64ExcelToJson: (base64: string, indexOrName: string | number, readMap?: any, cb?: (m: any, row: Row, rowNumber: number) => Promise<any>) => Promise<any[]>;
export declare const sheetToJson: (sheet: Worksheet, readMap?: any, cb?: (m: any, row: Row, rowNumber: number) => Promise<any>) => Promise<any[]>;
export declare const excelToJsonv2: (data: any, type: "string" | "file" | "base64" | "binary" | "buffer" | "array", indexOrName: string | number, range?: number, keyMap?: any, cb?: (m: any, rowNumber: number) => Promise<any>) => Promise<any[]>;
export declare const excelToJson: (data: string | Buffer | fs.ReadStream, indexOrName: string | number, readMap?: any, cb?: (m: any, row: Row, rowNumber: number) => Promise<any>) => Promise<any[]>;
export interface ExcelPage {
    sheetName: string;
    data: any[];
}
export declare const writeExcelFromJson: (pages: ExcelPage[], exportType?: string, fileNameOrStream?: any) => Promise<ExcelJS.Buffer>;
export declare const writeExcel: (workbook: Workbook, exportType?: string, fileNameOrStream?: any) => Promise<ExcelJS.Buffer>;
export declare const createExcelFromJson: (pages: ExcelPage[]) => {
    workbook: ExcelJS.Workbook;
    sheets: ExcelJS.Worksheet[];
};
