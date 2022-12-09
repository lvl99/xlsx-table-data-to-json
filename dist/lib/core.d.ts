/// <reference types="node" />
import * as xlsx from "xlsx";
import { GenericData, TableData, TableDataConfig } from "table-data-to-json/lib/core";
export interface XLSXParseConfig {
    url?: string;
    file?: Buffer | string;
    data?: xlsx.WorkBook;
    sheetIndex?: number;
    sheetName?: string;
    parsingOptions?: xlsx.ParsingOptions;
}
export interface XLSXTableDataOutput<Data = GenericData> {
    workbook: xlsx.WorkBook;
    sheet: xlsx.WorkSheet;
    tableData: TableData<Data>;
}
export interface XLSXTableDataConfig extends Partial<TableDataConfig>, XLSXParseConfig {
}
/**
 * Convert XLSX sheet from XLSX workbook to array table data.
 */
export declare function convertSheetToTableData<Data = GenericData>({ workbook, sheet, }: {
    workbook: xlsx.WorkBook;
    sheet: xlsx.WorkSheet;
}): TableData<Data>;
export declare function convertXLSXToTableData<Data = GenericData>({ file, url, data, sheetIndex, sheetName, parsingOptions, }: XLSXParseConfig): Promise<XLSXTableDataOutput<Data>>;
/**
 * Convert an XLSX local file path, remote HTTP URL, Buffer,
 * or xlsx.WorkBook object (or other such file format supported
 * by SheetJS) to JSON.
 *
 * Will convert the data to tableData first, then perform any
 * modifications on headers based on merged cells in the sheet,
 * then convert to JSON.
 */
export declare function convertXLSXToJSON<Data = GenericData, Output = object>({ file, url, data, sheetIndex, sheetName, parsingOptions, preset, headers, }: XLSXTableDataConfig): Promise<Output | Output[]>;
export default convertXLSXToJSON;
