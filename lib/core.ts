import * as xlsx from "xlsx";
import fetch from "node-fetch";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";
import get from "lodash/get";
import {
  GenericData,
  TableData,
  TableDataConfig,
} from "table-data-to-json/lib/core";
import convertTableDataToJSON from "table-data-to-json";
import {
  XLSX_DATE_FORMATS_TO_DAYJS_FORMATS,
  DEFAULT_XLSX_PARSING_OPTIONS,
} from "./constants";

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

export interface XLSXTableDataConfig
  extends Partial<TableDataConfig>,
    XLSXParseConfig {}

dayjs.extend(utc);

/**
 * Convert XLSX sheet from XLSX workbook to array table data.
 */
export function convertSheetToTableData<Data = GenericData>({
  workbook,
  sheet,
}: {
  workbook: xlsx.WorkBook;
  sheet: xlsx.WorkSheet;
}): TableData<Data> {
  const data: TableData<Data> = [];
  Object.entries(sheet).forEach(([cellRef, cell]) => {
    if (cellRef.indexOf("!") !== 0) {
      const coords = xlsx.utils.decode_cell(cellRef);
      if (!data[coords.r]) data[coords.r] = [];

      let cellValue = cell.v;
      switch (cell.t) {
        case "b":
          cellValue = ["true", "1", "y"].includes(`${cellValue}`.toLowerCase());
          break;

        case "d":
          if (cell.v instanceof Date) {
            cellValue = dayjs.utc(cell.v).toISOString();
          } else {
            cellValue = xlsx.SSF.parse_date_code(cell.v, {
              date1904: workbook.Workbook?.WBProps?.date1904,
            });
          }
          break;

        case "z":
          cellValue = null;
          break;

        case "n":
          if (
            cell.z &&
            Object.keys(XLSX_DATE_FORMATS_TO_DAYJS_FORMATS).includes(cell.z)
          ) {
            const parseDateFormat =
              XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z] instanceof Array
                ? XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z][0]
                : (XLSX_DATE_FORMATS_TO_DAYJS_FORMATS[cell.z] as string);

            cellValue = dayjs.utc(cell.w, parseDateFormat).toISOString();
          } else {
            cellValue = parseFloat(cellValue);
          }
          break;

        case "s":
        default:
          cellValue = `${cellValue}`;
      }

      data[coords.r][coords.c] = cellValue;
    }
  });

  return data;
}

export async function convertXLSXToTableData<Data = GenericData>({
  file,
  url,
  data,
  sheetIndex,
  sheetName,
  parsingOptions,
}: XLSXParseConfig): Promise<XLSXTableDataOutput<Data>> {
  const parseConfig = {
    ...DEFAULT_XLSX_PARSING_OPTIONS,
    sheets: sheetIndex || 0,
    ...parsingOptions,
  };

  let workbook: xlsx.WorkBook;
  if (data && data.Sheets) {
    workbook = data;
  } else if (url) {
    // Read remote file
    try {
      const contents = await (await fetch(url)).arrayBuffer();
      workbook = xlsx.read(contents, parseConfig);
    } catch (exp) {
      throw exp;
    }
  } else if (file) {
    // Read local file
    if (file instanceof Buffer) {
      workbook = xlsx.read(file, parseConfig);
    } else {
      workbook = xlsx.readFile(file, parseConfig);
    }
  } else {
    throw Error(
      "Please provide XLSX data, url or file (either file contents buffer or local path string)"
    );
  }

  const sheet: xlsx.Sheet = get(
    workbook.Sheets,
    sheetName || workbook.SheetNames[sheetIndex || 0]
  );
  const tableData: TableData<Data> = convertSheetToTableData<Data>({
    workbook,
    sheet,
  });

  return {
    workbook,
    sheet,
    tableData,
  };
}

/**
 * Convert an XLSX local file path, remote HTTP URL, Buffer,
 * or xlsx.WorkBook object (or other such file format supported
 * by SheetJS) to JSON.
 *
 * Will convert the data to tableData first, then perform any
 * modifications on headers based on merged cells in the sheet,
 * then convert to JSON.
 */
export async function convertXLSXToJSON<Data = GenericData, Output = object>({
  file,
  url,
  data,
  sheetIndex,
  sheetName,
  parsingOptions,
  preset,
  headers,
}: XLSXTableDataConfig): Promise<Output | Output[]> {
  const { sheet, tableData } = await convertXLSXToTableData<Data>({
    file,
    url,
    data,
    sheetIndex,
    sheetName,
    parsingOptions,
  });

  return convertTableDataToJSON<Data, Output>(tableData, {
    preset,
    headers,
    modifyHeaders: (headers) => {
      // Calculate header cell widths/heights for any
      // merged cells.
      if (sheet["!merges"] && sheet["!merges"].length) {
        headers.forEach((row) => {
          row.forEach((cell) => {
            // @ts-ignore
            sheet["!merges"].forEach((mergedCell) => {
              if (mergedCell.s.c === cell.c && mergedCell.s.r === cell.r) {
                cell.width = 1 + (mergedCell.e.c - mergedCell.s.c);
                cell.height = 1 + (mergedCell.e.r - mergedCell.s.r);
              }
            });
          });
        });
      }
    },
  });
}

export default convertXLSXToJSON;
