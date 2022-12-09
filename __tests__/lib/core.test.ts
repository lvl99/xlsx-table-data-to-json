import fs from "node:fs";
import path from "node:path";
import xlsx from "xlsx";
import fetch from "node-fetch";
import { DEFAULT_XLSX_PARSING_OPTIONS } from "../../lib/constants";
import { convertXLSXToTableData, convertXLSXToJSON } from "../../lib/core";

jest.mock("node-fetch");
const { Response } = jest.requireActual("node-fetch");

const testXLSXFilePath = path.join(__dirname, "./test.xlsx");
const testXLSXFileBuffer = fs.readFileSync(testXLSXFilePath);
const testXLSXWorkBook = xlsx.readFile(
  testXLSXFilePath,
  DEFAULT_XLSX_PARSING_OPTIONS
);

const testCSVFilePath = path.join(__dirname, "./test.csv");
const testCSVFileBuffer = fs.readFileSync(testCSVFilePath);

describe("convertXlsxToTableData", () => {
  it("should open XLSX file (local path string) and convert contents to array table data", async () => {
    const { tableData } = await convertXLSXToTableData({
      file: testXLSXFilePath,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should open XLSX file (Buffer) and convert contents to array table data", async () => {
    const { tableData } = await convertXLSXToTableData({
      file: testXLSXFileBuffer,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should open CSV file (local path string) and convert contents to array table data", async () => {
    const { tableData } = await convertXLSXToTableData({
      file: testCSVFilePath,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should open CSV file (Buffer) and convert contents to array table data", async () => {
    const { tableData } = await convertXLSXToTableData({
      file: testCSVFileBuffer,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should open XLSX file (remote HTTP URL) and convert contents to array table data", async () => {
    // @ts-ignore
    fetch.mockReturnValue(Promise.resolve(new Response(testXLSXFileBuffer)));

    const { tableData } = await convertXLSXToTableData({
      url: "https://example.com/table-data.xlsx",
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should use XLSX data (xlsx.WorkBook) and convert contents to array table data", async () => {
    const { tableData } = await convertXLSXToTableData({
      data: testXLSXWorkBook,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(4);
    expect(tableData).toMatchSnapshot();
  });

  it("should throw an error if no file, url or data given", async () => {
    await expect(convertXLSXToTableData({})).rejects.toThrowError(
      "Please provide XLSX data, url or file (either file contents buffer or local path string)"
    );
  });

  it("should throw an error if something happened with fetch", async () => {
    const errorMsg = "Testing error in node-fetch";

    // @ts-ignore
    fetch.mockReturnValue(Promise.reject(new Error(errorMsg)));

    await expect(
      convertXLSXToTableData({ url: "https://example.com/test.xlsx" })
    ).rejects.toThrowError(errorMsg);
  });

  it("should throw an error if data is not xlsx.WorkBook", async () => {
    await expect(
      convertXLSXToTableData({ data: {} as xlsx.WorkBook })
    ).rejects.toThrowError(
      "Please provide XLSX data, url or file (either file contents buffer or local path string)"
    );
  });

  it("should pick out a specific sheet by zero-based index", async () => {
    const { tableData } = await convertXLSXToTableData({
      data: testXLSXWorkBook,
      sheetIndex: 1,
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(3);
    expect(tableData).toMatchSnapshot();
  });

  it("should pick out a specific sheet by name", async () => {
    const { tableData } = await convertXLSXToTableData({
      data: testXLSXWorkBook,
      sheetName: "column",
    });
    expect(tableData).toBeInstanceOf(Array);
    expect(tableData).toHaveLength(3);
    expect(tableData).toMatchSnapshot();
  });

  it("should convert special Excel cells to JSON data", async () => {
    const originalTableData = [
      [
        "test",
        "",
        undefined,
        null,
        true,
        false,
        123456789,
        new Date("2022-11-13"),
      ],
    ];
    const sheet = xlsx.utils.aoa_to_sheet(originalTableData, {
      cellDates: true,
    });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, sheet, "Test 1");

    const { tableData } = await convertXLSXToTableData({
      data: wb,
    });
    expect(tableData).toHaveLength(1);
    expect(tableData[0]).toHaveLength(8);
    expect(tableData[0][0]).toBe("test");
    expect(tableData[0][1]).toBe("");
    expect(tableData[0][2]).toBeUndefined();
    expect(tableData[0][3]).toBeUndefined();
    expect(tableData[0][4]).toBe(true);
    expect(tableData[0][5]).toBe(false);
    expect(tableData[0][6]).toBe(123456789);
    expect(tableData[0][7]).toStrictEqual("2022-11-13T00:00:00.000Z");
  });
});

describe("convertXLSXToJSON", () => {
  it("should open XLSX file (local path string) and convert contents to array table data", async () => {
    const result = await convertXLSXToJSON({
      file: testXLSXFilePath,
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should open XLSX file (Buffer) and convert contents to array table data", async () => {
    const result = await convertXLSXToJSON({
      file: testXLSXFileBuffer,
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should open CSV file (local path string) and convert contents to array table data", async () => {
    const result = await convertXLSXToJSON({
      file: testCSVFilePath,
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should open CSV file (Buffer) and convert contents to array table data", async () => {
    const result = await convertXLSXToJSON({
      file: testCSVFileBuffer,
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should open XLSX file (remote HTTP URL) and convert contents to array table data", async () => {
    // @ts-ignore
    fetch.mockReturnValue(Promise.resolve(new Response(testXLSXFileBuffer)));

    const result = await convertXLSXToJSON({
      url: "https://example.com/table-data.xlsx",
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should use XLSX data (xlsx.Workbook) and convert contents to array table data", async () => {
    const result = await convertXLSXToJSON({
      data: testXLSXWorkBook,
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(3);
    expect(result).toMatchSnapshot();
  });

  it("should throw an error if no file, url or data given", async () => {
    await expect(convertXLSXToJSON({})).rejects.toThrowError(
      "Please provide XLSX data, url or file (either file contents buffer or local path string)"
    );
  });

  it("should throw an error if something happened with fetch", async () => {
    const errorMsg = "Testing error in node-fetch";

    // @ts-ignore
    fetch.mockReturnValue(Promise.reject(new Error(errorMsg)));

    await expect(
      convertXLSXToJSON({ url: "https://example.com/test.xlsx" })
    ).rejects.toThrowError(errorMsg);
  });

  it("should throw an error if data is not xlsx.WorkBook", async () => {
    await expect(
      convertXLSXToJSON({ data: {} as xlsx.WorkBook })
    ).rejects.toThrowError(
      "Please provide XLSX data, url or file (either file contents buffer or local path string)"
    );
  });

  it("should pick out a specific sheet by zero-based index", async () => {
    const result = await convertXLSXToJSON({
      file: testXLSXFileBuffer,
      sheetIndex: 1,
      preset: "row.column",
    });
    expect(result).toBeInstanceOf(Object);
    expect(result).toMatchSnapshot();
  });

  it("should merge header cells", async () => {
    const result = await convertXLSXToJSON({
      file: testXLSXFileBuffer,
      sheetIndex: 4,
      preset: "row.row",
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(1);
    expect(result).toMatchSnapshot();
  });

  it("should merge and nest header cells", async () => {
    const result = await convertXLSXToJSON({
      file: testXLSXFileBuffer,
      sheetIndex: 6,
      preset: "row.row",
    });
    expect(result).toBeInstanceOf(Array);
    expect(result).toHaveLength(2);
    expect(result).toMatchSnapshot();
  });
});
