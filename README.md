# XLSX Table Data to JSON

This package augments [`table-data-to-json`](https://www.npmjs.com/package/table-data-to-json)
with the ability to load and extract data from XLSX files using
[`xlsx`](https://www.npmjs.com/package/xlsx).

Thanks to the fully-featured nature of `xlsx`, this package
supports CSV files out of the box as well.

## Installation

```sh
  npm i xlsx-table-data-to-json
  yarn add xlsx-table-data-to-json
```

## Usage

Import and use `convertXLSXToJSON(options)` within your project like so:

```js
import convertXLSXToJSON from "xlsx-table-data-to-json";

// @NOTE: Unlike `table-data-to-json`, `xlsx-table-data-to-json` relies on
//        `async`, so will require `await` or `.then()`.
const jsonData = await convertXLSXToJSON(options);
```

Path to a local XLSX/CSV file can be given, or you can give it a `Buffer` directly:

```js
const jsonData = await convertXLSXToJSON({
  file: path.join(__dirname, "example.xlsx"),
});

const jsonData = await convertXLSXToJSON({
  file: fs.readFileSync(path.join(__dirname, "example.xlsx")),
});
```

URL to a remote XLSX/CSV file can be provided as well:

```js
const jsonData = await convertXLSXToJSON({
  url: "https://example.com/example.xlsx",
});
```

You can even give it an `xlsx.WorkBook` instance:

```js
const jsonData = await convertXLSXToJSON({
  data: xlsxWorkBook,
});
```

### Options

| Property                 | Type                                                                                                 | Description                                                                                                          |
| ------------------------ | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `options.file`           | String or `Buffer`                                                                                   | Either the path to a local file or a `Buffer` object itself.                                                         |
| `options.url`            | String                                                                                               | URL to a remote file.                                                                                                |
| `options.data`           | `xlsx.WorkBook`                                                                                      | The `xlsx.WorkBook` instance.                                                                                        |
| `options.sheetIndex`     | Number                                                                                               | A zero-index number which corresponds to the sheet you want to convert.                                              |
| `options.sheetName`      | String                                                                                               | The name of the sheet you want to convert.                                                                           |
| `options.parsingOptions` | `xlsx.ParsingOptions`                                                                                | Override or augment any further `xlsx` parsing options.                                                              |
| `preset`                 | String                                                                                               | Accepted values:<ul><li>`row`</li><li>`column`</li><li>`row.column`</li><li>`column.row`</li><li>`row.row`</li></ul> |
| `headers`                | [`TableDataConfigHeaders`](https://github.com/lvl99/table-data-to-json/blob/main/lib/core.ts#L26:32) | In case the presets don't cover your use-case, you can specify the headers here.                                     |

## Development

To download external dependencies:

```sh
  npm i
```

To run tests (using Jest):

```sh
  npm test
  npm run test:watch
```

## Contribute

Got cool ideas? Have questions or feedback? Found a bug? [Post an issue](https://github.com/lvl99/table-data-to-json/issues)

Added a feature? Fixed a bug? [Post a PR](https://github.com/lvl99/table-data-to-json/compare)

## License

[Apache 2.0](LICENSE.md)
