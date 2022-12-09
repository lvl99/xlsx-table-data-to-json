import { ParsingOptions } from "xlsx";

export const XLSX_DATE_FORMATS_TO_DAYJS_FORMATS: {
  [key: string]: string | string[];
} = {
  "m/d/yy": "M/D/YY",
  "d-mmm-yy": "D-MMM-YY",
  "d-mmm": "D-MMM",
  "mmm-yy": "MMM-YY",
  "h:mm AM/PM": "h:mm A",
  "h:mm:ss AM/PM": "h:mm:ss A",
  "h:mm": "h:mm",
  "h:mm:ss": "h:mm:ss",
  "m/d/yy h:mm": "M/D/YY h:mm",
  "mm:ss": "mm:ss",
  "[h]:mm:ss": ["[h]:mm:ss", "mm:ss"],
  "mmss.0": "mmss.SSS",
  '"上午/下午 "hh"時"mm"分"ss"秒 "': "A hh[時]mm[分]ss[秒]",
  "yyyy-mm-dd": "YYYY-MM-DD",
  "yyyy-mm-dd hh:mm:ss": "YYYY-MM-DD HH:mm:ss",
};

export const DEFAULT_XLSX_PARSING_OPTIONS: ParsingOptions = {
  cellNF: true,
  cellHTML: true,
  cellStyles: true,
};
