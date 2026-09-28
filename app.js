/*
 * Convert a Strava activities.csv export into a clean xlsx workbook.
 *
 * - Text fields may contain commas and line breaks; they are parsed as proper CSV.
 * - Columns in which every value is a plain number become numbers, columns in which
 *   every value is a Strava date become real Excel dates, everything else stays text.
 *
 * Works in the browser (window.StravaFixer) and in Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.StravaFixer = factory();
})(this, function () {
  "use strict";

  const NUMBER_RE = /^-?\d+(\.\d+)?$/;
  // Control characters that are not allowed in an xlsx file.
  const ILLEGAL_CHARACTERS_RE = /[\x00-\x08\x0b\x0c\x0e-\x1f]/g;
  const MAX_CELL_LENGTH = 32767;
  // The first distance column is display text with a unit per sport ("10,29" km,
  // "1.025" m), so it must never be read as a number.
  const DISPLAY_DISTANCE = 6;

  // "10 jun 2026, 16:17:48" (Dutch export) and "Jun 10, 2026, 4:17:48 PM" (English export).
  const DATE_NL_RE = /^(\d{1,2}) (\p{L}+)\.? (\d{4}),? (\d{1,2}):(\d{2}):(\d{2})$/u;
  const DATE_EN_RE = /^(\p{L}+)\.? (\d{1,2}), (\d{4}),? (\d{1,2}):(\d{2}):(\d{2})(?: ?([AP]M))?$/iu;
  const MONTHS = {
    jan: 0, feb: 1, mrt: 2, mar: 2, apr: 3, mei: 4, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, sept: 8, okt: 9, oct: 9, nov: 10, dec: 11,
  };
  const DATE_FORMAT = "dd-mm-yyyy hh:mm:ss";
  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

  /** An error with a message that can be shown to the user as is. */
  class FixError extends Error {}

  const NOT_STRAVA = "Dit is geen Strava-bestand. Sleep het bestand activities.csv hierheen.";

  /** Decode the file as UTF-8, falling back to Windows-1252. */
  function decodeText(bytes) {
    try {
      return { text: new TextDecoder("utf-8", { fatal: true }).decode(bytes), encoding: "utf-8" };
    } catch (e) {
      return { text: new TextDecoder("windows-1252").decode(bytes), encoding: "windows-1252" };
    }
  }

  /** Parse a Strava date as an Excel serial date number, or return null. */
  function parseDate(text) {
    let day, month, year, hour, minute, second, ampm;
    let m = DATE_NL_RE.exec(text);
    if (m) {
      [, day, month, year, hour, minute, second] = m;
    } else if ((m = DATE_EN_RE.exec(text))) {
      [, month, day, year, hour, minute, second, ampm] = m;
    } else {
      return null;
    }
    const monthIndex = MONTHS[month.toLowerCase()];
    if (monthIndex === undefined) return null;
    hour = Number(hour);
    if (ampm) hour = (hour % 12) + (ampm.toUpperCase() === "PM" ? 12 : 0);
    const time = Date.UTC(Number(year), monthIndex, Number(day), hour, Number(minute), Number(second));
    const check = new Date(time);
    if (check.getUTCDate() !== Number(day) || check.getUTCMonth() !== monthIndex) return null;
    return (time - EXCEL_EPOCH) / 86400000;
  }

  /** Decide per column whether it holds numbers, dates or text. */
  function columnTypes(records, width) {
    const types = [];
    for (let j = 0; j < width; j++) {
      const values = records.map((r) => r[j]).filter((v) => v !== "");
      if (j === DISPLAY_DISTANCE || values.length === 0) types.push("text");
      else if (values.every((v) => NUMBER_RE.test(v))) types.push("number");
      else if (values.every((v) => parseDate(v) !== null)) types.push("date");
      else types.push("text");
    }
    return types;
  }

  function cleanText(text, warnings, where) {
    // Excel uses a plain "\n" for a line break inside a cell.
    text = text.replace(/\r\n?/g, "\n").replace(ILLEGAL_CHARACTERS_RE, "");
    if (text.length > MAX_CELL_LENGTH) {
      warnings.push(`${where}: tekst ingekort tot ${MAX_CELL_LENGTH} tekens (maximum van Excel).`);
      text = text.slice(0, MAX_CELL_LENGTH);
    }
    return text;
  }

  function toCell(text, type, warnings, where) {
    if (text === "") return null;
    if (type === "number") return { t: "n", v: Number(text) };
    if (type === "date") return { t: "n", v: parseDate(text), z: DATE_FORMAT };
    return { t: "s", v: cleanText(text, warnings, where) };
  }

  function columnWidth(header, values, type) {
    if (type === "date") return 20;
    let longest = header.length;
    for (const v of values) {
      const firstLine = v.split("\n", 1)[0];
      if (firstLine.length > longest) longest = firstLine.length;
    }
    return Math.max(8, Math.min(50, longest + 2));
  }

  /**
   * Convert a Strava activities.csv. `data` is the file contents as a Uint8Array.
   * Returns the workbook plus a summary; throws FixError for unusable input.
   */
  function convertCsv(XLSX, Papa, data) {
    // An xlsx (zip) file starts with "PK".
    if (data[0] === 0x50 && data[1] === 0x4b) throw new FixError(NOT_STRAVA);

    const { text, encoding } = decodeText(data);
    const parsed = Papa.parse(text, { delimiter: ",", quoteChar: '"', skipEmptyLines: true });
    const [header, ...rows] = parsed.data;
    if (!header || header.length < 10 || rows.length === 0 || !/^\d+$/.test(rows[0][0])) {
      throw new FixError(NOT_STRAVA);
    }

    const width = header.length;
    const good = rows.filter((r) => r.length === width);
    const types = columnTypes(good, width);
    const warnings = [];
    const problems = [];

    const sheet = {};
    header.forEach((name, c) => {
      sheet[XLSX.utils.encode_cell({ r: 0, c })] = { t: "s", v: cleanText(name, warnings, "Kolomtitel") };
    });
    rows.forEach((record, i) => {
      const r = i + 1;
      const ok = record.length === width;
      if (!ok) {
        problems.push({ row: r + 1, fields: record.length, preview: record.slice(0, 3).join(" | ") });
      }
      record.forEach((value, c) => {
        const where = `Rij ${r + 1}, kolom ${c + 1}`;
        const cell = toCell(value, ok ? types[c] : "text", warnings, where);
        if (cell) sheet[XLSX.utils.encode_cell({ r, c })] = cell;
      });
    });

    const lastColumn = Math.max(width, ...rows.map((r) => r.length)) - 1;
    const range = { s: { r: 0, c: 0 }, e: { r: rows.length, c: lastColumn } };
    sheet["!ref"] = XLSX.utils.encode_range(range);
    sheet["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: width - 1 } }) };
    sheet["!cols"] = header.map((name, c) => ({ wch: columnWidth(name, good.map((r) => r[c]), types[c]) }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "activities");
    return { workbook, activities: rows.length, width, types, encoding, problems, warnings };
  }

  return { FixError, NOT_STRAVA, decodeText, parseDate, columnTypes, convertCsv };
});
