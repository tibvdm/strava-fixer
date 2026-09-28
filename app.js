/*
 * Repair logic for a Strava activities export that was pasted into Excel as one CSV
 * line per row (everything in column A).
 *
 * - Text fields containing newlines were split over several rows: rejoin them by
 *   parsing all lines together as CSV.
 * - UTF-8 text was misread as Windows-1252 ("okÃ©" instead of "oké"): undo that.
 * - One known corruption (lost quotes around description and distance) is repaired.
 *
 * Works in the browser (window.StravaFixer) and in Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.StravaFixer = factory();
})(this, function () {
  "use strict";

  // Windows-1252 characters in the 0x80-0x9F range, mapped back to their byte.
  const CP1252_BYTES = new Map([
    ["€", 0x80], ["‚", 0x82], ["ƒ", 0x83], ["„", 0x84], ["…", 0x85], ["†", 0x86],
    ["‡", 0x87], ["ˆ", 0x88], ["‰", 0x89], ["Š", 0x8a], ["‹", 0x8b], ["Œ", 0x8c],
    ["Ž", 0x8e], ["‘", 0x91], ["’", 0x92], ["“", 0x93], ["”", 0x94], ["•", 0x95],
    ["–", 0x96], ["—", 0x97], ["˜", 0x98], ["™", 0x99], ["š", 0x9a], ["›", 0x9b],
    ["œ", 0x9c], ["ž", 0x9e], ["Ÿ", 0x9f],
  ]);
  const UTF8 = new TextDecoder("utf-8", { fatal: true });

  const NUMBER_RE = /^-?\d+(\.\d+)?$/;
  // Control characters that are not allowed in an xlsx file.
  const ILLEGAL_CHARACTERS_RE = /[\x00-\x08\x0b\x0c\x0e-\x1f]/g;
  // Known corruption: description, elapsed time and distance merged, e.g. "\n,4458,16,03".
  const MERGED_FIELDS_RE = /^\s*,(\d+),(\d+,\d+)$/;
  const DESCRIPTION = 4;
  const ELAPSED_TIME = 5;

  /** An error with a message that can be shown to the user as is. */
  class FixError extends Error {}

  /** Read column A of the first sheet as a list of text lines. */
  function extractLines(XLSX, data) {
    let workbook;
    try {
      workbook = XLSX.read(data, { type: "array", dense: true });
    } catch (e) {
      throw new FixError("Dit bestand kan niet gelezen worden. Is het wel een Excel-bestand (.xlsx)?");
    }
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheet["!data"] || [];
    const lines = [];
    for (const row of rows) {
      const cells = row || [];
      if (cells.slice(1).some((c) => c && c.v !== undefined && c.v !== "")) {
        throw new FixError(
          "Dit bestand heeft al gegevens in meerdere kolommen. Waarschijnlijk is het al hersteld."
        );
      }
      const a = cells[0];
      lines.push(a && a.v !== undefined ? String(a.v) : "");
    }
    return lines;
  }

  /** Undo UTF-8 bytes that were decoded as Windows-1252, e.g. "okÃ©" -> "oké". */
  function fixMojibake(text) {
    const bytes = [];
    for (const ch of text) {
      const code = ch.codePointAt(0);
      if (code <= 0xff) bytes.push(code);
      else if (CP1252_BYTES.has(ch)) bytes.push(CP1252_BYTES.get(ch));
      else return text; // not mojibake: contains real non-Latin characters
    }
    try {
      return UTF8.decode(new Uint8Array(bytes));
    } catch (e) {
      return text; // already correct text, e.g. a genuine "é"
    }
  }

  function repairRecord(record, width) {
    if (record.length === width - 2) {
      const match = MERGED_FIELDS_RE.exec(record[DESCRIPTION]);
      if (match) {
        return [...record.slice(0, DESCRIPTION), "", match[1], match[2], ...record.slice(ELAPSED_TIME)];
      }
    }
    return record;
  }

  function toCellValue(text) {
    text = text.replace(ILLEGAL_CHARACTERS_RE, "");
    if (text === "") return null;
    if (NUMBER_RE.test(text)) return Number(text);
    return text;
  }

  /**
   * Repair a broken export. `data` is the file contents as an ArrayBuffer or Uint8Array.
   * Returns the repaired workbook plus a summary; throws FixError for unusable input.
   */
  function fixWorkbook(XLSX, Papa, data) {
    const lines = extractLines(XLSX, data);
    const parsed = Papa.parse(lines.join("\n"), { delimiter: ",", newline: "\n", quoteChar: '"' });
    const records = parsed.data.filter((r) => !(r.length === 1 && r[0] === ""));
    if (records.length < 2 || records[0].length < 10) {
      throw new FixError("Dit lijkt geen Strava-export te zijn: er werden geen activiteiten gevonden.");
    }

    const width = records[0].length;
    const problems = [];
    const rows = records.map((raw, i) => {
      const record = repairRecord(raw.map(fixMojibake), width);
      if (record.length !== width) {
        problems.push({ row: i + 1, fields: record.length, preview: record.slice(0, 3).join(" | ") });
      }
      return record.map(toCellValue);
    });

    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: width - 1 } }) };
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "activities");

    return { workbook, lineCount: lines.length, activities: rows.length - 1, width, problems };
  }

  return { FixError, extractLines, fixMojibake, repairRecord, toCellValue, fixWorkbook };
});
