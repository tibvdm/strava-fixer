# strava-fixer

A small web page that repairs a Strava activities export that was opened in Excel and ended up broken:

- every activity is one long comma-separated line in column A,
- activities whose description contains line breaks are split over several rows,
- special characters and emoji are garbled (`okÃ©` instead of `oké`).

**Use it here: https://tibvdm.github.io/strava-fixer/**

Drop the `.xlsx` file on the page and a repaired `<name>_fixed.xlsx` is downloaded, with one column per field.
The file is processed entirely in the browser; nothing is uploaded.

## How it works

`app.js` joins the lines back together, parses them as CSV (with [Papa Parse](https://www.papaparse.com/)),
undoes the UTF-8 → Windows-1252 mix-up and writes a new workbook (with [SheetJS](https://sheetjs.com/)).
Records that don't end up with as many fields as the header are listed on the page.

The page is plain static HTML, served by GitHub Pages from the `main` branch. The libraries are vendored in `vendor/`:

- SheetJS Community Edition 0.20.3 (Apache-2.0)
- Papa Parse 5.5.3 (MIT)

Never commit Strava exports to this repository: `.gitignore` excludes `*.xlsx` and `*.csv`.
