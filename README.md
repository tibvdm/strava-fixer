# strava-fixer

A small web page that converts a Strava `activities.csv` export into a clean Excel file.

**Use it here: https://tibvdm.github.io/strava-fixer/**

Drop `activities.csv` on the page and `activities.xlsx` is downloaded:

- one column per field, with descriptions that contain commas or line breaks kept intact,
- text, accents and emoji exactly as in the export (UTF-8),
- columns that only contain numbers as real numbers, the activity date as a real Excel date,
- nothing turned into a formula (`=…` or `- …` stays text).

The file is processed entirely in the browser; nothing is uploaded.

## How it works

`app.js` parses the CSV with [Papa Parse](https://www.papaparse.com/) and writes the workbook with
[SheetJS](https://sheetjs.com/). Rows that don't have as many fields as the header are listed on the page.
The first `Afstand` column is Strava's display text (km for most sports, metres for swimming) and stays text;
the second `Afstand` column holds the distance in metres as a number.

The page is plain static HTML, served by GitHub Pages from the `main` branch. The libraries are vendored in `vendor/`:

- SheetJS Community Edition 0.20.3 (Apache-2.0)
- Papa Parse 5.5.3 (MIT)

Never commit Strava exports to this repository: `.gitignore` excludes `*.xlsx` and `*.csv`.
