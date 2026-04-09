# school-project

Web viewer for the 25/26 college admission Excel analyzers.

## What it does

- Upload a filled `25수능` or `26수능` analyzer workbook
- Parse the workbook in the browser with SheetJS
- Show the result sheets as a searchable, filterable web UI
- Compare matched programs between the 25 and 26 analyzers when both files are loaded

## Current flow

1. Open the original analyzer in Excel or the Excel mobile app
2. Enter scores and save the file
3. Upload the saved workbook to this web app
4. Browse `이과` / `문과` results, inspect thresholds, and compare years

## Important files

- `index.html`: analyzer web UI
- `app.js`: workbook parser, filters, result explorer, compare view
- `vendor/xlsx.full.min.js`: browser bundle used to read `.xlsx` and `.xlsb`
- `wrangler.toml`: Cloudflare Pages configuration

## Local preview

- `npx wrangler pages dev .`

## Notes

- The app reads the calculated values already stored in the uploaded workbook.
- It does not modify the original Excel file.
