# Google Apps Script version

This version reads the live source tabs from a Google Sheets copy of the workbook each time the web app loads. It avoids hardcoding staff contacts and does not send the POC column to the browser.

## Files

- `Code.gs`: web app entry point and workbook data reader.
- `Index.html`: dashboard UI and client-side filtering, drill-down, and CSV export.

## Set it up

1. Upload `Copy of Teams & calling use cases.xlsx` to Google Drive and open/convert it as a Google Sheets file.
2. Confirm the tabs are named `Sheet1`, `Sheet3`, `Team wise`, and `vendor wise`.
3. In that spreadsheet, open **Extensions → Apps Script**.
4. Replace the default script with `Code.gs`. Add an HTML file named `Index` and paste in `Index.html`.
5. Save, then use **Deploy → New deployment → Web app**. Choose the execution and audience settings appropriate for your organization; do not expose internal cost data publicly.
6. Open the deployment URL. The dashboard reads the current sheet values when it loads.

The source workbook only has a single dated allocation snapshot. Actual invoice/payment history and call counts are not available, so the dashboard labels the three-month figure as an estimate and shows actuals/volume as unavailable.
