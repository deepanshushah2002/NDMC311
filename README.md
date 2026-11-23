# NDMC 311 Complaint Pendency Dashboard

A browser-based dashboard for turning an NDMC 311 complaint Excel export into department-wise and officer-wise pendency reports. It summarizes complaint aging, SLA status, and escalations, and exports reports as PDF and Excel files.

## Features

- Import `.xlsx` and `.xls` complaint workbooks and choose a worksheet.
- Map complaint date, department, assigned officer, and escalation columns.
- Set the report as-of date used to calculate complaint age.
- View total pending, within-SLA, and escalated counts.
- Group aging into 0–15, 16–30, 31–45, 46–90, and more than 90 days.
- Generate a department dashboard and officer-wise aging report.
- Highlight departments with complaints older than 90 days.
- Export dashboard and officer reports to PDF and Excel, including an officer pending-list workbook.
- Switch between light and dark themes.

## Run locally

No build step or package installation is required. The app loads its JavaScript libraries from public CDNs, so an internet connection is needed when opening it.

1. Download or clone this repository.
2. Open `index.html` in a modern browser.
3. Select or drag in an NDMC 311 `.xlsx` or `.xls` export.
4. Confirm the detected columns, choose the report date and worksheet, then select **Generate Report**.
5. Use the export buttons to download PDFs or Excel workbooks.

Alternatively, serve the folder locally:

```bash
python -m http.server 8000
```

Then open [http://localhost:8000](http://localhost:8000).

## Expected input columns

These columns are required (the app tries to detect them automatically):

| Field | Expected header |
| --- | --- |
| Complaint date | `Date` or `Complaint Date` |
| Department | `Department` |
| Assigned officer | `Currently Assigned To` |
| Escalation flag | `Is Escalated` |

The `Assigned To Mobile` (or `Mobile`) and `Current Status` columns are optional. The status field is used by the officer pivot workbook. If automatic detection chooses the wrong column, select the correct header in the mapping table before generating the report.

The source worksheet should contain the pending complaints to be reported. Every data row is counted as one complaint; the app uses the escalation flag to split within-SLA and escalated counts. It does not remove rows based on the status value.

## Reports and exports

- **Dashboard PDF** — overall and department-wise totals, SLA, escalation, and aging.
- **Officer-wise PDF** — complaint counts and aging by department and assigned officer.
- **Excel report** — `Dashboard`, `Officers Wise`, and `Source Data` sheets.
- **Officer pending-list Excel** — department/status summary and officer-wise pending list.

Age is calculated from the complaint date to the selected report date. Missing or unreadable dates are counted as 0 days and flagged in the dashboard checks. Complaints dated after the report date are also aged as 0 days and flagged. Aging ranges include the boundary day (for example, 15 days is in `0–15`, and 90 days is in `46–90`).

## Pricing and payment links

Plan names, prices, and QR image paths are configured in `promo.js`. Current listed prices are ₹199 for one month, ₹999 for six months, and ₹1,999 for one year. Choosing a plan opens `payment.html` with the matching QR image from `payment-assets/`. Payments are verified manually by an administrator; this project does not process or automatically verify payments.

## Project files

| File | Purpose |
| --- | --- |
| `index.html` | Main upload, summary, dashboard, and export interface |
| `pricing.html` | Pricing page and animated plan prices |
| `payment.html` | Selected plan summary and matching UPI QR code |
| `payment-assets/` | QR images for the monthly, 6-month, and yearly plans |
| `script.js` | Workbook parsing, calculations, dashboard rendering, and exports |
| `promo.js` | Shared pricing data and plan-selection popup |
| `styles.css` | Shared layout and visual styles |

## Publish with GitHub Pages

1. Push these project files to a GitHub repository.
2. In the repository, open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select the branch and `/ (root)` folder, then save.
5. Open the URL GitHub Pages provides. The main app is `index.html`; pricing is at `pricing.html`.

The workbook is read in the browser, and this project does not configure an application backend. The page does load third-party libraries and fonts from external CDNs.

## Technology

Plain HTML, CSS, and JavaScript, with SheetJS for reading workbooks, ExcelJS for Excel exports, jsPDF for PDFs, and html2canvas for rendering report layouts.
