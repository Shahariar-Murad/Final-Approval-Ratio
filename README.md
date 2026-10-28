# FundedNext Approval Ratio Dashboard

A React + Vite upload-first dashboard for FundedNext approval-ratio analysis. The dashboard starts empty and calculates the analysis from the transaction file you upload.

## Included
- Executive KPI dashboard
- With Hard Decline / Without Hard Decline switch
- Full country list ranked by approved revenue
- Cascaded country approval ratio using unique `merchantOrderId`
- Attempt-level PSP approval ratio (no `merchantOrderId` deduplication)
- Top-performing PSP by country
- Hard / Soft decline taxonomy
- PSP mapping rules
- Methodology page
- Full Excel export containing Executive Summary, Top 50 Countries, PSP Wise, Hard Soft Declines, Mapping Rules, Methodology, and Report Notes
- CSV export for country and PSP tables
- Upload-first landing page with drag-and-drop CSV/XLSX/XLS transaction upload
- Client-side parsing of raw transaction reports and compatible analysis workbooks

## Netlify / GitHub
1. Push this entire folder to a GitHub repository.
2. In Netlify, choose **Add new site → Import an existing project → GitHub**.
3. Build command: `npm run build`
4. Publish directory: `dist`
5. Node version 20+ is recommended.

No server or database is required. The dashboard runs entirely in the browser.

## Important calculation logic
- Country / cascaded approval: unique `merchantOrderId`; approved wins on retry/cascade.
- PSP approval: individual PSP attempts; do NOT deduplicate `merchantOrderId`.
- PSP approval = approved attempts / (approved + declined attempts).
- Top-performing PSP by country is attempt-level, with at least 20 attempts where available.
- CardPay → Unlimit.
- Innatech → Payabl.
- Payabl-Global = Innatech, non-India.
- Payabl-India = Innatech, India.
- PayPal-Global = PayPal, non-India.
- Paypal-India = PayPal, India.

## Interactive filters
After uploading a transaction file, the left filter panel supports:
- Date Range in GMT+3 (source transaction timestamps are treated as GMT+0 and shifted +3 hours)
- PSP multi-select with Include / Exclude modes
- MID Alias multi-select with Include / Exclude modes
- Country multi-select with Include / Exclude modes
- Select All / clear selections and removable selection chips
All dashboard KPIs, revenue, approval ratios, decline analysis, country ranking, PSP analysis and Excel export reflect the active filters.


## Timezone
- Source transaction timestamps are interpreted as GMT+0 when no explicit offset is present.
- The dashboard defaults to GMT+0.
- You can switch the analysis timezone to GMT+3, GMT+6, or other listed GMT offsets before applying date filters.
- The selected timezone is recorded in the Excel export.

## Country naming
- Dashboard and Excel use full country names such as United States, India, China, Canada and United Kingdom.
- ISO alpha-2 codes remain in a separate Country Code column.

## Excel export
- The export is designed as a presentable multi-sheet workbook with title rows, subtitles, filters, formatted percentages/currency, autofilters and the full country list.
- The previous unsupported freeze metadata has been removed to avoid Excel repair warnings.

## Large-file handling
- CSV files use Papa Parse streaming + Web Worker mode, so large CSV reports can be processed without loading the entire file into an XLSX workbook object.
- Large Excel workbooks (100 MB+) are intentionally rejected in the browser because XLSX parsing requires the workbook to be buffered in memory. Convert very large Excel reports to CSV for streaming mode.
