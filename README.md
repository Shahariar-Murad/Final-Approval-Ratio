# FundedNext Approval Ratio Dashboard

A React + Vite dashboard for the FundedNext approval-ratio analysis.

## Included
- Executive KPI dashboard
- With Hard Decline / Without Hard Decline switch
- Top 50 countries
- Cascaded country approval ratio using unique `merchantOrderId`
- Attempt-level PSP approval ratio (no `merchantOrderId` deduplication)
- Top-performing PSP by country
- Hard / Soft decline taxonomy
- PSP mapping rules
- Methodology page
- Full Excel export containing Executive Summary, Top 50 Countries, PSP Wise, Hard Soft Declines, Mapping Rules, Methodology, and Report Notes
- CSV export for country and PSP tables
- Client-side Excel upload for future analysis workbooks with the same sheet structure

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
