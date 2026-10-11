# InvoiceFlow

A responsive React dashboard for Excel invoices and two Microsoft 365 accounts.

## Run

Requires Node.js 22.12+ (validated on 24.19).

```sh
npm ci --cache /tmp/finance-npm-cache
npm run dev -- --port 5173
```

`npm run build` creates `dist/`; `npm run preview` serves the build. Deploy `dist/` to an HTTPS static host for Microsoft SSO. Localhost can be used for development.

## Features

- Excel/CSV import, downloadable template, invoice search and status filters.
- Paid, due and overdue status; overdue days use the Asia/Riyadh calendar date.
- Currency-specific totals (no conversion or summing different currencies).
- Mark paid/unpaid, reminders, escalations, editable email preview, dashboard action history.
- Two separately configured Microsoft 365 tenant/application connections using Microsoft MSAL popup SSO and Graph delegated Mail.Send.

## Excel format

The first worksheet is imported. Columns: `Invoice`, `Customer`, `Amount`, `Currency`, `Due date`, `Status`, `Email`.

Invoice, Customer, Amount and Due date are required. Dates must be actual Excel date cells or YYYY-MM-DD text. Amounts must be nonnegative numbers. Currency defaults to USD; status defaults to Due. Status accepts Paid, Due, Overdue or Unpaid; due/overdue is calculated from the due date. Invoice IDs must be unique. Email is optional for import but required for sending. Files are limited to 10 MB. Import replaces the invoice list only after all rows validate. History is retained by invoice ID, so use stable IDs across imports.

Invoices, connection identifiers and action history are stored in localStorage on this browser. Clearing browser data removes them. This version has no shared database, server-side scheduler or cross-device synchronization. Demo invoice email sending is disabled. History records dashboard actions, not mailbox conversations. Microsoft accepting a message does not prove delivery.

## Microsoft 365 setup

For each tenant, your administrator must register a Microsoft Entra ID application:

1. Add a **Single-page application** redirect URI matching the exact site origin (for local development, http://localhost:5173).
2. Configure delegated Graph `User.Read` and `Mail.Send` permissions. Grant consent as required by tenant policy.
3. Enter the Directory (tenant) ID and Application (client) ID in Settings.
4. Sign in with Microsoft using an account with a licensed Exchange Online mailbox. The first connection defaults to reminders, the second to escalations; the sender can be changed in the preview.

No password or client secret is entered or saved. MSAL stores authentication state in sessionStorage; expired sessions may require signing in again. Production requires HTTPS. Allow access to login.microsoftonline.com and graph.microsoft.com. Fonts use fonts.googleapis.com and fonts.gstatic.com, with system fallbacks.

Real sign-in and sending require your tenant configuration and have not been validated against a real mailbox. The UI does not claim a connection until Microsoft sign-in succeeds.

## Validation

```sh
npm test
npm run build
npm audit --audit-level=high
# With the dev server running and Chromium installed:
node tests/browser.mjs
```

`CHROMIUM_PATH` overrides /usr/bin/chromium; `APP_URL` overrides http://localhost:5173. Browser tests use a separate temporary browser profile. They exercise search, filters, demo email protection, payment/history persistence, both connection cards, template download and actual Excel UI import, and mobile layout.

The npm package named xlsx is aliased to @e965/xlsx 0.20.3, a repackaging of SheetJS, to avoid the known vulnerabilities in the old npm xlsx 0.18.5 release.
