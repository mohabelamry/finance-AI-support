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

One Microsoft Entra ID app registration can serve both the Finance and Escalation accounts:

1. Entra admin center → App registrations → New registration. Choose who can sign in (one organization, or any organization and personal Microsoft accounts).
2. Authentication → Add a platform → **Single-page application**, with redirect URI `https://<your-site>/auth/` (for local development, `http://localhost:5173/auth/`). Settings shows the exact value for the current site.
3. API permissions → Microsoft Graph → Delegated → `User.Read` and `Mail.Send`. Grant admin consent if your tenant requires it.
4. In Settings, paste the Application (client) ID. The Directory (tenant) ID is optional; enter it when the app is limited to a single organization.
5. Click Sign in with Microsoft and allow pop-ups for the site. Use an account with a licensed Exchange Online (or Outlook.com) mailbox.

Optional: set the Cloudflare build variables `VITE_MS_CLIENT_ID` (and `VITE_MS_TENANT_ID`) so the IDs are pre-filled for everyone and users only click Sign in.

Microsoft returns to `/auth/`, a small page that passes the sign-in result back to the dashboard (required by MSAL Browser v5 popups). No password or client secret is entered or saved. MSAL keeps authentication state in this browser's localStorage; Disconnect removes it. Production requires HTTPS. Allow access to login.microsoftonline.com and graph.microsoft.com. Fonts use fonts.googleapis.com and fonts.gstatic.com, with system fallbacks.

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
