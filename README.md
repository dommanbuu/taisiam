# TAI SIAM accounting pilot

เว็บไซต์บัญชีสำหรับ 4 สาขา: TAI SIAM Tachileik, TAI SIAM 2, MOK SIO และ TAI SIAM Murng Pan

## Cloudflare

Worker `taisiam` serves `prototype/dist`, authenticates testers on the server, and handles `/api/excel/*`. D1 stores the accounting ledger, import reviews and private source workbook chunks. Python and the local Codex spreadsheet runtime are not needed on Cloudflare. ExcelJS runs in the Worker; importing always requires explicit review before changing the ledger.

- Install: `npm ci`
- Test and package: `npm run test:cloud`
- Apply database migrations: `npm run db:migrate`
- Set or rotate the shared tester code: `npx wrangler secret put TESTER_ACCESS_CODE`
- Deploy: `npm run deploy`
- Cloudflare Git build root: repository root. Build command: `npm run deploy:check`. Deploy command: `npm run deploy`.

The D1 binding in `wrangler.jsonc` belongs to this deployment. A different Cloudflare account needs its own D1 database ID. Never upload `.local-data` or source workbooks as public static assets. Secrets must be set with Wrangler or Cloudflare's dashboard, never committed. A missing secret closes access. Sessions last eight hours; rotating the code invalidates existing sessions. All testers with the shared code have the same accounting access. The frontend role selector is only a UI demonstration.

## What works

- Daily sales from meters, expenses, fuel purchases, and monthly sales/expense views.
- Original sales workbook import, preserved source download, explicit date review and before/after comparison.
- Eight-sheet Excel export; editing `รายการเว็บ` and re-importing updates the selected original record after confirmation, retaining its earlier version.
- Revision checks prevent concurrent users silently overwriting each other.
- JSON ledger backup. Download original workbooks separately from file history.

## Pilot limits

Tank levels and other operations pages still use mock data. The five accounting categories remain draft formats pending the remaining source workbooks. Expense/purchase editing and payments on later dates are not implemented. Actual customer receipts and other income are still entered in the exported cash sheet; that sheet is not imported back. Only `รายการเว็บ` supports round trips. Cash totals are not profit and are not presented as final when receipt/payment data is incomplete.

For this small tester pilot, the ledger is an atomic D1 JSON snapshot capped at 1.5 MB, including record history. Each parsed import has the same bound; uploads are capped at 10 MB, expanded workbook metadata at 40 MB, 10,000 rows and 100 columns. The server refuses over-limit changes without replacing stored records. Before sustained daily production usage, migrate the ledger to normalized, paginated tables and add individual accounts/review permissions. This version does not identify the person behind the shared code. Successful and unsuccessful login attempts share a five-attempt/15-minute/IP window.

## Local development

Cloud runtime: create an ignored `.dev.vars` with `TESTER_ACCESS_CODE`, then `npm run db:migrate:local` and `npm run dev:cloud` (port 8787).

Original local-only runtime: `npm run dev` (port 4174). It keeps its own files and has no shared-code authentication. Its spreadsheet prerequisites are in [prototype/README.md](prototype/README.md). Local records are never automatically copied to the Cloudflare database.

`npm run test:cloud` uses a temporary in-memory D1 database and random test credentials. An optional original workbook can be checked locally with `node cloudflare/verify.mjs path/to/workbook.xlsx` after bundling; it is not uploaded to Cloudflare. The test checks auth, secure cookies, CSRF, concurrent writes, all implemented accounting mutations, Excel round trips, source download and login throttling.

The `util` alias prevents stream dependencies writing spreadsheet contents into debug logs. Dependency overrides pin patched compatible uuid and sharp releases; test the Worker bundle after dependency changes.
