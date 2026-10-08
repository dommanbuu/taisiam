# taisiam
taisaim website for manage finance and dashboard monitoring gas station of Taisiam 


## Local prototype

Application source and setup instructions are in [prototype/README.md](prototype/README.md).

```powershell
cd prototype
node server.cjs --port 4174
```

Open http://127.0.0.1:4174/#reports for branch accounting.

The Excel integration currently requires the local Python and Node dependencies described in the prototype README. Cloudflare static frontend deployment is configured; the local accounting API and Excel runtime have not yet been migrated. Authentication and role selection are still a prototype.

Private accounting records, source Excel files, generated reports, local screenshots, credentials and dependencies are excluded from this repository.


## Cloudflare Workers deployment (frontend preview)

Use the repository root as the Cloudflare Workers build root.

- Build command: leave empty (the HTML/CSS/JS are already in `prototype/dist`).
- Deploy command: `npm run deploy` (or `npx wrangler deploy`).
- Worker name: `taisiam` (must match `wrangler.jsonc`; adjust if your Cloudflare project uses another name).
- Dependency install: `npm ci`.

`wrangler.jsonc` explicitly sets `assets.directory` to `./prototype/dist`; no framework auto-detection is needed. `npm run deploy:check` validates packaging without publishing.

For Cloudflare Pages instead: framework None, root repository root, no build command, build output directory `prototype/dist`.

This only publishes the frontend prototype. `/api/excel/*`, Python processing and persistent local accounting files are NOT deployed. The accounting page displays an unavailable message until a cloud backend is implemented. Do not upload `.local-data`, source spreadsheets or the entire repository as public static assets.
