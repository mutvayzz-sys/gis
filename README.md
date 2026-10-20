# gis

Progress judgement checker for **gis.gcaplabs.com**. Load a progress sheet (Excel, CSV or pasted rows) and the page shows it as a spreadsheet, fixes adding-up mistakes, checks the judgements against the 20% Weak and 20% Outstanding limits, and builds simulated sheets with the fewest changes that would fit. Formatted Excel workbooks are built in the browser for download.

**No data is held in this repository or the page.** Everything a person loads stays in their own browser tab: nothing is uploaded, sent or saved, and closing the tab clears it.

The whole site is one self-contained page: `public/index.html`. It is served by a Cloudflare Worker using static assets (`wrangler.jsonc`), with no server code.

## Deploy with Cloudflare Workers

1. Cloudflare dashboard → **Workers & Pages** → **Create** → **Import a repository** → pick `mutvayzz-sys/gis`.
2. Build settings: leave the build command empty. Deploy command: `npx wrangler deploy`.
3. Deploy. The `routes` entry in `wrangler.jsonc` connects the Worker to `gis.gcaplabs.com` and creates the DNS record, as long as `gcaplabs.com` is on the same Cloudflare account.

Every push to `main` redeploys.

## Access

The page holds no student data, so it can be public. If you want only staff to use it, put it behind **Cloudflare Access**: Zero Trust → Access → Applications → Add → Self-hosted, domain `gis.gcaplabs.com`, and allow only the people or email domain who should see it. `public/_headers` also tells search engines not to index it.

## Updating the page

Replace `public/index.html` with the new version and push.
