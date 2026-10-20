# gis

English progress judgements for **gis.gcaplabs.com**: the original sheet, simulated sheets that meet the 20% Weak limit, the criteria, and tools to edit the data and build simulations.

The whole site is one self-contained page: `public/index.html`. It is served by a Cloudflare Worker using static assets (`wrangler.jsonc`), with no server code.

## Deploy with Cloudflare Workers

1. Cloudflare dashboard → **Workers & Pages** → **Create** → **Import a repository** → pick `mutvayzz-sys/gis`.
2. Build settings: leave the build command empty. Deploy command: `npx wrangler deploy`.
3. Deploy. The `routes` entry in `wrangler.jsonc` connects the Worker to `gis.gcaplabs.com` and creates the DNS record, as long as `gcaplabs.com` is on the same Cloudflare account.

Every push to `main` redeploys.

## Keep it private

The page holds real student progress data. Put it behind **Cloudflare Access**: Zero Trust → Access → Applications → Add → Self-hosted, domain `gis.gcaplabs.com`, and allow only the people or email domain who should see it. `public/_headers` also tells search engines not to index it.

## Updating the page

Replace `public/index.html` with the new version and push.
