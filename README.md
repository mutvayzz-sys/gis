# gis

**Progress Judgements** for gis.gcaplabs.com. People load a progress sheet (Excel, CSV or pasted rows) and the page:

1. reads the table, whatever its layout (columns can be matched by hand or by the assistant),
2. lets them set the rules (thresholds, points for phase totals, rounding, limits),
3. shows the actual results with adding-up mistakes fixed,
4. builds a clearly labelled simulated scenario with the fewest changes that meet the Weak and Outstanding limits,
5. downloads one Excel workbook: the original sheet, the corrected sheet and the simulated sheet in the same layout, a full comparison, judgement counts with both limit checks, total judgements from points, a corrections log, planning assumptions and the settings used. Percentages and judgements are formulas.

**No data is held in this repository or the page.** Sheets are read and processed in the person's browser. The only exception is the assistant: when someone asks it something, the parts of the sheet it needs are sent to Claude through `/api/chat`. Nothing is stored.

## Files

| Path | What it is |
|---|---|
| `src/page.html` | Page layout, styles and the six steps |
| `src/engine.js` | Reads tables, applies the rules, finds corrections, searches for the smallest simulation |
| `src/book.js` | Builds the Excel workbook (ExcelJS), copying the uploaded sheet's own formatting |
| `src/grid.js` | Draws a sheet on the page so it looks like Excel |
| `src/app.js` | The steps, uploads, settings, charts and downloads |
| `src/assistant.js` | The assistant: chat drawer, the page controls it may use, change log with Undo |
| `src/prompt.txt` | The assistant's instructions (used by the page and the Worker) |
| `worker/index.js` | Serves `public/` and relays `/api/chat` to Claude |
| `scripts/build.js` | Joins `src/` into `public/index.html` |

After changing anything in `src/`, run `npm run build` and commit `public/index.html` too.

## Deploy (Cloudflare Workers)

1. Workers & Pages, Create, Import a repository: `mutvayzz-sys/gis`.
2. Build command: `npm run build`. Deploy command: `npx wrangler deploy`.
3. Settings, Variables and Secrets: add a **Secret** `ANTHROPIC_API_KEY`. Without it the page still works; the assistant button stays hidden.
4. The `routes` entry in `wrangler.jsonc` connects the Worker to `gis.gcaplabs.com` and creates the DNS record.

The model is set by `MODEL` in `wrangler.jsonc`. `/api/chat` only accepts requests from the same site, caps request size and reply length, and only passes on the page's own tools. To limit who can use the assistant (and your API usage), put the site behind Cloudflare Access or add a rate-limiting rule for `/api/chat`.

Student sheets, reports and exports (`.xlsx`, `.csv`, `.pdf`) are blocked by `.gitignore` so they can't be committed by mistake.

## When the assistant can't do something

The assistant calls `/api/report`, which opens a GitHub issue labelled `assistant-gap` on this repo with a short summary (never sheet data). It needs a GitHub token:

1. GitHub → Settings → Developer settings → Fine-grained tokens → new token, repository `mutvayzz-sys/gis` only, permission **Issues: Read and write**.
2. Cloudflare → the `gis` Worker → Settings → Build → Variables and secrets → add **Secret** `GIS_GITHUB_TOKEN`.
3. Redeploy. The deploy command (`npm run deploy:ci`) copies build secrets into the Worker.

`/api/health` shows `"reports": true` once it is working. Reports are limited to 3 a minute per visitor, and assistant messages to 30 a minute.

## Attachments and Google Sheets

The assistant's paperclip takes spreadsheets (xlsx, xls, xlsb, ods, csv, Numbers), PDFs (scanned ones are sent as page pictures), Word (docx, old doc as rough text), PowerPoint (pptx), OpenDocument, RTF, HTML, text and pictures. Files are read in the visitor's browser; nothing is uploaded to this site or stored. A spreadsheet is loaded into the page; the text or pictures of the other files go to Claude with that message only.

A Google Sheets link (shared as "Anyone with the link") can be pasted on the first step or in the chat. `/api/gsheet` fetches it as .xlsx and passes it straight to the browser without storing it.

## Downloads

Step 6 offers the Excel workbook (every sheet with live formulas, a plain-language Summary sheet first, and an Assistant notes sheet when the assistant was used), a designed PDF report, and the assistant conversation as a Markdown file.

The PDF report (src/report.js) is built in the browser as one HTML document: cover with key findings, criteria, the sheet as received, corrections, the corrected sheet, tiles and bar charts, judgement counts, one landscape page per simulated option, and notes. On the site, `/api/pdf` prints it with Cloudflare Browser Run (binding `BROWSER`; the free plan allows 10 browser minutes a day), with JavaScript off and only Google Fonts allowed to load. Nothing is stored. Where that is not available (the Claude artifact), the browser's print dialog saves it as PDF. The assistant can put the same download buttons under its replies. All files are built in the browser.
