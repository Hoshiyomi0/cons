# Shopee Order Processing Website

A single-user cloud web application built around the existing Shopee Order Consolidation Apps Script logic (v10). The app adds a browser upload workflow and a copy/paste results table; it does not require Google Sheets.

## Current workflow

1. Sign in with the private app username/password.
2. Upload the Inventory Database Excel/CSV once. The latest upload replaces the saved database.
3. Upload each Shopee order export.
4. Review the consolidated table in the browser.
5. Use **Copy All** or **Copy Matched Only**, then paste into Google Sheets or Excel. Clipboard data is TSV, so fields split into columns. The existing blank spacer column is preserved.

Order uploads are held in memory during processing and are not saved to disk. The Inventory Database is saved as normalized JSON on the persistent data disk.

## Preserved and added business rules

- Uses the existing v10 product resolution, matching, categorization, and order grouping logic.
- Preserves the Pokémon Scarlet base-game/DLC validation. A `Scarlet [US]` variation plus `Switch Pokemon Scarlet` SKU reference must resolve to the base game, even if the parent listing title mentions *The Hidden Treasure of Area Zero*.
- Fujifilm film-only consumables map to the existing `Xbox Gaming Acc` output category. Fujifilm cameras, printers, cases, albums, stands, bags, and other hardware are excluded from this override.
- Adds a `Matched SKU` output column populated from the matched Inventory `SKU / Item Code`. It remains blank when there is no reliable match.
- Keeps the existing blank spacer column between Matched SKU and Notes.

## Input format

### Order file

- `.xlsx`, `.xls`, or `.csv`; maximum 25 MB.
- The app uses a worksheet named `Raw` when present, otherwise the first worksheet.
- Required raw layout: column A = Tracking Number, column B = Order ID, column C = Item Data.
- The item data text should retain the Shopee export format, including `Product Name:`, `Variation Name:`, `Price:`, `Quantity:`, and `SKU Reference No.:` labels.

### Inventory file

- `.xlsx`, `.xls`, or `.csv`; maximum 25 MB.
- Expected columns: `SKU / Item Code`, `Product Name`, `Category (ID or Name)`, `Tags`.
- The first worksheet is used. The app also recognizes common alternate header labels for SKU and product name.
- Uploading a new Inventory file replaces the previous saved database.

## Run locally

Requires Node.js 20 or newer.

```bash
cp .env.example .env
# Edit .env and set APP_USER and APP_PASSWORD.
npm install
npm start
```

Open `http://localhost:3000`. The browser will request the Basic Auth username/password.

Run the core regression tests:

```bash
npm test
```

The regression tests use a small anonymized Inventory fixture in `test/fixtures/`. The actual cloud app does not depend on this test fixture; it uses the Inventory file uploaded through the website.

## Cloud deployment on Render

This project includes `render.yaml` for a Node web service with a persistent disk mounted at `/var/data`.

1. Create a private Git repository and upload the contents of this project folder.
2. In Render, choose **New → Blueprint** and connect the repository.
3. Render reads `render.yaml`; select the paid Starter instance with the 1 GB persistent disk.
4. Set the `APP_USER` and `APP_PASSWORD` environment variables in the Render dashboard. Use a unique username and a long random password.
5. Deploy. Open the generated HTTPS URL and sign in.
6. Upload the Inventory Database file in the website before processing orders.

The app requires persistent storage because the Inventory Database must survive deploys and restarts. Do not change `DATA_DIR` from `/var/data` on Render unless the persistent disk mount is changed too.

## Security and operational notes

- The app uses HTTP Basic Authentication over HTTPS, suitable for a private single-user instance. Do not deploy without setting both credentials.
- The app does not retain order uploads; it only retains the Inventory Database.
- Uploaded files are limited to 25 MB and restricted to Excel/CSV extensions. Treat the website URL and credentials as private.
- This starter app does not yet provide password reset, multi-user accounts, audit logs, or automated backups. Back up the Inventory Database source file separately.
- Before using it for business-critical processing, test with representative order exports and compare output against the current Google Sheets workflow.

## Project structure

- `server.js` — authenticated web server and upload endpoints.
- `src/engine.js` — lightweight in-memory SpreadsheetApp compatibility layer.
- `src/shopee-v10-base.js` — v10 business logic with the Fujifilm and Matched SKU additions.
- `public/` — upload UI, results table, TSV copy actions.
- `test/` — regression tests for the critical matching/category behavior.
