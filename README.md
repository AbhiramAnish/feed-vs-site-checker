# Link / Price / Brand Checker

Verifies that for every row in your Excel feed, the live product page shows
the same **id (SKU)**, **price**, and **brand** (only when the brand column
is filled in for that row).

This uses a real headless Chromium browser via Playwright, so it behaves
like an actual visitor (this matters — some sites block plain HTTP requests
that don't look like a real browser).

## Setup (one-time)

```bash
npm install
npx playwright install chromium
```

The second command downloads the actual browser Playwright drives — it's
separate from `npm install` and only needs to be run once.

## Run

```bash
node check-links.js path/to/feed.xlsx
```

Optional: control how many pages run in parallel (default 8). Lower this if
the site starts blocking/rate-limiting you; raise it (carefully) for more
speed:

```bash
node check-links.js path/to/feed.xlsx --concurrency 5
```

## What it does

- Reads columns `id`, `price`, `link`, `availability`, `brand` from the first sheet.
- **Skips out-of-stock rows entirely** — only rows where `availability` is
  "in stock" (case-insensitive) are visited and checked. Out-of-stock rows
  are not navigated to, not counted as checked, and left untouched in the
  corrected Excel (see below).
- Visits every remaining (in-stock) `link` using a real browser page.
- Does **not** stop on a failed/mismatched row — keeps going through all rows.
- Prints only the rows with problems to the terminal, as they're found:
  - Broken link (non-200 response) or navigation failure/timeout
  - SKU/price/brand element missing on the page
  - ID mismatch / Price mismatch / Brand mismatch (brand only checked if
    the sheet has a brand value for that row)
- Prints a progress line every 100 rows.
- At the end prints a summary:
  ```
  ========== SUMMARY ==========
  Total rows in sheet:     37115
  Skipped (out of stock):  4200
  Checked (in stock):      32915
  Mismatches found:        12
  Matched OK:              32903
  Rows removed (broken):   3
  IDs corrected:           3
  Prices corrected:        9
  Availability corrected:  6
  ==============================
  ```

- Also saves full mismatch details to `mismatch-report.json` and
  `mismatch-report.csv` in this folder, as a backup in case terminal
  scrollback isn't enough for a sheet this size.
- **Generates a corrected Excel file**: `<yourfile>-corrected.xlsx`, an exact
  copy of your input, except:
  - Rows with a broken link (404, other bad HTTP status, or a timed-out/
    failed navigation) are **removed entirely** from this file.
  - For every other checked (in-stock) row: `id` is updated if the site's
    SKU differs from the sheet, and `price` is updated if the site's price
    differs (same "144 QAR" style format as the original cell).
  - `availability` is set to `"out of stock"` **only** when the price
    element specifically can't be found on the page (a missing SKU alone
    does not trigger this).
  Rows that were already out of stock in your original sheet are skipped
  entirely and left completely untouched.

## Tuning

- If you still see false "Broken link" results for pages that work fine in
  your own browser, the site may have stronger bot protection (e.g.
  Cloudflare challenge). Try lowering `--concurrency` first — most blocking
  is triggered by request speed/volume, not detection of Playwright itself.
- Timeout per page and retry count are set near the top of
  `check-links.js` (`NAV_TIMEOUT_MS`, `MAX_RETRIES`) if you want to adjust
  them.
