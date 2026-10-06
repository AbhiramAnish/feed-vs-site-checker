/**
 * check-links.js
 *
 * Reads an Excel feed (columns: id, price, link, brand [sometimes empty])
 * and verifies, for every row, that the live page at `link` shows the
 * same id (SKU), price, and brand (only when brand is present in the sheet).
 *
 * Page locators are loaded from testdatalocators.json (kept out of git).
 *
 * Setup (one-time):
 *   npm install
 *   npx playwright install chromium
 *
 * Usage:
 *   node check-links.js path/to/feed.xlsx
 *   node check-links.js path/to/feed.xlsx --concurrency 8
 *
 * Only rows with problems are printed to the terminal as they're found.
 * A running progress line is printed every 100 rows.
 * At the end, a summary count is printed, and full mismatch details are
 * saved to mismatch-report.json / mismatch-report.csv next to this script.
 */

const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ---------- CLI args ----------
const args = process.argv.slice(2);
const EXCEL_PATH = args[0];
const concurrencyFlagIdx = args.indexOf('--concurrency');
const CONCURRENCY = concurrencyFlagIdx !== -1 ? parseInt(args[concurrencyFlagIdx + 1], 10) : 8;
const NAV_TIMEOUT_MS = 30000;
const MAX_RETRIES = 1; // 1 retry on navigation failure/timeout

if (!EXCEL_PATH) {
  console.error('Usage: node check-links.js path/to/feed.xlsx [--concurrency N]');
  process.exit(1);
}

// ---------- Locators ----------
const LOCATORS = require('./testdatalocators.json');

function parseNumericPrice(val) {
  if (val === null || val === undefined || val === '') return null;
  const match = String(val).replace(/,/g, '').match(/[\d.]+/);
  return match ? parseFloat(match[0]) : null;
}

async function extractFromPage(page) {
  // SKU
  const skuText = (await page.locator(LOCATORS.sku).first().textContent().catch(() => null)) || '';
  const skuMatch = skuText.match(/(\d+)/);
  const foundId = skuMatch ? skuMatch[1] : null;

  // Price — only the element's direct text nodes, ignoring child elements.
  const priceText = await page
    .locator(LOCATORS.price)
    .first()
    .evaluate((el) => {
      let text = '';
      el.childNodes.forEach((node) => {
        if (node.nodeType === 3) text += node.textContent; // TEXT_NODE
      });
      return text.trim();
    })
    .catch(() => '');
  const priceMatch = priceText.match(/([\d,.]+)/);
  const foundPrice = priceMatch ? parseFloat(priceMatch[1].replace(/,/g, '')) : null;

  // Brand
  const foundBrand = ((await page.locator(LOCATORS.brand).first().textContent().catch(() => null)) || '').trim() || null;

  return {
    foundId,
    skuFound: skuText.trim().length > 0,
    foundPrice,
    priceFound: priceText.trim().length > 0,
    foundBrand,
  };
}

function formatPriceLikeOriginal(originalStr, newNumber) {
  // Keep whatever suffix/prefix text surrounded the number in the original
  // cell (e.g. a currency code). Falls back to a plain number if the
  // original didn't match the expected "<number><text>" shape.
  const str = String(originalStr);
  const match = str.match(/^(\s*)([\d.,]+)(.*)$/);
  if (!match) return String(newNumber);
  const [, lead, , trail] = match;
  return `${lead}${newNumber}${trail}`;
}

async function processRow(page, row, idx) {
  const expectedId = String(row.id).trim();
  const expectedPrice = parseNumericPrice(row.price);
  const rawBrand = row.brand !== undefined ? String(row.brand).trim() : '';
  const expectedBrand = rawBrand.length > 0 ? rawBrand.toLowerCase() : null; // "sometimes" check
  const link = String(row.link).trim();
  const availability = row.availability !== undefined ? String(row.availability).trim().toLowerCase() : '';

  const issues = [];
  let correctedId = null; // set when the site's id differs and was found
  let correctedPrice = null; // set when the site's price differs and was found
  let correctedAvailability = null; // set to 'out of stock' when the price element is missing
  let remove = false; // set true when the row should be dropped from the corrected Excel entirely

  const result = () => ({
    row: idx + 2,
    id: expectedId,
    link,
    issues,
    skipped: false,
    correctedId,
    correctedPrice,
    correctedAvailability,
    remove,
  });

  if (availability === 'out of stock') {
    return { ...result(), skipped: true };
  }

  if (!link) {
    issues.push('Missing link in sheet');
    remove = true;
    return result();
  }

  let response;
  let lastErr = null;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      response = await page.goto(link, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT_MS });
      lastErr = null;
      break;
    } catch (err) {
      lastErr = err;
    }
  }

  if (lastErr) {
    issues.push(`Navigation failed: ${lastErr.message}`);
    remove = true;
    return result();
  }

  if (response && !response.ok()) {
    issues.push(`Broken link (HTTP ${response.status()})`);
    remove = true;
    return result();
  }

  const { foundId, skuFound, foundPrice, priceFound, foundBrand } = await extractFromPage(page);

  // ID check
  if (!skuFound || foundId === null) {
    issues.push('SKU/id element not found on page');
  } else if (foundId !== expectedId) {
    issues.push(`ID mismatch (sheet=${expectedId}, site=${foundId})`);
    correctedId = foundId;
  }

  // Price check
  if (!priceFound || foundPrice === null) {
    issues.push('Price element not found on page');
    correctedAvailability = 'out of stock';
  } else if (expectedPrice === null) {
    issues.push(`Sheet price unparsable (sheet="${row.price}")`);
  } else if (foundPrice !== expectedPrice) {
    issues.push(`Price mismatch (sheet=${expectedPrice}, site=${foundPrice})`);
    correctedPrice = foundPrice;
  }

  // Brand check — only when sheet has a brand value for this row
  if (expectedBrand) {
    if (!foundBrand) {
      issues.push(`Brand element not found on page (sheet brand="${rawBrand}")`);
    } else if (foundBrand.toLowerCase() !== expectedBrand) {
      issues.push(`Brand mismatch (sheet=${rawBrand}, site=${foundBrand})`);
    }
  }

  return result();
}

async function run() {
  if (!fs.existsSync(EXCEL_PATH)) {
    console.error(`File not found: ${EXCEL_PATH}`);
    process.exit(1);
  }

  const workbook = XLSX.readFile(EXCEL_PATH);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  console.log(`Loaded ${rows.length} rows from ${EXCEL_PATH}`);
  console.log(`Concurrency: ${CONCURRENCY} browser pages\n`);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
  });

  // Block heavy resources — we only need the HTML/DOM, not images/fonts/css.
  await context.route('**/*', (route) => {
    const type = route.request().resourceType();
    if (['image', 'stylesheet', 'font', 'media'].includes(type)) {
      route.abort();
    } else {
      route.continue();
    }
  });

  const queue = rows.map((row, idx) => ({ row, idx }));
  let checked = 0; // only in-stock rows that were actually checked
  let skippedCount = 0; // out-of-stock rows, not checked
  let mismatchCount = 0;
  let priceCorrectionCount = 0;
  let idCorrectionCount = 0;
  let availabilityCorrectionCount = 0;
  let removedCount = 0;
  const mismatches = [];
  const removedIndices = new Set();
  const processedCount = rows.length;

  async function worker() {
    const page = await context.newPage();
    while (queue.length) {
      const item = queue.shift();
      if (!item) break;

      const result = await processRow(page, item.row, item.idx);

      if (result.skipped) {
        skippedCount++;
      } else {
        checked++;

        if (result.issues.length) {
          mismatchCount++;
          mismatches.push(result);
          console.log(
            `[MISMATCH] row=${result.row} id=${result.id}\n  link: ${result.link}\n  -> ${result.issues.join(' | ')}\n`
          );
        }

        if (result.remove) {
          removedIndices.add(item.idx);
          removedCount++;
        } else {
          // Apply corrections to the in-memory rows array (used to build the
          // corrected Excel file at the end).
          const originalRow = rows[item.idx];

          if (result.correctedId !== null && result.correctedId !== undefined) {
            originalRow.id = result.correctedId;
            idCorrectionCount++;
          }

          if (result.correctedPrice !== null && result.correctedPrice !== undefined) {
            originalRow.price = formatPriceLikeOriginal(originalRow.price, result.correctedPrice);
            priceCorrectionCount++;
          }

          if (result.correctedAvailability !== null && result.correctedAvailability !== undefined) {
            originalRow.availability = result.correctedAvailability;
            availabilityCorrectionCount++;
          }
        }
      }

      const done = checked + skippedCount;
      if (done % 100 === 0 || done === processedCount) {
        process.stdout.write(
          `Progress: ${done}/${processedCount} processed (${checked} checked, ${skippedCount} skipped as out of stock), ${mismatchCount} mismatches so far\r`
        );
      }
    }
    await page.close();
  }

  const workers = Array.from({ length: CONCURRENCY }, () => worker());
  await Promise.all(workers);

  await browser.close();

  console.log('\n\n========== SUMMARY ==========');
  console.log(`Total rows in sheet:     ${rows.length}`);
  console.log(`Skipped (out of stock):  ${skippedCount}`);
  console.log(`Checked (in stock):      ${checked}`);
  console.log(`Mismatches found:        ${mismatchCount}`);
  console.log(`Matched OK:              ${checked - mismatchCount}`);
  console.log(`Rows removed (broken):   ${removedCount}`);
  console.log(`IDs corrected:           ${idCorrectionCount}`);
  console.log(`Prices corrected:        ${priceCorrectionCount}`);
  console.log(`Availability corrected:  ${availabilityCorrectionCount}`);
  console.log('==============================\n');

  const outDir = __dirname;
  const jsonPath = path.join(outDir, 'mismatch-report.json');
  fs.writeFileSync(jsonPath, JSON.stringify(mismatches, null, 2));

  const csvLines = ['excel_row,id,link,issues'];
  for (const m of mismatches) {
    const issuesStr = m.issues.join(' | ').replace(/"/g, '""');
    csvLines.push(`${m.row},"${m.id}","${m.link}","${issuesStr}"`);
  }
  const csvPath = path.join(outDir, 'mismatch-report.csv');
  fs.writeFileSync(csvPath, csvLines.join('\n'));

  console.log(`Full mismatch details saved to:\n  ${jsonPath}\n  ${csvPath}`);

  // Build the corrected Excel: exact copy of the input, with id/price/
  // availability updated wherever the site's data differed, and rows with
  // a broken link (404/timeout/missing link) removed entirely. Out-of-stock/
  // unchecked rows are left exactly as they were.
  const finalRows = rows.filter((_, idx) => !removedIndices.has(idx));
  const originalName = path.basename(EXCEL_PATH, path.extname(EXCEL_PATH));
  const correctedPath = path.join(outDir, `${originalName}-corrected.xlsx`);
  const newSheet = XLSX.utils.json_to_sheet(finalRows, { header: Object.keys(rows[0] || {}) });
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, workbook.SheetNames[0]);
  XLSX.writeFile(newWorkbook, correctedPath);

  console.log(
    `Corrected Excel (id/price/availability synced, ${removedCount} broken-link rows removed) saved to:\n  ${correctedPath}`
  );
}

run();