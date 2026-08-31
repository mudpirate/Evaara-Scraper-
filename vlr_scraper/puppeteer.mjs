/**
 * vlr.gg scraper — Puppeteer version
 *
 * Usage:
 *   node puppeteer.mjs --mode results --pages 3
 *   node puppeteer.mjs --mode matches
 *   node puppeteer.mjs --mode results --headed --delay 3000 --out out.json --csv out.csv
 *
 * Modes:
 *   results  — completed match results (paginated)
 *   matches  — upcoming & live match schedule
 */
import puppeteer from "puppeteer";
import fs from "node:fs/promises";
import { parseMatches, toCsvRows, parseArgs } from "./parseMatches.js";

const args = parseArgs(process.argv.slice(2));
const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function pageUrl(mode, page) {
  if (mode === "matches") return "https://www.vlr.gg/matches";
  // results supports pagination via ?page=N
  return page > 1
    ? `https://www.vlr.gg/matches/results?page=${page}`
    : "https://www.vlr.gg/matches/results";
}

async function main() {
  console.log(`[puppeteer] mode=${args.mode} pages=${args.pages} delay=${args.delay}ms`);

  const browser = await puppeteer.launch({
    headless: args.headed ? false : "new",
    args: ["--no-sandbox", "--disable-blink-features=AutomationControlled"],
  });

  const all = [];
  try {
    const page = await browser.newPage();
    await page.setUserAgent(UA);
    await page.setExtraHTTPHeaders({ "Accept-Language": "en-US,en;q=0.9" });
    await page.setViewport({ width: 1366, height: 900 });

    for (let p = 1; p <= args.pages; p++) {
      const url = pageUrl(args.mode, p);
      console.log(`[puppeteer] fetching ${url}`);
      const resp = await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
      if (!resp || !resp.ok()) {
        console.warn(`[puppeteer] HTTP ${resp ? resp.status() : "?"} for ${url} — stopping.`);
        break;
      }

      // Wait for match items to render
      try {
        await page.waitForSelector("a.wf-module-item.match-item", { timeout: 15000 });
      } catch {
        console.warn(`[puppeteer] no match items found on ${url} — stopping.`);
        break;
      }
      // Give client-side hydration a moment
      await sleep(500);

      const matches = await page.evaluate(parseMatches);
      console.log(`[puppeteer] page ${p}: parsed ${matches.length} matches`);
      all.push(...matches);

      if (p < args.pages) await sleep(args.delay);
    }
  } finally {
    await browser.close();
  }

  await fs.writeFile(args.out, JSON.stringify(all, null, 2), "utf8");
  await fs.writeFile(args.csv, toCsvRows(all), "utf8");
  console.log(`[puppeteer] saved ${all.length} matches -> ${args.out} & ${args.csv}`);
}

main().catch((err) => {
  console.error("[puppeteer] failed:", err);
  process.exit(1);
});
