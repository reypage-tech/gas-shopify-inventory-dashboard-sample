'use strict';
/*
 * Builds docs/local_test_output.html and docs/local_test_output.png from the
 * actual output of the Node logic tests (fictional sample data, no live
 * Shopify store). Run with: node scripts/generate_docs_image.js
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadSamplePayloads } = require('../test/sample_data');

const root = path.join(__dirname, '..');
const docsDir = path.join(root, 'docs');

const code = fs.readFileSync(path.join(root, 'Code.gs'), 'utf8');
const mod = { exports: {} };
new Function('module', 'exports', code)(mod, mod.exports);
const g = mod.exports;

const payloads = loadSamplePayloads();
const goals = [
  { storeId: 7101001, storeName: 'Riverlane Downtown (sample)', revenueGoal: 500, lowStockThreshold: 5 },
  { storeId: 7101002, storeName: 'Riverlane Westside (sample)', revenueGoal: 500, lowStockThreshold: 5 }
];
const reportDate = '2026-10-29';

const model = g.buildSyncModel_(
  {
    orders: payloads.orders,
    locations: payloads.locations,
    products: payloads.products,
    inventory: payloads.inventory,
    assignees: payloads.assignees
  },
  goals,
  reportDate
);

const sentEmails = [];
global.MailApp = { sendEmail: m => sentEmails.push(m) };
g.sendLowStockAlerts_(model.lowStockByAssignee);

const testOutput = execFileSync('node', ['test/run.js'], { cwd: root, encoding: 'utf8' }).trim();

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function money(n) {
  return '$' + Number(n).toFixed(2);
}

const salesRows = model.dailySalesRows
  .slice()
  .sort((a, b) => a.storeName.localeCompare(b.storeName))
  .map(r => `
    <tr>
      <td>${esc(r.storeName)}</td>
      <td class="num">${money(r.revenue)}</td>
      <td class="num">${r.orderCount}</td>
      <td class="num">${money(r.revenueGoal)}</td>
      <td class="${r.meetsGoal ? 'ok' : 'bad'}">${r.meetsGoal ? 'meets goal' : 'below goal'}</td>
    </tr>`)
  .join('');

const invRows = model.inventoryRows
  .slice()
  .sort((a, b) => a.storeName.localeCompare(b.storeName) || a.sku.localeCompare(b.sku))
  .map(r => `
    <tr>
      <td>${esc(r.storeName)}</td>
      <td>${esc(r.sku)}</td>
      <td>${esc(r.productTitle)}</td>
      <td class="num">${r.qty}</td>
      <td class="num">${r.threshold}</td>
      <td class="${r.isLow ? 'bad' : 'ok'}">${r.isLow ? 'low stock' : 'ok'}</td>
    </tr>`)
  .join('');

const alertBlocks = sentEmails
  .map(m => `
    <div class="email">
      <div class="email-head">To: ${esc(m.to)} &middot; Subject: ${esc(m.subject)}</div>
      <pre>${esc(m.body)}</pre>
    </div>`)
  .join('');

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 28px;
    width: 904px;
    font-family: -apple-system, "Segoe UI", Helvetica, Arial, sans-serif;
    color: #1b1f24;
    background: #ffffff;
  }
  h1 { font-size: 18px; margin: 0 0 2px; }
  .sub { color: #57606a; font-size: 13px; margin: 0 0 18px; }
  .cmdline {
    background: #0d1117;
    color: #c9d1d9;
    font-family: "SFMono-Regular", Consolas, Menlo, monospace;
    font-size: 12.5px;
    padding: 10px 14px;
    border-radius: 6px;
    margin: 0 0 22px;
  }
  .cmdline .prompt { color: #7ee787; }
  h2 { font-size: 14px; margin: 0 0 8px; border-bottom: 1px solid #d0d7de; padding-bottom: 6px; }
  section { margin-bottom: 20px; }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th, td { text-align: left; padding: 5px 8px; border-bottom: 1px solid #eaeef2; }
  th { color: #57606a; font-weight: 600; }
  td.num { text-align: right; font-family: Consolas, Menlo, monospace; }
  td.ok { color: #1a7f37; }
  td.bad { color: #cf222e; font-weight: 600; }
  .email { border: 1px solid #d0d7de; border-radius: 6px; margin-bottom: 10px; }
  .email-head { background: #f6f8fa; padding: 6px 10px; font-size: 12px; border-bottom: 1px solid #d0d7de; }
  .email pre { margin: 0; padding: 10px; font-size: 12px; white-space: pre-wrap; font-family: Consolas, Menlo, monospace; }
</style>
</head>
<body>
  <h1>Local Node test output (fictional sample data)</h1>
  <p class="sub">Report date ${esc(reportDate)} &middot; no live Shopify store or Google account involved</p>
  <div class="cmdline"><span class="prompt">$</span> node test/run.js<br>${esc(testOutput)}</div>

  <section>
    <h2>Daily sales by store</h2>
    <table>
      <thead><tr><th>Store</th><th class="num">Revenue</th><th class="num">Orders</th><th class="num">Goal</th><th>Status</th></tr></thead>
      <tbody>${salesRows}</tbody>
    </table>
  </section>

  <section>
    <h2>Inventory</h2>
    <table>
      <thead><tr><th>Store</th><th>SKU</th><th>Product</th><th class="num">On hand</th><th class="num">Threshold</th><th>Status</th></tr></thead>
      <tbody>${invRows}</tbody>
    </table>
  </section>

  <section>
    <h2>Low-stock alert emails (mocked MailApp, one per assignee)</h2>
    ${alertBlocks}
  </section>
</body>
</html>`;

fs.mkdirSync(docsDir, { recursive: true });
const htmlPath = path.join(docsDir, 'local_test_output.html');
const pngPath = path.join(docsDir, 'local_test_output.png');
fs.writeFileSync(htmlPath, html);

execFileSync('google-chrome', [
  '--headless',
  '--disable-gpu',
  '--no-sandbox',
  '--hide-scrollbars',
  '--screenshot=' + pngPath,
  '--window-size=960,860',
  'file://' + htmlPath
], { stdio: 'inherit' });

console.log('Wrote ' + htmlPath + ' and ' + pngPath);
