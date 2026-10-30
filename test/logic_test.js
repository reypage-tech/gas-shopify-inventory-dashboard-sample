const assert = require('assert');
const path = require('path');
const { loadSamplePayloads } = require('./sample_data');

const mod = { exports: {} };
const code = require('fs').readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
new Function('module', 'exports', code)(mod, mod.exports);
const g = mod.exports;

const payloads = loadSamplePayloads();
const goals = [
  { storeId: 7101001, storeName: 'Riverlane Downtown (sample)', revenueGoal: 500, lowStockThreshold: 5 },
  { storeId: 7101002, storeName: 'Riverlane Westside (sample)', revenueGoal: 500, lowStockThreshold: 5 }
];
const reportDate = '2026-10-29';

let passed = 0;
function check(label, fn) {
  fn();
  passed += 1;
}

check('orderLocalDate respects America/New_York offset sample', () => {
  assert.strictEqual(g.orderLocalDate_('2026-10-29T14:22:11-05:00', 'America/New_York'), '2026-10-29');
  assert.strictEqual(g.orderLocalDate_('2026-10-28T19:10:00-05:00', 'America/New_York'), '2026-10-28');
});

check('aggregateDailySalesByStore sums only report date', () => {
  const locations = {};
  payloads.locations.locations.forEach(l => { locations[l.id] = { id: l.id, name: l.name }; });
  const rows = g.aggregateDailySalesByStore_(payloads.orders.orders, locations, goals, reportDate, 'America/New_York');
  const downtown = rows.find(r => r.storeId === 7101001);
  const west = rows.find(r => r.storeId === 7101002);
  assert.strictEqual(downtown.revenue, 98);
  assert.strictEqual(downtown.orderCount, 2);
  assert.strictEqual(west.revenue, 78);
  assert.strictEqual(west.orderCount, 1);
  assert.strictEqual(downtown.meetsGoal, false);
});

check('buildInventoryRows flags low stock', () => {
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
  const low = model.inventoryRows.filter(r => r.isLow);
  assert.strictEqual(low.length, 2);
  assert.ok(low.some(r => r.sku === 'RH-TOWEL-01' && r.storeId === 7101002));
  assert.ok(low.some(r => r.sku === 'RH-MUG-12' && r.storeId === 7101001));
});

check('groupLowStockByAssignee merges per email', () => {
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
  const grouped = model.lowStockByAssignee;
  assert.strictEqual(Object.keys(grouped).length, 2);
  assert.ok(grouped['alex.ops.west@example.com']);
  assert.ok(grouped['jordan.ops.downtown@example.com']);
  const body = g.buildAssigneeAlertEmail_('Alex (Ops West)', grouped['alex.ops.west@example.com'].items);
  assert.ok(body.includes('RH-TOWEL-01'));
  assert.ok(body.startsWith('Hello Alex'));
});

check('applyGoalUpdate changes thresholds and rebuild affects flags', () => {
  const updatedGoals = g.applyGoalUpdate_(goals, {
    storeId: 7101002,
    storeName: 'Riverlane Westside (sample)',
    revenueGoal: 400,
    lowStockThreshold: 2,
    updatedAt: new Date(),
    updatedBy: 'test'
  });
  const westGoal = updatedGoals.find(x => x.storeId === 7101002);
  assert.strictEqual(westGoal.lowStockThreshold, 2);
  const model = g.buildSyncModel_(
    {
      orders: payloads.orders,
      locations: payloads.locations,
      products: payloads.products,
      inventory: payloads.inventory,
      assignees: payloads.assignees
    },
    updatedGoals,
    reportDate
  );
  const towelWest = model.inventoryRows.find(r => r.sku === 'RH-TOWEL-01' && r.storeId === 7101002);
  assert.strictEqual(towelWest.isLow, false);
});

check('parseGoalFormAnswers reads form titles', () => {
  const item = (title, value) => ({ getItem: () => ({ getTitle: () => title }), getResponse: () => value });
  const answers = g.parseGoalFormAnswers_([
    item(g.GOAL_FIELDS.store, 'Riverlane Downtown (sample)'),
    item(g.GOAL_FIELDS.revenueGoal, '650'),
    item(g.GOAL_FIELDS.lowStock, '8')
  ]);
  assert.strictEqual(answers.storeName, 'Riverlane Downtown (sample)');
  assert.strictEqual(answers.revenueGoal, 650);
  assert.strictEqual(answers.lowStockThreshold, 8);
});

check('fetchShopify_ uses injected fetchImpl', () => {
  g.setFetchImpl_((url) => ({
    getResponseCode: () => 200,
    getContentText: () => JSON.stringify({ ok: true, url })
  }));
  const resp = g.fetchShopify_('https://example.myshopify.com/admin/api/2024-10/orders.json', {});
  assert.strictEqual(JSON.parse(resp.getContentText()).ok, true);
  g.setFetchImpl_(null);
});

const sent = [];
global.MailApp = {
  sendEmail: m => sent.push(m)
};

check('sendLowStockAlerts_ sends one email per assignee', () => {
  sent.length = 0;
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
  g.sendLowStockAlerts_(model.lowStockByAssignee);
  assert.strictEqual(sent.length, 2);
  assert.ok(sent.every(m => m.subject.indexOf('Low stock') >= 0));
});

check('header-only or missing sheets do not stop the dashboard read', () => {
  const sheet = (rows) => ({
    getLastRow: () => rows.length,
    getLastColumn: () => (rows[0] || []).length,
    getRange: (r, c, n, w) => ({ getValues: () => { if (n < 1) throw new Error('bad range'); return rows.slice(r - 1, r - 1 + n); } })
  });
  assert.deepStrictEqual(g.sheetToRows_(null), []);
  assert.deepStrictEqual(g.sheetToRows_(sheet([])), []);
  assert.deepStrictEqual(g.sheetToRows_(sheet([['date', 'store', 'sales']])), []);
  assert.strictEqual(g.sheetToRows_(sheet([['h'], ['a'], ['b']])).length, 2);
  assert.deepStrictEqual(g.readGoalsFromSheet_(g.sheetToRows_(sheet([['store', 'goal']]))), g.readGoalsFromSheet_([]));
  assert.ok(Array.isArray(g.readDailySalesForDashboard_([])));
  assert.ok(Array.isArray(g.readInventoryForDashboard_([])));
});

console.log('All logic tests passed (' + passed + ' checks)');
module.exports = { passed };
