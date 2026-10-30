/**
 * Riverlane Home Goods — retail ops dashboard (sample)
 *
 * Pulls Shopify Admin API shaped JSON into Google Sheets, tracks per-store daily sales
 * and inventory, sends low-stock emails grouped by assignee, accepts goal updates from a
 * Google Form, and serves a simple HtmlService dashboard.
 *
 * All store names, people, and API payloads are fictional. No live Shopify credentials
 * are stored in this repository. Run setup() once from the script editor, then syncRetailData().
 */

/** @type {null | function(string, Object=): *} */
var fetchImpl_ = null;

/** Replace UrlFetchApp for tests or offline demos. */
function setFetchImpl_(fn) {
  fetchImpl_ = fn;
}

function getFetchImpl_() {
  return fetchImpl_;
}

const CONFIG = {
  brandName: 'Riverlane Home Goods (sample)',
  timeZone: 'America/New_York',
  apiVersion: '2024-10',
  defaultRevenueGoal: 500,
  defaultLowStockThreshold: 5,
  alertFromName: 'Riverlane Ops Bot (sample)',
  sheets: {
    dailySales: 'DailySales',
    inventory: 'Inventory',
    goals: 'StoreGoals',
    syncLog: 'SyncLog'
  }
};

const GOAL_FIELDS = {
  store: 'Store location',
  revenueGoal: 'Daily revenue goal (USD)',
  lowStock: 'Low stock threshold (units)'
};

const DAILY_SALES_HEADERS = ['Report date', 'Store ID', 'Store name', 'Revenue', 'Order count', 'Revenue goal', 'Meets goal'];
const INVENTORY_HEADERS = ['Store ID', 'Store name', 'SKU', 'Product', 'Qty', 'Threshold', 'Low flag', 'Assignee name', 'Assignee email'];
const GOALS_HEADERS = ['Store ID', 'Store name', 'Revenue goal', 'Low stock threshold', 'Updated at', 'Updated by'];
const SYNC_LOG_HEADERS = ['Timestamp', 'Status', 'Message'];

/**
 * First-time setup: spreadsheet, goal form, triggers, default goals.
 */
function setup() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  const props = PropertiesService.getScriptProperties();
  let ss = openSpreadsheetOrNull_(props.getProperty('SHEET_ID'));
  let form = openFormOrNull_(props.getProperty('GOAL_FORM_ID'));

  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.brandName + ' — Operations');
    ss.setSpreadsheetTimeZone(CONFIG.timeZone);
    initSheets_(ss);
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!form) {
    form = createGoalForm_();
    props.setProperty('GOAL_FORM_ID', form.getId());
  }

  removeProjectTriggers_();
  ScriptApp.newTrigger('onGoalFormSubmit').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('runScheduledSync').timeBased().everyHours(6).create();

  seedDefaultGoals_(ss);
  console.log('Spreadsheet: ' + ss.getUrl());
  console.log('Goal form (respond): ' + form.getPublishedUrl());
  return { spreadsheetUrl: ss.getUrl(), formUrl: form.getPublishedUrl() };
}

function initSheets_(ss) {
  const names = ss.getSheets().map(s => s.getName());
  ensureSheet_(ss, CONFIG.sheets.dailySales, DAILY_SALES_HEADERS, names);
  ensureSheet_(ss, CONFIG.sheets.inventory, INVENTORY_HEADERS, names);
  ensureSheet_(ss, CONFIG.sheets.goals, GOALS_HEADERS, names);
  ensureSheet_(ss, CONFIG.sheets.syncLog, SYNC_LOG_HEADERS, names);
}

function ensureSheet_(ss, name, headers, existingNames) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedDefaultGoals_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.goals);
  if (sheet.getLastRow() > 1) return;
  const defaults = [
    [7101001, 'Riverlane Downtown (sample)', CONFIG.defaultRevenueGoal, CONFIG.defaultLowStockThreshold, new Date(), 'setup'],
    [7101002, 'Riverlane Westside (sample)', CONFIG.defaultRevenueGoal, CONFIG.defaultLowStockThreshold, new Date(), 'setup']
  ];
  sheet.getRange(2, 1, defaults.length, GOALS_HEADERS.length).setValues(defaults);
}

function createGoalForm_() {
  const form = FormApp.create(CONFIG.brandName + ' — Update store goals', true)
    .setDescription('Sample form only. Use fictional values.')
    .setCollectEmail(false);
  form.addListItem().setTitle(GOAL_FIELDS.store).setChoiceValues([
    'Riverlane Downtown (sample)',
    'Riverlane Westside (sample)'
  ]).setRequired(true);
  form.addTextItem().setTitle(GOAL_FIELDS.revenueGoal).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());
  form.addTextItem().setTitle(GOAL_FIELDS.lowStock).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());
  if (form.supportsAdvancedResponderPermissions() && !form.isPublished()) form.setPublished(true);
  return form;
}

/**
 * Menu entry and dashboard button: sync then optional alerts.
 */
function runSyncFromMenu() {
  const result = syncRetailData({ sendAlerts: true });
  return result;
}

function runScheduledSync() {
  syncRetailData({ sendAlerts: true });
}

/**
 * Fetch Shopify JSON (or use fetchImpl_), write sheets, optionally email alerts.
 * @param {{sendAlerts?: boolean}} options
 */
function syncRetailData(options) {
  options = options || {};
  const ss = getSpreadsheet_();
  const reportDate = formatReportDate_(new Date(), CONFIG.timeZone);

  let payloads;
  try {
    payloads = loadShopifyPayloads_();
  } catch (err) {
    appendSyncLog_(ss, 'ERROR', err.message);
    throw err;
  }

  const goals = readGoalsFromSheet_(sheetToRows_(ss.getSheetByName(CONFIG.sheets.goals)));
  const model = buildSyncModel_(payloads, goals, reportDate);

  writeDailySalesSheet_(ss.getSheetByName(CONFIG.sheets.dailySales), model.dailySalesRows);
  writeInventorySheet_(ss.getSheetByName(CONFIG.sheets.inventory), model.inventoryRows);
  propsSetLastSync_();

  if (options.sendAlerts) {
    sendLowStockAlerts_(model.lowStockByAssignee);
  }

  appendSyncLog_(ss, 'OK', 'Synced ' + model.dailySalesRows.length + ' store sales rows, ' +
    model.inventoryRows.length + ' inventory rows.');
  return { reportDate: reportDate, sales: model.dailySalesRows.length, inventory: model.inventoryRows.length };
}

function loadShopifyPayloads_() {
  const orders = fetchShopifyJson_('/orders.json?status=any&limit=250');
  const locations = fetchShopifyJson_('/locations.json');
  const products = fetchShopifyJson_('/products.json?limit=250');
  const locationIds = (locations.locations || []).map(l => l.id).join(',');
  const inventory = locationIds
    ? fetchShopifyJson_('/inventory_levels.json?location_ids=' + locationIds)
    : { inventory_levels: [] };
  const assignees = loadAssigneeCatalog_();
  return { orders, locations, products, inventory, assignees };
}

/** In production, assignees might live in a Sheet tab; sample uses Script property or defaults. */
function loadAssigneeCatalog_() {
  const raw = PropertiesService.getScriptProperties().getProperty('ASSIGNEE_JSON');
  if (raw) {
    try { return JSON.parse(raw).sku_assignees || []; } catch (e) { /* fall through */ }
  }
  return defaultAssigneeCatalog_();
}

function defaultAssigneeCatalog_() {
  return [
    { sku: 'RH-TOWEL-01', assignee_name: 'Alex (Ops West)', assignee_email: 'alex.ops.west@example.com' },
    { sku: 'RH-MUG-12', assignee_name: 'Jordan (Ops Downtown)', assignee_email: 'jordan.ops.downtown@example.com' },
    { sku: 'RH-CNDL-08', assignee_name: 'Sam (Merchandising)', assignee_email: 'sam.merch@example.com' }
  ];
}

function fetchShopifyJson_(path) {
  const shop = PropertiesService.getScriptProperties().getProperty('SHOPIFY_SHOP_DOMAIN');
  const token = PropertiesService.getScriptProperties().getProperty('SHOPIFY_ACCESS_TOKEN');
  if (!shop || !token) {
    throw new Error('Set SHOPIFY_SHOP_DOMAIN and SHOPIFY_ACCESS_TOKEN in Script properties, or inject fetchImpl_ in tests.');
  }
  const url = 'https://' + shop + '/admin/api/' + CONFIG.apiVersion + path;
  const response = fetchShopify_(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': token, Accept: 'application/json' }
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Shopify API ' + path + ' returned ' + code + ': ' + text.slice(0, 200));
  }
  return JSON.parse(text);
}

/**
 * @param {string} url
 * @param {Object} options
 * @return {{getResponseCode: function(): number, getContentText: function(): string}}
 */
function fetchShopify_(url, options) {
  if (fetchImpl_) {
    return fetchImpl_(url, options);
  }
  const resp = UrlFetchApp.fetch(url, options);
  return {
    getResponseCode: () => resp.getResponseCode(),
    getContentText: () => resp.getContentText()
  };
}

/**
 * Offline demo: pass parsed sample JSON (see test/loadSamplePayloads).
 * @param {Object} payloads
 * @param {{sendAlerts?: boolean}} options
 */
function syncRetailDataFromPayloads(payloads, options) {
  options = options || {};
  const ss = getSpreadsheet_();
  const reportDate = formatReportDate_(new Date(), CONFIG.timeZone);
  const goals = readGoalsFromSheet_(sheetToRows_(ss.getSheetByName(CONFIG.sheets.goals)));
  const model = buildSyncModel_(payloads, goals, reportDate);
  writeDailySalesSheet_(ss.getSheetByName(CONFIG.sheets.dailySales), model.dailySalesRows);
  writeInventorySheet_(ss.getSheetByName(CONFIG.sheets.inventory), model.inventoryRows);
  if (options.sendAlerts) sendLowStockAlerts_(model.lowStockByAssignee);
  appendSyncLog_(ss, 'OK', 'Sample payload sync');
  return model;
}

function onGoalFormSubmit(e) {
  const answers = parseGoalFormAnswers_(e.response.getItemResponses());
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.sheets.goals);
  const rows = sheetToRows_(sheet);
  const goals = readGoalsFromSheet_(rows);
  const store = findStoreByDisplayName_(answers.storeName);
  if (!store) {
    appendSyncLog_(ss, 'WARN', 'Unknown store in form: ' + answers.storeName);
    return;
  }
  const updated = applyGoalUpdate_(goals, {
    storeId: store.id,
    storeName: store.name,
    revenueGoal: answers.revenueGoal,
    lowStockThreshold: answers.lowStockThreshold,
    updatedAt: new Date(),
    updatedBy: 'Google Form'
  });
  writeGoalsSheet_(sheet, updated);
  syncRetailData({ sendAlerts: false });
}

function sendLowStockAlerts_(byAssignee) {
  Object.keys(byAssignee).forEach(email => {
    const bundle = byAssignee[email];
    MailApp.sendEmail({
      to: email,
      subject: '[' + CONFIG.brandName + '] Low stock items (' + bundle.items.length + ')',
      body: buildAssigneeAlertEmail_(bundle.assigneeName, bundle.items),
      name: CONFIG.alertFromName
    });
  });
}

function doGet() {
  const ss = getSpreadsheet_();
  const salesSheet = ss.getSheetByName(CONFIG.sheets.dailySales);
  const invSheet = ss.getSheetByName(CONFIG.sheets.inventory);
  const goalsSheet = ss.getSheetByName(CONFIG.sheets.goals);
  const reportDate = salesSheet.getLastRow() > 1 ? String(salesSheet.getRange(2, 1).getValue()) : formatReportDate_(new Date(), CONFIG.timeZone);
  const goals = readGoalsFromSheet_(sheetToRows_(goalsSheet));
  const salesRows = readDailySalesForDashboard_(sheetToRows_(salesSheet));
  const inventoryRows = readInventoryForDashboard_(sheetToRows_(invSheet));
  const data = {
    brandName: CONFIG.brandName,
    reportDate: reportDate,
    lastSync: PropertiesService.getScriptProperties().getProperty('LAST_SYNC') || 'never',
    salesRows: salesRows,
    inventoryRows: inventoryRows,
    goals: goals
  };
  const tpl = HtmlService.createTemplateFromFile('Dashboard');
  tpl.data = data;
  return tpl.evaluate()
    .setTitle(CONFIG.brandName + ' dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function removeProjectTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

function resetSample() {
  removeProjectTriggers_();
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log('Cleared script properties and triggers. Delete old spreadsheet and form manually if needed.');
}

/* ---- Sheet writers ---- */

function writeDailySalesSheet_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, DAILY_SALES_HEADERS.length).setValues([DAILY_SALES_HEADERS]).setFontWeight('bold');
  if (!rows.length) return;
  const values = rows.map(r => [
    r.reportDate, r.storeId, r.storeName, r.revenue, r.orderCount, r.revenueGoal, r.meetsGoal ? 'YES' : 'NO'
  ]);
  sheet.getRange(2, 1, values.length, DAILY_SALES_HEADERS.length).setValues(values);
}

function writeInventorySheet_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, INVENTORY_HEADERS.length).setValues([INVENTORY_HEADERS]).setFontWeight('bold');
  if (!rows.length) return;
  const values = rows.map(r => [
    r.storeId, r.storeName, r.sku, r.productTitle, r.qty, r.threshold, r.isLow ? 'LOW' : '',
    r.assigneeName, r.assigneeEmail
  ]);
  sheet.getRange(2, 1, values.length, INVENTORY_HEADERS.length).setValues(values);
}

function writeGoalsSheet_(sheet, goals) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, GOALS_HEADERS.length).setValues([GOALS_HEADERS]).setFontWeight('bold');
  const values = goals.map(g => [
    g.storeId, g.storeName, g.revenueGoal, g.lowStockThreshold, g.updatedAt, g.updatedBy
  ]);
  if (values.length) {
    sheet.getRange(2, 1, values.length, GOALS_HEADERS.length).setValues(values);
  }
}

function appendSyncLog_(ss, status, message) {
  const sheet = ss.getSheetByName(CONFIG.sheets.syncLog);
  sheet.appendRow([new Date(), status, message]);
}

function propsSetLastSync_() {
  PropertiesService.getScriptProperties().setProperty('LAST_SYNC', new Date().toISOString());
}

/* ---- Pure logic (Node tests) ---- */

/**
 * @param {Object} payloads
 * @param {Array<Object>} goals
 * @param {string} reportDate YYYY-MM-DD
 */
function buildSyncModel_(payloads, goals, reportDate) {
  const locations = indexLocations_(payloads.locations);
  const variantCatalog = buildVariantCatalog_(payloads.products, payloads.assignees);
  const dailySalesRows = aggregateDailySalesByStore_(
    payloads.orders.orders || [],
    locations,
    goals,
    reportDate,
    CONFIG.timeZone
  );
  const inventoryRows = buildInventoryRows_(
    (payloads.inventory && payloads.inventory.inventory_levels) || [],
    variantCatalog,
    goals,
    locations
  );
  const lowStockRows = inventoryRows.filter(r => r.isLow);
  const lowStockByAssignee = groupLowStockByAssignee_(lowStockRows);
  return { dailySalesRows, inventoryRows, lowStockByAssignee };
}

function indexLocations_(locationsPayload) {
  const map = {};
  (locationsPayload.locations || []).forEach(loc => {
    map[loc.id] = { id: loc.id, name: loc.name };
  });
  return map;
}

function buildVariantCatalog_(productsPayload, assigneeList) {
  const bySku = {};
  const assigneeBySku = {};
  (assigneeList || []).forEach(a => { assigneeBySku[a.sku] = a; });
  (productsPayload.products || []).forEach(p => {
    (p.variants || []).forEach(v => {
      if (!v.sku) return;
      const a = assigneeBySku[v.sku] || { assignee_name: 'Unassigned', assignee_email: 'ops-unassigned@example.com' };
      bySku[v.sku] = {
        sku: v.sku,
        productTitle: p.title,
        inventoryItemId: v.inventory_item_id,
        assigneeName: a.assignee_name,
        assigneeEmail: a.assignee_email
      };
    });
  });
  return { bySku, byInventoryItemId: invertInventoryItems_(bySku) };
}

function invertInventoryItems_(bySku) {
  const map = {};
  Object.keys(bySku).forEach(sku => {
    const row = bySku[sku];
    map[row.inventoryItemId] = row;
  });
  return map;
}

function goalForStore_(goals, storeId) {
  const g = goals.find(x => x.storeId === storeId);
  return g || {
    storeId: storeId,
    storeName: '',
    revenueGoal: CONFIG.defaultRevenueGoal,
    lowStockThreshold: CONFIG.defaultLowStockThreshold
  };
}

/**
 * Sum paid orders for reportDate (store local date) per location_id.
 */
function aggregateDailySalesByStore_(orders, locationsById, goals, reportDate, timeZone) {
  const totals = {};
  orders.forEach(order => {
    const localDate = orderLocalDate_(order.created_at, timeZone);
    if (localDate !== reportDate) return;
    const storeId = order.location_id;
    if (!totals[storeId]) totals[storeId] = { revenue: 0, orderCount: 0 };
    totals[storeId].orderCount += 1;
    (order.line_items || []).forEach(li => {
      totals[storeId].revenue += Number(li.price) * Number(li.quantity);
    });
  });

  return Object.keys(locationsById).map(idStr => {
    const storeId = Number(idStr);
    const loc = locationsById[storeId];
    const t = totals[storeId] || { revenue: 0, orderCount: 0 };
    const goal = goalForStore_(goals, storeId);
    const revenue = roundMoney_(t.revenue);
    return {
      reportDate,
      storeId,
      storeName: loc.name,
      revenue,
      orderCount: t.orderCount,
      revenueGoal: goal.revenueGoal,
      meetsGoal: revenue >= goal.revenueGoal
    };
  });
}

function buildInventoryRows_(levels, variantCatalog, goals, locationsById) {
  return levels.map(level => {
    const variant = variantCatalog.byInventoryItemId[level.inventory_item_id] || {
      sku: 'UNKNOWN',
      productTitle: 'Unknown item',
      assigneeName: 'Unassigned',
      assigneeEmail: 'ops-unassigned@example.com'
    };
    const storeId = level.location_id;
    const loc = locationsById[storeId] || { name: 'Unknown store' };
    const goal = goalForStore_(goals, storeId);
    const qty = Number(level.available);
    const threshold = goal.lowStockThreshold;
    return {
      storeId,
      storeName: loc.name,
      sku: variant.sku,
      productTitle: variant.productTitle,
      qty,
      threshold,
      isLow: qty < threshold,
      assigneeName: variant.assigneeName,
      assigneeEmail: variant.assigneeEmail
    };
  });
}

function groupLowStockByAssignee_(lowStockRows) {
  const out = {};
  lowStockRows.forEach(row => {
    const email = row.assigneeEmail;
    if (!out[email]) {
      out[email] = { assigneeName: row.assigneeName, items: [] };
    }
    out[email].items.push(row);
  });
  return out;
}

function buildAssigneeAlertEmail_(assigneeName, items) {
  const lines = [
    'Hello ' + assigneeName + ',',
    '',
    'The following items are below the configured threshold:',
    ''
  ];
  items.forEach(it => {
    lines.push('- ' + it.storeName + ' · ' + it.sku + ' (' + it.productTitle + '): ' +
      it.qty + ' on hand (threshold ' + it.threshold + ')');
  });
  lines.push('', 'This message was generated by the sample Riverlane ops script.');
  return lines.join('\n');
}

function parseGoalFormAnswers_(itemResponses) {
  const byTitle = {};
  itemResponses.forEach(ir => {
    byTitle[String(ir.getItem().getTitle()).trim()] = String(ir.getResponse() || '').trim();
  });
  return {
    storeName: byTitle[GOAL_FIELDS.store] || '',
    revenueGoal: Number(byTitle[GOAL_FIELDS.revenueGoal] || 0),
    lowStockThreshold: Number(byTitle[GOAL_FIELDS.lowStock] || 0)
  };
}

function readGoalsFromSheet_(rows) {
  if (!rows.length) return [];
  return rows.map(r => ({
    storeId: Number(r[0]),
    storeName: String(r[1]),
    revenueGoal: Number(r[2]),
    lowStockThreshold: Number(r[3]),
    updatedAt: r[4],
    updatedBy: String(r[5] || '')
  }));
}

function applyGoalUpdate_(goals, update) {
  let found = false;
  const next = goals.map(g => {
    if (g.storeId !== update.storeId) return g;
    found = true;
    return {
      storeId: update.storeId,
      storeName: update.storeName || g.storeName,
      revenueGoal: update.revenueGoal,
      lowStockThreshold: update.lowStockThreshold,
      updatedAt: update.updatedAt,
      updatedBy: update.updatedBy
    };
  });
  if (!found) {
    next.push({
      storeId: update.storeId,
      storeName: update.storeName,
      revenueGoal: update.revenueGoal,
      lowStockThreshold: update.lowStockThreshold,
      updatedAt: update.updatedAt,
      updatedBy: update.updatedBy
    });
  }
  return next;
}

function findStoreByDisplayName_(name) {
  const catalog = [
    { id: 7101001, name: 'Riverlane Downtown (sample)' },
    { id: 7101002, name: 'Riverlane Westside (sample)' }
  ];
  return catalog.find(s => s.name === name) || null;
}

function orderLocalDate_(iso, timeZone) {
  const d = new Date(iso);
  if (typeof Utilities !== 'undefined') {
    return Utilities.formatDate(d, timeZone, 'yyyy-MM-dd');
  }
  return new Intl.DateTimeFormat('en-CA', { timeZone: timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
}

function formatReportDate_(date, timeZone) {
  return orderLocalDate_(date.toISOString(), timeZone);
}

function roundMoney_(n) {
  return Math.round(Number(n) * 100) / 100;
}

function sheetToRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function readDailySalesForDashboard_(rows) {
  return rows.map(r => ({
    storeName: String(r[2]),
    revenue: Number(r[3]),
    orderCount: Number(r[4]),
    revenueGoal: Number(r[5]),
    meetsGoal: String(r[6]).toUpperCase() === 'YES'
  }));
}

function readInventoryForDashboard_(rows) {
  return rows.map(r => ({
    storeName: String(r[1]),
    sku: String(r[2]),
    productTitle: String(r[3]),
    qty: Number(r[4]),
    threshold: Number(r[5]),
    isLow: String(r[6]) === 'LOW',
    assigneeName: String(r[7])
  }));
}

/* ---- Helpers ---- */

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Missing SHEET_ID. Run setup() first.');
  return SpreadsheetApp.openById(id);
}

function openSpreadsheetOrNull_(id) {
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function openFormOrNull_(id) {
  if (!id) return null;
  try { return FormApp.openById(id); } catch (e) { return null; }
}

if (typeof module !== 'undefined') {
  module.exports = {
    CONFIG,
    GOAL_FIELDS,
    setFetchImpl_,
    getFetchImpl_,
    buildSyncModel_,
    aggregateDailySalesByStore_,
    buildInventoryRows_,
    groupLowStockByAssignee_,
    buildAssigneeAlertEmail_,
    parseGoalFormAnswers_,
    applyGoalUpdate_,
    readGoalsFromSheet_,
    orderLocalDate_,
    findStoreByDisplayName_,
    defaultAssigneeCatalog_,
    fetchShopify_,
    onGoalFormSubmit,
    sendLowStockAlerts_
  };
}
