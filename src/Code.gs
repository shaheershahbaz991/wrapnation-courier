// Wrapnation Courier Tracker: Google Sheet hi database hai.
const SHEET_NAME = 'Orders', DAILY_NAME = 'Daily', DEFAULT_CHARGE = 250;
const HEADERS = ['Order No','Date','Customer','Phone','City','Address','Items','Order Total','COD Amount',
  'Courier','Status','Delivered Amount','Delivered Date','Courier Charge','Net Amount','Note','Shopify ID'];
const C = {NO:1,DATE:2,NAME:3,PHONE:4,CITY:5,ADDR:6,ITEMS:7,TOTAL:8,COD:9,COURIER:10,STATUS:11,
  DELIVERED:12,DDATE:13,CHARGE:14,NET:15,NOTE:16,ID:17};
const TZ = () => Session.getScriptTimeZone();
const today_ = () => Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd');

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e4f3f8');
    [C.DATE, C.PHONE, C.DDATE, C.ID].forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
  }
  return sh;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Wrapnation Courier')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getSheetUrl() { return SpreadsheetApp.getActiveSpreadsheet().getUrl(); }


// Naye Shopify apps: CLIENT_ID + CLIENT_SECRET se token khud banta hai (24 ghante chalta hai).
// Purani custom app ho to sirf TOKEN (shpat_...) set karein.
function getToken_(p, shop) {
  const fixed = p.getProperty('TOKEN');
  if (fixed) return fixed;
  const cache = CacheService.getScriptCache(), hit = cache.get('shopify_token');
  if (hit) return hit;
  const id = p.getProperty('CLIENT_ID'), secret = p.getProperty('CLIENT_SECRET');
  if (!id || !secret) throw new Error('Script Properties mein CLIENT_ID aur CLIENT_SECRET set karein.');
  const res = UrlFetchApp.fetch('https://' + shop + '/admin/oauth/access_token', {
    method: 'post', contentType: 'application/x-www-form-urlencoded', muteHttpExceptions: true,
    payload: {grant_type: 'client_credentials', client_id: id, client_secret: secret}});
  if (res.getResponseCode() !== 200) throw new Error('Shopify token error: ' + res.getContentText());
  const t = JSON.parse(res.getContentText()).access_token;
  cache.put('shopify_token', t, 21000);
  return t;
}

// Shopify ke orders sheet mein NAHI jate. Pehle Inbox mein dikhte hain,
// sirf wo orders sheet mein aate hain jo aap "Local courier" chunte hain.
function fetchShopify_() {
  const p = PropertiesService.getScriptProperties(), shop = p.getProperty('SHOP');
  if (!shop) throw new Error('Script Properties mein SHOP set karein.');
  const token = getToken_(p, shop);
  const since = new Date(Date.now() - 30 * 864e5).toISOString();
  let url = 'https://' + shop + '/admin/api/2024-10/orders.json?status=any&fulfillment_status=unfulfilled&limit=250&created_at_min=' + encodeURIComponent(since);
  const out = [];
  while (url) {
    const res = UrlFetchApp.fetch(url, {headers: {'X-Shopify-Access-Token': token}, muteHttpExceptions: true});
    if (res.getResponseCode() !== 200) throw new Error('Shopify error: ' + res.getContentText());
    JSON.parse(res.getContentText()).orders.forEach(o => {
      if (o.cancelled_at) return;
      const a = o.shipping_address || {}, total = Number(o.total_price);
      out.push({id: String(o.id), no: o.name, createdAt: Utilities.formatDate(new Date(o.created_at), TZ(), 'yyyy-MM-dd'),
        name: a.name || ((o.customer || {}).first_name || ''), phone: a.phone || o.phone || (o.customer || {}).phone || '',
        city: a.city || '', address: [a.address1, a.address2].filter(Boolean).join(', '),
        items: (o.line_items || []).map(i => i.quantity + 'x ' + i.title).join(', '),
        total: total, cod: o.financial_status === 'paid' ? 0 : total});
    });
    const m = String(res.getHeaders()['Link'] || res.getHeaders()['link'] || '').match(/<([^>]+)>;\s*rel="next"/);
    url = m ? m[1] : null;
  }
  return out;
}

function sheetIds_() {
  const sh = getSheet_(), last = sh.getLastRow();
  return new Set(last > 1 ? sh.getRange(2, C.ID, last - 1, 1).getDisplayValues().flat() : []);
}

// Inbox: Shopify ke wo orders jo na sheet mein hain na skip kiye gaye
function getInbox() {
  const have = sheetIds_();
  const skipped = new Set(JSON.parse(PropertiesService.getScriptProperties().getProperty('SKIPPED') || '[]'));
  return fetchShopify_().filter(o => !have.has(o.id) && !skipped.has(o.id));
}

// Chune hue orders sheet mein likho (local courier wale)
function addOrders(list) {
  const have = sheetIds_();
  const rows = list.filter(o => !have.has(String(o.id))).map(o => [o.no, o.createdAt, o.name, o.phone, o.city,
    o.address, o.items, Number(o.total) || 0, Number(o.cod) || 0, '', 'New', '', '', DEFAULT_CHARGE, '', '', String(o.id)]);
  if (rows.length) { const sh = getSheet_(); sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows); }
  return rows.length;
}

// PostEx / baqi orders: Inbox se hata do (sheet mein kuch nahi likha jata)
function skipOrders(ids) {
  const p = PropertiesService.getScriptProperties();
  const all = JSON.parse(p.getProperty('SKIPPED') || '[]').concat(ids.map(String));
  p.setProperty('SKIPPED', JSON.stringify(all.slice(-400)));
  return true;
}

function getOrders() {
  const sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues().map(r => ({
    no: r[0], createdAt: r[1], name: r[2], phone: r[3], city: r[4], address: r[5], items: r[6],
    total: r[7], cod: r[8], courier: r[9], status: r[10] === 'Asked' ? 'Handed over' : (r[10] || 'New'), delivered: r[11], deliveredDate: r[12],
    charge: r[13], note: r[15], id: r[16]
  })).reverse();
}

function updateOrder(id, d) {
  const sh = getSheet_(), last = sh.getLastRow();
  const i = sh.getRange(2, C.ID, last - 1, 1).getDisplayValues().flat().indexOf(String(id));
  if (i < 0) throw new Error('Order nahi mila');
  const row = i + 2, del = d.status === 'Delivered', charge = Number(d.charge) || 0;
  const v = (c, x) => sh.getRange(row, c).setValue(x);
  v(C.COURIER, d.courier || ''); v(C.STATUS, d.status); v(C.NOTE, d.note || ''); v(C.PHONE, d.phone || '');
  v(C.CHARGE, charge);
  v(C.DELIVERED, del ? Number(d.amount) || 0 : '');
  v(C.DDATE, del ? today_() : '');
  v(C.NET, del ? (Number(d.amount) || 0) - charge : '');
  return true;
}

// Roz raat: Daily sheet mein hisaab likho (naye orders sheet mein nahi aate)
function dailyJob() { writeDaily_(); }

function writeDaily_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dh = ss.getSheetByName(DAILY_NAME);
  if (!dh) {
    dh = ss.insertSheet(DAILY_NAME);
    dh.appendRow(['Date','New','Handed over','Delivered','Returned','COD with courier','Delivered today','Charges today','Net today']);
    dh.setFrozenRows(1); dh.getRange('A:A').setNumberFormat('@');
  }
  const sh = getSheet_(), last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues() : [];
  const t = today_(), n = {New:0, 'Handed over':0, Delivered:0, Returned:0};
  let cod = 0, got = 0, ch = 0;
  data.forEach(r => {
    let s = r[C.STATUS - 1] || 'New'; if (s === 'Asked') s = 'Handed over'; n[s] = (n[s] || 0) + 1;
    if (s === 'Handed over') cod += Number(r[C.COD - 1]) || 0;
    if (s === 'Delivered' && r[C.DDATE - 1] === t) { got += Number(r[C.DELIVERED - 1]) || 0; ch += Number(r[C.CHARGE - 1]) || 0; }
  });
  const line = [t, n.New, n['Handed over'], n.Delivered, n.Returned, cod, got, ch, got - ch];
  const dl = dh.getLastRow();
  const at = dl > 1 ? dh.getRange(2, 1, dl - 1, 1).getDisplayValues().flat().indexOf(t) : -1;
  if (at >= 0) dh.getRange(at + 2, 1, 1, line.length).setValues([line]); else dh.appendRow(line);
}

// Ek baar chala dein: roz raat 11 baje Daily hisaab likha jayega
function createTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(23).create();
}

// Label settings (sender, default weight, dimensions, tracking prefix)
function getLabelSettings() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('LABEL') || 'null');
}
function saveLabelSettings(s) {
  PropertiesService.getScriptProperties().setProperty('LABEL', JSON.stringify(s));
  return true;
}

// ---- Shopify fulfillment: courier ko dene par order Fulfilled ----
// Scopes: read_orders + read/write_merchant_managed_fulfillment_orders
function shopifyFetch_(path, opt) {
  const p = PropertiesService.getScriptProperties(), shop = p.getProperty('SHOP');
  let res;
  for (let t = 0; t < 2; t++) {
    const o = Object.assign({headers: {'X-Shopify-Access-Token': getToken_(p, shop)}, muteHttpExceptions: true}, opt || {});
    res = UrlFetchApp.fetch('https://' + shop + '/admin/api/2024-10/' + path, o);
    const code = res.getResponseCode();
    // Purana cached token naye scopes ke baghair ho sakta hai: cache saaf karke ek baar dobara
    if ((code === 401 || code === 403) && t === 0 && !p.getProperty('TOKEN')) { CacheService.getScriptCache().remove('shopify_token'); continue; }
    break;
  }
  return res;
}

function markGiven_(id) {
  const sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return;
  const i = sh.getRange(2, C.ID, last - 1, 1).getDisplayValues().flat().indexOf(String(id));
  if (i < 0) return;
  const cell = sh.getRange(i + 2, C.STATUS);
  if (!cell.getValue() || cell.getValue() === 'New') cell.setValue('Handed over');
}

function fulfillOrder(id, tracking, notify) {
  const r1 = shopifyFetch_('orders/' + id + '/fulfillment_orders.json');
  if (r1.getResponseCode() !== 200)
    throw new Error('Shopify (' + r1.getResponseCode() + '): ' + r1.getContentText().slice(0, 200) + ' | Scopes aur app approval check karein.');
  const fos = JSON.parse(r1.getContentText()).fulfillment_orders.filter(f => f.status === 'open' || f.status === 'in_progress');
  if (!fos.length) { markGiven_(id); return {ok: true, msg: 'Shopify par pehle se fulfilled hai'}; }
  const body = {fulfillment: {notify_customer: !!notify,
    tracking_info: {number: String(tracking || ''), company: 'Wrapnation Courier'},
    line_items_by_fulfillment_order: fos.map(f => ({fulfillment_order_id: f.id}))}};
  const r2 = shopifyFetch_('fulfillments.json', {method: 'post', contentType: 'application/json', payload: JSON.stringify(body)});
  if (r2.getResponseCode() !== 200 && r2.getResponseCode() !== 201)
    throw new Error('Fulfill nahi hua (' + r2.getResponseCode() + '): ' + r2.getContentText().slice(0, 250));
  markGiven_(id);
  return {ok: true, msg: 'Shopify par Fulfilled ✅ (' + tracking + ')'};
}

// Ek baar chalayein: purane "Asked" status Sheet mein "Handed over" ban jayenge
function migrateAskedToHandedOver() {
  const sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return 0;
  const rg = sh.getRange(2, C.STATUS, last - 1, 1), v = rg.getValues();
  let n = 0;
  v.forEach(r => { if (r[0] === 'Asked') { r[0] = 'Handed over'; n++; } });
  rg.setValues(v);
  return n;
}
