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

// Shopify se naye orders sheet mein (purana status kabhi overwrite nahi hota)
function syncOrders() {
  const p = PropertiesService.getScriptProperties();
  const shop = p.getProperty('SHOP');
  if (!shop) throw new Error('Script Properties mein SHOP set karein.');
  const token = getToken_(p, shop);
  const sh = getSheet_(), last = sh.getLastRow();
  const have = new Set(last > 1 ? sh.getRange(2, C.ID, last - 1, 1).getDisplayValues().flat() : []);
  const since = new Date(Date.now() - 30 * 864e5).toISOString();
  let url = 'https://' + shop + '/admin/api/2024-10/orders.json?status=any&limit=250&created_at_min=' + encodeURIComponent(since);
  const rows = [];
  while (url) {
    const res = UrlFetchApp.fetch(url, {headers: {'X-Shopify-Access-Token': token}, muteHttpExceptions: true});
    if (res.getResponseCode() !== 200) throw new Error('Shopify error: ' + res.getContentText());
    JSON.parse(res.getContentText()).orders.forEach(o => {
      if (o.cancelled_at || have.has(String(o.id))) return;
      const a = o.shipping_address || {}, total = Number(o.total_price);
      rows.push([o.name, Utilities.formatDate(new Date(o.created_at), TZ(), 'yyyy-MM-dd'),
        a.name || ((o.customer || {}).first_name || ''), a.phone || o.phone || (o.customer || {}).phone || '',
        a.city || '', [a.address1, a.address2].filter(Boolean).join(', '),
        (o.line_items || []).map(i => i.quantity + 'x ' + i.title).join(', '),
        total, o.financial_status === 'paid' ? 0 : total, '', 'New', '', '', DEFAULT_CHARGE, '', '', String(o.id)]);
    });
    const m = String(res.getHeaders()['Link'] || res.getHeaders()['link'] || '').match(/<([^>]+)>;\s*rel="next"/);
    url = m ? m[1] : null;
  }
  rows.reverse(); // purane pehle, naye neeche
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  return rows.length;
}

function getOrders() {
  const sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues().map(r => ({
    no: r[0], createdAt: r[1], name: r[2], phone: r[3], city: r[4], address: r[5], items: r[6],
    total: r[7], cod: r[8], courier: r[9], status: r[10] || 'New', delivered: r[11], deliveredDate: r[12],
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

// Roz raat: naye orders lao aur Daily sheet mein hisaab likho
function dailyJob() { syncOrders(); writeDaily_(); }

function writeDaily_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dh = ss.getSheetByName(DAILY_NAME);
  if (!dh) {
    dh = ss.insertSheet(DAILY_NAME);
    dh.appendRow(['Date','New','Asked','Delivered','Returned','COD with courier','Delivered today','Charges today','Net today']);
    dh.setFrozenRows(1); dh.getRange('A:A').setNumberFormat('@');
  }
  const sh = getSheet_(), last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues() : [];
  const t = today_(), n = {New:0, Asked:0, Delivered:0, Returned:0};
  let cod = 0, got = 0, ch = 0;
  data.forEach(r => {
    const s = r[C.STATUS - 1] || 'New'; n[s] = (n[s] || 0) + 1;
    if (s === 'Asked') cod += Number(r[C.COD - 1]) || 0;
    if (s === 'Delivered' && r[C.DDATE - 1] === t) { got += Number(r[C.DELIVERED - 1]) || 0; ch += Number(r[C.CHARGE - 1]) || 0; }
  });
  const line = [t, n.New, n.Asked, n.Delivered, n.Returned, cod, got, ch, got - ch];
  const dl = dh.getLastRow();
  const at = dl > 1 ? dh.getRange(2, 1, dl - 1, 1).getDisplayValues().flat().indexOf(t) : -1;
  if (at >= 0) dh.getRange(at + 2, 1, 1, line.length).setValues([line]); else dh.appendRow(line);
}

// Ek baar chala dein: roz raat 11 baje dailyJob chalega
function createTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(23).create();
}
