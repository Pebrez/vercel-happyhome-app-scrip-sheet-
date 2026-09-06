/** ============================================================================
 *  HAPPY HOME — DASHBOARD WEB APP  (Google Apps Script)
 *  ----------------------------------------------------------------------------
 *  Wuxuu akhriyaa xogta sheet-yada 'Sales' iyo 'Expenses' ee spreadsheet-ka
 *  uu kuxidhan yahay, wuxuuna kuu soo saaraa dashboard web ah oo leh:
 *   - Login (Admin + Staff) — Users sheet
 *   - Garaafyo badan (iibka, kharashka, faa'iidada, qaybaha, iibiyayaasha, degmooyinka)
 *   - Filter: sanad / bil / maalin
 * ============================================================================ */

var SHEET_SALES = 'Sales';
var SHEET_EXPENSES = 'Expenses';
var SHEET_USERS = 'Users';
var SHEET_ORDERS = 'Orders';
var SHEET_ORDER_ITEMS = 'OrderItems';
var SESSION_HOURS = 6;

/* ---- Tiirarka Orders (1-based): dalab kasta hal saf ---- */
var O = { id: 1, date: 2, supplier: 3, notes: 4, status: 5, receivedDate: 6, cargo: 7, otherCosts: 8, monthStart: 9, year: 10 };
/* ---- Tiirarka OrderItems (1-based): alaab kasta hal saf, ku xidhan Order ID ---- */
var OI = { orderId: 1, product: 2, qty: 3, unitCost: 4, total: 5 };

/* ---- Tiirarka Sales (1-based) — kaliya haddii header-ka la heli waayo ---- */
var S = {
  invoiceId: 1, paper: 2, date: 3, customer: 4, phone: 5, custType: 6,
  location: 7, city: 8, district: 9, salesperson: 10, payMethod: 11,
  payStatus: 12, gross: 13, discType: 14, discInput: 15, discAmt: 16,
  net: 17, initPay: 18, extraPay: 19, totalPaid: 20, balance: 21,
  dueDate: 22, orderStatus: 23, cancelReason: 24, delivery: 25,
  custDelivCost: 26, compDelivCost: 27, custStatus: 28, monthStart: 29,
  year: 30, completedSales: 31, completedOrders: 32, receipt: 33
};

/* ---- Tiirarka Expenses (1-based) ---- */
var X = { id: 1, date: 2, category: 3, newCategory: 4, description: 5, amount: 6, paidBy: 7, payMethod: 8, monthStart: 9, year: 10 };

/* ============================== WEB APP ============================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Happy Home — Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= REST API (Vercel proxy → doPost) =================
   Marka dashboard-ku Vercel ku shaqeeyo, call-yadu waxay marayaan
   /api/data (Vercel function) → halkan (doPost). JSON: {method, args}
   Kaliya methods-ka whitelist-ka ah ayaa la aqbalo; token-ka ayaa
   function kasta gudihiisa la xaqiijinayaa (auth_). */
var API_ROUTES = {
  login:            function (a) { return login(String(a[0] || ''), String(a[1] || '')); },
  logout:           function (a) { return logout(a[0]); },
  whoAmI:           function (a) { return whoAmI(a[0]); },
  getMeta:          function (a) { return getMeta(a[0]); },
  getData:          function (a) { return getData(a[0], a[1]); },
  addSale:          function (a) { return addSale(a[0], a[1]); },
  addExpense:       function (a) { return addExpense(a[0], a[1]); },
  createOrder:      function (a) { return createOrder(a[0], a[1]); },
  receiveOrder:     function (a) { return receiveOrder(a[0], a[1], a[2], a[3]); },
  cancelOrder:      function (a) { return cancelOrder(a[0], a[1]); },
  getReceipt:       function (a) { return getReceipt(a[0], a[1]); },
  searchCustomer:   function (a) { return searchCustomer(a[0], a[1]); },
  saveSettings:     function (a) { return saveSettings(a[0], a[1]); },
  changePassword:   function (a) { return changePassword(a[0], a[1], a[2]); },
  saveGeminiKey:    function (a) { return saveGeminiKey(a[0], a[1]); },
  hasGeminiKey:     function (a) { return hasGeminiKey(a[0]); },
  scanInvoiceImage: function (a) { return scanInvoiceImage(a[0], a[1], a[2]); },
  scanExpenseImage: function (a) { return scanExpenseImage(a[0], a[1], a[2]); },
  aiChat:           function (a) { return aiChat(a[0], a[1], a[2], a[3]); }
};

function doPost(e) {
  try {
    var body = e && e.postData && e.postData.contents ? e.postData.contents : '';
    var req = {};
    try { req = JSON.parse(body); } catch (err) { return jsonOut_({ ok: false, msg: 'JSON body aan sax ahayn.' }); }
    var method = String(req.method || '');
    var route = API_ROUTES[method];
    if (!route) return jsonOut_({ ok: false, msg: 'Method aan la garanayn: ' + method });
    var args = Object.prototype.toString.call(req.args) === '[object Array]' ? req.args : [];
    return jsonOut_(route(args));
  } catch (err) {
    return jsonOut_({ ok: false, msg: 'Khalad server-ka: ' + err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (id) return SpreadsheetApp.openById(id);
    throw new Error('Kuma xidhna spreadsheet. Ka fur faylka Google Sheets kadibna dib u day deploy-ga.');
  }
}

/* ============================== USERS / LOGIN ============================== */

function setupUsers() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) sh = ss.insertSheet(SHEET_USERS);
  if (sh.getLastRow() === 0) {
    sh.getRange('A1:D1').setValues([['Username', 'Password', 'Full Name', 'Role']])
      .setFontWeight('bold').setBackground('#1f6b3b').setFontColor('#ffffff');
  }
  ensureUser_(sh, 'admin', 'admin123', 'Admin', 'Admin');
  ensureUser_(sh, 'staff', 'staff123', 'Staff', 'Staff');
  ensureUser_(sh, 'asma', 'admin123', 'Asma (Admin)', 'Admin');
  // Auto-register the Google account that runs Setup as Admin (username = email-kaaga)
  var ownerEmail = '';
  try { ownerEmail = String(Session.getEffectiveUser().getEmail() || ''); } catch (e) {}
  if (ownerEmail && ownerEmail.indexOf('@') > 0) {
    ensureUser_(sh, ownerEmail, 'admin123', 'Admin (Google Account)', 'Admin');
  }
  return 'Users sheet waa diyaar. Account-yada: admin/admin123, staff/staff123, asma/admin123 (Admin). Google account-kaaga (' + (ownerEmail || 'lama helin') + ') waa Admin — badal passwords-ka!';
}

function ensureUser_(sh, username, password, fullName, role) {
  var vals = sh.getDataRange().getValues();
  var uname = String(username).toLowerCase();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0] || '').toLowerCase() === uname) return;
  }
  sh.appendRow([username, password, fullName, role]);
}

function setupOrders() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) sh = ss.insertSheet(SHEET_ORDERS);
  if (sh.getLastRow() === 0) {
    sh.getRange('A1:J1').setValues([['Order ID', 'Order Date', 'Supplier', 'Notes', 'Status',
      'Received Date', 'Cargo Cost', 'Other Costs', 'Month Start', 'Year']])
      .setFontWeight('bold').setBackground('#1f4b84').setFontColor('#ffffff');
  }
  var sh2 = ss.getSheetByName(SHEET_ORDER_ITEMS);
  if (!sh2) sh2 = ss.insertSheet(SHEET_ORDER_ITEMS);
  if (sh2.getLastRow() === 0) {
    sh2.getRange('A1:E1').setValues([['Order ID', 'Product', 'Qty', 'Unit Cost', 'Total Cost']])
      .setFontWeight('bold').setBackground('#1f4b84').setFontColor('#ffffff');
  }
  return 'Orders + OrderItems sheets waa diyaar.';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Happy Home')
    .addItem('1. Setup (Samee Users sheet)', 'setupUsers')
    .addItem('2. Setup (Samee Orders sheets)', 'setupOrders')
    .addToUi();
}

function login(username, password) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) return { ok: false, msg: 'Users sheet ma jiro. Fur faylka, menu "Happy Home" → "1. Setup".' };
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var u = String(vals[i][0] || '').toLowerCase();
    var p = String(vals[i][1] || '');
    if (u === String(username || '').toLowerCase() && p === String(password || '')) {
      var user = { user: String(vals[i][0]), name: String(vals[i][2] || vals[i][0]), role: String(vals[i][3] || 'Staff') };
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('hh_' + token, JSON.stringify(user), SESSION_HOURS * 3600);
      return { ok: true, token: token, user: user };
    }
  }
  return { ok: false, msg: 'Username ama password waa khalad.' };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('hh_' + token);
  return { ok: true };
}

function whoAmI(token) {
  return auth_(token);
}

function auth_(token) {
  if (!token) return null;
  var v = CacheService.getScriptCache().get('hh_' + token);
  return v ? JSON.parse(v) : null;
}

/* ============================== HELPERS ============================== */

function findHeaderRow_(sheet, headerText, maxScan) {
  maxScan = maxScan || 10;
  if (sheet.getLastRow() < 1) return -1;
  var scan = Math.min(maxScan, sheet.getLastRow());
  var vals = sheet.getRange(1, 1, scan, 1).getValues();
  for (var i = 0; i < scan; i++) {
    if (String(vals[i][0]).toLowerCase().indexOf(String(headerText).toLowerCase()) !== -1) return i + 1;
  }
  return -1;
}

function buildColMap_(sheet, headerRow, maxCols) {
  var hdr = sheet.getRange(headerRow, 1, 1, maxCols).getValues()[0];
  var map = {};
  for (var i = 0; i < hdr.length; i++) {
    var k = String(hdr[i] || '').trim().toLowerCase();
    if (k) map[k] = i + 1;
  }
  return map;
}

function toDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000));
  var s = String(v || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var m2 = s.match(/^(\d{1,2})[\/\-.]?(\d{1,2})[\/\-.]?(\d{2,4})/);
  if (m2) {
    var y = Number(m2[3]); if (y < 100) y += 2000;
    var p1 = Number(m2[1]), p2 = Number(m2[2]);
    var mo, da;
    if (p1 > 12) { mo = p2; da = p1; }        // maalin-hore: 16-08-2026
    else if (p2 > 12) { mo = p1; da = p2; }   // bil-hore (US): 08/16/2026
    else { mo = p1; da = p2; }                // bil-maal: 08-16-2026
    return new Date(y, mo - 1, da);
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function num_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function fmtDate_(d) {
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  return d.getFullYear() + '-' + mm + '-' + dd;
}

/* ============================== READ DATA ============================== */

function readSales_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_SALES);
  if (!sh) return [];
  var hdrRow = findHeaderRow_(sh, 'System Invoice ID', 10);
  if (hdrRow < 1) hdrRow = 4;
  var map = buildColMap_(sh, hdrRow, 33);
  function col(key, fb) { return map[String(key).toLowerCase()] || fb; }
  var cInv = col('system invoice id', S.invoiceId), cPaper = col('paper invoice no.', S.paper),
    cDate = col('sale date*', S.date), cCust = col('customer name*', S.customer),
    cPhone = col('phone*', S.phone), cType = col('customer type*', S.custType),
    cLoc = col('location option*', S.location),
    cSales = col('salesperson*', S.salesperson), cDist = col('district', S.district),
    cStatus = col('order status*', S.orderStatus), cNet = col('net sales', S.net),
    cGross = col('gross sales*', S.gross), cDisc = col('discount amount', S.discAmt),
    cPaid = col('total paid', S.totalPaid), cBal = col('balance / resta', S.balance),
    cDeliv = col('customer delivery amount', S.custDelivCost), cComp = col('company delivery amount', S.compDelivCost);
  var start = hdrRow + 1;
  var last = sh.getLastRow();
  if (last < start) return [];
  var end = Math.min(last, start + 2000);
  var vals = sh.getRange(start, 1, end - start + 1, 33).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var d = toDate_(r[cDate - 1]);
    if (!d) continue;
    var status = String(r[cStatus - 1] || '').trim();
    var net = num_(r[cNet - 1]);
    rows.push({
      date: d, dateStr: fmtDate_(d),
      invoiceId: String(r[cInv - 1] || '').trim(),
      paper: String(r[cPaper - 1] || '').trim(),
      location: String(r[cLoc - 1] || '').trim(),
      customer: String(r[cCust - 1] || '').trim(),
      phone: String(r[cPhone - 1] || '').trim(),
      custType: String(r[cType - 1] || '').trim(),
      salesperson: String(r[cSales - 1] || '').trim(),
      district: String(r[cDist - 1] || '').trim(),
      status: status,
      net: net,
      gross: num_(r[cGross - 1]),
      discount: num_(r[cDisc - 1]),
      totalPaid: num_(r[cPaid - 1]),
      balance: num_(r[cBal - 1]),
      custDeliv: num_(r[cDeliv - 1]),
      compDeliv: num_(r[cComp - 1]),
      completed: status.toLowerCase() === 'completed'
    });
  }
  return rows;
}

function readExpenses_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_EXPENSES);
  if (!sh) return [];
  var hdrRow = findHeaderRow_(sh, 'Expense ID', 10);
  if (hdrRow < 1) hdrRow = 4;
  var map = buildColMap_(sh, hdrRow, 10);
  function col(key, fb) { return map[String(key).toLowerCase()] || fb; }
  var cDate = col('expense date*', X.date), cCat = col('category*', X.category),
    cNew = col('new category (if other)', X.newCategory),
    cDesc = col('description', X.description), cAmt = col('amount*', X.amount),
    cPaidBy = col('paid by', X.paidBy);
  var start = hdrRow + 1;
  var last = sh.getLastRow();
  if (last < start) return [];
  var end = Math.min(last, start + 2000);
  var vals = sh.getRange(start, 1, end - start + 1, 10).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var d = toDate_(r[cDate - 1]);
    if (!d) continue;
    var cat = String(r[cNew - 1] || '').trim() || String(r[cCat - 1] || '').trim() || 'Other';
    rows.push({
      date: d, dateStr: fmtDate_(d),
      category: cat,
      description: String(r[cDesc - 1] || '').trim(),
      amount: num_(r[cAmt - 1]),
      paidBy: String(r[cPaidBy - 1] || '').trim()
    });
  }
  return rows;
}

function readOrderItems_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDER_ITEMS);
  if (!sh) return {};
  var last = sh.getLastRow();
  if (last < 2) return {};
  var end = Math.min(last, 5000);
  var vals = sh.getRange(2, 1, end - 1, 5).getValues();
  var byOrder = {};
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var oid = String(r[OI.orderId - 1] || '').trim();
    if (!oid) continue;
    if (!byOrder[oid]) byOrder[oid] = [];
    byOrder[oid].push({
      product: String(r[OI.product - 1] || '').trim(),
      qty: num_(r[OI.qty - 1]),
      unitCost: num_(r[OI.unitCost - 1]),
      total: num_(r[OI.total - 1])
    });
  }
  return byOrder;
}

function readOrders_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var end = Math.min(last, 3000);
  var vals = sh.getRange(2, 1, end - 1, 10).getValues();
  var itemsMap = readOrderItems_();
  var rows = [];
  var today = new Date(); today.setHours(0, 0, 0, 0);
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var d = toDate_(r[O.date - 1]);
    if (!d) continue;
    var id = String(r[O.id - 1] || '').trim();
    var status = String(r[O.status - 1] || 'Pending').trim();
    var recD = toDate_(r[O.receivedDate - 1]);
    var items = itemsMap[id] || [];
    var total = 0;
    for (var j = 0; j < items.length; j++) total += items[j].total;
    var days = 0;
    if (status.toLowerCase() === 'pending') {
      days = Math.max(0, Math.round((today - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000));
    } else if (recD) {
      days = Math.max(0, Math.round((new Date(recD.getFullYear(), recD.getMonth(), recD.getDate()) - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000));
    }
    rows.push({
      id: id,
      date: d, dateStr: fmtDate_(d),
      supplier: String(r[O.supplier - 1] || '').trim(),
      notes: String(r[O.notes - 1] || '').trim(),
      status: status,
      receivedDate: recD, receivedDateStr: recD ? fmtDate_(recD) : '',
      cargo: num_(r[O.cargo - 1]),
      otherCosts: num_(r[O.otherCosts - 1]),
      items: items, total: Math.round(total * 100) / 100,
      days: days
    });
  }
  return rows;
}

/* ============================== META ============================== */

function getMeta(token) {
  if (!auth_(token)) return { error: 'login' };
  var sales = readSales_();
  var exp = readExpenses_();
  var years = {};
  var i;
  for (i = 0; i < sales.length; i++) years[sales[i].date.getFullYear()] = 1;
  for (i = 0; i < exp.length; i++) years[exp[i].date.getFullYear()] = 1;
  var yList = Object.keys(years).map(Number).sort(function (a, b) { return b - a; });
  if (!yList.length) yList = [new Date().getFullYear()];

  var ss = getSpreadsheet_();
  var dailyTarget = 500, monthlyTarget = 15000, profitRate = 0.3;
  var set = ss.getSheetByName('Settings');
  if (set) {
    var tv = set.getRange('B4:B6').getValues();
    if (num_(tv[0][0]) > 0) dailyTarget = num_(tv[0][0]);
    if (num_(tv[1][0]) > 0) monthlyTarget = num_(tv[1][0]);
    if (num_(tv[2][0]) > 0) profitRate = num_(tv[2][0]);
  }
  var people = {}, cats = {};
  for (i = 0; i < sales.length; i++) if (sales[i].salesperson) people[sales[i].salesperson] = 1;
  for (i = 0; i < exp.length; i++) if (exp[i].category) cats[exp[i].category] = 1;
  var props = null;
  try { props = PropertiesService.getScriptProperties(); } catch (e) { props = null; }
  var bizName = 'Happy Home', bizLogo = '';
  if (props) {
    bizName = props.getProperty('BIZ_NAME') || 'Happy Home';
    bizLogo = props.getProperty('BIZ_LOGO') || '';
  }
  return {
    years: yList,
    dailyTarget: dailyTarget,
    monthlyTarget: monthlyTarget,
    profitRate: profitRate,
    salespeople: Object.keys(people).sort(),
    categories: Object.keys(cats).sort(),
    currency: 'USD',
    bizName: bizName,
    bizLogo: bizLogo,
    theme: props ? (props.getProperty('THEME') || 'classic') : 'classic'
  };
}

/* ============================== DATA ============================== */

function inPeriod_(d, f) {
  if (d.getFullYear() !== f.year) return false;
  if (f.month > 0 && (d.getMonth() + 1) !== f.month) return false;
  if (f.day > 0 && d.getDate() !== f.day) return false;
  return true;
}

function getData(token, filter) {
  if (!auth_(token)) return { error: 'login' };
  filter = filter || {};
  var year = num_(filter.year) || new Date().getFullYear();
  var month = num_(filter.month) || 0;   // 0 = dhamaan bilaha
  var day = num_(filter.day) || 0;       // 0 = dhamaan maalmaha

  var sales = readSales_();
  var exp = readExpenses_();
  var poList = readOrders_();
  var i;

  var sRows = [], eRows = [];
  for (i = 0; i < sales.length; i++) if (inPeriod_(sales[i].date, { year: year, month: month, day: day })) sRows.push(sales[i]);
  for (i = 0; i < exp.length; i++) if (inPeriod_(exp[i].date, { year: year, month: month, day: day })) eRows.push(exp[i]);

  var totalSales = 0, orders = 0;
  for (i = 0; i < sRows.length; i++) { if (sRows[i].completed) { totalSales += sRows[i].net; orders++; } }
  var totalExp = 0;
  for (i = 0; i < eRows.length; i++) totalExp += eRows[i].amount;
  /* Cargo-ga iyo lacagta alaabta la dalbaday laguma darsan kharashka/faa'iidada guud (Sales - Expenses) —
     waxaa loo xisaabiyaa gooni ahaan hoos (qeybta "purchases"), sida aad dalbatay. */
  var net = totalSales - totalExp;

  var meta = getMeta(token);
  var dailyTarget = meta.dailyTarget, monthlyTarget = meta.monthlyTarget, profitRate = meta.profitRate;

  /* ---- Kala qaybin (breakdowns) ---- */
  var catMap = {}, distMap = {}, typeMap = { 'New': 0, 'Returning': 0 }, custMap = {}, custPhones = {};
  var newPhones = {}, retPhones = {};
  var personInfo = {}; // salesperson -> {sales, orders, discounts, discountAmt}
  var debtTotal = 0, debtCustomers = {}, pendingCount = 0;
  var discountTotal = 0, discountOrders = 0;
  var deliveryTotal = 0, deliveryOrders = 0;
  for (i = 0; i < sRows.length; i++) {
    var s = sRows[i];
    if (!s.completed) { if (s.status.toLowerCase() === 'pending') pendingCount++; continue; }
    var pn = s.salesperson || '(aan la aqoon)';
    if (!personInfo[pn]) personInfo[pn] = { sales: 0, orders: 0, discounts: 0, discountAmt: 0 };
    personInfo[pn].sales += s.net;
    personInfo[pn].orders += 1;
    if (s.discount > 0) { personInfo[pn].discounts += 1; personInfo[pn].discountAmt += s.discount; discountTotal += s.discount; discountOrders++; }
    distMap[s.district || '(aan la aqoon)'] = (distMap[s.district || '(aan la aqoon)'] || 0) + s.net;
    var t = (s.custType || '').toLowerCase();
    if (t.indexOf('return') !== -1) { typeMap['Returning'] += s.net; if (s.phone) retPhones[s.phone] = 1; }
    else { typeMap['New'] += s.net; if (s.phone) newPhones[s.phone] = 1; }
    if (s.customer) custMap[s.customer] = (custMap[s.customer] || 0) + s.net;
    if (s.phone) custPhones[s.phone] = 1;
    if (s.balance > 0) { debtTotal += s.balance; debtCustomers[s.customer || s.phone || '?'] = 1; }
    if (s.custDeliv > 0) { deliveryTotal += s.custDeliv; deliveryOrders++; }
  }
  for (i = 0; i < eRows.length; i++) {
    catMap[eRows[i].category || 'Other'] = (catMap[eRows[i].category || 'Other'] || 0) + eRows[i].amount;
  }

  function toSortedArr(map) {
    var out = [];
    for (var k in map) if (map.hasOwnProperty(k)) out.push({ label: k, value: Math.round(map[k] * 100) / 100 });
    out.sort(function (a, b) { return b.value - a.value; });
    return out;
  }

  var salespeople = [];
  for (var pk in personInfo) if (personInfo.hasOwnProperty(pk)) {
    var pi = personInfo[pk];
    salespeople.push({
      label: pk,
      sales: Math.round(pi.sales * 100) / 100,
      orders: pi.orders,
      discounts: pi.discounts,
      discountAmt: Math.round(pi.discountAmt * 100) / 100
    });
  }
  salespeople.sort(function (a, b) { return b.sales - a.sales; });

  var byDiscount = toSortedArr((function () {
    var m = {};
    for (var k in personInfo) if (personInfo.hasOwnProperty(k) && personInfo[k].discountAmt > 0) m[k] = personInfo[k].discountAmt;
    return m;
  })());

  var topCustomers = toSortedArr(custMap).slice(0, 5);

  /* ---- Series (chart data) ---- */
  var labels = [], salesSeries = [], expSeries = [], targetSeries = [];
  var monthNames = ['Janaayo', 'Febraayo', 'Maarso', 'Abriil', 'Maajo', 'Juun', 'Luuliyo', 'Ogos', 'Sebteembar', 'Oktoobar', 'Noofeembar', 'Diseembar'];

  if (month > 0) {
    var dim = new Date(year, month, 0).getDate();
    var dSales = {}, dExp = {};
    for (i = 0; i < sRows.length; i++) if (sRows[i].completed) dSales[sRows[i].date.getDate()] = (dSales[sRows[i].date.getDate()] || 0) + sRows[i].net;
    for (i = 0; i < eRows.length; i++) dExp[eRows[i].date.getDate()] = (dExp[eRows[i].date.getDate()] || 0) + eRows[i].amount;
    for (var d = 1; d <= dim; d++) {
      labels.push(String(d));
      salesSeries.push(Math.round((dSales[d] || 0) * 100) / 100);
      expSeries.push(Math.round((dExp[d] || 0) * 100) / 100);
      targetSeries.push(dailyTarget);
    }
  } else {
    var mSales = {}, mExp = {};
    for (i = 0; i < sRows.length; i++) if (sRows[i].completed) mSales[sRows[i].date.getMonth()] = (mSales[sRows[i].date.getMonth()] || 0) + sRows[i].net;
    for (i = 0; i < eRows.length; i++) mExp[eRows[i].date.getMonth()] = (mExp[eRows[i].date.getMonth()] || 0) + eRows[i].amount;
    for (var m = 0; m < 12; m++) {
      labels.push(monthNames[m]);
      salesSeries.push(Math.round((mSales[m] || 0) * 100) / 100);
      expSeries.push(Math.round((mExp[m] || 0) * 100) / 100);
      targetSeries.push(monthlyTarget);
    }
  }

  var netSeries = [];
  for (var ni = 0; ni < salesSeries.length; ni++) netSeries.push(Math.round((salesSeries[ni] - expSeries[ni]) * 100) / 100);

  /* ---- Dalabaadka Dibadda / Overseas Orders - qeyb gooni ah, taxane bishii oo dhamaystiran sanadka ----
     - "Ordered" (lacagta la dalbaday) waxaa lagu xisaabiyaa bisha DALABKA (order date).
     - "Cargo + Other" waxaa lagu xisaabiyaa bisha AY ALAABTU TIMAADO (received date). */
  var pOrderedSeries = [0,0,0,0,0,0,0,0,0,0,0,0], pCargoSeries = [0,0,0,0,0,0,0,0,0,0,0,0];
  var yearOrdered = 0, yearCargo = 0;
  var pendingOrders = [], receivedOrders = [], cancelledOrders = [];
  var pendingTotal = 0;
  for (i = 0; i < poList.length; i++) {
    var ord = poList[i];
    var st = String(ord.status).toLowerCase();
    if (ord.date.getFullYear() === year && st !== 'cancelled') {
      pOrderedSeries[ord.date.getMonth()] = Math.round((pOrderedSeries[ord.date.getMonth()] + ord.total) * 100) / 100;
      yearOrdered += ord.total;
    }
    if (st === 'received' && ord.receivedDate && ord.receivedDate.getFullYear() === year) {
      var cm = ord.receivedDate.getMonth();
      pCargoSeries[cm] = Math.round((pCargoSeries[cm] + ord.cargo + ord.otherCosts) * 100) / 100;
      yearCargo += (ord.cargo + ord.otherCosts);
    }
    var orow = { id: ord.id, date: ord.dateStr, supplier: ord.supplier, notes: ord.notes,
      status: ord.status, receivedDate: ord.receivedDateStr, cargo: ord.cargo, otherCosts: ord.otherCosts,
      items: ord.items, total: ord.total, days: ord.days };
    if (st === 'pending') { pendingOrders.push(orow); pendingTotal += ord.total; }
    else if (st === 'received') receivedOrders.push(orow);
    else cancelledOrders.push(orow);
  }
  pendingOrders.sort(function (a, b) { return b.days - a.days; });
  receivedOrders.sort(function (a, b) { return (b.receivedDate || '').localeCompare(a.receivedDate || ''); });
  receivedOrders = receivedOrders.slice(0, 30);
  cancelledOrders = cancelledOrders.slice(0, 20);

  /* ---- Maalmaha Ugu Shaqada Badan (weekday) - saddex dhinac: Iibka, Dalabyada, Kharashka ---- */
  var wdNames = ['Axad','Isniin','Talaado','Arbaco','Khamiis','Jimce','Sabti'];
  var wdSales = [0,0,0,0,0,0,0], wdOrders = [0,0,0,0,0,0,0], wdExp = [0,0,0,0,0,0,0];
  for (i = 0; i < sRows.length; i++) {
    if (!sRows[i].completed) continue;
    var wd1 = sRows[i].date.getDay();
    wdSales[wd1] = Math.round((wdSales[wd1] + sRows[i].net) * 100) / 100;
    wdOrders[wd1]++;
  }
  for (i = 0; i < eRows.length; i++) {
    var wd2 = eRows[i].date.getDay();
    wdExp[wd2] = Math.round((wdExp[wd2] + eRows[i].amount) * 100) / 100;
  }
  var bestWeekday = wdNames[wdSales.indexOf(Math.max.apply(null, wdSales))];

  /* ---- Xogta ugu dhow (recent) ---- */
  var recentSales = sRows.filter(function (r) { return r.completed; })
    .sort(function (a, b) { return b.date - a.date; }).slice(0, 50)
    .map(function (r) {
      return { date: r.dateStr, invoiceId: r.invoiceId, customer: r.customer, salesperson: r.salesperson, district: r.district, net: r.net, balance: r.balance };
    });
  var recentExp = eRows.sort(function (a, b) { return b.date - a.date; }).slice(0, 50)
    .map(function (r) { return { date: r.dateStr, category: r.category, description: r.description, amount: r.amount, paidBy: r.paidBy }; });

  var periodLabel;
  if (day > 0) periodLabel = fmtDate_(new Date(year, month - 1, day));
  else if (month > 0) periodLabel = monthNames[month - 1] + ' ' + year;
  else periodLabel = 'Sanadka ' + year;

  /* ---- Analysis extras ---- */
  var bestMonth = { label: '-', value: 0 };
  var monthsWithData = 0, targetsAchieved = 0;
  if (month > 0) {
    monthsWithData = 1;
    targetsAchieved = totalSales >= monthlyTarget ? 1 : 0;
    bestMonth = { label: monthNames[month - 1], value: totalSales };
  } else {
    var mIdx;
    for (mIdx = 0; mIdx < 12; mIdx++) {
      var mv = mSales[mIdx] || 0;
      if (mv > 0) {
        monthsWithData++;
        if (mv >= monthlyTarget) targetsAchieved++;
        if (mv > bestMonth.value) bestMonth = { label: monthNames[mIdx], value: mv };
      }
    }
  }
  var todaySales = 0;
  var now = new Date();
  for (i = 0; i < sRows.length; i++) {
    if (sRows[i].completed && sRows[i].date.getFullYear() === now.getFullYear() &&
        sRows[i].date.getMonth() === now.getMonth() && sRows[i].date.getDate() === now.getDate()) {
      todaySales += sRows[i].net;
    }
  }

  return {
    periodLabel: periodLabel,
    kpi: {
      sales: Math.round(totalSales * 100) / 100,
      expenses: Math.round(totalExp * 100) / 100,
      net: Math.round(net * 100) / 100,
      profit30: Math.round(net * 100) / 100,
      orders: orders,
      avgDay: orders > 0 ? Math.round((totalSales / orders) * 100) / 100 : 0,
      margin: totalSales > 0 ? Math.round((net / totalSales) * 10000) / 100 : 0,
      dailyTarget: dailyTarget,
      monthlyTarget: monthlyTarget,
      profitRate: profitRate
    },
    labels: labels,
    salesSeries: salesSeries,
    expSeries: expSeries,
    targetSeries: targetSeries,
    netSeries: netSeries,
    byCategory: toSortedArr(catMap),
    salespeople: salespeople,
    byPerson: salespeople.map(function (p) { return { label: p.label, value: p.sales }; }),
    byDistrict: toSortedArr(distMap),
    byDiscount: byDiscount,
    byType: typeMap,
    topCustomers: topCustomers,
    analysis: {
      customers: Object.keys(newPhones).length + Object.keys(retPhones).length,
      newCount: Object.keys(newPhones).length,
      returningCount: Object.keys(retPhones).length,
      pending: pendingCount,
      completedOrders: orders,
      debt: Math.round(debtTotal * 100) / 100,
      debtCustomers: Object.keys(debtCustomers).length,
      discountTotal: Math.round(discountTotal * 100) / 100,
      discountOrders: discountOrders,
      delivery: Math.round(deliveryTotal * 100) / 100,
      deliveryOrders: deliveryOrders,
      avgOrder: orders > 0 ? Math.round((totalSales / orders) * 100) / 100 : 0,
      bestPerson: salespeople.length ? salespeople[0] : null,
      bestDistrict: toSortedArr(distMap)[0] || null,
      bestMonth: bestMonth,
      targetsAchieved: targetsAchieved,
      targetsTotal: monthsWithData,
      todaySales: Math.round(todaySales * 100) / 100,
      todayTarget: dailyTarget,
      todayPct: dailyTarget > 0 ? Math.round((todaySales / dailyTarget) * 100) : 0
    },
    recentSales: recentSales,
    recentExp: recentExp,
    weekday: { labels: wdNames, sales: wdSales, orders: wdOrders, expenses: wdExp, best: bestWeekday },
    orders: {
      monthLabels: monthNames,
      orderedSeries: pOrderedSeries,
      cargoSeries: pCargoSeries,
      yearOrdered: Math.round(yearOrdered * 100) / 100,
      yearCargo: Math.round(yearCargo * 100) / 100,
      pendingCount: pendingOrders.length,
      pendingTotal: Math.round(pendingTotal * 100) / 100,
      receivedCount: receivedOrders.length,
      pending: pendingOrders,
      received: receivedOrders,
      cancelled: cancelledOrders
    }
  };
}

/* ============================== WRITE / LOOKUP ============================== */

function findEmptyRow_(sh, fromRow, maxScan) {
  if (sh.getLastRow() < fromRow) return fromRow;
  var maxRow = Math.min(sh.getLastRow(), fromRow + maxScan);
  var vals = sh.getRange(fromRow, 1, maxRow - fromRow + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || '').trim();
    if (!v) return fromRow + i;
  }
  return maxRow + 1;
}

function addSale(token, s) {
  if (!auth_(token)) return { error: 'login' };
  s = s || {};
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_SALES);
  if (!sh) return { ok: false, msg: 'Sheet Sales ma jiro.' };
  var hdrRow = findHeaderRow_(sh, 'System Invoice ID', 10);
  if (hdrRow < 1) hdrRow = 4;
  var map = buildColMap_(sh, hdrRow, 33);
  function col(key, fb) { return map[String(key).toLowerCase()] || fb; }
  var d = toDate_(s.date);
  if (!d || !s.customer || !s.salesperson) return { ok: false, msg: 'Date, macmiil iyo iibiyaha waa qasab.' };
  var gross = num_(s.gross);
  var discType = String(s.discType || 'None');
  var discInput = num_(s.discInput);
  var discAmt = discType.indexOf('Percent') !== -1 ? Math.round(gross * discInput) / 100 : (discType.indexOf('Amount') !== -1 ? discInput : 0);
  var net = Math.max(0, Math.round((gross - discAmt) * 100) / 100);
  var totalPaid = Math.max(0, Math.min(num_(s.initial), net));
  var balance = Math.max(0, Math.round((net - totalPaid) * 100) / 100);
  var status = String(s.status || 'Completed');
  var completed = status.toLowerCase() === 'completed';
  var invoiceId = 'INV-' + d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2) + '-' + String(Math.floor(Math.random() * 9000) + 1000);
  var row = [
    invoiceId, String(s.paper || ''), d, String(s.customer), String(s.phone || ''),
    String(s.custType || 'New'), String(s.location || ''), '', '',
    String(s.salesperson), String(s.payMethod || 'Cash'), String(s.payStatus || 'Paid'),
    gross, discType, discType === 'None' ? '' : discInput, Math.round(discAmt * 100) / 100,
    net, totalPaid, 0, totalPaid, balance, s.dueDate ? toDate_(s.dueDate) : '',
    status, '', 'No', 0, 0, String(s.custType || 'New'),
    new Date(d.getFullYear(), d.getMonth(), 1), d.getFullYear(), completed ? net : 0, completed ? 1 : 0, ''
  ];
  var targetRow = findEmptyRow_(sh, hdrRow + 1, 2000);
  sh.getRange(targetRow, 1, 1, 33).setValues([row]);
  return { ok: true, invoiceId: invoiceId, msg: 'Iibka waa la kaydiyay: ' + invoiceId };
}

function addExpense(token, e) {
  if (!auth_(token)) return { error: 'login' };
  e = e || {};
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_EXPENSES);
  if (!sh) return { ok: false, msg: 'Sheet Expenses ma jiro.' };
  var hdrRow = findHeaderRow_(sh, 'Expense ID', 10);
  if (hdrRow < 1) hdrRow = 4;
  var d = toDate_(e.date);
  if (!d || !e.category || num_(e.amount) <= 0) return { ok: false, msg: 'Date, qaybta iyo qadarka waa qasab.' };
  var expId = 'EXP-' + (Math.floor(Math.random() * 900000) + 100000);
  var row = [
    expId, d, String(e.category), String(e.newCat || ''), String(e.description || ''),
    num_(e.amount), String(e.paidBy || 'Admin'), String(e.payMethod || 'Cash'),
    new Date(d.getFullYear(), d.getMonth(), 1), d.getFullYear()
  ];
  var targetRow = findEmptyRow_(sh, hdrRow + 1, 2000);
  sh.getRange(targetRow, 1, 1, 10).setValues([row]);
  return { ok: true, msg: 'Kharashka waa la kaydiyay.' };
}

function createOrder(token, o) {
  if (!auth_(token)) return { error: 'login' };
  o = o || {};
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  var shI = ss.getSheetByName(SHEET_ORDER_ITEMS);
  if (!sh || !shI) { setupOrders(); sh = ss.getSheetByName(SHEET_ORDERS); shI = ss.getSheetByName(SHEET_ORDER_ITEMS); }
  var d = toDate_(o.date);
  if (!d) return { ok: false, msg: 'Taariikhda dalabka waa qasab.' };
  var items = o.items || [];
  var cleanItems = [];
  for (var k = 0; k < items.length; k++) {
    var it = items[k];
    var qty = num_(it.qty), unitCost = num_(it.unitCost);
    var product = String(it.product || '').trim();
    if (!product || !(qty > 0)) continue;
    cleanItems.push({ product: product, qty: qty, unitCost: unitCost, total: Math.round(qty * unitCost * 100) / 100 });
  }
  if (!cleanItems.length) return { ok: false, msg: 'Ugu yaraan hal alaab (Product + Qty) waa qasab.' };
  var id = 'ORD-' + d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2) + '-' + String(Math.floor(Math.random() * 9000) + 1000);
  /* Xilliga la dalbanayo: cargo iyo kharashaadka kale lama gelinayo weli — waxay imanayaan marka alaabtu timaado (Receive). */
  sh.appendRow([id, d, String(o.supplier || ''), String(o.notes || ''), 'Pending', '', 0, 0,
    new Date(d.getFullYear(), d.getMonth(), 1), d.getFullYear()]);
  for (var m = 0; m < cleanItems.length; m++) {
    var ci = cleanItems[m];
    shI.appendRow([id, ci.product, ci.qty, ci.unitCost, ci.total]);
  }
  return { ok: true, id: id, msg: 'Dalabka waa la kaydiyay: ' + id };
}

function receiveOrder(token, id, cargo, otherCosts) {
  if (!auth_(token)) return { error: 'login' };
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return { ok: false, msg: 'Orders sheet ma jiro.' };
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][O.id - 1]) === String(id)) {
      sh.getRange(i + 1, O.status).setValue('Received');
      sh.getRange(i + 1, O.receivedDate).setValue(new Date());
      sh.getRange(i + 1, O.cargo).setValue(num_(cargo));
      sh.getRange(i + 1, O.otherCosts).setValue(num_(otherCosts));
      return { ok: true, msg: 'Dalabka waa la xaqiijiyay in la heshay — cargo iyo kharashaadka kale waa la duubay.' };
    }
  }
  return { ok: false, msg: 'Dalabka lama helin.' };
}

function cancelOrder(token, id) {
  if (!auth_(token)) return { error: 'login' };
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return { ok: false, msg: 'Orders sheet ma jiro.' };
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][O.id - 1]) === String(id)) {
      sh.getRange(i + 1, O.status).setValue('Cancelled');
      return { ok: true, msg: 'Dalabka waa la joojiyay.' };
    }
  }
  return { ok: false, msg: 'Dalabka lama helin.' };
}

/* ============================== SCAN INVOICE (GEMINI VISION - HEERSARE) ============================== */

function saveGeminiKey(token, key) {
  if (!auth_(token)) return { error: 'login' };
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', String(key || '').trim());
  return { ok: true, msg: 'Gemini API Key waa la kaydiyay.' };
}

function hasGeminiKey(token) {
  if (!auth_(token)) return { error: 'login' };
  var k = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return { ok: true, has: !!(k && k.length > 10) };
}

function scanInvoiceImage(token, base64, mimeType) {
  if (!auth_(token)) return { error: 'login' };
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, msg: 'Gemini API Key lama helin. Tag Settings → geli API Key-ga si scan-ku heersare ugu shaqeeyo.' };

  var prompt = 'Waxaad tahay khabiir ku takhasusay akhrinta invoice-yada iyo receipt-yada la qoro gacanta (Somali/English handwriting, khad qalin). ' +
    'Sawirka soo socda waa invoice ganacsi. Akhri si taxadar leh, gaar ahaan tirooyinka (numbers) ee saf kasta ee alaabta (Qty, U.Price, Total) ' +
    'iyo qiimayaasha hoose (PRICE, DISCOUNT, TOTAL, DELIVERY). Soo celi JSON KELIYA — ma jirto sharaxaad, ma jirto ``` code fence — qaab sax ah:\n' +
    '{"customer":"","address":"","salesperson":"","date":"YYYY-MM-DD","phone":"","items":[{"name":"","qty":0,"unitPrice":0,"total":0}],' +
    '"price":0,"discountType":"percent|amount|none","discountValue":0,"delivery":0,"total":0}\n' +
    'Haddii goobi maqan tahay ka dhig 0 (tiro) ama "" (qoraal). "discountType" waa "percent" haddii lambarka discount-ku leeyahay "%", ' +
    '"amount" haddii uu lacag toos ah yahay, "none" haddii aan discount jirin. "salesperson" waa magaca qofka iibiyaha haddii rasiidhku ku qoran yahay, haddii kale "".';

  var payload = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }
    ] }],
    generationConfig: { temperature: 0.1 }
  };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(key);
  var res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, msg: 'Khalad internet/Gemini: ' + e.message };
  }
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200) return { ok: false, msg: 'Gemini khalad (' + code + '): ' + body.slice(0, 250) };

  var json;
  try { json = JSON.parse(body); } catch (e) { return { ok: false, msg: 'Gemini jawaab qaldan (JSON).' }; }
  var text = '';
  try { text = json.candidates[0].content.parts[0].text; } catch (e) { return { ok: false, msg: 'Gemini ma soo celin natiijo la akhrin karo — sawirka mar kale tijaabi.' }; }
  text = text.replace(/```json/gi, '').replace(/```/g, '').trim();
  var data;
  try { data = JSON.parse(text); } catch (e) { return { ok: false, msg: 'Ma fahmin natiijada Gemini.', raw: text.slice(0, 300) }; }

  /* ---- Hubinta xisaabta (verification): Qty × U.Price = Total (saf kasta), Total-yada = Price, Price - Discount + Delivery ?= Total qoran ---- */
  var items = data.items || [];
  var itemMismatches = [];
  var itemsSum = 0;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var qty = num_(it.qty), unitPrice = num_(it.unitPrice), written = num_(it.total);
    var expected = Math.round(qty * unitPrice * 100) / 100;
    itemsSum += (written > 0 ? written : expected);
    if (Math.abs(expected - written) > 0.05) {
      itemMismatches.push({ row: i + 1, name: it.name, qty: qty, unitPrice: unitPrice, writtenTotal: written, expectedTotal: expected });
    }
  }
  itemsSum = Math.round(itemsSum * 100) / 100;
  var price = num_(data.price);
  var priceMismatch = null;
  if (price > 0 && Math.abs(itemsSum - price) > 0.05) priceMismatch = { itemsSum: itemsSum, writtenPrice: price };

  var base = price > 0 ? price : itemsSum;
  var discType = String(data.discountType || 'none').toLowerCase();
  var discVal = num_(data.discountValue);
  var delivery = num_(data.delivery);
  var writtenTotal = num_(data.total);

  /* Isku day labada nooc ee discount-ka (percent/amount) - si aan u ogaano midka la mid ah Total-ka qoran, si loo hubiyo in discount-ka la ilaawin (proceso). */
  var asPercent = Math.round((base - base * discVal / 100) * 100) / 100;
  var asAmount = Math.round((base - discVal) * 100) / 100;
  var matched = false, matchedLabel = discType;
  [{ l: 'percent', v: asPercent }, { l: 'amount', v: asAmount }, { l: 'none', v: base }].forEach(function (c) {
    if (Math.abs(Math.round((c.v + delivery) * 100) / 100 - writtenTotal) <= 0.05 || Math.abs(c.v - writtenTotal) <= 0.05) {
      matched = true; matchedLabel = c.l;
    }
  });
  var totalMismatch = null;
  if (writtenTotal > 0 && !matched) {
    totalMismatch = { writtenTotal: writtenTotal, base: base, asPercentGuess: asPercent, asAmountGuess: asAmount, delivery: delivery };
  }

  return {
    ok: true,
    data: data,
    verification: {
      itemMismatches: itemMismatches,
      priceMismatch: priceMismatch,
      totalMismatch: totalMismatch,
      discountGuess: matched ? matchedLabel : discType,
      clean: itemMismatches.length === 0 && !priceMismatch && !totalMismatch
    }
  };
}

function scanExpenseImage(token, base64, mimeType) {
  if (!auth_(token)) return { error: 'login' };
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, msg: 'Gemini API Key lama helin. Tag Settings → geli API Key-ga si scan-ku heersare ugu shaqeeyo.' };

  var prompt = 'Waxaad tahay khabiir ku takhasusay akhrinta rasiidhada kharashaadka (Somali/English). ' +
    'Sawirka soo socda waa rasiidh kharash (receipt/expense). Akhri si taxadar leh: taariikhda, qaybta/category-ga, sharaxaad, qadarka/lacagta, magaca meesha/dhaxasho, iyo habka lacag bixinta. ' +
    'Sooceli JSON KELIYA — ma jirto sharaxaad, ma jirto ``` code fence — qaab sax ah:\n' +
    '{"date":"YYYY-MM-DD","category":"","description":"","amount":0,"vendor":"","payMethod":"Cash|EVC Plus|Bank|Card"}\n' +
    'Haddii goobi maqan tahay ka dhig 0 (tiro) ama "" (qoraal). Category waa mid ka mid ah: Kirada, Shaqaalaha, Gaadiidka, Korontada, Xayeysiiska, Qalabka, Office Supplies, Cashuur, Other.';

  var payload = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }
    ] }],
    generationConfig: { temperature: 0.1 }
  };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(key);
  var res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, msg: 'Khalad internet/Gemini: ' + e.message };
  }
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200) return { ok: false, msg: 'Gemini khalad (' + code + '): ' + body.slice(0, 250) };

  var json;
  try { json = JSON.parse(body); } catch (e) { return { ok: false, msg: 'Gemini jawaab qaldan (JSON).' }; }
  var text = '';
  try { text = json.candidates[0].content.parts[0].text; } catch (e) { return { ok: false, msg: 'Gemini ma soo celin natiijo la akhrin karo — sawirka mar kale tijaabi.' }; }
  text = text.replace(/```json/gi, '').replace(/```/g, '').trim();
  var data;
  try { data = JSON.parse(text); } catch (e) { return { ok: false, msg: 'Ma fahmin natiijada Gemini.', raw: text.slice(0, 300) }; }

  return { ok: true, data: data };
}

function getReceipt(token, q) {
  if (!auth_(token)) return { error: 'login' };
  q = String(q || '').trim().toLowerCase();
  if (!q) return { ok: false, msg: 'Geli System Invoice ID ama Paper Invoice No.' };
  var sales = readSales_();
  for (var i = 0; i < sales.length; i++) {
    var s = sales[i];
    if (s.invoiceId.toLowerCase() === q || s.paper.toLowerCase() === q) {
      return { ok: true, r: {
        invoiceId: s.invoiceId, paper: s.paper, date: s.dateStr, customer: s.customer,
        phone: s.phone, location: s.location, gross: s.gross, discount: s.discount,
        net: s.net, totalPaid: s.totalPaid, balance: s.balance, status: s.status,
        salesperson: s.salesperson
      } };
    }
  }
  return { ok: false, msg: 'Invoice lama helin.' };
}

function searchCustomer(token, q) {
  if (!auth_(token)) return { error: 'login' };
  q = String(q || '').trim().toLowerCase();
  if (!q) return { ok: true, items: [] };
  var sales = readSales_();
  var out = [];
  for (var i = 0; i < sales.length; i++) {
    var s = sales[i];
    if (s.customer.toLowerCase().indexOf(q) !== -1 || s.phone.toLowerCase().indexOf(q) !== -1) {
      out.push({ date: s.dateStr, invoiceId: s.invoiceId, customer: s.customer, phone: s.phone, net: s.net, balance: s.balance, status: s.status });
    }
  }
  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  return { ok: true, items: out.slice(0, 20) };
}

/* ============================== SETTINGS ============================== */

function changePassword(token, oldPw, newPw) {
  if (!auth_(token)) return { error: 'login' };
  var u = auth_(token);
  oldPw = String(oldPw || '');
  newPw = String(newPw || '');
  if (!oldPw || newPw.length < 5) return { ok: false, msg: 'Password-ka cusub waa inuu ahaadaa ugu yaraan 5 xaraf.' };
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) return { ok: false, msg: 'Users sheet ma jiro.' };
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === String(u.user).toLowerCase()) {
      if (String(vals[i][1]) !== oldPw) return { ok: false, msg: 'Password-ka hadda waa khalad.' };
      sh.getRange(i + 1, 2).setValue(newPw);
      return { ok: true, msg: 'Password waa la badalay.' };
    }
  }
  return { ok: false, msg: 'User lama helin.' };
}

function saveSettings(token, s) {
  if (!auth_(token)) return { error: 'login' };
  var u = auth_(token);
  if (String(u.role).toLowerCase() !== 'admin') return { ok: false, msg: 'Admin kaliya ayaa settings-ka badali kara.' };
  s = s || {};
  var props;
  try { props = PropertiesService.getScriptProperties(); } catch (e) { return { ok: false, msg: 'Settings lama helin (PropertiesService).' }; }
  var name = String(s.name || '').trim();
  if (name) props.setProperty('BIZ_NAME', name);
  var logo = String(s.logo || '');
  if (logo) {
    if (logo.length > 8000) return { ok: false, msg: 'Sawirka aad buu u weyn yahay — isku day sawar yar (96px).' };
    props.setProperty('BIZ_LOGO', logo);
  } else if (s.clearLogo) {
    props.deleteProperty('BIZ_LOGO');
  }
  if (s.theme) props.setProperty('THEME', String(s.theme));
  return { ok: true, msg: 'Settings waa la kaydiyay.' };
}
function fasaxBixi() {
  UrlFetchApp.fetch("https://www.google.com");
}

/* ================= AI Chat (xog-wareeji + xisaab-xaqiijin) =================
   Waxay isticmaashaa xogta KPI ee dashboard-ka ay bixiso (context),
   taariikhda sheekaysiga (history), iyo GEMINI_API_KEY (Script Properties). */
function aiChat(token, message, context, history) {
  if (!auth_(token)) return { error: 'login' };
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { reply: 'Gemini API Key lama helin. Fur Settings → geli API Key-ga, ama isticmaal Vercel deployment-ka (GEMINI_API_KEY env).' };

  var sysPrompt = 'Waxaad tahay AI assistant-ka ganacsiga Happy Home (alaab guri iibisa). Ka jawaab su\'aalaha xogta ganacsiga, Somali ama English. ' +
    'Xisaabta sax: Faa\'iidada = Iibka - Kharashka; Margin = Faa\'iidada / Iibka * 100; Net = Gross - Discount; Balance = Net - TotalPaid. ' +
    'Tirooyinka la siyo isticmaal haddii xog la heli karo; haddii xisaab qaldan la arko, sax oo sharax sababta. Jawaab kooban oo faahfaahsan qor.';

  var contents = [];
  (history || []).slice(-10).forEach(function (h) {
    if (h && h.text) {
      contents.push({ role: (h.role === 'assistant' ? 'model' : 'user'), parts: [{ text: String(h.text).slice(0, 2000) }] });
    }
  });
  var userText = (context ? 'XOGTA GANACSIGA HADDA:\n' + context + '\n\nSU\'AASHA: ' : '') + String(message || '');
  contents.push({ role: 'user', parts: [{ text: userText }] });

  var payload = {
    system_instruction: { parts: [{ text: sysPrompt }] },
    contents: contents,
    generationConfig: { temperature: 0.3, maxOutputTokens: 1024 }
  };

  try {
    var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(key), {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { reply: 'Khalad Gemini: ' + res.getContentText().slice(0, 200) };
    var j = JSON.parse(res.getContentText());
    var text = '';
    try { text = j.candidates[0].content.parts[0].text; } catch (err) { return { reply: 'Gemini ma soo celin jawaab cad — mar kale isku day.' }; }
    return { reply: text };
  } catch (e2) {
    return { reply: 'Khalad: ' + e2.message };
  }
}