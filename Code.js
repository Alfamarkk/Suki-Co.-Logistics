/**
 * Suki Co. MVP backend (Google Apps Script)
 *
 * SETUP
 *  1. Create a NEW Google Sheet > Extensions > Apps Script (binds the script to the Sheet).
 *  2. Run setupDatabase() once from the editor and approve the permissions (Sheets + Drive).
 *  3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *  4. After every code change: Deploy > Manage deployments > Edit > New version.
 *
 * Functions whose names end with "_" are private. The browser cannot call them.
 * Every other function (except the four public ones) needs a session token from login().
 *
 * FILES in the Apps Script editor (names must match exactly, without .html):
 *   Code.gs, Index, Style, ViewSuki, ViewPartner, ViewToda, ViewAdmin
 */


const TZ = 'Asia/Manila';

// Edit these lists to match your barangay. They feed the apply forms and the TODA route matching.
const AREAS = ['Tulay na Bato', 'Purok 1', 'Purok 2', 'Purok 3', 'Purok 4', 'Purok 5'];
const CATEGORIES = ['Order Issues', 'Payment / Transaction', 'Delivery Issues', 'Account / Login', 'App / System Bug', 'Partner Complaint'];
const ROLE_LABEL = { suki: 'Suki', grocery: 'Partner', toda: 'TODA', admin: 'Developer' };

const SCHEMA = {
  Users:        ['UserID', 'Username', 'PasswordHash', 'Name', 'Role', 'Phone', 'Area', 'Address', 'Status', 'DateRegistered'],
  Groceries:    ['GroceryID', 'Name', 'OwnerUserID', 'Area', 'Address', 'Phone', 'GcashNo', 'GcashName', 'PhotoID', 'Status'],
  Products:     ['ProductID', 'GroceryID', 'Name', 'Price', 'Stock'],
  Riders:       ['RiderID', 'Name', 'TodaName', 'PlateNo', 'Phone', 'Route', 'PhotoID', 'Status'],
  Orders:       ['OrderID', 'UserID', 'GroceryID', 'Items', 'Total', 'Status', 'RiderID', 'Area', 'Address',
                 'PaymentStatus', 'ReceiptID', 'PayRef', 'ProofID', 'Timestamp', 'DeliveredAt'],
  Applications: ['AppID', 'Type', 'Name', 'Username', 'PasswordHash', 'Phone', 'Area', 'Address', 'Business',
                 'GcashNo', 'GcashName', 'Route', 'TodaName', 'PlateNo', 'PhotoID', 'Status', 'Date'],
  Reports:      ['ReportID', 'Category', 'ReportedByID', 'ReporterRole', 'Description', 'OrderID', 'Date', 'Status', 'DevNotes']
};

// Stored as plain text so phone numbers keep their leading 0.
const TEXT_COLS = ['Username', 'PasswordHash', 'Phone', 'GcashNo', 'PayRef', 'Route', 'PlateNo'];

/* ================================================================== */
/* Page                                                                */
/* ================================================================== */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Suki Co.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGet(request){
  return HtmlService.createTemplateFromFile('Index').evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/* ================================================================== */
/* Setup (run once from the editor)                                    */
/* ================================================================== */

function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());
  if (!props.getProperty('DEV_KEY')) props.setProperty('DEV_KEY', 'SARI-DEV-2026');

  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(name => {
    const cols = SCHEMA[name];
    let sh = ss.getSheetByName(name);
    if (sh) {
      // A tab with the same name but different columns (from an older version) is kept and renamed, never deleted.
      const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].filter(String);
      if (head.join('|') !== cols.join('|')) {
        sh.setName('old_' + name + '_' + Utilities.formatDate(new Date(), TZ, 'MMdd_HHmm'));
        sh = null;
      }
    }
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#c9d1dc');
    sh.setFrozenRows(1);
    cols.forEach((c, i) => {
      if (TEXT_COLS.indexOf(c) > -1) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });

  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  const seed = seedData_();
  Object.keys(seed).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh.getLastRow() < 2) sh.getRange(2, 1, seed[name].length, SCHEMA[name].length).setValues(seed[name]);
  });

  Logger.log('Suki Co. is ready.\n' +
    'Developer sign-in: username "developer", password "sari2026", access key "' + props.getProperty('DEV_KEY') + '".\n' +
    'Demo accounts (password 1234): suki1, grocery1, toda1.\n' +
    'Change the developer password and DEV_KEY (Project Settings > Script properties) before real use.');
}

function daysAgo_(n) { return new Date(Date.now() - n * 86400000); }

function seedData_() {
  const pw = hash_('1234');
  return {
    Users: [
      ['U01', 'developer', hash_('sari2026'), 'Developer',            'admin',   '',            '',              '',                  'Active', daysAgo_(40)],
      ['U02', 'suki1',     pw, 'Maria Santos',        'suki',    '09171112222', 'Purok 3',       'Blk 2 Lot 5',       'Active', daysAgo_(12)],
      ['U03', 'suki2',     pw, 'Pedro Reyes',         'suki',    '09177778888', 'Purok 1',       'Beside the chapel', 'Active', daysAgo_(5)],
      ['U04', 'grocery1',  pw, 'Kim Gonzales',        'grocery', '09173334444', 'Tulay na Bato', 'Near the bridge',   'Active', daysAgo_(30)],
      ['U05', 'toda1',     pw, 'Juan Dela Cruz',      'toda',    '09175556666', 'Purok 1',       'Purok 1',           'Active', daysAgo_(25)],
      ['U06', 'grocery2',  pw, 'Tomas Reyes',         'grocery', '09179990000', 'Purok 2',       'Corner of Rizal St', 'Active', daysAgo_(28)],
      ['U07', 'toda2',     pw, 'Ana Lopez',           'toda',    '09171213434', 'Purok 4',       'Purok 4',           'Active', daysAgo_(20)]
    ],
    Groceries: [
      ['G01', 'Gonzales Supermarket', 'U04', 'Tulay na Bato', 'Near the bridge',    '09173334444', '09173334444', 'K. Gonzales', '', 'Active'],
      ['G02', 'Mang Tomas Store',     'U06', 'Purok 2',       'Corner of Rizal St', '09179990000', '09179990000', 'T. Reyes',    '', 'Active']
    ],
    Products: [
      ['P01', 'G01', 'Rice (5kg)',          260, 40],
      ['P02', 'G01', 'Eggs (tray)',         210, 25],
      ['P03', 'G01', 'Cooking oil (1L)',     95, 30],
      ['P04', 'G01', 'Canned sardines',      28, 100],
      ['P05', 'G01', 'Instant noodles',      14, 200],
      ['P06', 'G02', 'Sugar (1kg)',          72, 50],
      ['P07', 'G02', 'Coffee 3-in-1 (10s)',  45, 60],
      ['P08', 'G02', 'Laundry soap',         38, 40],
      ['P09', 'G02', 'Bottled water (1L)',   25, 80]
    ],
    Riders: [
      ['U05', 'Juan Dela Cruz', 'Tulay na Bato TODA', 'Tricycle 014', '09175556666', 'Tulay na Bato, Purok 1, Purok 3', '', 'Active'],
      ['U07', 'Ana Lopez',      'Tulay na Bato TODA', 'Tricycle 027', '09171213434', 'Purok 2, Purok 4, Purok 5',        '', 'Active']
    ],
    Orders: [
      [1001, 'U02', 'G01', 'Rice (5kg) x1, Eggs (tray) x1', 470, 'Delivered', 'U05', 'Purok 3', 'Blk 2 Lot 5',
       'Verified', '', '', '', daysAgo_(2), daysAgo_(2)],
      [1002, 'U03', 'G02', 'Sugar (1kg) x2, Coffee 3-in-1 (10s) x3', 279, 'Delivered', 'U05', 'Purok 1', 'Beside the chapel',
       'Verified', '', '', '', daysAgo_(1), daysAgo_(1)]
    ],
    Applications: [
      ['A01', 'suki',    'Liza Mendoza',    'liza',    pw, '09175657878', 'Purok 5',       'Near the basketball court', '', '', '', '', '', '', '', 'Pending', daysAgo_(1)],
      ['A02', 'grocery', 'Rosa Villanueva', 'rosav',   pw, '09170001111', 'Purok 4',       'Main road',  'Villanueva Mini Mart', '09170001111', 'R. Villanueva', '', '', '', '', 'Pending', daysAgo_(1)],
      ['A03', 'toda',    'Mario Bautista',  'mario',   pw, '09172223333', 'Purok 2',       'Purok 2',    '', '', '', 'Purok 2, Purok 3', 'Tulay na Bato TODA', 'Tricycle 033', '', 'Pending', daysAgo_(0)]
    ],
    Reports: [
      ['R01', 'Delivery Issues',  'U02', 'suki',    'The rider arrived 40 minutes late.',                           1001, daysAgo_(1), 'New',       ''],
      ['R02', 'App / System Bug', 'U03', 'suki',    'My order did not show up after placing it. I had to refresh.', '',   daysAgo_(2), 'In Review', 'Checking the refresh interval.'],
      ['R03', 'Payment / Transaction', 'U04', 'grocery', 'The GCash receipt amount did not match the order total.', 1002, daysAgo_(3), 'Resolved',  'Checkout now shows the exact amount to send.']
    ]
  };
}

/* ================================================================== */
/* Private helpers                                                     */
/* ================================================================== */

function bad_(msg) { return { ok: false, msg: msg }; }

function hash_(pw) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || 'sari';
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8));
}

function fmt_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'MMM d, h:mm a') : v; }
function fmtDay_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'MMM d, yyyy') : v; }

// Reads a tab into an array of objects keyed by header. Dates become text unless raw is true,
// because google.script.run cannot send Date objects to the browser.
function readAll_(name, raw) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '". Run setupDatabase() first.');
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((k, i) => { o[k] = raw ? r[i] : fmt_(r[i]); });
    return o;
  });
}

function index_(rows, key) {
  const m = {};
  rows.forEach(r => { m[r[key]] = r; });
  return m;
}

function append_(name, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(head.map(k => obj[k] === undefined ? '' : obj[k]));
}

function updateRow_(name, id, changes) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const v = sh.getDataRange().getValues();
  const head = v[0];
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      Object.keys(changes).forEach(k => {
        const c = head.indexOf(k);
        if (c > -1) sh.getRange(i + 1, c + 1).setValue(changes[k]);
      });
      return true;
    }
  }
  return false;
}

// prefix '' gives plain numbers from `start` (orders). Otherwise U08, G03, R04, and so on.
function nextId_(name, prefix, start) {
  const first = SCHEMA[name][0];
  const nums = readAll_(name).map(r => parseInt(String(r[first]).replace(/\D/g, ''), 10) || 0);
  const n = Math.max.apply(null, nums.concat([start - 1])) + 1;
  return prefix ? prefix + String(n).padStart(2, '0') : n;
}

function lc_(s) { return String(s || '').trim().toLowerCase(); }

function routeHas_(route, area) {
  return String(route || '').split(',').map(s => s.trim()).indexOf(area) > -1;
}

/* ---- sessions ---------------------------------------------------- */

function newSession_(data) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, JSON.stringify(data), 21600); // 6 hours
  return token;
}

function auth_(token, roles) {
  const raw = token && CacheService.getScriptCache().get('s:' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  const s = JSON.parse(raw);
  if (roles && roles.indexOf(s.role) === -1) throw new Error('NOT_ALLOWED');
  return s;
}

/* ---- images (Drive) ---------------------------------------------- */

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('UPLOAD_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ } }
  const f = DriveApp.createFolder('SariExpress_Uploads');
  props.setProperty('UPLOAD_FOLDER', f.getId());
  return f;
}

function imgError_(dataUrl, what) {
  const m = String(dataUrl || '').match(/^data:image\/(jpeg|png|webp);base64,(.+)$/);
  if (!m) return 'Attach ' + what + ' (JPG or PNG).';
  if (m[2].length > 4000000) return what.charAt(0).toUpperCase() + what.slice(1) + ' is too large. Pick a smaller photo.';
  return '';
}

function saveImage_(dataUrl, label) {
  const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], label + '-' + Date.now());
  const f = folder_().createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain may block link sharing */ }
  return f.getId();
}

/* ---- orders ------------------------------------------------------ */

function decorate_(orders) {
  const users = index_(readAll_('Users'), 'UserID');
  const groc = index_(readAll_('Groceries'), 'GroceryID');
  const riders = index_(readAll_('Riders'), 'RiderID');
  return orders.map(o => {
    const u = users[o.UserID] || {};
    const g = groc[o.GroceryID] || {};
    const r = riders[o.RiderID] || {};
    return {
      id: o.OrderID, userId: o.UserID, customer: u.Name || '', phone: u.Phone || '',
      address: o.Address, area: o.Area,
      groceryId: o.GroceryID, store: g.Name || '', storeArea: g.Area || '', storeAddress: g.Address || '',
      items: o.Items, total: Number(o.Total), status: o.Status,
      riderId: o.RiderID, rider: r.Name || '', riderPhone: r.Phone || '',
      payment: o.PaymentStatus, receiptId: o.ReceiptID, payRef: o.PayRef, proofId: o.ProofID,
      time: o.Timestamp, deliveredAt: o.DeliveredAt
    };
  });
}

function ordersWhere_(test) {
  return decorate_(readAll_('Orders').filter(test)).reverse(); // newest first
}

// Riders do not need the customer's payment receipt.
function forRider_(list) {
  return list.map(o => { o.receiptId = ''; o.payRef = ''; return o; });
}

function restock_(order) {
  const prods = readAll_('Products').filter(p => p.GroceryID === order.GroceryID);
  String(order.Items).split(', ').forEach(part => {
    const m = part.match(/^(.*) x(\d+)$/);
    if (!m) return;
    const p = prods.find(x => x.Name === m[1]);
    if (p) updateRow_('Products', p.ProductID, { Stock: Number(p.Stock) + Number(m[2]) });
  });
}

function appRow_(a) {
  return {
    id: a.AppID, type: a.Type, name: a.Name, username: a.Username, phone: a.Phone, area: a.Area, address: a.Address,
    business: a.Business, gcashNo: a.GcashNo, gcashName: a.GcashName, route: a.Route,
    todaName: a.TodaName, plateNo: a.PlateNo, photoId: a.PhotoID, status: a.Status, date: a.Date
  };
}

/* ================================================================== */
/* Public (no session): config, apply, sign in                         */
/* ================================================================== */

function getConfig() {
  return { areas: AREAS, categories: CATEGORIES };
}

function submitApplication(a) {
  const type = a.type;
  if (['suki', 'grocery', 'toda'].indexOf(type) === -1) return bad_('Choose how you want to apply.');
  const clean = s => String(s || '').trim().slice(0, 120);
  const phoneOk = p => /^09\d{9}$/.test(p);
  const name = clean(a.name);
  const username = lc_(clean(a.username));
  const phone = clean(a.phone).replace(/[\s-]/g, '');
  const area = clean(a.area);
  const address = clean(a.address);

  if (!name) return bad_('Enter your full name.');
  if (!phoneOk(phone)) return bad_('Enter an 11-digit mobile number that starts with 09.');
  if (AREAS.indexOf(area) === -1) return bad_('Choose your area.');
  if (!/^[a-z0-9._]{4,20}$/.test(username)) return bad_('Username must be 4 to 20 letters, numbers, dots or underscores.');
  if (String(a.password || '').length < 6) return bad_('Password must be at least 6 characters.');

  const app = {
    Type: type, Name: name, Username: username, PasswordHash: hash_(a.password), Phone: phone,
    Area: area, Address: address, Status: 'Pending', Date: new Date()
  };

  if (type === 'suki') {
    if (!address) return bad_('Enter your address or a landmark near your house.');
  }
  if (type === 'grocery') {
    const gno = clean(a.gcashNo).replace(/[\s-]/g, '');
    if (!clean(a.business)) return bad_('Enter your store name.');
    if (!address) return bad_('Enter your store address.');
    if (!phoneOk(gno)) return bad_('Enter the 11-digit GCash number that customers will pay to.');
    if (!clean(a.gcashName)) return bad_('Enter the name on your GCash account.');
    const err = imgError_(a.photo, 'a photo of your store');
    if (err) return bad_(err);
    app.Business = clean(a.business); app.GcashNo = gno; app.GcashName = clean(a.gcashName);
  }
  if (type === 'toda') {
    const route = (a.route || []).filter(r => AREAS.indexOf(r) > -1);
    if (!clean(a.todaName)) return bad_('Enter your TODA name.');
    if (!clean(a.plateNo)) return bad_('Enter your tricycle or plate number.');
    if (!route.length) return bad_('Choose at least one area on your route.');
    const err = imgError_(a.photo, 'a photo of you and your tricycle');
    if (err) return bad_(err);
    app.TodaName = clean(a.todaName); app.PlateNo = clean(a.plateNo); app.Route = route.join(', ');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const taken = readAll_('Users').some(u => lc_(u.Username) === username) ||
      readAll_('Applications').some(x => lc_(x.Username) === username && x.Status !== 'Declined');
    if (taken) return bad_('That username is already taken. Try another one.');
    if (app.PhotoID === undefined && a.photo) app.PhotoID = saveImage_(a.photo, 'app-' + type);
    app.AppID = nextId_('Applications', 'A', 1);
    append_('Applications', app);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function login(username, password, role) {
  const uname = lc_(username);
  const h = hash_(password);
  const user = readAll_('Users').find(x => lc_(x.Username) === uname && x.PasswordHash === h);

  if (user) {
    if (user.Role === 'admin') return bad_('Developer accounts sign in from Developer access.');
    if (user.Role !== role) return bad_('This is a ' + ROLE_LABEL[user.Role] + ' account. Go back and choose ' + ROLE_LABEL[user.Role] + '.');
    if (user.Status !== 'Active') return bad_('This account is suspended. Contact the developer team.');
    const g = user.Role === 'grocery' ? readAll_('Groceries').find(x => x.OwnerUserID === user.UserID) : null;
    const s = { id: user.UserID, role: user.Role, name: g ? g.Name : user.Name, groceryId: g ? g.GroceryID : '' };
    return { ok: true, token: newSession_(s), id: s.id, role: s.role, name: s.name, groceryId: s.groceryId, owner: user.Name };
  }

  const app = readAll_('Applications').find(x => lc_(x.Username) === uname && x.PasswordHash === h);
  if (app && app.Status === 'Pending') {
    return { ok: false, pending: true, msg: 'Your application is still under review. You can sign in once the developer team approves it.' };
  }
  if (app && app.Status === 'Declined') {
    return bad_('Your application was declined. You can apply again with the same username.');
  }
  return bad_('Wrong username or password.');
}

function adminLogin(username, password, key) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + lc_(username);
  if (Number(cache.get(failKey) || 0) >= 5) return bad_('Too many attempts. Wait 10 minutes and try again.');
  const devKey = PropertiesService.getScriptProperties().getProperty('DEV_KEY');
  const user = readAll_('Users').find(x => x.Role === 'admin' && lc_(x.Username) === lc_(username) && x.PasswordHash === hash_(password));
  if (!user || !devKey || String(key).trim() !== devKey) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
    return bad_('Sign-in details are incorrect.');
  }
  cache.remove(failKey);
  const s = { id: user.UserID, role: 'admin', name: user.Name, groceryId: '' };
  return { ok: true, token: newSession_(s), id: s.id, role: 'admin', name: s.name, groceryId: '' };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s:' + token);
  return true;
}

/* ================================================================== */
/* My Suki (customer)                                                  */
/* ================================================================== */

function getStores(token) {
  auth_(token, ['suki']);
  const products = readAll_('Products');
  return readAll_('Groceries').filter(g => g.Status === 'Active').map(g => ({
    id: g.GroceryID, name: g.Name, area: g.Area, address: g.Address, photoId: g.PhotoID,
    products: products.filter(p => p.GroceryID === g.GroceryID).length
  }));
}

function getProducts(token, groceryId) {
  auth_(token, ['suki']);
  return readAll_('Products').filter(p => p.GroceryID === groceryId)
    .map(p => ({ id: p.ProductID, name: p.Name, price: Number(p.Price), stock: Number(p.Stock) }));
}

function getCheckoutInfo(token, groceryId) {
  const s = auth_(token, ['suki']);
  const u = readAll_('Users').find(x => x.UserID === s.id) || {};
  const g = readAll_('Groceries').find(x => x.GroceryID === groceryId && x.Status === 'Active');
  if (!g) return null;
  return { area: u.Area, address: u.Address, store: g.Name, gcashNo: g.GcashNo, gcashName: g.GcashName };
}

// The server recalculates prices and checks stock, so the browser cannot change the total.
function placeOrder(token, o) {
  const s = auth_(token, ['suki']);
  const g = readAll_('Groceries').find(x => x.GroceryID === o.groceryId && x.Status === 'Active');
  if (!g) return bad_('This store is not taking orders right now.');
  const area = String(o.area || '');
  const address = String(o.address || '').trim().slice(0, 160);
  if (AREAS.indexOf(area) === -1 || !address) return bad_('Enter your delivery area and address.');
  const err = imgError_(o.receipt, 'your GCash receipt');
  if (err) return bad_(err);
  const lines = (o.cart || []).filter(l => Number(l.qty) > 0);
  if (!lines.length) return bad_('Your cart is empty.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const prods = readAll_('Products').filter(p => p.GroceryID === o.groceryId);
    let total = 0;
    const parts = [];
    for (const line of lines) {
      const p = prods.find(x => String(x.ProductID) === String(line.id));
      const qty = Math.floor(Number(line.qty));
      if (!p) return bad_('A product in your cart is no longer available.');
      if (qty > Number(p.Stock)) return bad_('Not enough stock for ' + p.Name + '. Only ' + p.Stock + ' left.');
      total += Number(p.Price) * qty;
      parts.push(p.Name + ' x' + qty);
    }
    const receiptId = saveImage_(o.receipt, 'receipt');
    lines.forEach(line => {
      const p = prods.find(x => String(x.ProductID) === String(line.id));
      updateRow_('Products', p.ProductID, { Stock: Number(p.Stock) - Math.floor(Number(line.qty)) });
    });
    const id = nextId_('Orders', '', 1001);
    append_('Orders', {
      OrderID: id, UserID: s.id, GroceryID: o.groceryId, Items: parts.join(', '), Total: total, Status: 'Placed',
      RiderID: '', Area: area, Address: address, PaymentStatus: 'To verify', ReceiptID: receiptId,
      PayRef: String(o.payRef || '').trim().slice(0, 30), ProofID: '', Timestamp: new Date(), DeliveredAt: ''
    });
    return { ok: true, id: id, total: total };
  } finally {
    lock.releaseLock();
  }
}

function getMyOrders(token) {
  const s = auth_(token, ['suki']);
  return ordersWhere_(o => o.UserID === s.id);
}

function cancelOrder(token, orderId) {
  const s = auth_(token, ['suki']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = readAll_('Orders').find(x => String(x.OrderID) === String(orderId) && x.UserID === s.id);
    if (!o) return bad_('Order not found.');
    if (o.Status !== 'Placed') return bad_('A rider already accepted this order, so it can no longer be cancelled here.');
    updateRow_('Orders', orderId, { Status: 'Cancelled' });
    restock_(o);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ================================================================== */
/* Grocery partner                                                     */
/* ================================================================== */

function getGroceryOrders(token) {
  const s = auth_(token, ['grocery']);
  return ordersWhere_(o => o.GroceryID === s.groceryId);
}

// Kapartner dashboard: counters, this week's sales (Mon to Sun, verified GCash payments), recent orders.
function getPartnerDashboard(token) {
  const s = auth_(token, ['grocery']);
  const g = readAll_('Groceries').find(x => x.GroceryID === s.groceryId) || {};
  const owner = readAll_('Users').find(x => x.UserID === s.id) || {};
  const mine = readAll_('Orders', true).filter(o => o.GroceryID === s.groceryId);
  const live = mine.filter(o => o.Status !== 'Cancelled');
  const TRANSIT = ['Accepted', 'Picked Up', 'On the Way'];

  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1 = Monday ... 7 = Sunday
  const week = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(now.getTime() + (i - (dow - 1)) * 86400000);
    week.push({ key: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), day: Utilities.formatDate(d, TZ, 'EEEE'),
                date: Utilities.formatDate(d, TZ, 'MMM d'), total: 0, orders: 0, future: i > dow - 1 });
  }
  live.forEach(o => {
    if (!(o.Timestamp instanceof Date) || o.PaymentStatus !== 'Verified') return;
    const w = week.find(x => x.key === Utilities.formatDate(o.Timestamp, TZ, 'yyyy-MM-dd'));
    if (w) { w.total += Number(o.Total) || 0; w.orders++; }
  });

  return {
    store: g.Name || s.name, owner: owner.Name || '', photoId: g.PhotoID || '',
    today: Utilities.formatDate(now, TZ, 'MMMM d, yyyy'),
    stats: {
      total: live.length,
      transit: live.filter(o => TRANSIT.indexOf(o.Status) > -1).length,
      delivered: live.filter(o => o.Status === 'Delivered').length,
      pending: live.filter(o => o.Status === 'Placed').length
    },
    toVerify: live.filter(o => o.PaymentStatus === 'To verify').length,
    week: week,
    recent: ordersWhere_(o => o.GroceryID === s.groceryId).slice(0, 5)
  };
}

function getMyProducts(token) {
  const s = auth_(token, ['grocery']);
  const g = readAll_('Groceries').find(x => x.GroceryID === s.groceryId) || {};
  return {
    store: { name: g.Name, gcashNo: g.GcashNo, gcashName: g.GcashName },
    products: readAll_('Products').filter(p => p.GroceryID === s.groceryId)
      .map(p => ({ id: p.ProductID, name: p.Name, price: Number(p.Price), stock: Number(p.Stock) }))
  };
}

function addProduct(token, p) {
  const s = auth_(token, ['grocery']);
  const name = String(p.name || '').trim().slice(0, 80);
  if (!name || !(Number(p.price) > 0) || !(Number(p.stock) >= 0)) return bad_('Enter a name, a price above 0, and a stock of 0 or more.');
  append_('Products', {
    ProductID: nextId_('Products', 'P', 1), GroceryID: s.groceryId,
    Name: name, Price: Number(p.price), Stock: Math.floor(Number(p.stock))
  });
  return { ok: true };
}

function updateStock(token, productId, stock) {
  const s = auth_(token, ['grocery']);
  if (!(Number(stock) >= 0)) return bad_('Stock must be 0 or more.');
  const p = readAll_('Products').find(x => x.ProductID === productId && x.GroceryID === s.groceryId);
  if (!p) return bad_('Product not found.');
  return { ok: updateRow_('Products', productId, { Stock: Math.floor(Number(stock)) }) };
}

// The store checks the GCash receipt. A rejected receipt cancels the order and returns the stock.
function verifyPayment(token, orderId, decision) {
  const s = auth_(token, ['grocery']);
  if (['Verified', 'Rejected'].indexOf(decision) === -1) return bad_('Invalid choice.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = readAll_('Orders').find(x => String(x.OrderID) === String(orderId) && x.GroceryID === s.groceryId);
    if (!o) return bad_('Order not found.');
    if (o.PaymentStatus !== 'To verify') return bad_('This payment was already checked.');
    if (decision === 'Verified') {
      updateRow_('Orders', orderId, { PaymentStatus: 'Verified' });
    } else {
      updateRow_('Orders', orderId, { PaymentStatus: 'Rejected', Status: 'Cancelled' });
      restock_(o);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ================================================================== */
/* TODA rider                                                          */
/* ================================================================== */

function getAvailable(token, showAll) {
  const s = auth_(token, ['toda']);
  const rider = readAll_('Riders').find(r => String(r.RiderID) === s.id) || {};
  const rows = ordersWhere_(o => o.Status === 'Placed' && !o.RiderID);
  rows.forEach(o => { o.onRoute = routeHas_(rider.Route, o.area); });
  return { route: rider.Route || '', rows: forRider_(showAll ? rows : rows.filter(o => o.onRoute)) };
}

function getMyDeliveries(token) {
  const s = auth_(token, ['toda']);
  return forRider_(ordersWhere_(o => String(o.RiderID) === s.id));
}

function acceptDelivery(token, orderId) {
  const s = auth_(token, ['toda']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = readAll_('Orders').find(x => String(x.OrderID) === String(orderId));
    if (!o || o.Status !== 'Placed' || o.RiderID) return bad_('Another rider already accepted this delivery, or the order was cancelled.');
    updateRow_('Orders', orderId, { Status: 'Accepted', RiderID: s.id });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function updateDelivery(token, orderId, newStatus) {
  const s = auth_(token, ['toda']);
  const flow = { 'Picked Up': 'Accepted', 'On the Way': 'Picked Up' };
  if (!flow[newStatus]) return bad_('Invalid step.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = readAll_('Orders').find(x => String(x.OrderID) === String(orderId) && String(x.RiderID) === s.id);
    if (!o) return bad_('Order not found.');
    if (o.Status !== flow[newStatus]) return bad_('This order is already ' + o.Status + '.');
    if (newStatus === 'Picked Up' && o.PaymentStatus !== 'Verified') return bad_('Wait until the store confirms the GCash payment.');
    updateRow_('Orders', orderId, { Status: newStatus });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function completeDelivery(token, orderId, photo) {
  const s = auth_(token, ['toda']);
  const err = imgError_(photo, 'a photo of the delivered order');
  if (err) return bad_(err);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = readAll_('Orders').find(x => String(x.OrderID) === String(orderId) && String(x.RiderID) === s.id);
    if (!o) return bad_('Order not found.');
    if (o.Status !== 'On the Way') return bad_('Start the delivery before marking it delivered.');
    updateRow_('Orders', orderId, { Status: 'Delivered', ProofID: saveImage_(photo, 'proof'), DeliveredAt: new Date() });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ================================================================== */
/* Reports (suki, grocery and TODA can send; developer reviews)        */
/* ================================================================== */

function submitReport(token, r) {
  const s = auth_(token, ['suki', 'grocery', 'toda']);
  const desc = String(r.description || '').trim().slice(0, 1000);
  if (!desc) return bad_('Describe what went wrong.');
  if (CATEGORIES.indexOf(r.category) === -1) return bad_('Choose a category.');
  append_('Reports', {
    ReportID: nextId_('Reports', 'R', 1), Category: r.category, ReportedByID: s.id, ReporterRole: s.role,
    Description: desc, OrderID: String(r.orderId || '').trim().slice(0, 12), Date: new Date(), Status: 'New', DevNotes: ''
  });
  return { ok: true };
}

/* ================================================================== */
/* Developer                                                           */
/* ================================================================== */

function getAdminStats(token) {
  auth_(token, ['admin']);
  const pending = readAll_('Applications').filter(a => a.Status === 'Pending');
  const count = t => pending.filter(a => a.Type === t).length;
  return {
    suki: readAll_('Users').filter(u => u.Role === 'suki').length,
    groceries: readAll_('Groceries').filter(g => g.Status === 'Active').length,
    toda: readAll_('Riders').filter(r => r.Status === 'Active').length,
    orders: readAll_('Orders').length,
    pending: { suki: count('suki'), grocery: count('grocery'), toda: count('toda') },
    openReports: readAll_('Reports').filter(r => r.Status !== 'Resolved').length,
    apps: pending.slice(0, 8).map(appRow_)
  };
}

// Approving creates the account, and for partners also the store or rider profile.
function reviewApplication(token, appId, decision) {
  auth_(token, ['admin']);
  if (['Approved', 'Declined'].indexOf(decision) === -1) return bad_('Invalid decision.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const a = readAll_('Applications').find(x => x.AppID === appId);
    if (!a || a.Status !== 'Pending') return bad_('This application was already reviewed.');
    if (decision === 'Approved') {
      if (readAll_('Users').some(u => lc_(u.Username) === lc_(a.Username))) return bad_('That username already has an account.');
      const userId = nextId_('Users', 'U', 1);
      append_('Users', {
        UserID: userId, Username: a.Username, PasswordHash: a.PasswordHash, Name: a.Name, Role: a.Type,
        Phone: a.Phone, Area: a.Area, Address: a.Address, Status: 'Active', DateRegistered: new Date()
      });
      if (a.Type === 'grocery') {
        append_('Groceries', {
          GroceryID: nextId_('Groceries', 'G', 1), Name: a.Business, OwnerUserID: userId, Area: a.Area, Address: a.Address,
          Phone: a.Phone, GcashNo: a.GcashNo, GcashName: a.GcashName, PhotoID: a.PhotoID, Status: 'Active'
        });
      }
      if (a.Type === 'toda') {
        append_('Riders', {
          RiderID: userId, Name: a.Name, TodaName: a.TodaName, PlateNo: a.PlateNo, Phone: a.Phone,
          Route: a.Route, PhotoID: a.PhotoID, Status: 'Active'
        });
      }
    }
    updateRow_('Applications', appId, { Status: decision });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// kind: 'suki' | 'grocery' | 'toda'
function getPartners(token, kind) {
  auth_(token, ['admin']);
  const apps = readAll_('Applications').filter(a => a.Type === kind).map(appRow_);
  const pending = apps.filter(a => a.status === 'Pending');
  const declined = apps.filter(a => a.status === 'Declined');
  const orders = readAll_('Orders');
  let active, stats;

  if (kind === 'suki') {
    const counts = {};
    orders.forEach(o => { counts[o.UserID] = (counts[o.UserID] || 0) + 1; });
    const weekAgo = Date.now() - 7 * 86400000;
    active = readAll_('Users', true).filter(u => u.Role === 'suki').map(u => ({
      id: u.UserID, name: u.Name, phone: u.Phone, area: u.Area, status: u.Status,
      registered: fmtDay_(u.DateRegistered), orders: counts[u.UserID] || 0,
      isNew: u.DateRegistered instanceof Date && u.DateRegistered.getTime() > weekAgo
    }));
    stats = [
      ['Total users', active.length],
      ['Active users', active.filter(r => r.status === 'Active').length],
      ['New this week', active.filter(r => r.isNew).length],
      ['Pending applications', pending.length]
    ];
  } else if (kind === 'grocery') {
    const products = readAll_('Products');
    const users = index_(readAll_('Users'), 'UserID');
    active = readAll_('Groceries').map(g => ({
      id: g.GroceryID, name: g.Name, owner: (users[g.OwnerUserID] || {}).Name || '', area: g.Area, phone: g.Phone, status: g.Status,
      products: products.filter(p => p.GroceryID === g.GroceryID).length,
      orders: orders.filter(o => o.GroceryID === g.GroceryID).length
    }));
    stats = [
      ['Total groceries', active.length],
      ['Active', active.filter(r => r.status === 'Active').length],
      ['Pending applications', pending.length],
      ['Total products', products.length]
    ];
  } else {
    active = readAll_('Riders').map(r => ({
      id: r.RiderID, name: r.Name, toda: r.TodaName, plate: r.PlateNo, route: r.Route, phone: r.Phone, status: r.Status,
      deliveries: orders.filter(o => String(o.RiderID) === String(r.RiderID) && o.Status === 'Delivered').length
    }));
    stats = [
      ['Total riders', active.length],
      ['Active', active.filter(r => r.status === 'Active').length],
      ['Pending applications', pending.length],
      ['Total deliveries', active.reduce((n, r) => n + r.deliveries, 0)]
    ];
  }
  return { stats: stats, active: active, pending: pending, declined: declined };
}

// One shape for all three kinds so the drawer can render them the same way.
function getDetail(token, kind, id) {
  auth_(token, ['admin']);
  if (kind === 'suki') {
    const u = readAll_('Users', true).find(x => x.UserID === id && x.Role === 'suki');
    if (!u) return null;
    return {
      title: u.Name, status: u.Status, photoId: '', historyTitle: 'Order history',
      profile: [['User ID', u.UserID], ['Phone', u.Phone], ['Area', u.Area], ['Address', u.Address], ['Registered', fmtDay_(u.DateRegistered)]],
      orders: ordersWhere_(o => o.UserID === id)
    };
  }
  if (kind === 'grocery') {
    const g = readAll_('Groceries').find(x => x.GroceryID === id);
    if (!g) return null;
    const owner = readAll_('Users').find(x => x.UserID === g.OwnerUserID) || {};
    return {
      title: g.Name, status: g.Status, photoId: g.PhotoID, historyTitle: 'Order history',
      profile: [['Store ID', g.GroceryID], ['Owner', owner.Name], ['Phone', g.Phone], ['Area', g.Area], ['Address', g.Address],
                ['GCash', g.GcashNo + (g.GcashName ? ' (' + g.GcashName + ')' : '')]],
      products: readAll_('Products').filter(p => p.GroceryID === id)
        .map(p => ({ id: p.ProductID, name: p.Name, price: Number(p.Price), stock: Number(p.Stock) })),
      orders: ordersWhere_(o => o.GroceryID === id)
    };
  }
  const r = readAll_('Riders').find(x => String(x.RiderID) === String(id));
  if (!r) return null;
  return {
    title: r.Name, status: r.Status, photoId: r.PhotoID, historyTitle: 'Delivery history',
    profile: [['Rider ID', r.RiderID], ['TODA', r.TodaName], ['Tricycle no.', r.PlateNo], ['Route', r.Route], ['Phone', r.Phone]],
    orders: ordersWhere_(o => String(o.RiderID) === String(id))
  };
}

// status: 'Active' | 'Suspended'. Suspending a partner also blocks their login.
function setPartnerStatus(token, kind, id, status) {
  auth_(token, ['admin']);
  if (['Active', 'Suspended'].indexOf(status) === -1) return bad_('Invalid status.');
  if (kind === 'suki') return { ok: updateRow_('Users', id, { Status: status }) };
  if (kind === 'grocery') {
    const g = readAll_('Groceries').find(x => x.GroceryID === id);
    if (!g) return bad_('Store not found.');
    updateRow_('Groceries', id, { Status: status });
    updateRow_('Users', g.OwnerUserID, { Status: status });
    return { ok: true };
  }
  updateRow_('Riders', id, { Status: status });
  updateRow_('Users', id, { Status: status });
  return { ok: true };
}

function getReports(token) {
  auth_(token, ['admin']);
  const users = index_(readAll_('Users'), 'UserID');
  return readAll_('Reports').map(r => ({
    id: r.ReportID, category: r.Category, by: (users[r.ReportedByID] || {}).Name || r.ReportedByID,
    role: r.ReporterRole, description: r.Description, orderId: r.OrderID, date: r.Date, status: r.Status, notes: r.DevNotes
  })).reverse();
}

function updateReport(token, reportId, status, notes) {
  auth_(token, ['admin']);
  if (['New', 'In Review', 'Resolved'].indexOf(status) === -1) return bad_('Invalid status.');
  return { ok: updateRow_('Reports', reportId, { Status: status, DevNotes: String(notes || '').slice(0, 1000) }) };
}


/* ================================================================== */
/* Fix-it tool (run from the editor, never from the web app)           */
/* ================================================================== */

/**
 * Run this if developer sign-in says "incorrect".
 * Step 1 prints which part does not match. Step 2 resets the developer account to known values:
 *   username: developer   password: sari2026   access key: SARI-DEV-2026
 * Open View > Logs (or Execution log) after running it.
 */
function resetDeveloper() {
  const props = PropertiesService.getScriptProperties();
  const log = [];

  // Step 1: diagnose
  const hadSalt = !!props.getProperty('SALT');
  log.push('SALT saved in Script properties: ' + (hadSalt ? 'yes' : 'NO'));
  log.push('DEV_KEY saved in Script properties: ' + (props.getProperty('DEV_KEY') ? 'yes (' + props.getProperty('DEV_KEY') + ')' : 'NO'));
  let dev = null;
  try {
    dev = readAll_('Users').find(r => r.Role === 'admin' && lc_(r.Username) === 'developer');
  } catch (e) {
    log.push('Users tab problem: ' + e.message);
  }
  log.push('Developer row in Users tab: ' + (dev ? 'yes (' + dev.UserID + ')' : 'NO'));
  if (dev) log.push('Stored password matches "sari2026": ' + (dev.PasswordHash === hash_('sari2026') ? 'yes' : 'NO'));

  // Step 2: reset
  if (!hadSalt) {
    props.setProperty('SALT', Utilities.getUuid());
    log.push('Created a new SALT. Demo accounts (suki1, grocery1, toda1) will not accept "1234" anymore.');
  }
  props.setProperty('DEV_KEY', 'SARI-DEV-2026');
  if (dev) {
    updateRow_('Users', dev.UserID, { PasswordHash: hash_('sari2026'), Status: 'Active' });
  } else {
    append_('Users', {
      UserID: nextId_('Users', 'U', 1), Username: 'developer', PasswordHash: hash_('sari2026'), Name: 'Developer',
      Role: 'admin', Status: 'Active', DateRegistered: new Date()
    });
  }
  CacheService.getScriptCache().remove('fail:developer');
  log.push('Reset done. Sign in with username "developer", password "sari2026", access key "SARI-DEV-2026".');
  Logger.log(log.join('\n'));
}
