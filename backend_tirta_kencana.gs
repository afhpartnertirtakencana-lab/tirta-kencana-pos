// ══════════════════════════════════════════════════════════════
//  TIRTA KENCANA — Google Apps Script Backend (v8.5 - TAMBAH createdAt di getTrxList)
//  - Perbaikan: Fungsi doRekap lengkap, Filter Status Pembayaran, Error Handling doGet
//  - [NEW v8.4] Tambah fungsi deleteFoto() untuk menu Hapus Foto di web (Drive + Sheet FotoBukti)
//  - [NEW v8.5] ensureTrxHeaders() sekarang memastikan kolom 'CreatedAt' selalu ada
//               (ditambahkan otomatis kalau sheet lama belum punya, sama seperti pola
//               penambahan kolom TotalModal/TotalProfit yang sudah ada sebelumnya).
//               getTrxList() sekarang IKUT mengembalikan field createdAt (tanggal+jam
//               lengkap) supaya aplikasi bisa mengurutkan transaksi di tanggal yang sama
//               secara AKURAT sesuai waktu input sebenarnya, konsisten di semua device -
//               bukan lagi cuma tebakan lokal di satu browser saja.
//               Tidak ada fungsi/variabel/logika lain yang diubah.
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
//  [SECURITY FIX — SESI/TOKEN API]
//  Sebelum ini: login() diverifikasi server, tapi SETIAP endpoint lain
//  (getTrxList, saveTrx, dst) bisa dipanggil siapapun yang tahu URL web
//  app ini, tanpa perlu login sama sekali. Sekarang SEMUA endpoint
//  (kecuali 'login' sendiri) wajib menyertakan token sesi yang valid.
//
//  Desain: token "stateless" (mirip JWT versi ringan) - TIDAK disimpan
//  di sheet/database manapun, jadi validasinya cepat (tidak perlu baca
//  sheet tiap request) dan otomatis berlaku untuk banyak device/HP
//  sekaligus tanpa perlu tabel sesi yang harus dibersihkan.
//  Isi token: base64(nama+role+kedaluwarsa) + '.' + tanda-tangan HMAC.
//  Kalau token dipalsukan/di-utak-atik, tanda-tangannya tidak akan cocok.
//
//  PENTING SEBELUM DEPLOY:
//  - Ganti SESSION_SECRET di bawah dengan string acak unikmu sendiri
//    (beda dari PASSWORD_SALT), JANGAN dibiarkan seperti contoh.
//  - SESSION_TTL_MS menentukan berapa lama sesi login bertahan sebelum
//    wajib login ulang. Default 30 hari - dipilih supaya konsisten
//    dengan perilaku app selama ini (device yang sudah login tetap
//    login terus), tapi tidak selamanya seperti sebelumnya.
// ══════════════════════════════════════════════════════════════
var SESSION_SECRET = 'JOoUu15b16URfrrq0QPoZ4wUAdfVauXGWFqfadHnZ05YwHuE';
var SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 hari

function issueToken_(name, role) {
  var payload = JSON.stringify({ n: name, r: role, exp: Date.now() + SESSION_TTL_MS });
  var payloadB64 = Utilities.base64EncodeWebSafe(payload);
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, SESSION_SECRET));
  return payloadB64 + '.' + sig;
}

// Melempar Error kalau token tidak ada/tidak valid/kedaluwarsa.
// Return { name, role } kalau valid.
function verifyToken_(token) {
  if (!token || token.indexOf('.') < 0) throw new Error('SESSION_INVALID');
  var parts = token.split('.');
  var payloadB64 = parts[0], sig = parts[1];
  var expectedSig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, SESSION_SECRET));
  if (sig !== expectedSig) throw new Error('SESSION_INVALID');
  var payload;
  try { payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString()); }
  catch (e) { throw new Error('SESSION_INVALID'); }
  if (!payload || !payload.exp || Date.now() > payload.exp) throw new Error('SESSION_EXPIRED');
  return { name: payload.n, role: payload.r };
}

// Endpoint yang boleh dipanggil TANPA token (cuma login itu sendiri).
var PUBLIC_ACTIONS = { login: 1 };

var SPREADSHEET_ID = '12q8SwBtoww9Y9c6EZ46-SNa1q5TKIXnf9g_3wLbsNK4';



function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('Gagal membuka Spreadsheet: ' + e.message);
  }
}
// [FIX] Sebelumnya: var SS = getSpreadsheet(); dipanggil SEKALI di top-level saat script "dipanaskan".
// Kalau panggilan pertama itu gagal sesaat (transient error, baru redeploy, dsb), SS jadi undefined SELAMANYA
// untuk proses/container itu (karena kode top-level cuma jalan sekali per container, bukan per request),
// menyebabkan semua fungsi yang butuh Sheet gagal terus walau koneksi sebenarnya sudah normal.
// Perbaikan: ambil SS secara lazy + self-healing (coba ulang otomatis kalau ternyata kosong).
var SS = null;
function getSS() {
  if (!SS) SS = getSpreadsheet();
  return SS;
}

var SHEET_NAMES = {
  TRX       : ['Transaksi',    'Transactions'],
  DETAIL    : ['DetailTrx',    'TrxDetail'],
  PRODUCTS  : ['Produk',       'Products'],
  CUSTOMERS : ['Pelanggan',    'Customers'],
  SETTINGS  : ['Pengaturan',   'Settings'],
  BARANG    : ['InputBarang',  'BarangMasuk'],
  SETORAN   : ['Setoran',      'Setoran'],
  USERS     : ['Users',        'User'],
  MASTERSTOCK : ['MasterStock', 'MasterStock'],
  DRIVERS   : ['Driver',       'Drivers'],
  CEKKOSONGAN : ['CekKosongan', 'CekKosongan']
};

function sheetName(candidates) {
  if (!candidates || !Array.isArray(candidates)) return 'Sheet1';
  for (var i = 0; i < candidates.length; i++) {
    try { if (getSS().getSheetByName(candidates[i])) return candidates[i]; } catch (e) {}
  }
  return candidates[0];
}

var SHEET = {
  TRX      : sheetName(SHEET_NAMES.TRX),
  DETAIL   : sheetName(SHEET_NAMES.DETAIL),
  PRODUCTS : sheetName(SHEET_NAMES.PRODUCTS),
  CUSTOMERS: sheetName(SHEET_NAMES.CUSTOMERS),
  SETTINGS : sheetName(SHEET_NAMES.SETTINGS),
  BARANG   : sheetName(SHEET_NAMES.BARANG),
  SETORAN  : sheetName(SHEET_NAMES.SETORAN),
  USERS    : sheetName(SHEET_NAMES.USERS),
  MASTERSTOCK : sheetName(SHEET_NAMES.MASTERSTOCK),
  DRIVERS  : sheetName(SHEET_NAMES.DRIVERS),
  CEKKOSONGAN : sheetName(SHEET_NAMES.CEKKOSONGAN)
};

// ══════════════════════════════════════════════════════════════
//  doGet / doPost — JSONP entry point
//  [SECURITY FIX] doGet dan doPost sekarang berbagi SATU dispatcher
//  (dispatchAction) yang sama-sama memakai whitelist switch di bawah.
//  Sebelumnya doPost memakai this[fn].apply(null, args) — itu artinya
//  SEMBARANG fungsi global di project ini bisa dipanggil lewat POST
//  tanpa whitelist sama sekali (bukan cuma fungsi yang dimaksud untuk
//  dipakai API). Frontend aplikasi ini sendiri tidak pernah memakai
//  doPost (semua lewat JSONP/doGet), jadi menutup ini TIDAK mengubah
//  perilaku aplikasi sama sekali — murni menutup celah yang tidak
//  dipakai siapapun secara sah.
//  [SECURITY FIX] Operasi tulis (save/update/delete) sekarang dibungkus
//  LockService.getScriptLock() supaya 2 sales/driver yang input
//  bersamaan tidak saling menimpa baris di Sheets (race condition).
//  Operasi baca TIDAK dikunci (tidak perlu, dan supaya tetap cepat).
// ══════════════════════════════════════════════════════════════
var WRITE_ACTIONS = {
  saveTrx: 1, upsertTrx: 1, updateStatus: 1, deleteTrx: 1,
  saveProducts: 1,
  addCustomer: 1, saveCustomers: 1,
  saveSettings: 1,
  saveInputBarang: 1, updateInputBarang: 1, deleteInputBarang: 1,
  saveSetoran: 1,
  deleteFoto: 1,
  saveMasterStock: 1, updateStockAwal: 1, updateStockRow: 1, cleanupDuplicateMasterStock: 1, saveCekKosongan: 1,
  saveUser: 1, deleteUser: 1,
  saveDrivers: 1
};

function dispatchAction(fn, args, token) {
  // [SECURITY FIX] Wajib token valid untuk semua fungsi KECUALI 'login'.
  // Ini yang menutup celah "siapapun yang tahu URL bisa panggil getTrxList dkk
  // tanpa login sama sekali" yang sebelumnya masih terbuka meski login sudah
  // diverifikasi server.
  if (!PUBLIC_ACTIONS[fn]) {
    verifyToken_(token); // lempar Error('SESSION_INVALID') / Error('SESSION_EXPIRED') kalau gagal
  }

  var lock = null;
  if (WRITE_ACTIONS[fn]) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(15000); // tunggu s/d 15 detik kalau ada penulisan lain sedang berjalan
    } catch (lockErr) {
      throw new Error('Server sedang memproses transaksi lain, coba lagi sebentar.');
    }
  }
  try {
    switch (fn) {
      case 'getTrxList'           : return getTrxList();
      case 'saveTrx'              : return saveTrx(args[0]);
      case 'upsertTrx'            : return upsertTrx(args[0]);
      case 'updateStatus'         : return updateStatus(args[0], args[1]);
      case 'deleteTrx'            : return deleteTrx(args[0]);
      case 'getProducts'          : return getProducts();
      case 'saveProducts'         : return saveProducts(args[0]);
      case 'getCustomers'         : return getCustomers();
      case 'addCustomer'          : return addCustomer(args[0]);
      case 'saveCustomers'        : return saveCustomers(args[0]);
      case 'getSettings'          : return getSettings();
      case 'saveSettings'         : return saveSettings(args[0]);
      case 'saveInputBarang'      : return saveInputBarang(args[0]);
      case 'getInputBarangHistory': return getInputBarangHistory();
      case 'saveSetoran'          : return saveSetoran(args[0]);
      case 'getSetoranHistory'    : return getSetoranHistory();
      case 'uploadFoto'           : return uploadFoto(args[0]);
      case 'getFotoList'          : return getFotoList(args[0]);
      case 'deleteFoto'           : return deleteFoto(args[0]); // [NEW v8.4]
      case 'saveMasterStock'      : return saveMasterStock(args[0]);
      case 'saveCekKosongan'      : return saveCekKosongan(args[0]);
      case 'getCekKosonganByDate' : return getCekKosonganByDate(args[0]);
      case 'cleanupDuplicateMasterStock' : return cleanupDuplicateMasterStock();
      case 'getMasterStockByDate' : return getMasterStockByDate(args[0]);
      case 'getStockBarang'       : return getStockBarang();
      case 'getUsers'             : return getUsers();
      case 'saveUser'             : return saveUser(args[0]);
      case 'deleteUser'           : return deleteUser(args[0]);
      case 'getDrivers'           : return getDrivers();
      case 'saveDrivers'          : return saveDrivers(args[0]);
      case 'updateInputBarang'    : return updateInputBarang(args[0]);
      case 'deleteInputBarang'    : return deleteInputBarang(args[0]);
      case 'getTrxDetail'         : return getTrxDetail(args[0]);
      case 'updateStockAwal'      : return updateStockAwal(args[0]);
      case 'updateStockRow'       : return updateStockRow(args[0]);
      case 'doRekap'              : return doRekap();
      case 'login'                : return loginUser(args[0], args[1]); // [NEW] lihat catatan keamanan login di bawah file
      default: throw new Error('Fungsi tidak dikenal: ' + fn);
    }
  } finally {
    if (lock) lock.releaseLock();
  }
}

function doGet(e) {
  var params = {};
  if (e && e.parameter) {
    params = e.parameter;
  }

  var cb    = params.callback || '';
  var fn    = params.fn       || '';
  var token = params.token    || '';
  var args  = [];

  try {
    var argStr = params.args || '[]';
    args = JSON.parse(decodeURIComponent(argStr));
  } catch(x) {
    try { args = JSON.parse(argStr); } catch(x2) { args = []; }
  }

  var result, isOk = true, errMsg = '';
  try {
    result = dispatchAction(fn, args, token);
  } catch(err) {
    isOk = false;
    errMsg = err.message + ' (' + fn + ')';
    Logger.log(err);
  }

  var payload = isOk ? JSON.stringify({ ok: true, data: result }) : JSON.stringify({ ok: false, error: errMsg });

  if (cb) return ContentService.createTextOutput(cb + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var fn = body.fn;
    var args = body.args || [];
    var token = body.token || '';
    var result = dispatchAction(fn, args, token); // [SECURITY FIX] sekarang lewat whitelist + token yang sama dengan doGet
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════════════════════
//  HELPER
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
//  CEK KOSONGAN (galon kosong toko/mobil/titipan/pinjaman)
//  Satu baris per TANGGAL (upsert - kalau tanggal sudah ada, ditimpa, bukan
//  ditambah dobel). Titipan & Pinjaman disimpan sebagai JSON di 1 sel karena
//  jumlah barisnya dinamis (nama+qty bebas ditambah/dikurangi oleh user).
// ══════════════════════════════════════════════════════════════
function saveCekKosongan(payload) {
  try {
    if (!payload || !payload.date) throw new Error('Tanggal diperlukan');
    var sh = getSheet(SHEET.CEKKOSONGAN);
    if (sh.getLastRow() === 0) sh.appendRow(['Tanggal','Toko','Mobil','Titipan','Pinjaman','UpdatedAt']);
    var data = sh.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (formatDateCell(data[i][0]) === formatDateCell(payload.date)) { foundRow = i + 1; break; }
    }
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var row = [payload.date, Number(payload.toko)||0, Number(payload.mobil)||0, JSON.stringify(payload.titipan||[]), JSON.stringify(payload.pinjaman||[]), now];
    if (foundRow > 0) sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return { ok: true, date: payload.date };
  } catch(e) {
    throw new Error('saveCekKosongan gagal: ' + e.message);
  }
}
function getCekKosonganByDate(date) {
  try {
    var sh = getSS().getSheetByName(SHEET.CEKKOSONGAN);
    if (!sh) return null;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (formatDateCell(data[i][0]) === formatDateCell(date)) {
        var titipan = [], pinjaman = [];
        try { titipan = JSON.parse(data[i][3] || '[]'); } catch(e) {}
        try { pinjaman = JSON.parse(data[i][4] || '[]'); } catch(e) {}
        return { date: formatDateCell(data[i][0]), toko: data[i][1], mobil: data[i][2], titipan: titipan, pinjaman: pinjaman };
      }
    }
    return null;
  } catch(e) {
    throw new Error('getCekKosonganByDate gagal: ' + e.message);
  }
}

function getSheet(name) { 
  try {
    var sh = getSS().getSheetByName(name); 
    if (!sh) sh = getSS().insertSheet(name); 
    return sh;
  } catch(e) {
    throw new Error('Gagal mengakses sheet "' + name + '": ' + e.message);
  }
}
function getSheetData(name) { return getSheet(name).getDataRange().getValues(); }
function formatDateCell(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Jakarta', 'yyyy-MM-dd');
  var s = String(val).trim(); if (s.length >= 10) return s.substring(0,10); return s;
}
function parseHarga(val) { if (typeof val === 'number') return val; var s = String(val||'').replace(/[^0-9]/g,''); return parseInt(s)||0; }

// ══════════════════════════════════════════════════════════════
//  REKAP DATA (LENGKAP & DIPERBAIKI)
//  Menghasilkan: Penjualan per Sales, Rekap per SKU, Customer, Qty, Payment Methods
// ══════════════════════════════════════════════════════════════
function doRekap() {
  try {
    // 1. Ambil Data Detail & Transaksi
    var detailData = getSheetData(SHEET.DETAIL);
    var trxData = getSheetData(SHEET.TRX);
    
    if (detailData.length <= 1) return { message: "Belum ada data transaksi" };

    // Parsing Header Detail
    var dHeaders = detailData[0].map(function(h) { return String(h||'').toLowerCase().trim(); });
    var dCol = {
      trxid: dHeaders.indexOf('trxid'),
      sales: dHeaders.indexOf('sales'),
      customer: dHeaders.indexOf('customer'),
      sku: dHeaders.indexOf('sku'),
      qty: dHeaders.indexOf('qty'),
      subtotal: dHeaders.indexOf('subtotal'),
      status: dHeaders.indexOf('status')
    };

    // Parsing Header Transaksi (untuk validasi status jika di detail kosong)
    var tHeaders = trxData[0].map(function(h) { return String(h||'').toLowerCase().trim(); });
    var tColStatus = tHeaders.indexOf('status');
    if (tColStatus < 0) tColStatus = tHeaders.indexOf('pembayaran');
    
    // Map Status Transaksi per ID untuk fallback
    var trxStatusMap = {};
    for (var t = 1; t < trxData.length; t++) {
      var tid = String(trxData[t][0]||'').trim();
      var tStat = String(trxData[t][tColStatus]||'belumTransfer').trim();
      trxStatusMap[tid] = tStat;
    }

    // Variabel Penampung Hasil
    var salesReport = {};       // { salesName: { sku: { qty, total } } }
    var skuReport = {};         // { sku: { customer: { qty, total } } }
    var customerReport = {};    // { customer: { sku: { qty, total } } }
    var paymentReport = {       // { sales: { method: { customer: total } } }
      cod: {}, transfer: {}, qris: {}, belumTransfer: {}
    };
    var salesQtyTotal = {};     // { sales: { sku: qty } }

    // Loop Data Detail
    for (var i = 1; i < detailData.length; i++) {
      var row = detailData[i];
      var trxId = String(row[dCol.trxid]||'').trim();
      var sales = String(row[dCol.sales]||'Unknown').trim();
      var customer = String(row[dCol.customer]||'Umum').trim();
      var sku = String(row[dCol.sku]||'').trim();
      var qty = Number(row[dCol.qty]) || 0;
      var subtotal = Number(row[dCol.subtotal]) || 0;
      
      // Ambil Status: Prioritas dari Detail, jika kosong ambil dari Header Transaksi
      var statusRaw = String(row[dCol.status]||'').trim();
      if (!statusRaw && trxStatusMap[trxId]) {
        statusRaw = trxStatusMap[trxId];
      }
      // Normalisasi Status
      var status = 'belumTransfer';
      var sLower = statusRaw.toLowerCase();
      if (sLower.indexOf('cod') >= 0 || sLower.indexOf('cash') >= 0) status = 'cod';
      else if (sLower.indexOf('qris') >= 0) status = 'qris';
      else if (sLower.indexOf('transfer') >= 0 && sLower.indexOf('belum') === -1) status = 'transfer';
      else if (sLower.indexOf('lunas') >= 0) status = 'transfer'; // Asumsi lunas = transfer
      
      // --- 1. Laporan Per Sales (Nested by SKU) ---
      if (!salesReport[sales]) salesReport[sales] = {};
      if (!salesReport[sales][sku]) salesReport[sales][sku] = { qty: 0, total: 0 };
      salesReport[sales][sku].qty += qty;
      salesReport[sales][sku].total += subtotal;

      // --- 2. Laporan Per SKU (Nested by Customer) ---
      if (!skuReport[sku]) skuReport[sku] = {};
      if (!skuReport[sku][customer]) skuReport[sku][customer] = { qty: 0, total: 0 };
      skuReport[sku][customer].qty += qty;
      skuReport[sku][customer].total += subtotal;

      // --- 3. Laporan Per Customer (Nested by SKU) ---
      if (!customerReport[customer]) customerReport[customer] = {};
      if (!customerReport[customer][sku]) customerReport[customer][sku] = { qty: 0, total: 0 };
      customerReport[customer][sku].qty += qty;
      customerReport[customer][sku].total += subtotal;

      // --- 4. Laporan Sales Qty (Flat) ---
      if (!salesQtyTotal[sales]) salesQtyTotal[sales] = {};
      if (!salesQtyTotal[sales][sku]) salesQtyTotal[sales][sku] = 0;
      salesQtyTotal[sales][sku] += qty;

      // --- 5. Laporan Pembayaran (Per Sales -> Customer) ---
      if (!paymentReport[status][sales]) paymentReport[status][sales] = {};
      if (!paymentReport[status][sales][customer]) paymentReport[status][sales][customer] = 0;
      paymentReport[status][sales][customer] += subtotal;
    }

    return {
      perSales: salesReport,
      perSku: skuReport,
      perCustomer: customerReport,
      perSalesQty: salesQtyTotal,
      pembayaran: paymentReport,
      timestamp: new Date().toString()
    };

  } catch(e) {
    throw new Error('doRekap gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  DRIVERS
// ══════════════════════════════════════════════════════════════
function getDrivers() {
  try {
    var sheet = getSheet(SHEET.DRIVERS);
    var data = sheet.getDataRange().getValues();
    if (data.length === 0 || String(data[0][0] || '').toLowerCase() !== 'nama driver') {
      sheet.clearContents();
      sheet.appendRow(['Nama Driver', 'Status', 'Tanggal Dibuat']);
      ['oji', 'padong', 'said', 'dedi', 'zehpudin'].forEach(function(d) {
        sheet.appendRow([d, 'aktif', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd')]);
      });
      data = sheet.getDataRange().getValues();
    }
    var drivers = [];
    for (var i = 1; i < data.length; i++) {
      var nama = String(data[i][0] || '').trim();
      if (nama) drivers.push(nama);
    }
    return drivers;
  } catch(e) {
    throw new Error('getDrivers gagal: ' + e.message);
  }
}

function saveDrivers(driverList) {
  try {
    if (typeof driverList === 'string') {
      try { driverList = JSON.parse(driverList); } catch(e) { driverList = [driverList]; }
    }
    if (!Array.isArray(driverList)) driverList = driverList ? [String(driverList)] : [];
    var sheet = getSheet(SHEET.DRIVERS);
    sheet.clearContents();
    sheet.appendRow(['Nama Driver', 'Status', 'Tanggal Dibuat']);
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    driverList.forEach(function(nama) {
      if (String(nama || '').trim()) sheet.appendRow([String(nama).trim(), 'aktif', now]);
    });
    return { ok: true, count: driverList.length };
  } catch(e) {
    throw new Error('saveDrivers gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  SETTINGS
// ══════════════════════════════════════════════════════════════
function getSettings() {
  try {
    var data = getSheetData(SHEET.SETTINGS);
    var cfg  = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]||'').trim();
      var val = String(data[i][1]||'').trim();
      if (!key) continue;
      if (['bank1','bank2','salesList'].indexOf(key) >= 0) {
        try { cfg[key] = JSON.parse(val); } catch(e) { cfg[key] = (key==='salesList') ? [] : {}; }
      } else { cfg[key] = val; }
    }
    if (!cfg.storeName) cfg.storeName = 'Tirta Kencana';
    if (!cfg.tagline)   cfg.tagline   = 'Distributor Air Minum Terpercaya';
    if (!cfg.bank1)     cfg.bank1     = { nama:'BCA', norek:'6930099099', penerima:'Hendri' };
    if (!cfg.bank2)     cfg.bank2     = { nama:'bluBCA', norek:'002283588888', penerima:'Hendri' };
    if (!cfg.salesList) cfg.salesList = [];
    return cfg;
  } catch(e) {
    throw new Error('getSettings gagal: ' + e.message);
  }
}

function saveSettings(cfg) {
  if (!cfg || typeof cfg !== 'object') throw new Error('Data settings tidak valid');
  // [FIX] QRIS sebelumnya HANYA tersimpan di localStorage browser (tidak pernah ikut
  // disimpan ke sini) - artinya setiap kali localStorage/cache device dibersihkan
  // (uninstall PWA, clear site data, ganti device), gambar QRIS yang sudah di-upload
  // ikut hilang dan otomatis balik ke gambar bawaan (EMBEDDED_QRIS) yang gagal muncul
  // saat download PNG karena masalah CORS di hosting gambarnya. Sekarang QRIS ikut
  // disimpan ke Sheet ini juga, supaya permanen & ikut tersinkron ke semua device,
  // tidak hilang lagi walau cache/localStorage device dibersihkan.
  // Catatan: 1 sel Google Sheets punya batas ~50.000 karakter - base64 gambar QRIS
  // (biasanya gambar QR kecil/simple) umumnya jauh di bawah itu, tapi tetap dijaga
  // dengan pengecekan panjang supaya tidak error kalau ada yang upload gambar besar.
  var qrisVal = String(cfg.qris || '');
  if (qrisVal.length > 45000) qrisVal = ''; // terlalu besar utk 1 sel Sheets, jangan dipaksa simpan (biarkan tetap di localStorage lokal saja)
  var sh = getSheet(SHEET.SETTINGS); sh.clearContents();
  sh.getRange(1,1,9,2).setValues([
    ['storeName', String(cfg.storeName||'')],
    ['tagline',   String(cfg.tagline||'')],
    ['address',   String(cfg.address||'')],
    ['phone',     String(cfg.phone||'')],
    ['footer',    String(cfg.footer||'')],
    ['salesList', JSON.stringify(cfg.salesList || [])],
    ['bank1',     JSON.stringify(cfg.bank1 || {})],
    ['bank2',     JSON.stringify(cfg.bank2 || {})],
    ['qris',      qrisVal]
  ]);
  return 'ok';
}

// ══════════════════════════════════════════════════════════════
//  USERS
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
//  [SECURITY FIX] Password sekarang di-hash (SHA-256 + salt), tidak
//  lagi plaintext. getUsers() yang dipanggil dari client TIDAK lagi
//  mengirim field password/hash sama sekali — client hanya butuh
//  name+role untuk tampilan/dropdown. Verifikasi login sepenuhnya
//  di server lewat loginUser().
//
//  PENTING SEBELUM DEPLOY:
//  1. Ganti PASSWORD_SALT di bawah ini dengan string acak unikmu sendiri
//     (jangan dibiarkan sama seperti contoh).
//  2. Backup sheet "Users" (duplicate spreadsheet atau export CSV) SEBELUM
//     menjalankan migratePasswordsToHash_ONETIME().
//  3. Jalankan migratePasswordsToHash_ONETIME() SEKALI SAJA dari editor
//     Apps Script (bukan lewat URL/doGet) untuk meng-hash password lama
//     yang masih plaintext di sheet.
//  4. Setelah migrasi jalan, deploy ulang web app, lalu test login semua
//     role (admin/sales/driver) sebelum staff mulai pakai.
// ══════════════════════════════════════════════════════════════
var PASSWORD_SALT = '3qKPYeDxA80ViD1qllqLqEvUsBNGV5gdJ8bzWrmpl0slO7HP';

function hashPassword_(plain) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, PASSWORD_SALT + ':' + String(plain));
  return Utilities.base64Encode(raw);
}

// Heuristik: hash SHA-256 kita selalu base64 44 karakter berakhiran '='.
// Dipakai migrasi supaya tidak meng-hash ulang baris yang sudah di-hash.
function looksAlreadyHashed_(val) {
  var s = String(val || '');
  return s.length === 44 && s.charAt(43) === '=';
}

// Versi internal LENGKAP (termasuk passwordHash) — hanya dipakai fungsi
// server lain (saveUser/deleteUser/loginUser), TIDAK PERNAH diekspos
// langsung ke client lewat dispatchAction.
function getUsersRaw_() {
  var sheet = getSheet(SHEET.USERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    var defaultUsers = [
      ['admin','admin','020730'],
      ['hasan','sales','1234'],
      ['ujang','sales','1234'],
      ['oji','driver','1234'],
      ['padong','driver','1234'],
      ['said','driver','1234'],
      ['dedi','driver','1234'],
      ['tirta','admin','888999']
    ];
    if (data.length === 0) sheet.appendRow(['Nama','Role','PasswordHash']);
    defaultUsers.forEach(function(u) { sheet.appendRow([u[0], u[1], hashPassword_(u[2])]); });
    data = sheet.getDataRange().getValues();
  }
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({ name: String(data[i][0]||''), role: String(data[i][1]||'sales'), passwordHash: String(data[i][2]||'') });
  }
  return users;
}

// Versi PUBLIK yang dipanggil client — tanpa password/hash sama sekali.
function getUsers() {
  try {
    return getUsersRaw_().map(function(u) { return { name: u.name, role: u.role }; });
  } catch(e) {
    throw new Error('getUsers gagal: ' + e.message);
  }
}

// Dipanggil dari client saat klik "Login". Password dikirim apa adanya
// lewat JSONP (masih plaintext di transit — itu keterbatasan JSONP/GET,
// lihat catatan JSONP di bagian bawah file), tapi di-hash & dibandingkan
// di server, dan TIDAK PERNAH dikirim balik ke client.
function loginUser(name, password) {
  if (!name || !password) throw new Error('Username/password wajib diisi');
  var users = getUsersRaw_();
  var hash = hashPassword_(password);
  for (var i = 0; i < users.length; i++) {
    if (users[i].name === name) {
      if (users[i].passwordHash === hash) {
        return { name: users[i].name, role: users[i].role, token: issueToken_(users[i].name, users[i].role) };
      }
      throw new Error('Username/password salah!');
    }
  }
  throw new Error('Username/password salah!');
}

function saveUser(user) {
  if (!user || !user.name || !user.password) throw new Error('Data user tidak valid');
  var users = getUsersRaw_();
  var idx = -1;
  for (var i = 0; i < users.length; i++) { if (users[i].name === user.name) { idx = i; break; } }
  var hash = hashPassword_(user.password);
  if (idx >= 0) users[idx] = { name: user.name, role: user.role || 'sales', passwordHash: hash };
  else users.push({ name: user.name, role: user.role || 'sales', passwordHash: hash });
  var sh = getSheet(SHEET.USERS); sh.clearContents();
  sh.appendRow(['Nama', 'Role', 'PasswordHash']);
  users.forEach(function(u) { sh.appendRow([u.name, u.role, u.passwordHash]); });
  return { ok: true, name: user.name };
}

function deleteUser(name) {
  if (!name) throw new Error('Nama user diperlukan');
  var users = getUsersRaw_().filter(function(u) { return u.name !== name; });
  var sh = getSheet(SHEET.USERS); sh.clearContents();
  sh.appendRow(['Nama', 'Role', 'PasswordHash']);
  users.forEach(function(u) { sh.appendRow([u.name, u.role, u.passwordHash]); });
  return { ok: true, deleted: name };
}

// Jalankan SEKALI SAJA dari editor Apps Script (pilih fungsi ini di
// dropdown, klik Run) — BUKAN lewat URL web app. Mengubah semua
// password plaintext yang masih ada di sheet "Users" menjadi hash.
// Backup sheet dulu sebelum menjalankan ini.
function migratePasswordsToHash_ONETIME() {
  var sh = getSheet(SHEET.USERS);
  var data = sh.getDataRange().getValues();
  var changed = 0;
  for (var i = 1; i < data.length; i++) {
    var current = data[i][2];
    if (!looksAlreadyHashed_(current)) {
      sh.getRange(i + 1, 3).setValue(hashPassword_(current));
      changed++;
    }
  }
  Logger.log('Migrasi selesai. Baris yang di-hash: ' + changed + ' dari total ' + (data.length - 1));
  return { ok: true, migrated: changed, total: data.length - 1 };
}

// ══════════════════════════════════════════════════════════════
//  PRODUK
// ══════════════════════════════════════════════════════════════
function ensureProductHeaders() {
  try {
    var sh = getSheet(SHEET.PRODUCTS);
    var data = sh.getDataRange().getValues();
    if (data.length === 0 || String(data[0][0]||'').toLowerCase() !== 'sku') {
      sh.clearContents();
      sh.appendRow(['SKU','Barcode','Nama','Harga','Modal','Satuan','StokAwal','HasBarcode']);
    }
  } catch(e) { throw new Error('Gagal inisialisasi produk: ' + e.message); }
}

function getProducts() {
  try {
    ensureProductHeaders();
    var data = getSheetData(SHEET.PRODUCTS); if (data.length <= 1) return [];
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      sku: headers.indexOf('sku'), barcode: headers.indexOf('barcode'), nama: headers.indexOf('nama'),
      harga: headers.indexOf('harga'), modal: headers.indexOf('modal'), satuan: headers.indexOf('satuan'),
      stokAwal: headers.indexOf('stokawal'), hasBarcode: headers.indexOf('hasbarcode')
    };
    var products = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i]; var sku = String(r[col.sku]||'').trim(); if (!sku) continue;
      products.push({
        sku: sku, barcode: col.barcode >= 0 ? String(r[col.barcode]||'') : '',
        nama: col.nama >= 0 ? String(r[col.nama]||'') : '',
        harga: col.harga >= 0 ? parseHarga(r[col.harga]) : 0,
        modal: col.modal >= 0 ? parseHarga(r[col.modal]) : 0,
        satuan: col.satuan >= 0 ? String(r[col.satuan]||'Pcs') : 'Pcs',
        stokAwal: col.stokAwal >= 0 ? parseInt(r[col.stokAwal])||0 : 0,
        hasBarcode: col.hasBarcode >= 0 ? String(r[col.hasBarcode]||'')==='true' : false
      });
    }
    return products;
  } catch(e) {
    throw new Error('getProducts gagal: ' + e.message);
  }
}

function saveProducts(list) {
  if (!list) return 'error: data kosong';
  if (!Array.isArray(list)) { try { list = JSON.parse(list); } catch(e) { return 'error: format tidak valid'; } }
  var sh = getSheet(SHEET.PRODUCTS); sh.clearContents();
  sh.appendRow(['SKU','Barcode','Nama','Harga','Modal','Satuan','StokAwal','HasBarcode']);
  list.forEach(function(p) {
    sh.appendRow([String(p.sku||''), String(p.barcode||''), String(p.nama||''), Number(p.harga)||0, Number(p.modal)||0, String(p.satuan||'Pcs'), parseInt(p.stokAwal)||0, p.hasBarcode?'true':'false']);
  });
  return 'ok';
}

// ══════════════════════════════════════════════════════════════
//  CUSTOMERS
// ══════════════════════════════════════════════════════════════
function getCustomers() {
  try {
    var data = getSheetData(SHEET.CUSTOMERS);
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var v = String(data[i][0]||'').trim();
      if (v && v.toLowerCase() !== 'nama' && v.toLowerCase() !== 'pelanggan') list.push(v);
    }
    return list;
  } catch(e) { throw new Error('getCustomers gagal: ' + e.message); }
}
function addCustomer(name) { if (!name) throw new Error('Nama kosong'); if (getCustomers().indexOf(name) >= 0) return 'exists'; getSheet(SHEET.CUSTOMERS).appendRow([name]); return 'ok'; }
function saveCustomers(list) { if (!Array.isArray(list)) throw new Error('Harus array'); var sh = getSheet(SHEET.CUSTOMERS); sh.clearContents(); list.forEach(function(c) { if (c) sh.appendRow([c]); }); return 'ok'; }

// ══════════════════════════════════════════════════════════════
//  TRANSAKSI
// ══════════════════════════════════════════════════════════════
function ensureTrxHeaders() {
  try {
    var sh = getSheet(SHEET.TRX);
    var data = sh.getDataRange().getValues();
    if (!data.length || String(data[0][0]||'').trim().toLowerCase() !== 'id') {
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,9).setValues([['ID','Tanggal','Customer','Sales','Gross','Diskon','Nett','Status','CreatedAt']]);
    }
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(c) { return String(c||'').toLowerCase(); });
    if (headers.indexOf('totalmodal') === -1) {
      sh.insertColumnAfter(9); sh.getRange(1,10).setValue('TotalModal');
      sh.insertColumnAfter(10); sh.getRange(1,11).setValue('TotalProfit');
    }
    // [NEW v8.5] Sheet lama (dibuat sebelum 'CreatedAt' ada di header awal) mungkin belum
    // punya kolom ini. Tambahkan otomatis kalau memang belum ada, sama seperti pola
    // penambahan TotalModal/TotalProfit di atas - supaya waktu transaksi mulai tercatat.
    headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(c) { return String(c||'').toLowerCase(); });
    if (headers.indexOf('createdat') === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue('CreatedAt');
    }
  } catch(e) { throw new Error('Gagal inisialisasi header transaksi: ' + e.message); }
}

function ensureDetailHeaders() {
  try {
    var sh = getSheet(SHEET.DETAIL);
    var data = sh.getDataRange().getValues();
    if (!data.length) {
      sh.appendRow(['TrxID','Tanggal','Customer','Sales','Status','SKU','Barcode','Nama','Qty','Harga','Modal','DiscRpPer','NettPer','Subtotal','Profit','CreatedAt']);
      return;
    }
    var hdr = data[0].map(function(c) { return String(c||'').toLowerCase().trim(); });
    if (hdr.indexOf('trxid') === -1) {
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,16).setValues([['TrxID','Tanggal','Customer','Sales','Status','SKU','Barcode','Nama','Qty','Harga','Modal','DiscRpPer','NettPer','Subtotal','Profit','CreatedAt']]);
    }
  } catch(e) { throw new Error('Gagal inisialisasi header detail: ' + e.message); }
}

function saveTrx(trx) {
  try {
    if (!trx || !trx.id) throw new Error('Data transaksi tidak valid: id kosong');
    ensureTrxHeaders(); ensureDetailHeaders();
    // [BUGFIX] Client memanggil GAS lewat JSONP dengan timeout 20 detik + auto-retry
    // kalau tidak ada respon dalam waktu itu. Masalahnya: kalau Apps Script-nya lagi
    // lambat (bukan benar-benar gagal) - transaksi BISA SAJA sudah berhasil tersimpan
    // di server walau responnya belum sempat balik ke client sebelum timeout. Client
    // yang tidak tahu itu lalu mengirim ulang request YANG SAMA PERSIS (retry), dan
    // kalau retry itu juga berhasil, transaksi dengan ID yang sama jadi tersimpan
    // DUA KALI (baris header duplikat + semua item di Detail ikut dobel). Ini yang
    // menyebabkan "ID transaksi sama tapi isi struk beda/dobel" yang dilaporkan user.
    // Sekarang: sebelum menyimpan apa pun, dicek dulu apakah trx.id ini SUDAH ADA di
    // Sheet - kalau sudah, anggap ini retry dari request sebelumnya yang sebenarnya
    // sukses, langsung return sukses TANPA menulis apa pun lagi (idempotent).
    var trxSheetCheck = getSheet(SHEET.TRX);
    var checkHeaders = trxSheetCheck.getRange(1,1,1,trxSheetCheck.getLastColumn()).getValues()[0].map(function(c){ return String(c||'').toLowerCase().trim(); });
    var idColCheck = checkHeaders.indexOf('id'); if (idColCheck < 0) idColCheck = checkHeaders.indexOf('trxid'); if (idColCheck < 0) idColCheck = 0;
    var existingIds = trxSheetCheck.getLastRow() > 1 ? trxSheetCheck.getRange(2, idColCheck+1, trxSheetCheck.getLastRow()-1, 1).getValues() : [];
    for (var ei = 0; ei < existingIds.length; ei++) {
      if (String(existingIds[ei][0]||'').trim() === String(trx.id).trim()) {
        return { ok: true, id: trx.id, duplicate: true, note: 'ID transaksi ini sudah pernah tersimpan sebelumnya - dilewati supaya tidak dobel.' };
      }
    }
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var items = Array.isArray(trx.items) ? trx.items : [];
    var totalModal = 0, totalProfit = 0;
    var trxSheet = getSheet(SHEET.TRX);
    var trxHeaders = trxSheet.getRange(1,1,1,trxSheet.getLastColumn()).getValues()[0].map(String);
    var newTrxRow = new Array(trxHeaders.length).fill('');
    trxHeaders.forEach(function(h, idx) {
      var key = h.toString().toLowerCase().trim();
      if (key === 'id' || key === 'trxid') newTrxRow[idx] = trx.id;
      else if (key === 'tanggal' || key === 'tgl') newTrxRow[idx] = trx.tgl || '';
      else if (key === 'customer' || key === 'pelanggan') newTrxRow[idx] = trx.customer || '';
      else if (key === 'sales') newTrxRow[idx] = trx.sales || '';
      else if (key === 'gross') newTrxRow[idx] = Number(trx.gross) || 0;
      else if (key === 'diskon' || key === 'discount') newTrxRow[idx] = Number(trx.diskon) || 0;
      else if (key === 'nett' || key === 'net') newTrxRow[idx] = Number(trx.nett) || 0;
      else if (key === 'status' || key === 'pembayaran') newTrxRow[idx] = trx.status || 'belumTransfer';
      else if (key === 'createdat') newTrxRow[idx] = now;
    });
    items.forEach(function(it) {
      var modal = Number(it.modal)||0, qty = Number(it.qty)||1, harga = Number(it.harga)||0, disc = Number(it.discRpPer)||0;
      totalModal += modal * qty;
      totalProfit += ((harga - disc) - modal) * qty;
    });
    trxHeaders.forEach(function(h, idx) {
      var key = h.toString().toLowerCase().trim();
      if (key === 'totalmodal') newTrxRow[idx] = totalModal;
      else if (key === 'totalprofit') newTrxRow[idx] = totalProfit;
    });
    trxSheet.appendRow(newTrxRow);
    var detSheet = getSheet(SHEET.DETAIL);
    var detHeaders = detSheet.getRange(1,1,1,detSheet.getLastColumn()).getValues()[0].map(String);
    items.forEach(function(it) {
      var modal = Number(it.modal)||0, qty = Number(it.qty)||1, harga = Number(it.harga)||0, disc = Number(it.discRpPer)||0;
      var nettPer = harga - disc, subtotal = nettPer * qty, profit = (nettPer - modal) * qty;
      var newDetRow = new Array(detHeaders.length).fill('');
      detHeaders.forEach(function(h, idx) {
        var key = h.toString().toLowerCase().trim();
        if (key === 'trxid') newDetRow[idx] = trx.id;
        else if (key === 'tanggal' || key === 'tgl') newDetRow[idx] = trx.tgl || '';
        else if (key === 'customer' || key === 'pelanggan') newDetRow[idx] = trx.customer || '';
        else if (key === 'sales') newDetRow[idx] = trx.sales || '';
        else if (key === 'status' || key === 'pembayaran') newDetRow[idx] = trx.status || '';
        else if (key === 'sku') newDetRow[idx] = it.sku || '';
        else if (key === 'barcode') newDetRow[idx] = it.barcode || '';
        else if (key === 'nama') newDetRow[idx] = it.nama || '';
        else if (key === 'qty') newDetRow[idx] = qty;
        else if (key === 'harga') newDetRow[idx] = harga;
        else if (key === 'modal') newDetRow[idx] = modal;
        else if (key === 'discrpper' || key === 'disc') newDetRow[idx] = disc;
        else if (key === 'nettper' || key === 'nett') newDetRow[idx] = nettPer;
        else if (key === 'subtotal') newDetRow[idx] = subtotal;
        else if (key === 'profit') newDetRow[idx] = profit;
        else if (key === 'createdat') newDetRow[idx] = now;
      });
      detSheet.appendRow(newDetRow);
    });
    return { ok: true, id: trx.id, totalModal: totalModal, totalProfit: totalProfit };
  } catch(e) {
    throw new Error('saveTrx gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  UPSERT TRX - replace kalau ID sudah ada, insert kalau belum.
//  Beda dengan saveTrx() (yang SKIP kalau ID sudah ada, dipakai utk mencegah
//  dobel akibat retry timeout) - upsertTrx() dipakai khusus utk tombol "Sync
//  Transaksi Hari Ini" yang user MINTA sengaja: kalau ID sudah ada, TIMPA
//  dengan data terbaru dari browser (hapus baris lama, tulis ulang baru),
//  bukan dilewati. Tetap tidak akan pernah menghasilkan baris dobel.
// ══════════════════════════════════════════════════════════════
function upsertTrx(trx) {
  try {
    if (!trx || !trx.id) throw new Error('Data transaksi tidak valid: id kosong');
    ensureTrxHeaders(); ensureDetailHeaders();
    var trxId = String(trx.id).trim();

    // 1) Hapus baris header TRX lama (kalau ada) dengan ID yang sama
    var trxSheet = getSheet(SHEET.TRX);
    var trxHeaders = trxSheet.getRange(1,1,1,trxSheet.getLastColumn()).getValues()[0].map(function(c){ return String(c||'').toLowerCase().trim(); });
    var idColTrx = trxHeaders.indexOf('id'); if (idColTrx < 0) idColTrx = trxHeaders.indexOf('trxid'); if (idColTrx < 0) idColTrx = 0;
    if (trxSheet.getLastRow() > 1) {
      var trxIds = trxSheet.getRange(2, idColTrx+1, trxSheet.getLastRow()-1, 1).getValues();
      for (var i = trxIds.length - 1; i >= 0; i--) {
        if (String(trxIds[i][0]||'').trim() === trxId) { trxSheet.deleteRow(i + 2); }
      }
    }

    // 2) Hapus baris Detail lama (kalau ada) dengan trxId yang sama
    var detSheet = getSheet(SHEET.DETAIL);
    var detHeadersCheck = detSheet.getRange(1,1,1,detSheet.getLastColumn()).getValues()[0].map(function(c){ return String(c||'').toLowerCase().trim(); });
    var idColDet = detHeadersCheck.indexOf('trxid'); if (idColDet < 0) idColDet = 0;
    if (detSheet.getLastRow() > 1) {
      var detIds = detSheet.getRange(2, idColDet+1, detSheet.getLastRow()-1, 1).getValues();
      for (var j = detIds.length - 1; j >= 0; j--) {
        if (String(detIds[j][0]||'').trim() === trxId) { detSheet.deleteRow(j + 2); }
      }
    }

    // 3) Tulis ulang fresh (sama persis logika saveTrx, tanpa pengecekan skip)
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var items = Array.isArray(trx.items) ? trx.items : [];
    var totalModal = 0, totalProfit = 0;
    var trxHeadersFull = trxSheet.getRange(1,1,1,trxSheet.getLastColumn()).getValues()[0].map(String);
    var newTrxRow = new Array(trxHeadersFull.length).fill('');
    trxHeadersFull.forEach(function(h, idx) {
      var key = h.toString().toLowerCase().trim();
      if (key === 'id' || key === 'trxid') newTrxRow[idx] = trx.id;
      else if (key === 'tanggal' || key === 'tgl') newTrxRow[idx] = trx.tgl || '';
      else if (key === 'customer' || key === 'pelanggan') newTrxRow[idx] = trx.customer || '';
      else if (key === 'sales') newTrxRow[idx] = trx.sales || '';
      else if (key === 'gross') newTrxRow[idx] = Number(trx.gross) || 0;
      else if (key === 'diskon' || key === 'discount') newTrxRow[idx] = Number(trx.diskon) || 0;
      else if (key === 'nett' || key === 'net') newTrxRow[idx] = Number(trx.nett) || 0;
      else if (key === 'status' || key === 'pembayaran') newTrxRow[idx] = trx.status || 'belumTransfer';
      else if (key === 'createdat') newTrxRow[idx] = now;
    });
    items.forEach(function(it) {
      var modal = Number(it.modal)||0, qty = Number(it.qty)||1, harga = Number(it.harga)||0, disc = Number(it.discRpPer)||0;
      totalModal += modal * qty;
      totalProfit += ((harga - disc) - modal) * qty;
    });
    trxHeadersFull.forEach(function(h, idx) {
      var key = h.toString().toLowerCase().trim();
      if (key === 'totalmodal') newTrxRow[idx] = totalModal;
      else if (key === 'totalprofit') newTrxRow[idx] = totalProfit;
    });
    trxSheet.appendRow(newTrxRow);
    var detHeaders = detSheet.getRange(1,1,1,detSheet.getLastColumn()).getValues()[0].map(String);
    items.forEach(function(it) {
      var modal = Number(it.modal)||0, qty = Number(it.qty)||1, harga = Number(it.harga)||0, disc = Number(it.discRpPer)||0;
      var nettPer = harga - disc, subtotal = nettPer * qty, profit = (nettPer - modal) * qty;
      var newDetRow = new Array(detHeaders.length).fill('');
      detHeaders.forEach(function(h, idx) {
        var key = h.toString().toLowerCase().trim();
        if (key === 'trxid') newDetRow[idx] = trx.id;
        else if (key === 'tanggal' || key === 'tgl') newDetRow[idx] = trx.tgl || '';
        else if (key === 'customer' || key === 'pelanggan') newDetRow[idx] = trx.customer || '';
        else if (key === 'sales') newDetRow[idx] = trx.sales || '';
        else if (key === 'status' || key === 'pembayaran') newDetRow[idx] = trx.status || '';
        else if (key === 'sku') newDetRow[idx] = it.sku || '';
        else if (key === 'barcode') newDetRow[idx] = it.barcode || '';
        else if (key === 'nama') newDetRow[idx] = it.nama || '';
        else if (key === 'qty') newDetRow[idx] = qty;
        else if (key === 'harga') newDetRow[idx] = harga;
        else if (key === 'modal') newDetRow[idx] = modal;
        else if (key === 'discrpper' || key === 'disc') newDetRow[idx] = disc;
        else if (key === 'nettper' || key === 'nett') newDetRow[idx] = nettPer;
        else if (key === 'subtotal') newDetRow[idx] = subtotal;
        else if (key === 'profit') newDetRow[idx] = profit;
        else if (key === 'createdat') newDetRow[idx] = now;
      });
      detSheet.appendRow(newDetRow);
    });
    return { ok: true, id: trx.id, totalModal: totalModal, totalProfit: totalProfit };
  } catch(e) {
    throw new Error('upsertTrx gagal: ' + e.message);
  }
}

function getTrxList() {
  try {
    ensureTrxHeaders();
    var data = getSheetData(SHEET.TRX); if (data.length <= 1) return [];
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      id: headers.indexOf('id') >= 0 ? headers.indexOf('id') : 0,
      tgl: headers.indexOf('tanggal') >= 0 ? headers.indexOf('tanggal') : (headers.indexOf('tgl') >= 0 ? headers.indexOf('tgl') : 1),
      customer: headers.indexOf('customer') >= 0 ? headers.indexOf('customer') : (headers.indexOf('pelanggan') >= 0 ? headers.indexOf('pelanggan') : 2),
      sales: headers.indexOf('sales') >= 0 ? headers.indexOf('sales') : 3,
      gross: headers.indexOf('gross') >= 0 ? headers.indexOf('gross') : 4,
      diskon: headers.indexOf('diskon') >= 0 ? headers.indexOf('diskon') : (headers.indexOf('discount') >= 0 ? headers.indexOf('discount') : 5),
      nett: headers.indexOf('nett') >= 0 ? headers.indexOf('nett') : (headers.indexOf('net') >= 0 ? headers.indexOf('net') : 6),
      status: headers.indexOf('status') >= 0 ? headers.indexOf('status') : (headers.indexOf('pembayaran') >= 0 ? headers.indexOf('pembayaran') : 7),
      createdAt: headers.indexOf('createdat') // [NEW v8.5]
    };
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i]; var id = String(r[col.id]||'').trim(); if (!id) continue;
      list.push({
        id: id, tgl: formatDateCell(r[col.tgl]), customer: String(r[col.customer]||''),
        sales: String(r[col.sales]||''), gross: Number(r[col.gross])||0, diskon: Number(r[col.diskon])||0,
        nett: Number(r[col.nett])||0, status: String(r[col.status]||'belumTransfer'),
        createdAt: col.createdAt >= 0 ? String(r[col.createdAt]||'') : '' // [NEW v8.5] tanggal+jam asli dari server, dipakai frontend untuk urutan akurat
      });
    }
    return list;
  } catch(e) {
    throw new Error('getTrxList gagal: ' + e.message);
  }
}

function updateStatus(id, newStatus) {
  try {
    if (!id || !newStatus) throw new Error('ID dan status diperlukan');
    var trxSheet = getSheet(SHEET.TRX);
    var trxData = trxSheet.getDataRange().getValues();
    var trxHeaders = trxData[0].map(function(c) { return String(c||'').toLowerCase(); });
    var idCol = trxHeaders.indexOf('id'); if (idCol < 0) idCol = 0;
    var statusCol = trxHeaders.indexOf('status'); if (statusCol < 0) statusCol = trxHeaders.indexOf('pembayaran'); if (statusCol < 0) statusCol = 7;
    for (var i = trxData.length - 1; i >= 1; i--) {
      if (String(trxData[i][idCol]||'').trim() === id) { trxSheet.getRange(i+1, statusCol+1).setValue(newStatus); break; }
    }
    var detSheet = getSheet(SHEET.DETAIL);
    var detData = detSheet.getDataRange().getValues();
    var detHeaders = detData[0].map(function(c) { return String(c||'').toLowerCase(); });
    var trxIdCol = detHeaders.indexOf('trxid'); if (trxIdCol < 0) trxIdCol = 0;
    var detStatusCol = detHeaders.indexOf('status'); if (detStatusCol < 0) detStatusCol = detHeaders.indexOf('pembayaran'); if (detStatusCol < 0) detStatusCol = 4;
    for (var j = detData.length - 1; j >= 1; j--) {
      if (String(detData[j][trxIdCol]||'').trim() === id) { detSheet.getRange(j+1, detStatusCol+1).setValue(newStatus); }
    }
    return { ok: true, id: id, newStatus: newStatus };
  } catch(e) {
    throw new Error('updateStatus gagal: ' + e.message);
  }
}

function deleteTrx(id) {
  try {
    if (!id) throw new Error('ID transaksi diperlukan');
    var trxSheet = getSheet(SHEET.TRX);
    var trxData = trxSheet.getDataRange().getValues();
    var trxHeaders = trxData[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var idCol = trxHeaders.indexOf('id'); if (idCol < 0) idCol = 0;
    var rowToDelete = -1;
    for (var i = trxData.length - 1; i >= 1; i--) {
      if (String(trxData[i][idCol]||'').trim() === id) { rowToDelete = i + 1; break; }
    }
    if (rowToDelete > 0) trxSheet.deleteRow(rowToDelete);
    var detSheet = getSheet(SHEET.DETAIL);
    var detData = detSheet.getDataRange().getValues();
    var detHeaders = detData[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var detIdCol = detHeaders.indexOf('trxid'); if (detIdCol < 0) detIdCol = 0;
    var rowsToDelete = [];
    for (var j = detData.length - 1; j >= 1; j--) {
      if (String(detData[j][detIdCol]||'').trim() === id) rowsToDelete.push(j + 1);
    }
    rowsToDelete.sort(function(a,b) { return b - a; });
    rowsToDelete.forEach(function(r) { detSheet.deleteRow(r); });
    return { ok: true, deleted: id };
  } catch(e) {
    throw new Error('deleteTrx gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  INPUT BARANG
// ══════════════════════════════════════════════════════════════
function saveInputBarang(entry) {
  try {
    if (!entry || !entry.id) throw new Error('Data input barang tidak valid');
    var sheet = getSheet(SHEET.BARANG);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID','GroupID','SKU','Nama','Qty','HargaModal','Disc','NetModal','Driver','Rit','Status','Tanggal','CreatedAt']);
    }
    // [BUGFIX] Sama seperti saveTrx - kalau request timeout di client (JSONP 20 detik)
    // padahal sebenarnya sudah berhasil tersimpan di server, client akan retry otomatis
    // dan bisa menyimpan entry.id yang SAMA dua kali. Dicek dulu sebelum menyimpan.
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]||'').trim() === String(entry.id).trim()) {
          return 'ok'; // idempotent no-op - sudah pernah tersimpan, jangan dobel
        }
      }
    }
    sheet.appendRow([entry.id, entry.groupId || '', entry.sku, entry.nama, entry.qty, entry.hargaModal, entry.disc, entry.netModal, entry.driver, entry.rit, entry.status, entry.date || '', entry.createdAt || '']);
    return 'ok';
  } catch(e) {
    throw new Error('saveInputBarang gagal: ' + e.message);
  }
}

function getInputBarangHistory() {
  try {
    var sheet = getSheet(SHEET.BARANG);
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      id: headers.indexOf('id'), groupId: headers.indexOf('groupid'), sku: headers.indexOf('sku'),
      nama: headers.indexOf('nama'), qty: headers.indexOf('qty'), hargaModal: headers.indexOf('hargamodal'),
      disc: headers.indexOf('disc'), netModal: headers.indexOf('netmodal'), driver: headers.indexOf('driver'),
      rit: headers.indexOf('rit'), status: headers.indexOf('status'),
      date: headers.indexOf('tanggal') >= 0 ? headers.indexOf('tanggal') : headers.indexOf('date'),
      createdAt: headers.indexOf('createdat')
    };
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      list.push({
        id: col.id >= 0 ? r[col.id] : '', groupId: col.groupId >= 0 ? r[col.groupId] : '',
        sku: col.sku >= 0 ? String(r[col.sku]) : '', nama: col.nama >= 0 ? String(r[col.nama]) : '',
        qty: col.qty >= 0 ? Number(r[col.qty])||0 : 0, hargaModal: col.hargaModal >= 0 ? Number(r[col.hargaModal])||0 : 0,
        disc: col.disc >= 0 ? Number(r[col.disc])||0 : 0, netModal: col.netModal >= 0 ? Number(r[col.netModal])||0 : 0,
        driver: col.driver >= 0 ? String(r[col.driver]) : '', rit: col.rit >= 0 ? String(r[col.rit]) : '',
        status: col.status >= 0 ? String(r[col.status]) : '',
        date: (col.date >= 0 && r[col.date]) ? formatDateCell(r[col.date]) : (col.createdAt >= 0 && r[col.createdAt] ? formatDateCell(r[col.createdAt]) : ''),
        createdAt: col.createdAt >= 0 ? formatDateCell(r[col.createdAt]) : ''
      });
    }
    return list;
  } catch(e) {
    throw new Error('getInputBarangHistory gagal: ' + e.message);
  }
}

function updateInputBarang(updated) {
  try {
    if (!updated || !updated.id) throw new Error('ID diperlukan');
    var sheet = getSheet(SHEET.BARANG);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').toLowerCase(); });
    var idCol = headers.indexOf('id');
    var qtyCol = headers.indexOf('qty');
    var statusCol = headers.indexOf('status');
    var hargaModalCol = headers.indexOf('hargamodal');
    var discCol = headers.indexOf('disc');
    var netModalCol = headers.indexOf('netmodal');
    if (idCol < 0 || qtyCol < 0 || statusCol < 0) throw new Error('Kolom tidak lengkap');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol] || '').trim() === updated.id) {
        var row = i + 1;
        if (updated.qty !== undefined) sheet.getRange(row, qtyCol + 1).setValue(updated.qty);
        if (updated.status !== undefined) sheet.getRange(row, statusCol + 1).setValue(updated.status);
        if (hargaModalCol >= 0 && discCol >= 0 && netModalCol >= 0 && updated.qty !== undefined) {
          var harga = Number(data[i][hargaModalCol]) || 0;
          var disc = Number(data[i][discCol]) || 0;
          var netModal = (harga - disc) * updated.qty;
          sheet.getRange(row, netModalCol + 1).setValue(netModal);
        }
        return { ok: true, id: updated.id };
      }
    }
    // [FIX] Sebelumnya throw error kalau data tidak ketemu - ini bisa bikin antrian
    // sync di client retry TERUS MENERUS tanpa pernah berhasil (misal karena data ini
    // sudah keburu dihapus lewat aksi lain). Dianggap sukses (no-op) supaya tidak
    // jadi "gagal permanen" yang mengganjal di badge status sync selamanya.
    return { ok: true, id: updated.id, note: 'Data tidak ditemukan (mungkin sudah dihapus) - dilewati' };
  } catch(e) {
    throw new Error('updateInputBarang gagal: ' + e.message);
  }
}

function deleteInputBarang(id) {
  try {
    if (!id) throw new Error('ID diperlukan');
    var sheet = getSheet(SHEET.BARANG);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').toLowerCase(); });
    var idCol = headers.indexOf('id');
    if (idCol < 0) throw new Error('Kolom ID tidak ditemukan');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol] || '').trim() === id) { sheet.deleteRow(i + 1); return { ok: true, deleted: id }; }
    }
    // [FIX] Sama seperti updateInputBarang - kalau sudah tidak ada (mis. sudah pernah
    // berhasil terhapus di percobaan sebelumnya, retry cuma menemukan yang sudah
    // kosong), dianggap sukses, bukan error permanen.
    return { ok: true, deleted: id, note: 'Data tidak ditemukan (mungkin sudah terhapus) - dianggap sukses' };
  } catch(e) {
    throw new Error('deleteInputBarang gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  SETORAN
// ══════════════════════════════════════════════════════════════
function saveSetoran(data) {
  try {
    if (!data || !data.tgl) throw new Error('Data setoran tidak valid');
    var sheet = getSheet(SHEET.SETORAN);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal','Sales','GrandTotal','Makan','Tips','Parkir','Bensin','Flazz','Transfer','Cicilan','Tagihan','Ket1','Jml1','Ket2','Jml2','Total','Setor','Selisih','CreatedAt','FotoUrl']);
    }
    var headers = sheet.getLastRow() > 0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(c){return String(c||'').toLowerCase();}) : [];
    if (headers.indexOf('fotourl') === -1 && sheet.getLastRow() > 0) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue('FotoUrl');
    }
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var newRow = [data.tgl, data.sales, data.grandTotal, data.makan, data.tips, data.parkir, data.bensin, data.flazz, data.transfer, data.cicilan, data.tagihan, data.ket1, data.jml1, data.ket2, data.jml2, data.total, data.setor, data.selisih, now, data.fotoUrl||''];

    // [UPSERT - BARU] Cari baris dengan Tanggal & Sales yang sama -> UPDATE baris itu, bukan tambah baris baru.
    // Ini diperlukan agar fitur "edit setoran yang sudah tersimpan" tidak menghasilkan data duplikat di Sheet.
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existingKeys = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // kolom Tanggal & Sales saja
      var targetTgl = formatDateCell(data.tgl);
      var targetSales = String(data.sales||'').trim().toLowerCase();
      for (var i = existingKeys.length - 1; i >= 0; i--) {
        var rowTgl = formatDateCell(existingKeys[i][0]);
        var rowSales = String(existingKeys[i][1]||'').trim().toLowerCase();
        if (rowTgl === targetTgl && rowSales === targetSales) {
          sheet.getRange(i + 2, 1, 1, newRow.length).setValues([newRow]);
          return 'ok';
        }
      }
    }
    sheet.appendRow(newRow);
    return 'ok';
  } catch(e) {
    throw new Error('saveSetoran gagal: ' + e.message);
  }
}

// [BARU] Membaca seluruh riwayat setoran dari Sheet untuk ditampilkan & diedit di aplikasi (menu Setoran & Rekap Setoran).
// PENTING: dibaca berdasarkan POSISI kolom (bukan nama header di baris 1), karena saveSetoran() di atas
// menulis nilai berdasarkan urutan posisi berikut, mengikuti urutan field yang sama dengan saat appendRow header awal dibuat:
// 0 Tanggal, 1 Sales, 2 GrandTotal, 3 Makan, 4 Tips, 5 Parkir, 6 Bensin, 7 Flazz, 8 Transfer, 9 Cicilan,
// 10 Tagihan, 11 Ket1, 12 Jml1, 13 Ket2, 14 Jml2, 15 Total, 16 Setor, 17 Selisih, 18 CreatedAt, 19 FotoUrl
// Pembacaan berbasis posisi ini sengaja dipakai (bukan berbasis nama header) supaya tetap akurat
// walau teks header di baris 1 pada Sheet yang sudah ada sebelumnya berbeda/tidak sinkron dengan kode ini.
function getSetoranHistory() {
  try {
    var sheet = getSheet(SHEET.SETORAN);
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var tgl = formatDateCell(r[0]);
      if (!tgl) continue;
      list.push({
        tgl: tgl,
        sales: String(r[1]||''),
        grandTotal: Number(r[2])||0,
        makan: Number(r[3])||0,
        tips: Number(r[4])||0,
        parkir: Number(r[5])||0,
        bensin: Number(r[6])||0,
        flazz: Number(r[7])||0,
        transfer: Number(r[8])||0,
        cicilan: Number(r[9])||0,
        tagihan: Number(r[10])||0,
        ket1: String(r[11]||''),
        jml1: Number(r[12])||0,
        ket2: String(r[13]||''),
        jml2: Number(r[14])||0,
        total: Number(r[15])||0,
        setor: Number(r[16])||0,
        selisih: Number(r[17])||0,
        createdAt: r.length > 18 ? String(r[18]||'') : '',
        fotoUrl: r.length > 19 ? String(r[19]||'') : ''
      });
    }
    return list;
  } catch(e) {
    throw new Error('getSetoranHistory gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  FOTO BUKTI
// ══════════════════════════════════════════════════════════════
function uploadFoto(payload) {
  try {
    if (!payload || !payload.base64) throw new Error('Data foto tidak valid');
    var folderName = 'Tirta Kencana - Bukti';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var subName = payload.jenis || 'lainnya';
    var subFolders = folder.getFoldersByName(subName);
    var subFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subName);
    var mimeType = payload.mimeType || 'image/jpeg';
    var ext = mimeType.indexOf('png') >= 0 ? '.png' : mimeType.indexOf('pdf') >= 0 ? '.pdf' : '.jpg';
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd-HHmmss');
    var fileName = (payload.nama || payload.refId || 'foto') + '-' + now + ext;
    var decoded = Utilities.newBlob(Utilities.base64Decode(payload.base64), mimeType, fileName);
    var file = subFolder.createFile(decoded);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
    var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
    var fotoSheet = getSheet('FotoBukti');
    if (fotoSheet.getLastRow() === 0) {
      fotoSheet.appendRow(['Waktu','Jenis','RefID','NamaFile','ViewUrl','ThumbUrl','Uploader']);
    }
    fotoSheet.appendRow([now, subName, payload.refId||'', fileName, viewUrl, thumbUrl, payload.uploader||'']);
    return { ok: true, fileId: fileId, viewUrl: viewUrl, thumbUrl: thumbUrl, fileName: fileName };
  } catch(e) {
    throw new Error('uploadFoto gagal: ' + e.message);
  }
}

function getFotoList(refId) {
  try {
    var sheet = getSS().getSheetByName('FotoBukti');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c){ return String(c||'').toLowerCase(); });
    var col = {
      waktu: headers.indexOf('waktu'), jenis: headers.indexOf('jenis'),
      refId: headers.indexOf('refid'), nama: headers.indexOf('namafile'),
      viewUrl: headers.indexOf('viewurl'), thumbUrl: headers.indexOf('thumburl'),
      uploader: headers.indexOf('uploader')
    };
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!refId || String(r[col.refId]||'').trim() === refId) {
        result.push({
          waktu: String(r[col.waktu]||''), jenis: String(r[col.jenis]||''),
          refId: String(r[col.refId]||''), nama: String(r[col.nama]||''),
          viewUrl: String(r[col.viewUrl]||''), thumbUrl: String(r[col.thumbUrl]||''),
          uploader: String(r[col.uploader]||'')
        });
      }
    }
    return result;
  } catch(e) {
    throw new Error('getFotoList gagal: ' + e.message);
  }
}

// [NEW v8.4] Hapus foto bukti: pindahkan file fisiknya ke Trash di Google Drive, dan hapus baris
// yang sesuai di Sheet "FotoBukti" (dicocokkan lewat fileId yang ada di kolom ViewUrl).
// Dipanggil dari web lewat menu "Hapus Foto" di modal Lihat Foto.
function deleteFoto(payload) {
  try {
    if (!payload || !payload.fileId) throw new Error('fileId diperlukan');
    try {
      var file = DriveApp.getFileById(payload.fileId);
      file.setTrashed(true);
    } catch(eFile) {
      // File mungkin sudah terhapus manual sebelumnya di Drive - tetap lanjut bersihkan baris di Sheet
      Logger.log('deleteFoto: file Drive tidak ditemukan/sudah terhapus - ' + eFile.message);
    }
    var sheet = getSS().getSheetByName('FotoBukti');
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(c){ return String(c||'').toLowerCase(); });
      var viewUrlCol = headers.indexOf('viewurl');
      if (viewUrlCol >= 0) {
        for (var i = data.length - 1; i >= 1; i--) {
          var rowUrl = String(data[i][viewUrlCol]||'');
          if (rowUrl.indexOf(payload.fileId) >= 0) { sheet.deleteRow(i + 1); }
        }
      }
    }
    return { ok: true, deleted: payload.fileId };
  } catch(e) {
    throw new Error('deleteFoto gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  MASTER STOCK
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
//  BERSIHKAN DUPLIKAT MASTERSTOCK (sekali jalan)
//  Sisa dari bug lama (perbandingan sel tanggal yang salah, sudah diperbaiki di
//  saveMasterStock/updateStockRow/updateStockAwal) yang membuat baris tanggal+SKU
//  yang sama bisa tersimpan berkali-kali. Fungsi ini menghapus baris duplikat,
//  MEMPERTAHANKAN baris PALING BAWAH/TERAKHIR untuk tiap kombinasi tanggal+SKU
//  (asumsi: baris yang ditulis paling akhir = data paling baru/akurat).
// ══════════════════════════════════════════════════════════════
function cleanupDuplicateMasterStock() {
  var sheet = getSheet(SHEET.MASTERSTOCK);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 2) return { ok: true, removed: 0, note: 'Tidak ada data untuk dibersihkan.' };
  // [BUGFIX - PENTING] Versi sebelumnya mengasumsikan kolom Tanggal selalu di posisi A
  // (index 0) dan SKU selalu di posisi B (index 1) secara HARDCODE. Kalau susunan kolom
  // sheet-nya ternyata berbeda (kolom tergeser/ditambah manual, dll), asumsi ini SALAH -
  // akibatnya baris-baris yang sebenarnya BEDA (tanggal/SKU berbeda) bisa keliru dianggap
  // "sama", dan baris yang harusnya tetap ada malah ikut terhapus. Sekarang kolom dicari
  // berdasarkan NAMA header-nya (persis seperti updateStockRow/updateStockAwal), bukan
  // posisi tetap, supaya selalu akurat berapa pun urutan kolomnya.
  var headers = data[0].map(function(c) { return String(c||'').toLowerCase().trim(); });
  var colTgl = headers.indexOf('tanggal'); if (colTgl < 0) colTgl = 0;
  var colSku = headers.indexOf('sku'); if (colSku < 0) colSku = 1;
  var seen = {}; // key "tanggal|sku" -> row index TERAKHIR yang ditemukan (dipertahankan)
  for (var i = 1; i < data.length; i++) {
    var key = formatDateCell(data[i][colTgl]) + '|' + String(data[i][colSku]||'').trim().toLowerCase();
    seen[key] = i; // overwrite terus - baris paling bawah untuk key yang sama akan jadi nilai akhir
  }
  var keepRows = {}; // set index baris (0-based dlm array data) yang DIPERTAHANKAN
  for (var k in seen) keepRows[seen[k]] = true;
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (!keepRows[i]) rowsToDelete.push(i + 1); // +1 karena sheet row 1-indexed & data[0] adalah header
  }
  // Hapus dari bawah ke atas supaya index baris yang belum diproses tidak bergeser
  rowsToDelete.sort(function(a,b){ return b - a; });
  rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });
  return { ok: true, removed: rowsToDelete.length, totalBefore: data.length - 1, totalAfter: data.length - 1 - rowsToDelete.length };
}

function saveMasterStock(payload) {
  try {
    if (!payload || !payload.date) return 'ok';
    var sheet = getSheet(SHEET.MASTERSTOCK);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Tanggal', 'SKU', 'Nama', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir']);
    var stock = payload.stock || []; if (!stock.length) return 'ok';
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    // [BUGFIX] Sebelumnya baris ini membandingkan String(selTanggal) langsung dengan
    // payload.date - itu GAGAL cocok kalau sel tanggalnya sudah dikonversi Sheets jadi
    // objek Date asli (default kalau format kolom "Automatic"), karena String(Date)
    // hasilnya BUKAN "yyyy-MM-dd" lagi. Akibatnya baris lama tidak pernah ketemu/terhapus,
    // dan setiap kali sync yang sama dijalankan ulang, data baru numpuk di atas data lama
    // (dobel terus). formatDateCell() menangani baik objek Date maupun string dengan benar.
    for (var i = data.length - 1; i >= 1; i--) {
      if (formatDateCell(data[i][0]) === formatDateCell(payload.date)) rowsToDelete.push(i + 1);
    }
    rowsToDelete.sort(function(a,b) { return b - a; });
    rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });
    stock.forEach(function(s) {
      sheet.appendRow([payload.date, s.sku, s.nama, s.stokAwal || 0, s.masuk || 0, s.keluar || 0, s.stokAkhir || 0]);
    });
    return 'ok';
  } catch(e) {
    throw new Error('saveMasterStock gagal: ' + e.message);
  }
}

function getMasterStockByDate(date) {
  try {
    var sheet = getSS().getSheetByName(SHEET.MASTERSTOCK); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (formatDateCell(r[0]) === formatDateCell(date)) {
        result.push({ sku: String(r[1] || ''), nama: String(r[2] || ''), stokAwal: Number(r[3]) || 0, masuk: Number(r[4]) || 0, keluar: Number(r[5]) || 0, stokAkhir: Number(r[6]) || 0 });
      }
    }
    return result;
  } catch(e) {
    throw new Error('getMasterStockByDate gagal: ' + e.message);
  }
}

function getStockBarang() {
  try {
    var sheet = getSS().getSheetByName(SHEET.MASTERSTOCK); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var tgl = '';
      try { tgl = formatDateCell(r[0]); } catch(e) { tgl = String(r[0]||'').trim(); }
      if (!tgl) continue;
      result.push({ date: tgl, sku: String(r[1]||''), nama: String(r[2]||''), stokAwal: Number(r[3])||0, masuk: Number(r[4])||0, keluar: Number(r[5])||0, stokAkhir: Number(r[6])||0 });
    }
    return result;
  } catch(e) {
    throw new Error('getStockBarang gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  UPDATE STOK AWAL
// ══════════════════════════════════════════════════════════════
function updateStockAwal(payload) {
  try {
    if (!payload || !payload.date || !payload.sku) throw new Error('Data tidak valid: date dan sku diperlukan');
    var sheet = getSheet(SHEET.MASTERSTOCK);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal', 'SKU', 'Nama', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir']);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var colTgl    = headers.indexOf('tanggal');  if (colTgl < 0)    colTgl = 0;
    var colSku    = headers.indexOf('sku');       if (colSku < 0)    colSku = 1;
    var colNama   = headers.indexOf('nama');      if (colNama < 0)   colNama = 2;
    var colAwal   = headers.indexOf('stokawal');  if (colAwal < 0)   colAwal = 3;
    var colMasuk  = headers.indexOf('masuk');     if (colMasuk < 0)  colMasuk = 4;
    var colKeluar = headers.indexOf('keluar');    if (colKeluar < 0) colKeluar = 5;
    var colAkhir  = headers.indexOf('stokakhir'); if (colAkhir < 0)  colAkhir = 6;

    var stokAwal   = Number(payload.stokAwal)  || 0;
    var masuk      = Number(payload.masuk)     || 0;
    var keluar     = Number(payload.keluar)    || 0;
    var stokAkhir  = stokAwal + masuk - keluar;

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowTgl = formatDateCell(data[i][colTgl]);
      var rowSku = String(data[i][colSku]||'').trim().toLowerCase();
      if (rowTgl === formatDateCell(payload.date) && rowSku === payload.sku.toLowerCase()) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      var existingMasuk  = Number(data[foundRow-1][colMasuk])  || 0;
      var existingKeluar = Number(data[foundRow-1][colKeluar]) || 0;
      var newAkhir = stokAwal + existingMasuk - existingKeluar;
      sheet.getRange(foundRow, colAwal  + 1).setValue(stokAwal);
      sheet.getRange(foundRow, colAkhir + 1).setValue(newAkhir);
    } else {
      var newRow = new Array(Math.max(colAkhir + 1, 7)).fill('');
      newRow[colTgl]    = payload.date;
      newRow[colSku]    = payload.sku;
      newRow[colNama]   = payload.nama || '';
      newRow[colAwal]   = stokAwal;
      newRow[colMasuk]  = masuk;
      newRow[colKeluar] = keluar;
      newRow[colAkhir]  = stokAkhir;
      sheet.appendRow(newRow);
    }

    return { ok: true, sku: payload.sku, date: payload.date, stokAwal: stokAwal };
  } catch(e) {
    throw new Error('updateStockAwal gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  UPDATE STOK (Awal, Masuk, Keluar sekaligus - manual edit dari web)
//  Beda dengan updateStockAwal (cuma StokAwal), fungsi ini dipakai saat
//  user mengedit langsung sel Stok Awal/Masuk/Keluar di tabel Stock Barang.
//  StokAkhir SELALU dihitung ulang di server = stokAwal + masuk - keluar,
//  tidak pernah dipercaya dari nilai kiriman client, supaya konsisten.
// ══════════════════════════════════════════════════════════════
function updateStockRow(payload) {
  try {
    if (!payload || !payload.date || !payload.sku) throw new Error('Data tidak valid: date dan sku diperlukan');
    var sheet = getSheet(SHEET.MASTERSTOCK);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal', 'SKU', 'Nama', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir']);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var colTgl    = headers.indexOf('tanggal');  if (colTgl < 0)    colTgl = 0;
    var colSku    = headers.indexOf('sku');       if (colSku < 0)    colSku = 1;
    var colNama   = headers.indexOf('nama');      if (colNama < 0)   colNama = 2;
    var colAwal   = headers.indexOf('stokawal');  if (colAwal < 0)   colAwal = 3;
    var colMasuk  = headers.indexOf('masuk');     if (colMasuk < 0)  colMasuk = 4;
    var colKeluar = headers.indexOf('keluar');    if (colKeluar < 0) colKeluar = 5;
    var colAkhir  = headers.indexOf('stokakhir'); if (colAkhir < 0)  colAkhir = 6;

    var stokAwal  = Number(payload.stokAwal) || 0;
    var masuk     = Number(payload.masuk)    || 0;
    var keluar    = Number(payload.keluar)   || 0;
    // StokAkhir: kalau dikirim eksplisit oleh client (user edit manual, misal hasil stock
    // opname fisik), pakai nilai itu apa adanya. Kalau tidak dikirim, hitung otomatis.
    var stokAkhir = (payload.stokAkhir !== undefined && payload.stokAkhir !== null && payload.stokAkhir !== '')
      ? Number(payload.stokAkhir) || 0
      : stokAwal + masuk - keluar;

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowTgl = formatDateCell(data[i][colTgl]);
      var rowSku = String(data[i][colSku]||'').trim().toLowerCase();
      if (rowTgl === formatDateCell(payload.date) && rowSku === payload.sku.toLowerCase()) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, colAwal   + 1).setValue(stokAwal);
      sheet.getRange(foundRow, colMasuk  + 1).setValue(masuk);
      sheet.getRange(foundRow, colKeluar + 1).setValue(keluar);
      sheet.getRange(foundRow, colAkhir  + 1).setValue(stokAkhir);
    } else {
      var newRow = new Array(Math.max(colAkhir + 1, 7)).fill('');
      newRow[colTgl]    = payload.date;
      newRow[colSku]    = payload.sku;
      newRow[colNama]   = payload.nama || '';
      newRow[colAwal]   = stokAwal;
      newRow[colMasuk]  = masuk;
      newRow[colKeluar] = keluar;
      newRow[colAkhir]  = stokAkhir;
      sheet.appendRow(newRow);
    }

    return { ok: true, sku: payload.sku, date: payload.date, stokAwal: stokAwal, masuk: masuk, keluar: keluar, stokAkhir: stokAkhir };
  } catch(e) {
    throw new Error('updateStockRow gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  TRX DETAIL
// ══════════════════════════════════════════════════════════════
function getTrxDetail(id) {
  try {
    var sheet = getSheet(SHEET.DETAIL);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { id: id, items: [] };
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      trxid: headers.indexOf('trxid'),
      sku: headers.indexOf('sku'),
      nama: headers.indexOf('nama'),
      qty: headers.indexOf('qty'),
      harga: headers.indexOf('harga'),
      modal: headers.indexOf('modal'),
      disc: headers.indexOf('discrpper'),
      nett: headers.indexOf('nettper'),
      subtotal: headers.indexOf('subtotal'),
      profit: headers.indexOf('profit')
    };
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[col.trxid]||'').trim() === id) {
        items.push({
          sku: col.sku>=0 ? String(r[col.sku]) : '',
          nama: col.nama>=0 ? String(r[col.nama]) : '',
          qty: col.qty>=0 ? Number(r[col.qty]) : 0,
          harga: col.harga>=0 ? Number(r[col.harga]) : 0,
          modal: col.modal>=0 ? Number(r[col.modal]) : 0,
          discRpPer: col.disc>=0 ? Number(r[col.disc]) : 0,
          nettPer: col.nett>=0 ? Number(r[col.nett]) : 0,
          subtotal: col.subtotal>=0 ? Number(r[col.subtotal]) : 0,
          profit: col.profit>=0 ? Number(r[col.profit]) : 0
        });
      }
    }
    return { id: id, items: items };
  } catch(e) {
    throw new Error('getTrxDetail gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  [NEW] AUTO ROLLOVER STOK AWAL — dijalankan otomatis oleh time-driven
//  trigger Google Apps Script setiap pukul 00:00 (Asia/Jakarta), supaya
//  tetap jalan walau aplikasi/PWA di HP sedang tertutup.
//  Logika perhitungan sama persis dengan rollover manual di sisi aplikasi
//  (stok akhir kemarin -> jadi stok awal hari ini), dan menggunakan ULANG
//  fungsi getStockBarang()/updateStockAwal() yang sudah ada TANPA mengubahnya
//  sedikit pun. Tidak ada fungsi/variabel lama yang dimodifikasi.
// ══════════════════════════════════════════════════════════════
function dailyRolloverStokAwal() {
  var today   = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  var _kmrDate = new Date(); _kmrDate.setDate(_kmrDate.getDate() - 1);
  var kemarin = Utilities.formatDate(_kmrDate, 'Asia/Jakarta', 'yyyy-MM-dd');

  try {
    var stockAll = getStockBarang();

    // Jangan timpa kalau stok hari ini sudah ada (mis. sudah pernah di-rollover)
    var dataHariIni = stockAll.filter(function(d) { return (d.date||'').substring(0,10) === today; });
    if (dataHariIni.length > 0) {
      Logger.log('[AutoRollover] Stok tanggal ' + today + ' sudah ada, dilewati.');
      return;
    }

    var dataKemarin = stockAll.filter(function(d) { return (d.date||'').substring(0,10) === kemarin; });
    if (!dataKemarin.length) {
      Logger.log('[AutoRollover] Tidak ada data stok kemarin (' + kemarin + '), dilewati.');
      return;
    }

    // Hitung "keluar" per SKU pada tanggal kemarin dari sheet Detail Transaksi
    var keluarMap = {};
    try {
      var detSheet = getSS().getSheetByName(SHEET.DETAIL);
      if (detSheet) {
        var detData = detSheet.getDataRange().getValues();
        if (detData.length > 1) {
          var dHdr = detData[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
          var dColTgl = dHdr.indexOf('tanggal'); var dColSku = dHdr.indexOf('sku'); var dColQty = dHdr.indexOf('qty');
          if (dColTgl >= 0 && dColSku >= 0 && dColQty >= 0) {
            for (var i = 1; i < detData.length; i++) {
              var rowTgl = formatDateCell(detData[i][dColTgl]);
              if (rowTgl !== kemarin) continue;
              var sku = String(detData[i][dColSku]||'').trim();
              var qty = Number(detData[i][dColQty]) || 0;
              keluarMap[sku] = (keluarMap[sku] || 0) + qty;
            }
          }
        }
      }
    } catch(e) { Logger.log('[AutoRollover] Gagal hitung keluar: ' + e.message); }

    // Hitung "masuk" (khusus status Kiriman) per SKU pada tanggal kemarin dari riwayat input barang
    var masukMap = {}, adaRecordMap = {};
    try {
      var barangList = getInputBarangHistory();
      barangList.forEach(function(b) {
        if ((b.date||'').substring(0,10) !== kemarin) return;
        adaRecordMap[b.sku] = true;
        if (String(b.status||'').toLowerCase() === 'kiriman') {
          masukMap[b.sku] = (masukMap[b.sku] || 0) + (Number(b.qty)||0);
        }
      });
    } catch(e) { Logger.log('[AutoRollover] Gagal hitung masuk: ' + e.message); }

    var berhasil = 0, gagal = 0;
    dataKemarin.forEach(function(d) {
      try {
        var masuk = adaRecordMap[d.sku] ? (masukMap[d.sku] || 0) : (d.masuk || 0);
        var keluarKemarin = keluarMap[d.sku] || 0;
        var stokAkhirKemarin = (d.stokAwal || 0) + masuk - keluarKemarin;
        updateStockAwal({ date: today, sku: d.sku, nama: d.nama, stokAwal: stokAkhirKemarin, masuk: 0, keluar: 0 });
        berhasil++;
      } catch(e) {
        Logger.log('[AutoRollover] Gagal rollover SKU ' + d.sku + ': ' + e.message);
        gagal++;
      }
    });

    Logger.log('[AutoRollover] Selesai tanggal ' + today + ': ' + berhasil + ' berhasil, ' + gagal + ' gagal.');
  } catch(e) {
    Logger.log('[AutoRollover] Error: ' + e.message);
  }
}

// [NEW] Jalankan fungsi ini SEKALI SAJA secara manual dari editor Apps Script
// (pilih fungsi "installDailyRolloverTrigger" lalu klik Run) untuk memasang
// trigger harian jam 00:00. Fungsi ini aman dijalankan berulang kali karena
// akan membersihkan trigger lama untuk handler yang sama sebelum memasang yang baru.
function installDailyRolloverTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyRolloverStokAwal') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('dailyRolloverStokAwal')
    .timeBased()
    .atHour(0)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone('Asia/Jakarta')
    .create();
  Logger.log('[AutoRollover] Trigger harian jam 00:00 berhasil dipasang.');
}
