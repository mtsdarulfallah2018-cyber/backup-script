/**
 * e-Sarpras Madrasah — Backend Multi-Role & Password + Barcode Generator
 * MTs Darul Fallah
 */

const MADRASAH_NAME = 'MTs Darul Fallah';
const SPREADSHEET_ID = '1Uu0du42TXQUEyXFOE-UU0t9k5NTmA03Vq77PhNOAWzo';

const SHEET_NAMES = {
  ASSETS: 'Aset',
  MAINTENANCE: 'Usulan',
  PROCUREMENT: 'Pengadaan',
  BOOKINGS: 'Peminjaman',
  USERS: 'Pengguna',
};

const HEADERS = {
  Aset: ['ID', 'Nama', 'Lokasi', 'Kategori', 'Jumlah', 'Tahun', 'Kondisi', 'Barcode'],
  Usulan: ['ID', 'Aset', 'Keluhan', 'DiajukanOleh', 'Tanggal', 'Prioritas', 'Status', 'Catatan'],
  Pengadaan: ['ID', 'Barang', 'Alasan', 'Biaya', 'DiajukanOleh', 'Status', 'Catatan'],
  Peminjaman: ['ID', 'Fasilitas', 'Tanggal', 'Waktu', 'Keperluan', 'Pemohon', 'Status', 'Catatan'],
  Pengguna: ['Email', 'Nama', 'Peran', 'Password'],
};

function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.keys(HEADERS).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[name]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    } else {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    }
  });

  seedIfEmpty(ss);
  return ss.getUrl();
}

function seedIfEmpty(ss) {
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() <= 1) {
    const defaultUsers = [
      ['admin@madrasah.sch.id', 'Admin Sarpras', 'admin', 'admin123'],
      ['kepala@madrasah.sch.id', 'Kepala Madrasah', 'kepala', 'kepala123'],
      ['guru@madrasah.sch.id', 'Ust. Fauzan (Guru)', 'guru', 'guru123']
    ];
    usersSheet.getRange(2, 1, defaultUsers.length, defaultUsers[0].length).setValues(defaultUsers);
  }

  const assets = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (assets.getLastRow() <= 1) {
    const rows = [
      ['A1', 'Meja Siswa', 'Ruang Kelas VII A–IX C', 'Furnitur', 180, 2022, 'Baik', 'A1'],
      ['A2', 'Unit Komputer', 'Lab Komputer', 'Elektronik', 20, 2021, 'Baik', 'A2'],
      ['A3', 'Proyektor', 'Ruang Guru', 'Elektronik', 3, 2018, 'Rusak Berat', 'A3']
    ];
    assets.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name) {
  let ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    setup();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i] !== undefined ? String(row[i]).trim() : ''));
    return obj;
  });
}

function safeGetSheetData(sheetName) {
  try {
    const res = sheetToObjects(getSheet(sheetName));
    return Array.isArray(res) ? res : [];
  } catch (e) {
    return [];
  }
}

// ----------------- AUTENTIKASI -----------------

function loginUser(email, password) {
  setup();
  const users = safeGetSheetData(SHEET_NAMES.USERS);
  const found = users.find(u => 
    String(u.Email).toLowerCase().trim() === String(email).toLowerCase().trim() && 
    String(u.Password).trim() === String(password).trim()
  );

  if (!found) {
    throw new Error('Email atau Password salah!');
  }

  return {
    email: found.Email,
    name: found.Nama,
    role: found.Peran
  };
}

function getBootstrapData(userSession) {
  setup();
  return {
    madrasah: MADRASAH_NAME,
    user: userSession,
    assets: safeGetSheetData(SHEET_NAMES.ASSETS),
    maintenance: safeGetSheetData(SHEET_NAMES.MAINTENANCE),
    procurement: safeGetSheetData(SHEET_NAMES.PROCUREMENT),
    bookings: safeGetSheetData(SHEET_NAMES.BOOKINGS),
    users: safeGetSheetData(SHEET_NAMES.USERS).map(u => ({ Email: u.Email, Nama: u.Nama, Peran: u.Peran }))
  };
}

// ----------------- INVENTARIS / ASET (DENGAN BARCODE) -----------------

function addAsset(form) {
  const sheet = getSheet(SHEET_NAMES.ASSETS);
  const nextNum = sheet.getLastRow();
  const id = 'A' + nextNum;
  const barcodeValue = id; // Kode Barcode menggunakan ID Aset unik secara otomatis

  sheet.appendRow([id, form.name, form.location, form.category, form.qty, form.year, form.condition, barcodeValue]);
  return safeGetSheetData(SHEET_NAMES.ASSETS);
}

// ----------------- USULAN PERBAIKAN -----------------

function addMaintenance(form) {
  const sheet = getSheet(SHEET_NAMES.MAINTENANCE);
  const id = 'U' + (sheet.getLastRow());
  sheet.appendRow([id, form.asset, form.issue, '-', new Date().toLocaleDateString('id-ID'), form.priority, 'Diajukan', '']);
  return safeGetSheetData(SHEET_NAMES.MAINTENANCE);
}

function advanceMaintenanceStatus(id) {
  const sheet = getSheet(SHEET_NAMES.MAINTENANCE);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const currentStatus = values[i][6];
      const newStatus = currentStatus === 'Diajukan' ? 'Diproses' : 'Selesai';
      sheet.getRange(i + 1, 7).setValue(newStatus);
      break;
    }
  }
  return safeGetSheetData(SHEET_NAMES.MAINTENANCE);
}

function reviewMaintenance(id, status, note) {
  const sheet = getSheet(SHEET_NAMES.MAINTENANCE);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 7).setValue(status);
      sheet.getRange(i + 1, 8).setValue(note || '-');
      break;
    }
  }
  return safeGetSheetData(SHEET_NAMES.MAINTENANCE);
}

// ----------------- PENGADAAN -----------------

function addProcurement(form) {
  const sheet = getSheet(SHEET_NAMES.PROCUREMENT);
  const id = 'P' + (sheet.getLastRow());
  sheet.appendRow([id, form.item, form.reason, form.cost, '-', 'Diajukan', '']);
  return safeGetSheetData(SHEET_NAMES.PROCUREMENT);
}

function decideProcurement(id, decision, note) {
  const sheet = getSheet(SHEET_NAMES.PROCUREMENT);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(decision);
      sheet.getRange(i + 1, 7).setValue(note || '-');
      break;
    }
  }
  return safeGetSheetData(SHEET_NAMES.PROCUREMENT);
}

function markProcurementPurchased(id) {
  return decideProcurement(id, 'Dibeli', '');
}

// ----------------- PEMINJAMAN -----------------

function addBooking(form) {
  const sheet = getSheet(SHEET_NAMES.BOOKINGS);
  const id = 'B' + (sheet.getLastRow());
  sheet.appendRow([id, form.facility, form.date, form.time, form.purpose, '-', 'Diajukan', '']);
  return safeGetSheetData(SHEET_NAMES.BOOKINGS);
}

function decideBooking(id, decision, note) {
  const sheet = getSheet(SHEET_NAMES.BOOKINGS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 7).setValue(decision);
      sheet.getRange(i + 1, 8).setValue(note || '-');
      break;
    }
  }
  return safeGetSheetData(SHEET_NAMES.BOOKINGS);
}

// ----------------- MANAJEMEN PENGGUNA -----------------

function addUser(form) {
  const sheet = getSheet(SHEET_NAMES.USERS);
  sheet.appendRow([form.email, form.name, form.role, form.password]);
  return safeGetSheetData(SHEET_NAMES.USERS).map(u => ({ Email: u.Email, Nama: u.Nama, Peran: u.Peran }));
}

function updateUserRole(email, role) {
  const sheet = getSheet(SHEET_NAMES.USERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(email).toLowerCase()) {
      sheet.getRange(i + 1, 3).setValue(role);
      break;
    }
  }
  return safeGetSheetData(SHEET_NAMES.USERS).map(u => ({ Email: u.Email, Nama: u.Nama, Peran: u.Peran }));
}

// ----------------- WEB APP -----------------

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('e-Sarpras — ' + MADRASAH_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}