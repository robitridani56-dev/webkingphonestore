/**
 * ============================================================
 *  PRISMA — Backend Google Apps Script (database produk)
 * ============================================================
 *  Script ini TIDAK perlu dibuat dari dalam Spreadsheet.
 *  Spreadsheet-nya dibuat otomatis oleh script ini sendiri,
 *  lalu ID-nya disimpan supaya dipakai terus di request berikutnya.
 *
 *  Cara pakai:
 *  1. Buka https://script.google.com/ > New project.
 *  2. Hapus isi default, tempel SELURUH isi file ini.
 *  3. Di dropdown fungsi (sebelah tombol Run), pilih "setup",
 *     lalu klik "Run". Baris pertama akan minta izin akses —
 *     klik "Review permissions" > pilih akun Google kamu > Allow.
 *     Fungsi ini otomatis membuat spreadsheet, membuat sheet
 *     "Products", dan mengisi 2 baris produk contoh supaya
 *     sheet-nya langsung terisi (bukan kosong).
 *  4. Buka tab "Executions" (ikon jam di sisi kiri) untuk melihat
 *     log-nya, atau lihat panel "Execution log" di bawah editor —
 *     di situ akan tertulis ID Spreadsheet dan link untuk membukanya.
 *  5. Klik Deploy > New deployment > pilih tipe "Web app".
 *       - Execute as   : Me
 *       - Who has access: Anyone
 *  6. Klik Deploy, copy URL Web App yang muncul (diakhiri /exec).
 *  7. Tempel URL itu ke variabel GAS_URL di index.html (toko & admin).
 *
 *  CATATAN: panel admin di index.html#admin TIDAK memakai password.
 *  Siapa pun yang tahu link "...#admin" bisa menambah/mengubah/
 *  menghapus produk. Kalau nanti butuh proteksi lagi, kasih tahu saja.
 * ============================================================
 */

const SHEET_NAME = 'Products';
const SPREADSHEET_NAME = 'PRISMA - Database Produk';
const PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';

/**
 * Jalankan fungsi ini secara manual sekali dari editor Apps Script
 * (dropdown fungsi > pilih "setup" > Run). Fungsi ini akan:
 *  - membuat spreadsheet baru (kalau belum ada),
 *  - membuat sheet "Products" dengan header kolom,
 *  - mengisi baris contoh produk supaya sheet-nya langsung terisi,
 *  - mencetak ID / link spreadsheet ke Execution log.
 */
function setup() {
  try {
    const sheet = getSheet_();
    const ss = sheet.getParent();
    const url = ss.getUrl();
    const id = ss.getId();
    const jumlahProduk = sheet.getLastRow() - 1; // dikurangi baris header
    Logger.log('BERHASIL. Spreadsheet siap dipakai.');
    Logger.log('Spreadsheet ID : ' + id);
    Logger.log('Buka di        : ' + url);
    Logger.log('Jumlah produk saat ini di sheet: ' + jumlahProduk);
    return { spreadsheetId: id, spreadsheetUrl: url, jumlahProduk: jumlahProduk };
  } catch (err) {
    Logger.log('GAGAL menjalankan setup(): ' + err.message);
    Logger.log(err.stack);
    throw err;
  }
}

/**
 * Jalankan fungsi ini kalau setup() sudah pernah dijalankan tapi
 * spreadsheet-nya tidak ketemu / rusak / kehapus. Ini akan menghapus
 * ID spreadsheet yang tersimpan supaya setup() berikutnya membuat
 * spreadsheet yang benar-benar baru.
 */
function resetDatabase() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SPREADSHEET_ID);
  Logger.log('ID spreadsheet lama sudah dihapus dari Script Properties.');
  Logger.log('Jalankan setup() lagi untuk membuat spreadsheet baru.');
}

function doGet(e) {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const products = rows
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
  return jsonResponse_({ status: 'ok', products: products });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ status: 'error', message: 'Payload tidak valid' });
  }

  const sheet = getSheet_();

  switch (body.action) {
    case 'verify':
      return jsonResponse_({ status: 'ok' });

    case 'info': {
      const ss = getOrCreateSpreadsheet_();
      return jsonResponse_({ status: 'ok', spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl() });
    }

    case 'add': {
      const id = 'p_' + new Date().getTime();
      sheet.appendRow([
        id, body.brand || '', body.name || '', body.specs || '',
        body.price || '', body.old || '', body.badge || '', body.rating || '4.5', body.image || ''
      ]);
      return jsonResponse_({ status: 'ok', id: id });
    }

    case 'update': {
      const row = findRowById_(sheet, body.id);
      if (!row) return jsonResponse_({ status: 'error', message: 'Produk tidak ditemukan' });
      sheet.getRange(row, 1, 1, 9).setValues([[
        body.id, body.brand || '', body.name || '', body.specs || '',
        body.price || '', body.old || '', body.badge || '', body.rating || '4.5', body.image || ''
      ]]);
      return jsonResponse_({ status: 'ok' });
    }

    case 'delete': {
      const row = findRowById_(sheet, body.id);
      if (!row) return jsonResponse_({ status: 'error', message: 'Produk tidak ditemukan' });
      sheet.deleteRow(row);
      return jsonResponse_({ status: 'ok' });
    }

    default:
      return jsonResponse_({ status: 'error', message: 'Aksi tidak dikenali' });
  }
}

/**
 * Ambil spreadsheet yang sudah dibuat sebelumnya (lewat ID yang
 * tersimpan di Script Properties), atau buat baru kalau belum ada
 * sama sekali. ID spreadsheet yang baru dibuat otomatis disimpan
 * supaya request berikutnya memakai spreadsheet yang sama.
 */
function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(PROP_SPREADSHEET_ID);

  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      // ID tersimpan tapi spreadsheet-nya sudah tidak ada/terhapus -> buat ulang.
    }
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  Logger.log('Spreadsheet baru dibuat otomatis.');
  Logger.log('Spreadsheet ID : ' + ss.getId());
  Logger.log('Buka di        : ' + ss.getUrl());
  return ss;
}

function getSheet_() {
  const ss = getOrCreateSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // Hapus sheet default kosong "Sheet1" kalau spreadsheet baru saja dibuat.
    const defaultSheet = ss.getSheetByName('Sheet1');
    sheet = ss.insertSheet(SHEET_NAME);
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
    sheet.appendRow(['id', 'brand', 'name', 'specs', 'price', 'old', 'badge', 'rating', 'image']);
    // Contoh baris awal (boleh dihapus/diubah langsung dari admin panel)
    sheet.appendRow(['p_seed1', 'Samsung', 'Galaxy S25 Ultra', '256GB · 12GB RAM', 'Rp 21.999.000', '', 'Terlaris', '4.9', '']);
    sheet.appendRow(['p_seed2', 'Apple', 'iPhone 16 Pro Max', '256GB · Titanium', 'Rp 24.499.000', '', 'Baru', '4.9', '']);
  } else {
    // Migrasi: kalau sheet sudah ada dari sebelumnya dan belum punya kolom "image", tambahkan.
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    if (headers.indexOf('image') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('image');
    }
  }
  return sheet;
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return null;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
