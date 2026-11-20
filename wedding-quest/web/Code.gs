/**
 * Wedding Quest — server Google Apps Script.
 * Halaman web (Index.html) disajikan oleh script ini, dan semua data
 * disimpan di sheet "wq_data" pada Google Sheet tempat script ini dibuat.
 *
 * 1) Ganti PIN di bawah (boleh angka atau kata). Kosongkan ('') kalau tidak mau pakai PIN.
 * 2) Jalankan fungsi setup sekali dari editor untuk memberi izin.
 * 3) Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 */
const PIN = '1234';

const COLLECTIONS = ['config', 'prog', 'custom', 'budget', 'guests', 'vendors'];
const SHEET_NAME = 'wq_data';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Wedding Quest')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function setup() {
  sheet_();
  return 'Siap. Sheet "' + SHEET_NAME + '" sudah dibuat.';
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A:B').setNumberFormat('@'); // keep ids as text
    sh.appendRow(['collection', 'id', 'json', 'updated']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function check_(pin) {
  if (PIN && String(pin) !== String(PIN)) throw new Error('PIN');
}

function rev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('rev') || 0);
}

/** All documents, or {same: true} when nothing changed since `since`. */
function getAll(pin, since) {
  check_(pin);
  const rev = rev_();
  if (since && Number(since) === rev) return { rev: rev, same: true };
  const values = sheet_().getDataRange().getValues();
  const docs = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) docs.push([String(values[i][0]), String(values[i][1]), String(values[i][2])]);
  }
  return { rev: rev, docs: docs };
}

/** Create, replace, or (json === '') delete one document. */
function put(pin, collection, id, json) {
  check_(pin);
  collection = String(collection);
  id = String(id);
  if (COLLECTIONS.indexOf(collection) < 0) throw new Error('Koleksi tidak dikenal');
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) throw new Error('ID tidak valid');
  if (json && String(json).length > 45000) throw new Error('Data terlalu besar');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    const values = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === collection && String(values[i][1]) === id) { row = i + 1; break; }
    }
    if (!json) {
      if (row > 0) sh.deleteRow(row);
    } else if (row > 0) {
      sh.getRange(row, 3, 1, 2).setValues([[String(json), new Date()]]);
    } else {
      sh.appendRow([collection, id, String(json), new Date()]);
    }
    const rev = rev_() + 1;
    PropertiesService.getScriptProperties().setProperty('rev', String(rev));
    return rev;
  } finally {
    lock.releaseLock();
  }
}
