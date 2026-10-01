/****************************************************
 * RSVP & UCAPAN - Nia & Bima  (Google Apps Script)
 * CARA PAKAI:
 * 1. Buat Google Spreadsheet baru (mis. "RSVP Nia Bima").
 * 2. Buka Extensions > Apps Script, hapus isi editor, paste SELURUH file ini, Save.
 * 3. Deploy > New deployment > tipe "Web app":
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - Deploy, lalu izinkan akses (Continue > pilih akun > Advanced > Go to... > Allow).
 * 4. Salin URL Web App (.../exec), tempel ke RSVP_API_URL di index.html.
 * Kolom sheet otomatis dibuat: Timestamp | Nama | Kehadiran | Jumlah Tamu | Ucapan
 ****************************************************/

function _sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'Jumlah Tamu', 'Ucapan']);
  }
  return sh;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET ?action=list -> daftar ucapan (terlama dulu; client membalik jadi terbaru dulu)
function doGet(e) {
  var sh = _sheet();
  if (sh.getLastRow() < 2) {
    return _json({ status: 'ok', count: 0, data: [] });
  }
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var data = [];
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][1] && !rows[i][4]) continue; // lewati baris kosong
    data.push({
      ts: rows[i][0] instanceof Date ? rows[i][0].toISOString() : String(rows[i][0] || ''),
      nama: String(rows[i][1] || ''),
      kehadiran: String(rows[i][2] || ''),
      tamu: String(rows[i][3] || ''),
      ucapan: String(rows[i][4] || '')
    });
  }
  return _json({ status: 'ok', count: data.length, data: data });
}

// POST {nama, kehadiran, tamu, ucapan} -> simpan baris baru
function doPost(e) {
  var d = {};
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return _json({ status: 'error', message: 'Data tidak valid' });
  }
  var nama = String(d.nama || '').trim().slice(0, 80);
  var kehadiran = String(d.kehadiran || '').trim().slice(0, 20);
  var tamu = String(d.tamu || '').trim().slice(0, 20);
  var ucapan = String(d.ucapan || '').trim().slice(0, 1000);

  if (!nama) {
    return _json({ status: 'error', message: 'Nama wajib diisi.' });
  }
  if (ucapan.length < 2) {
    return _json({ status: 'error', message: 'Ucapan minimal 2 karakter.' });
  }

  _sheet().appendRow([new Date(), nama, kehadiran, tamu, ucapan]);
  return _json({ status: 'ok' });
}
