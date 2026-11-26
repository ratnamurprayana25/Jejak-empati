/**
 * Jejak Empati — penampung data di Google Sheets
 * ------------------------------------------------
 * Buka script.new, tempel seluruh isi file ini (atau lewat Ekstensi → Apps Script di Sheet),
 * lalu Terapkan → Deployment baru → Aplikasi web
 *   Jalankan sebagai : Saya
 *   Yang memiliki akses : Siapa saja
 * Salin URL aplikasi web (berakhiran /exec) dan masukkan ke game.
 *
 * Sheet "Check-in" : satu baris per check-in pemain
 * Sheet "Skor"     : satu baris per permainan (diperbarui saat game selesai)
 */

// ID Google Sheet tujuan (dari link docs.google.com/spreadsheets/d/<ID>/edit)
const SHEET_ID = '11FZcv8muGlwp75d-6BE36xGC0H2DhDuGpoZPstBUIVE';

const TAB_CHECKIN = 'Check-in';
const TAB_SKOR = 'Skor';

const HEAD_CHECKIN = ['Waktu', 'Nomor helm', 'Sudah mencoba', 'Perasaan', 'Tersulit',
  'Pertanyaan refleksi', 'Refleksi', 'ID permainan'];
const HEAD_SKOR = ['ID permainan', 'Pemain', 'Skor', 'XP', 'Gelar', 'Warga ditenangkan',
  'Ricuh', 'Selesai', 'Mulai', 'Terakhir diperbarui'];

/** Jalankan sekali dari editor (opsional) untuk membuat kedua tab beserta judul kolomnya. */
function setup() {
  tab_(TAB_CHECKIN, HEAD_CHECKIN);
  tab_(TAB_SKOR, HEAD_SKOR);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.type === 'checkin') saveCheckin_(d);
    else if (d.type === 'run') saveRun_(d);
    else return json_({ ok: false, error: 'tipe tidak dikenal' });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** GET ...?board=1 → papan skor (100 teratas) dalam bentuk JSON. */
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.board) {
    return json_({ ok: true, info: 'Jejak Empati aktif. Tambahkan ?board=1 untuk papan skor.' });
  }
  const sh = tab_(TAB_SKOR, HEAD_SKOR);
  const n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, board: [] });
  const rows = sh.getRange(2, 1, n, HEAD_SKOR.length).getValues();
  const board = rows
    .filter(r => validId_(String(r[0])) && String(r[1]))
    .map(r => ({
      id: String(r[0]),
      name: String(r[1]).slice(0, 20),
      score: Number(r[2]) || 0,
      xp: Number(r[3]) || 0,
      rank: String(r[4] || '').slice(0, 30),
      calmed: Number(r[5]) || 0,
      riots: Number(r[6]) || 0,
      done: r[7] === 'Ya',
      day: 1,
      ts: r[9] instanceof Date ? r[9].getTime() : Date.now()
    }))
    .sort((a, b) => b.score - a.score || b.xp - a.xp)
    .slice(0, 100);
  return json_({ ok: true, board: board });
}

/* ---------- penyimpanan ---------- */

function saveCheckin_(d) {
  const sh = tab_(TAB_CHECKIN, HEAD_CHECKIN);
  sh.appendRow([
    new Date(),
    "'" + txt_(d.helm, 10).replace(/^'/, ''), // tetap teks, supaya 07 tidak jadi 7
    txt_(d.sudah_mencoba, 80),
    txt_(d.perasaan, 40),
    txt_(d.tersulit, 80),
    txt_(d.pertanyaan_refleksi, 160),
    txt_(d.refleksi, 350),
    validId_(d.id) ? d.id : ''
  ]);
}

function saveRun_(d) {
  if (!validId_(d.id)) throw new Error('ID tidak valid');
  const score = num_(d.score, 0, 20000);
  const xp = num_(d.xp, 0, 99999);
  if (score === null || xp === null) throw new Error('Skor tidak valid');
  const sh = tab_(TAB_SKOR, HEAD_SKOR);
  const now = new Date();
  const row = [
    d.id,
    txt_(d.name, 20),
    score,
    xp,
    txt_(d.rank, 30),
    num_(d.calmed, 0, 999) || 0,
    num_(d.riots, 0, 999) || 0,
    d.done ? 'Ya' : 'Belum'
  ];
  const found = findRow_(sh, d.id);
  if (found) {
    sh.getRange(found, 1, 1, row.length).setValues([row]);
    sh.getRange(found, 10).setValue(now);
  } else {
    sh.appendRow(row.concat([now, now]));
  }
}

/* ---------- bantuan ---------- */

function tab_(name, head) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const hit = sh.getRange(2, 1, n, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function validId_(id) {
  return typeof id === 'string' && /^[a-z0-9]{6,20}$/.test(id);
}

/** Bersihkan teks dan cegah rumus spreadsheet (=, +, -, @) dari input pemain. */
function txt_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function num_(v, min, max) {
  const n = Number(v);
  if (!isFinite(n) || n < min || n > max) return null;
  return Math.round(n);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
