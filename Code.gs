/**
 * REKAP BARANG BEKAS — backend Apps Script
 * -----------------------------------------
 * Sheet dipakai (dibuat otomatis): Entries, Users, ActivityLog, Config.
 * Foto disimpan di folder Drive "RekapBarangBekas_Foto".
 *
 * PERAN AKUN (kolom Role di sheet "Users"):
 *  - admin : akses penuh + kelola akun + log aktivitas + reset/backup
 *  - input : cuma bisa menambah data baru + lampirkan foto
 *  - guest : cuma bisa melihat data
 *
 * CARA DEPLOY:
 * 1. Buka Google Sheet baru. Extensions > Apps Script, tempel seluruh isi file ini.
 * 2. GANTI nilai SECRET di bawah dengan teks acak milikmu sendiri.
 * 3. Jalankan fungsi "setup" sekali dari editor (dropdown fungsi > setup > Run),
 *    izinkan akses saat diminta. Ini langsung membuat semua sheet yang dibutuhkan.
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 5. Copy URL /exec, pasang di aplikasi frontend.
 * 6. Buka sheet "Users", ganti password admin default (admin/admin123).
 * 7. (Opsional) Buka sheet "Config", isi kolom Value untuk baris "NotifyEmail"
 *    dengan alamat email yang ingin menerima notifikasi saat data direset.
 * 8. Setiap ubah kode ini, buat "New version" lewat Manage deployments.
 */

const SHEET_NAME = "Entries";
const USERS_SHEET_NAME = "Users";
const LOG_SHEET_NAME = "ActivityLog";
const CONFIG_SHEET_NAME = "Config";
const PHOTO_FOLDER_NAME = "RekapBarangBekas_Foto";
const HEADERS = [
  "ID", "Tanggal", "LB", "NamaPart", "KodeBarang", "Jumlah",
  "Satuan", "Mekanik", "FotoKondisiURL", "FotoPasangURL", "CreatedAt",
  "FotoBanBaruKananURL", "FotoBanBaruKiriURL",
  "FotoBanKodeBaruKananURL", "FotoBanKodeBaruKiriURL",
  "FotoBanBekasKananURL", "FotoBanBekasKiriURL",
  "FotoBanKodeBekasKananURL", "FotoBanKodeBekasKiriURL",
  // Ban belakang (jumlah=4): 4 posisi baru (kanan/kiri x luar/dalam) + kode-nya,
  // tapi foto BEKAS cuma per sisi (kanan/kiri), sementara KODE bekas per posisi
  // (kanan/kiri x luar/dalam) — total tetap 14 foto.
  "FotoBanRearBaruKananLuarURL", "FotoBanRearBaruKananDalamURL",
  "FotoBanRearBaruKiriLuarURL", "FotoBanRearBaruKiriDalamURL",
  "FotoBanRearKodeBaruKananLuarURL", "FotoBanRearKodeBaruKananDalamURL",
  "FotoBanRearKodeBaruKiriLuarURL", "FotoBanRearKodeBaruKiriDalamURL",
  "FotoBanRearBekasKananURL", "FotoBanRearBekasKiriURL",
  "FotoBanRearKodeBekasKananLuarURL", "FotoBanRearKodeBekasKananDalamURL",
  "FotoBanRearKodeBekasKiriLuarURL", "FotoBanRearKodeBekasKiriDalamURL",
];
// Peta nama "slot" foto -> nomor kolom di sheet Entries (1-based).
const SLOT_COLUMNS = {
  kondisi: 9, pasang: 10,
  banBaruKanan: 12, banBaruKiri: 13,
  banKodeBaruKanan: 14, banKodeBaruKiri: 15,
  banBekasKanan: 16, banBekasKiri: 17,
  banKodeBekasKanan: 18, banKodeBekasKiri: 19,
  banRearBaruKananLuar: 20, banRearBaruKananDalam: 21,
  banRearBaruKiriLuar: 22, banRearBaruKiriDalam: 23,
  banRearKodeBaruKananLuar: 24, banRearKodeBaruKananDalam: 25,
  banRearKodeBaruKiriLuar: 26, banRearKodeBaruKiriDalam: 27,
  banRearBekasKanan: 28, banRearBekasKiri: 29,
  banRearKodeBekasKananLuar: 30, banRearKodeBekasKananDalam: 31,
  banRearKodeBekasKiriLuar: 32, banRearKodeBekasKiriDalam: 33,
};
const USER_HEADERS = ["Username", "Password", "Role"];
const LOG_HEADERS = ["Timestamp", "Username", "Role", "Action", "Detail"];
const CONFIG_HEADERS = ["Key", "Value"];
const ROLES = ["admin", "input", "guest"];

// GANTI dengan teks acak milikmu sendiri sebelum deploy!
const SECRET = "GANTI-DENGAN-TEKS-RAHASIA-ACAK-MILIKMU";
const TOKEN_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 hari
const LOG_MAX_ROWS = 500; // baris log lama otomatis dipangkas melebihi ini

// ---------------- sheets ----------------

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Selalu samakan baris header dengan skema HEADERS saat ini (bukan cuma
    // menambah kolom kalau kurang). Data dibaca berdasarkan TEKS header, jadi
    // kalau nama kolom pernah diganti (mis. saat struktur ban belakang disusun
    // ulang) tapi baris header lama tidak ikut diperbarui, hasilnya data yang
    // sudah ada di kolom itu jadi seperti "hilang" saat dibaca kembali.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sheet.getRange("B2:B").setNumberFormat("@");
  return sheet;
}

function getUsersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(USERS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(USER_HEADERS);
    sheet.setFrozenRows(1);
    sheet.appendRow(["admin", "admin123", "admin"]); // SEGERA GANTI
  }
  return sheet;
}

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(LOG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LOG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONFIG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.appendRow(["NotifyEmail", ""]);
  }
  return sheet;
}

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

/**
 * Jalankan SEKALI secara manual dari editor Apps Script untuk langsung
 * membuat semua sheet yang dibutuhkan tanpa perlu buka situsnya dulu.
 */
function setup() {
  getSheet_();
  getUsersSheet_();
  getLogSheet_();
  getConfigSheet_();
  Logger.log("Selesai. Cek sheet Entries, Users, ActivityLog, Config.");
}

function sheetToObjects_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;
    const obj = {};
    headers.forEach((h, idx) => (obj[h] = row[idx]));
    rows.push({
      id: String(obj.ID),
      tanggal: formatDate_(obj.Tanggal),
      lb: String(obj.LB || ""),
      namaPart: String(obj.NamaPart || ""),
      kodeBarang: String(obj.KodeBarang || ""),
      jumlah: String(obj.Jumlah || ""),
      satuan: String(obj.Satuan || ""),
      mekanik: String(obj.Mekanik || ""),
      fotoKondisi: obj.FotoKondisiURL || null,
      fotoPasang: obj.FotoPasangURL || null,
      fotoBan: {
        baruKanan: obj.FotoBanBaruKananURL || null,
        baruKiri: obj.FotoBanBaruKiriURL || null,
        kodeBaruKanan: obj.FotoBanKodeBaruKananURL || null,
        kodeBaruKiri: obj.FotoBanKodeBaruKiriURL || null,
        bekasKanan: obj.FotoBanBekasKananURL || null,
        bekasKiri: obj.FotoBanBekasKiriURL || null,
        kodeBekasKanan: obj.FotoBanKodeBekasKananURL || null,
        kodeBekasKiri: obj.FotoBanKodeBekasKiriURL || null,
      },
      fotoBanRear: {
        baruKananLuar: obj.FotoBanRearBaruKananLuarURL || null,
        baruKananDalam: obj.FotoBanRearBaruKananDalamURL || null,
        baruKiriLuar: obj.FotoBanRearBaruKiriLuarURL || null,
        baruKiriDalam: obj.FotoBanRearBaruKiriDalamURL || null,
        kodeBaruKananLuar: obj.FotoBanRearKodeBaruKananLuarURL || null,
        kodeBaruKananDalam: obj.FotoBanRearKodeBaruKananDalamURL || null,
        kodeBaruKiriLuar: obj.FotoBanRearKodeBaruKiriLuarURL || null,
        kodeBaruKiriDalam: obj.FotoBanRearKodeBaruKiriDalamURL || null,
        bekasKanan: obj.FotoBanRearBekasKananURL || null,
        bekasKiri: obj.FotoBanRearBekasKiriURL || null,
        kodeBekasKananLuar: obj.FotoBanRearKodeBekasKananLuarURL || null,
        kodeBekasKananDalam: obj.FotoBanRearKodeBekasKananDalamURL || null,
        kodeBekasKiriLuar: obj.FotoBanRearKodeBekasKiriLuarURL || null,
        kodeBekasKiriDalam: obj.FotoBanRearKodeBekasKiriDalamURL || null,
      },
      rowIndex: i + 1,
    });
  }
  return rows;
}

function formatDate_(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd");
  }
  return String(val);
}

function findRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function findUserRow_(sheet, username) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) return i + 1;
  }
  return -1;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ---------------- activity log ----------------

function logActivity_(username, role, action, detail) {
  try {
    const sheet = getLogSheet_();
    sheet.appendRow([new Date(), username || "", role || "", action || "", detail || ""]);
    const last = sheet.getLastRow();
    if (last > LOG_MAX_ROWS + 1) {
      sheet.deleteRows(2, last - LOG_MAX_ROWS - 1); // buang baris terlama, sisakan LOG_MAX_ROWS
    }
  } catch (err) {
    // jangan sampai kegagalan logging menggagalkan aksi utama
  }
}

function listActivity_() {
  const sheet = getLogSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const [ts, username, role, action, detail] = values[i];
    rows.push({
      timestamp: ts instanceof Date ? ts.toISOString() : String(ts),
      username: String(username || ""),
      role: String(role || ""),
      action: String(action || ""),
      detail: String(detail || ""),
    });
  }
  return rows.reverse(); // terbaru duluan
}

// ---------------- config ----------------

function getConfig_(key) {
  const sheet = getConfigSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) return String(values[i][1] || "");
  }
  return "";
}

function setConfig_(key, value) {
  const sheet = getConfigSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function sendResetNotification_(byUsername, backupSheet) {
  try {
    const email = getConfig_("NotifyEmail");
    if (!email) return;
    MailApp.sendEmail({
      to: email,
      subject: "Rekap Barang Bekas — Data direset",
      body:
        "Reset data dilakukan oleh: " + byUsername + "\n" +
        "Waktu: " + new Date().toString() + "\n" +
        "Cadangan otomatis tersimpan di sheet: " + backupSheet + "\n\n" +
        "Kalau ini bukan kamu yang melakukan, segera cek sheet Users dan ganti semua password.",
    });
  } catch (err) {
    // kuota MailApp habis atau error lain — jangan gagalkan proses reset
  }
}

// ---------------- auth ----------------

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
}

function signToken_(username, role) {
  const payload = JSON.stringify({ u: username, r: role, t: Date.now() });
  const payloadB64 = b64url_(Utilities.newBlob(payload).getBytes());
  const sig = Utilities.computeHmacSha256Signature(payloadB64, SECRET);
  const sigB64 = b64url_(sig);
  return payloadB64 + "." + sigB64;
}

function verifyToken_(token) {
  try {
    if (!token || token.indexOf(".") === -1) return null;
    const parts = token.split(".");
    const payloadB64 = parts[0];
    const sigB64 = parts[1];
    const expectedSig = b64url_(Utilities.computeHmacSha256Signature(payloadB64, SECRET));
    if (expectedSig !== sigB64) return null;
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString());
    if (!payload.t || Date.now() - payload.t > TOKEN_MAX_AGE_MS) return null;
    return { username: payload.u, role: payload.r };
  } catch (err) {
    return null;
  }
}

function login_(username, password) {
  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [u, p, role] = values[i];
    if (String(u).trim() === String(username).trim() && String(p) === String(password)) {
      const cleanUser = String(u).trim();
      const cleanRole = String(role).trim().toLowerCase();
      logActivity_(cleanUser, cleanRole, "login", "Login berhasil");
      return { ok: true, username: cleanUser, role: cleanRole, token: signToken_(cleanUser, cleanRole) };
    }
  }
  return { ok: false, error: "Username atau password salah." };
}

function verifyPassword_(username, password) {
  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [u, p] = values[i];
    if (String(u).trim() === String(username).trim() && String(p) === String(password)) return true;
  }
  return false;
}

function backupEntriesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyyMMdd_HHmmss");
  const name = ("Backup_" + stamp).slice(0, 90);
  const copy = sheet.copyTo(ss);
  copy.setName(name);
  return name;
}

function requireRole_(token, allowedRoles) {
  const auth = verifyToken_(token);
  if (!auth) throw new Error("Sesi tidak valid atau sudah habis, silakan login ulang.");
  if (allowedRoles.indexOf(auth.role) === -1) {
    throw new Error("Akun kamu (" + auth.role + ") tidak punya izin untuk aksi ini.");
  }
  return auth;
}

// ---------------- user management (admin) ----------------

function listUsers_() {
  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    rows.push({ username: String(values[i][0]), role: String(values[i][2]).toLowerCase() });
  }
  return rows;
}

function countAdmins_(sheet) {
  const values = sheet.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2]).trim().toLowerCase() === "admin") n++;
  }
  return n;
}

// ---- GET: baca data rekap (dipakai semua peran yang sudah login) ----
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "list";
  try {
    if (action === "list") {
      return jsonOut_({ ok: true, entries: sheetToObjects_() });
    }
    return jsonOut_({ ok: false, error: "Aksi tidak dikenal: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---- POST: login serta semua aksi tulis/baca yang butuh otorisasi ----
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = body.action;

    if (action === "login") {
      return jsonOut_(login_(body.username || "", body.password || ""));
    }

    const sheet = getSheet_();

    if (action === "add") {
      const auth = requireRole_(body.token, ["admin", "input"]);
      const id = Utilities.getUuid();
      const now = new Date();
      sheet.appendRow([
        id, body.tanggal || "", body.lb || "", body.namaPart || "", body.kodeBarang || "",
        body.jumlah || "", body.satuan || "", body.mekanik || "", "", "", now,
      ]);
      logActivity_(auth.username, auth.role, "add", "Tambah: " + (body.namaPart || "-") + " (LB " + (body.lb || "-") + ")");
      return jsonOut_({ ok: true, id: id });
    }

    // Impor banyak baris sekaligus dalam satu request (dipakai fitur Impor CSV/XLSX)
    // supaya tidak perlu ratusan request terpisah yang rawan putus di tengah jalan.
    if (action === "addBatch") {
      const auth = requireRole_(body.token, ["admin", "input"]);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      if (!rows.length) return jsonOut_({ ok: false, error: "Tidak ada baris untuk diimpor." });
      const now = new Date();
      const ids = [];
      const matrix = rows.map((r) => {
        const id = Utilities.getUuid();
        ids.push(id);
        return [
          id, r.tanggal || "", r.lb || "", r.namaPart || "", r.kodeBarang || "",
          r.jumlah || "", r.satuan || "", r.mekanik || "", "", "", now,
        ];
      });
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
      logActivity_(auth.username, auth.role, "addBatch", "Impor " + matrix.length + " baris sekaligus");
      return jsonOut_({ ok: true, ids: ids, count: matrix.length });
    }

    if (action === "update") {
      const auth = requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const fieldToCol = { tanggal: 2, lb: 3, namaPart: 4, kodeBarang: 5, jumlah: 6, satuan: 7, mekanik: 8 };
      const changed = [];
      Object.keys(body.fields || {}).forEach((key) => {
        const col = fieldToCol[key];
        if (col) {
          sheet.getRange(row, col).setValue(body.fields[key]);
          changed.push(key);
        }
      });
      logActivity_(auth.username, auth.role, "update", "Edit id=" + body.id + " field=" + changed.join(","));
      return jsonOut_({ ok: true });
    }

    if (action === "delete") {
      const auth = requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const namaPart = sheet.getRange(row, 4).getValue();
      sheet.deleteRow(row);
      logActivity_(auth.username, auth.role, "delete", "Hapus: " + namaPart + " (id=" + body.id + ")");
      return jsonOut_({ ok: true });
    }

    if (action === "uploadPhoto") {
      const auth = requireRole_(body.token, ["admin", "input"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const col = SLOT_COLUMNS[body.slot];
      if (!col) return jsonOut_({ ok: false, error: "Slot foto tidak dikenal: " + body.slot });

      const matches = String(body.dataUrl || "").match(/^data:(image\/[a-zA-Z]+);base64,(.+)$/);
      if (!matches) return jsonOut_({ ok: false, error: "Format foto tidak valid" });
      const mime = matches[1];
      const base64 = matches[2];
      const bytes = Utilities.base64Decode(base64);
      const ext = mime.split("/")[1] || "jpg";
      const fileName = `${body.id}_${body.slot}.${ext}`;
      const blob = Utilities.newBlob(bytes, mime, fileName);

      const folder = getPhotoFolder_();
      const existing = folder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      const publicUrl = `https://lh3.googleusercontent.com/d/${fileId}`;

      sheet.getRange(row, col).setValue(publicUrl);
      logActivity_(auth.username, auth.role, "uploadPhoto", "Unggah foto " + body.slot + " id=" + body.id);
      return jsonOut_({ ok: true, url: publicUrl });
    }

    if (action === "removePhoto") {
      const auth = requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const col = SLOT_COLUMNS[body.slot];
      if (!col) return jsonOut_({ ok: false, error: "Slot foto tidak dikenal: " + body.slot });
      sheet.getRange(row, col).setValue("");
      logActivity_(auth.username, auth.role, "removePhoto", "Hapus foto " + body.slot + " id=" + body.id);
      return jsonOut_({ ok: true });
    }

    if (action === "backup") {
      const auth = requireRole_(body.token, ["admin"]);
      const backupSheet = backupEntriesSheet_();
      logActivity_(auth.username, auth.role, "backup", "Backup manual: " + backupSheet);
      return jsonOut_({ ok: true, backupSheet: backupSheet });
    }

    if (action === "resetAll") {
      const auth = requireRole_(body.token, ["admin"]);
      if (!verifyPassword_(auth.username, body.password || "")) {
        return jsonOut_({ ok: false, error: "Password salah, reset dibatalkan." });
      }
      const backupSheet = backupEntriesSheet_();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
      }
      logActivity_(auth.username, auth.role, "resetAll", "RESET SEMUA DATA. Backup: " + backupSheet);
      sendResetNotification_(auth.username, backupSheet);
      return jsonOut_({ ok: true, backupSheet: backupSheet });
    }

    // ---- log aktivitas (admin) ----
    if (action === "listActivity") {
      requireRole_(body.token, ["admin"]);
      return jsonOut_({ ok: true, logs: listActivity_() });
    }

    // ---- kelola akun (admin) ----
    if (action === "listUsers") {
      requireRole_(body.token, ["admin"]);
      return jsonOut_({ ok: true, users: listUsers_() });
    }

    if (action === "addUser") {
      const auth = requireRole_(body.token, ["admin"]);
      const username = String(body.username || "").trim();
      const password = String(body.password || "");
      const role = String(body.role || "").trim().toLowerCase();
      if (!username || !password) return jsonOut_({ ok: false, error: "Username dan password wajib diisi." });
      if (ROLES.indexOf(role) === -1) return jsonOut_({ ok: false, error: "Role tidak valid." });
      const usersSheet = getUsersSheet_();
      if (findUserRow_(usersSheet, username) !== -1) return jsonOut_({ ok: false, error: "Username sudah dipakai." });
      usersSheet.appendRow([username, password, role]);
      logActivity_(auth.username, auth.role, "addUser", "Tambah akun: " + username + " (" + role + ")");
      return jsonOut_({ ok: true });
    }

    if (action === "updateUser") {
      const auth = requireRole_(body.token, ["admin"]);
      const username = String(body.username || "").trim();
      const usersSheet = getUsersSheet_();
      const row = findUserRow_(usersSheet, username);
      if (row === -1) return jsonOut_({ ok: false, error: "Akun tidak ditemukan." });
      const fields = body.fields || {};
      if (fields.role) {
        const role = String(fields.role).trim().toLowerCase();
        if (ROLES.indexOf(role) === -1) return jsonOut_({ ok: false, error: "Role tidak valid." });
        if (role !== "admin" && String(usersSheet.getRange(row, 3).getValue()).toLowerCase() === "admin" && countAdmins_(usersSheet) <= 1) {
          return jsonOut_({ ok: false, error: "Tidak bisa mengubah role admin terakhir." });
        }
        usersSheet.getRange(row, 3).setValue(role);
      }
      if (fields.password) {
        usersSheet.getRange(row, 2).setValue(String(fields.password));
      }
      logActivity_(auth.username, auth.role, "updateUser", "Update akun: " + username + " (" + Object.keys(fields).join(",") + ")");
      return jsonOut_({ ok: true });
    }

    if (action === "deleteUser") {
      const auth = requireRole_(body.token, ["admin"]);
      const username = String(body.username || "").trim();
      if (username.toLowerCase() === auth.username.toLowerCase()) {
        return jsonOut_({ ok: false, error: "Tidak bisa menghapus akun sendiri saat sedang login." });
      }
      const usersSheet = getUsersSheet_();
      const row = findUserRow_(usersSheet, username);
      if (row === -1) return jsonOut_({ ok: false, error: "Akun tidak ditemukan." });
      if (String(usersSheet.getRange(row, 3).getValue()).toLowerCase() === "admin" && countAdmins_(usersSheet) <= 1) {
        return jsonOut_({ ok: false, error: "Tidak bisa menghapus admin terakhir." });
      }
      usersSheet.deleteRow(row);
      logActivity_(auth.username, auth.role, "deleteUser", "Hapus akun: " + username);
      return jsonOut_({ ok: true });
    }

    // ---- config (admin) ----
    if (action === "getConfig") {
      requireRole_(body.token, ["admin"]);
      return jsonOut_({ ok: true, config: { notifyEmail: getConfig_("NotifyEmail") } });
    }

    if (action === "setConfig") {
      const auth = requireRole_(body.token, ["admin"]);
      if (typeof body.notifyEmail === "string") setConfig_("NotifyEmail", body.notifyEmail.trim());
      logActivity_(auth.username, auth.role, "setConfig", "Update pengaturan notifikasi email");
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: "Aksi tidak dikenal: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
