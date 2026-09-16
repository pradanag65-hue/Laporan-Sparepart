/**
 * REKAP BARANG BEKAS — backend Apps Script
 * -----------------------------------------
 * Database: sheet "Entries" + sheet "Users" (akun login) di spreadsheet
 *           yang sama dengan script ini (bikin script ini lewat
 *           Extensions > Apps Script di dalam Google Sheet-nya).
 * Foto    : disimpan sebagai file di folder Drive "RekapBarangBekas_Foto",
 *           URL publik-nya disimpan di kolom FotoKondisiURL / FotoPasangURL.
 *
 * PERAN AKUN (kolom Role di sheet "Users"):
 *  - admin : akses penuh (lihat, tambah, edit, hapus, reset)
 *  - input : cuma bisa menambah data baru + lampirkan foto miliknya sendiri,
 *            tidak bisa mengedit atau menghapus data yang sudah tersimpan
 *  - guest : cuma bisa melihat data, tidak bisa menambah/mengedit/menghapus
 *
 * CARA DEPLOY:
 * 1. Buka Google Sheet baru, kasih nama mis. "Rekap Barang Bekas - DB".
 * 2. Extensions > Apps Script, hapus isi default, tempel seluruh isi file ini.
 * 3. GANTI nilai SECRET di bawah dengan teks acak milikmu sendiri (jangan
 *    dibiarkan nilai default) — ini dipakai untuk menandatangani sesi login.
 * 4. Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Klik Deploy, izinkan akses (authorize) saat diminta.
 * 6. Copy URL Web App yang muncul (diakhiri /exec) — itu yang dipakai
 *    di pengaturan aplikasi frontend.
 * 7. Buka sheet "Users" yang otomatis terbuat, akan ada 1 baris akun admin
 *    default: username "admin", password "admin123", role "admin".
 *    SEGERA GANTI password itu, dan tambah baris akun lain (role: admin /
 *    input / guest) sesuai kebutuhan tim kamu.
 * 8. Setiap kali code ini diubah, buat "New deployment" lagi (atau Manage
 *    deployments > edit > New version) supaya perubahan aktif.
 */

const SHEET_NAME = "Entries";
const USERS_SHEET_NAME = "Users";
const PHOTO_FOLDER_NAME = "RekapBarangBekas_Foto";
const HEADERS = [
  "ID", "Tanggal", "LB", "NamaPart", "KodeBarang", "Jumlah",
  "Satuan", "Mekanik", "FotoKondisiURL", "FotoPasangURL", "CreatedAt",
];
const USER_HEADERS = ["Username", "Password", "Role"];

// GANTI dengan teks acak milikmu sendiri sebelum deploy!
const SECRET = "GANTI-DENGAN-TEKS-RAHASIA-ACAK-MILIKMU";
const TOKEN_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 hari

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  // Paksa kolom Tanggal (B) selalu teks biasa, supaya Sheets tidak pernah
  // otomatis mengubahnya jadi objek Date (yang bisa bergeser tanggal karena zona waktu).
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
    // Akun admin default supaya bisa login pertama kali — SEGERA GANTI.
    sheet.appendRow(["admin", "admin123", "admin"]);
  }
  return sheet;
}

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

/**
 * Jalankan fungsi ini SEKALI secara manual dari editor Apps Script
 * (pilih "setup" di dropdown fungsi lalu klik Run) untuk langsung
 * membuat sheet "Entries" dan "Users" tanpa perlu buka situsnya dulu.
 * Google akan minta izin akses saat pertama kali dijalankan — klik Allow.
 */
function setup() {
  getSheet_();
  getUsersSheet_();
  Logger.log("Selesai. Cek sheet 'Entries' dan 'Users' di spreadsheet ini.");
}

function sheetToObjects_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // skip baris kosong
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
    if (String(values[i][0]) === String(id)) return i + 1; // 1-based row number
  }
  return -1;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
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

// Mengembalikan {username, role} kalau valid, atau null kalau tidak.
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
      return { ok: true, username: String(u).trim(), role: String(role).trim().toLowerCase(), token: signToken_(String(u).trim(), String(role).trim().toLowerCase()) };
    }
  }
  return { ok: false, error: "Username atau password salah." };
}

// Melempar error kalau token tidak valid / rolenya tidak diizinkan.
function requireRole_(token, allowedRoles) {
  const auth = verifyToken_(token);
  if (!auth) throw new Error("Sesi tidak valid atau sudah habis, silakan login ulang.");
  if (allowedRoles.indexOf(auth.role) === -1) {
    throw new Error("Akun kamu (" + auth.role + ") tidak punya izin untuk aksi ini.");
  }
  return auth;
}

// ---- GET: baca data (semua peran yang sudah login boleh baca) ----
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "list";
  try {
    if (action === "list") {
      return jsonOut_({ ok: true, entries: sheetToObjects_() });
    }
    return jsonOut_({ ok: false, error: "Aksi tidak dikenal: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---- POST: login serta tulis data (add / update / delete / uploadPhoto) ----
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action;

    if (action === "login") {
      return jsonOut_(login_(body.username || "", body.password || ""));
    }

    const sheet = getSheet_();

    if (action === "add") {
      requireRole_(body.token, ["admin", "input"]);
      const id = Utilities.getUuid();
      const now = new Date();
      sheet.appendRow([
        id,
        body.tanggal || "",
        body.lb || "",
        body.namaPart || "",
        body.kodeBarang || "",
        body.jumlah || "",
        body.satuan || "",
        body.mekanik || "",
        "",
        "",
        now,
      ]);
      return jsonOut_({ ok: true, id: id });
    }

    if (action === "update") {
      requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const fieldToCol = {
        tanggal: 2, lb: 3, namaPart: 4, kodeBarang: 5,
        jumlah: 6, satuan: 7, mekanik: 8,
      };
      Object.keys(body.fields || {}).forEach((key) => {
        const col = fieldToCol[key];
        if (col) sheet.getRange(row, col).setValue(body.fields[key]);
      });
      return jsonOut_({ ok: true });
    }

    if (action === "delete") {
      requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      sheet.deleteRow(row);
      return jsonOut_({ ok: true });
    }

    if (action === "uploadPhoto") {
      requireRole_(body.token, ["admin", "input"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });

      const matches = String(body.dataUrl || "").match(/^data:(image\/[a-zA-Z]+);base64,(.+)$/);
      if (!matches) return jsonOut_({ ok: false, error: "Format foto tidak valid" });
      const mime = matches[1];
      const base64 = matches[2];
      const bytes = Utilities.base64Decode(base64);
      const ext = mime.split("/")[1] || "jpg";
      const fileName = `${body.id}_${body.slot}.${ext}`;
      const blob = Utilities.newBlob(bytes, mime, fileName);

      const folder = getPhotoFolder_();
      // hapus file lama dengan nama sama kalau ada, supaya tidak menumpuk
      const existing = folder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      const publicUrl = `https://lh3.googleusercontent.com/d/${fileId}`;

      const col = body.slot === "kondisi" ? 9 : 10; // FotoKondisiURL / FotoPasangURL
      sheet.getRange(row, col).setValue(publicUrl);

      return jsonOut_({ ok: true, url: publicUrl });
    }

    if (action === "removePhoto") {
      requireRole_(body.token, ["admin"]);
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      const col = body.slot === "kondisi" ? 9 : 10;
      sheet.getRange(row, col).setValue("");
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: "Aksi tidak dikenal: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
