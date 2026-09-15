/**
 * REKAP BARANG BEKAS — backend Apps Script
 * -----------------------------------------
 * Database: sheet "Entries" di spreadsheet yang sama dengan script ini
 *           (bikin script ini lewat Extensions > Apps Script di dalam
 *           Google Sheet-nya, supaya otomatis "bound" ke sheet itu).
 * Foto    : disimpan sebagai file di folder Drive "RekapBarangBekas_Foto",
 *           URL publik-nya disimpan di kolom FotoKondisiURL / FotoPasangURL.
 *
 * CARA DEPLOY:
 * 1. Buka Google Sheet baru, kasih nama mis. "Rekap Barang Bekas - DB".
 * 2. Buat sheet bernama persis "Entries" (huruf besar E), baris pertama
 *    diisi header (script akan membuatnya otomatis kalau kosong).
 * 3. Extensions > Apps Script, hapus isi default, tempel seluruh isi file ini.
 * 4. Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Klik Deploy, izinkan akses (authorize) saat diminta.
 * 6. Copy URL Web App yang muncul (diakhiri /exec) — itu yang dipakai
 *    di pengaturan aplikasi frontend (VITE / localStorage apiUrl).
 * 7. Setiap kali code ini diubah, buat "New deployment" lagi (atau Manage
 *    deployments > edit > New version) supaya perubahan aktif.
 */

const SHEET_NAME = "Entries";
const PHOTO_FOLDER_NAME = "RekapBarangBekas_Foto";
const HEADERS = [
  "ID", "Tanggal", "LB", "NamaPart", "KodeBarang", "Jumlah",
  "Satuan", "Mekanik", "FotoKondisiURL", "FotoPasangURL", "CreatedAt",
];

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

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
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

// ---- GET: baca data ----
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

// ---- POST: tulis data (add / update / delete / uploadPhoto) ----
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action;
    const sheet = getSheet_();

    if (action === "add") {
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
      const row = findRowById_(sheet, body.id);
      if (row === -1) return jsonOut_({ ok: false, error: "ID tidak ditemukan" });
      sheet.deleteRow(row);
      return jsonOut_({ ok: true });
    }

    if (action === "uploadPhoto") {
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
