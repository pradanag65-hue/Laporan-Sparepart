import { useState, useEffect, useRef, useMemo, useCallback, Fragment } from "react";
import {
  Plus, Trash2, Upload, Download, Printer, Image as ImageIcon,
  FileSpreadsheet, Loader2, Camera, X, ClipboardList, ScanLine,
  RotateCcw, ChevronDown, Settings, Link2, CheckCircle2,
} from "lucide-react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { api, getApiUrl, setApiUrl } from "./api.js";
import "./styles.css";

const SATUAN_OPTIONS = ["PC", "LITER", "SET", "UNIT", "METER", "BUAH"];

function uid() {
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function formatTanggalID(iso) {
  if (!iso) return "-";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("id-ID", { day: "2-digit", month: "long", year: "numeric" });
}

function compressImage(file, maxDim = 1280, quality = 0.78) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Gagal membaca file"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("File bukan gambar yang valid"));
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function normalizeTanggal(val) {
  if (val == null || val === "") return todayISO();
  if (typeof val === "number" && val > 20000 && val < 60000) {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    let [, dd, mm, yy] = dmy;
    if (yy.length === 2) yy = "20" + yy;
    return `${yy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return todayISO();
}

function findField(row, candidates) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const hit = keys.find((k) => k.toLowerCase().replace(/[^a-z0-9]/g, "").includes(cand));
    if (hit) return row[hit];
  }
  return "";
}

function rowToFields(row) {
  return {
    tanggal: normalizeTanggal(findField(row, ["tanggal", "tgl"])),
    lb: String(findField(row, ["lb", "lamb", "hari", "no"])).trim(),
    namaPart: String(findField(row, ["namasparepart", "namapart", "sparepart"])).trim(),
    kodeBarang: String(findField(row, ["kodebarang", "kode"])).trim(),
    jumlah: String(findField(row, ["jumlah"])).trim(),
    satuan: String(findField(row, ["satuan"])).trim().toUpperCase() || "PC",
    mekanik: String(findField(row, ["mekanik"])).trim(),
  };
}

const emptyDraft = () => ({
  tanggal: todayISO(), lb: "", namaPart: "", kodeBarang: "", jumlah: "", satuan: "PC", mekanik: "",
});

export default function App() {
  const [apiUrlInput, setApiUrlInput] = useState(getApiUrl());
  const [connected, setConnected] = useState(!!getApiUrl());
  const [entries, setEntries] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [tab, setTab] = useState("input");
  const [draft, setDraft] = useState(emptyDraft());
  const [lampiranDate, setLampiranDate] = useState("");
  const [toast, setToast] = useState(null);
  const [busy, setBusy] = useState(false);
  const [confirmReset, setConfirmReset] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const importRef = useRef(null);

  const notify = useCallback((msg, tone = "ok") => {
    setToast({ msg, tone });
    window.clearTimeout(notify._t);
    notify._t = window.setTimeout(() => setToast(null), 3200);
  }, []);

  const loadEntries = useCallback(async () => {
    if (!getApiUrl()) return;
    setLoaded(false);
    setLoadError("");
    try {
      const list = await api.listEntries();
      setEntries(list);
    } catch (err) {
      console.error(err);
      setLoadError(String(err.message || err));
    } finally {
      setLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (connected) loadEntries();
  }, [connected, loadEntries]);

  function saveApiUrl() {
    if (!apiUrlInput.trim()) return;
    setApiUrl(apiUrlInput);
    setConnected(true);
    setShowSettings(false);
  }

  const tanggalList = useMemo(() => {
    const set = new Set(entries.map((e) => e.tanggal));
    return Array.from(set).sort((a, b) => (a < b ? 1 : -1));
  }, [entries]);

  useEffect(() => {
    if (!lampiranDate && tanggalList.length) setLampiranDate(tanggalList[0]);
  }, [tanggalList, lampiranDate]);

  const sortedEntries = useMemo(
    () =>
      [...entries].sort((a, b) => {
        if (a.tanggal !== b.tanggal) return a.tanggal < b.tanggal ? 1 : -1;
        return (a.lb || "").localeCompare(b.lb || "", undefined, { numeric: true });
      }),
    [entries]
  );

  const lampiranEntries = useMemo(
    () => sortedEntries.filter((e) => e.tanggal === lampiranDate),
    [sortedEntries, lampiranDate]
  );

  async function addEntry() {
    if (!draft.namaPart.trim()) {
      notify("Isi nama spare part dulu.", "err");
      return;
    }
    setBusy(true);
    try {
      const id = await api.addEntry(draft);
      setEntries((prev) => [{ id, ...draft, fotoKondisi: null, fotoPasang: null }, ...prev]);
      setDraft({ ...emptyDraft(), tanggal: draft.tanggal });
      notify("Data ditambahkan ke Google Sheet.");
    } catch (err) {
      console.error(err);
      notify("Gagal menyimpan: " + err.message, "err");
    } finally {
      setBusy(false);
    }
  }

  async function deleteEntry(id) {
    setBusy(true);
    try {
      await api.deleteEntry(id);
      setEntries((prev) => prev.filter((e) => e.id !== id));
      notify("Data dihapus.");
    } catch (err) {
      notify("Gagal menghapus: " + err.message, "err");
    } finally {
      setBusy(false);
    }
  }

  function updateFieldLocal(id, field, value) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, [field]: value } : e)));
  }

  async function commitField(id, field, value) {
    try {
      await api.updateEntry(id, { [field]: value });
    } catch (err) {
      notify("Gagal menyimpan perubahan: " + err.message, "err");
    }
  }

  async function attachPhoto(id, slot, file) {
    if (!file) return;
    setBusy(true);
    try {
      const dataUrl = await compressImage(file);
      const url = await api.uploadPhoto(id, slot, dataUrl);
      const key = slot === "kondisi" ? "fotoKondisi" : "fotoPasang";
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, [key]: url } : e)));
      notify("Foto tersimpan ke Google Drive.");
    } catch (err) {
      console.error(err);
      notify("Gagal mengunggah foto: " + err.message, "err");
    } finally {
      setBusy(false);
    }
  }

  async function removePhoto(id, slot) {
    try {
      await api.removePhoto(id, slot);
      const key = slot === "kondisi" ? "fotoKondisi" : "fotoPasang";
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, [key]: null } : e)));
    } catch (err) {
      notify("Gagal menghapus foto: " + err.message, "err");
    }
  }

  async function handleImport(file) {
    if (!file) return;
    setBusy(true);
    try {
      const ext = file.name.split(".").pop().toLowerCase();
      let rows = [];
      if (ext === "csv") {
        rows = await new Promise((resolve, reject) => {
          Papa.parse(file, { header: true, skipEmptyLines: true, complete: (r) => resolve(r.data), error: reject });
        });
      } else {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      }
      const fieldsList = rows.map(rowToFields).filter((f) => f.namaPart || f.kodeBarang);
      if (!fieldsList.length) {
        notify("Tidak ada baris yang bisa dibaca dari file ini.", "err");
        return;
      }
      let count = 0;
      for (const fields of fieldsList) {
        const id = await api.addEntry(fields);
        setEntries((prev) => [{ id, ...fields, fotoKondisi: null, fotoPasang: null }, ...prev]);
        count++;
      }
      notify(`${count} baris berhasil diimpor ke Google Sheet.`);
    } catch (err) {
      console.error(err);
      notify("Gagal mengimpor: " + err.message, "err");
    } finally {
      setBusy(false);
      if (importRef.current) importRef.current.value = "";
    }
  }

  function exportData() {
    if (!entries.length) {
      notify("Belum ada data untuk diekspor.", "err");
      return;
    }
    const rows = sortedEntries.map((e) => ({
      TANGGAL: formatTanggalID(e.tanggal),
      LB: e.lb,
      "NAMA SPARE PART": e.namaPart,
      "KODE BARANG": e.kodeBarang,
      "JUMLAH SPARE PART BEKAS": e.jumlah,
      SATUAN: e.satuan,
      MEKANIK: e.mekanik,
      "FOTO KONDISI": e.fotoKondisi || "-",
      "FOTO PEMASANGAN": e.fotoPasang || "-",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 16 }, { wch: 6 }, { wch: 34 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 40 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Barang Bekas");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `rekap-barang-bekas-${todayISO()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    notify("Rekap data diunduh.");
  }

  async function resetAll() {
    setBusy(true);
    try {
      for (const e of entries) await api.deleteEntry(e.id);
      setEntries([]);
      notify("Semua data direset.");
    } catch (err) {
      notify("Gagal mereset data: " + err.message, "err");
    } finally {
      setBusy(false);
      setConfirmReset(false);
    }
  }

  function triggerPrint() {
    if (!lampiranEntries.length) {
      notify("Pilih tanggal yang punya data dulu.", "err");
      return;
    }
    window.print();
  }

  // ---- onboarding: belum ada URL Apps Script ----
  if (!connected) {
    return (
      <div className="onboard">
        <div className="onboard-card">
          <div className="brand-mark"><ScanLine size={18} /></div>
          <h1>Hubungkan ke Google Sheet</h1>
          <p>
            Tempel URL Web App Apps Script (diakhiri <code>/exec</code>) yang sudah kamu deploy
            dari sheet database rekap barang bekas.
          </p>
          <input
            type="text"
            placeholder="https://script.google.com/macros/s/.../exec"
            value={apiUrlInput}
            onChange={(e) => setApiUrlInput(e.target.value)}
          />
          <button className="btn primary" onClick={saveApiUrl}>
            <Link2 size={16} /> Simpan &amp; hubungkan
          </button>
        </div>
      </div>
    );
  }

  if (!loaded) {
    return (
      <div className="loading-screen">
        <Loader2 className="spin" size={22} />
        <span>Memuat data dari Google Sheet…</span>
      </div>
    );
  }

  if (loadError) {
    return (
      <div className="loading-screen">
        <span className="err-text">Gagal terhubung: {loadError}</span>
        <div className="btn-row">
          <button className="btn ghost" onClick={loadEntries}>Coba lagi</button>
          <button className="btn ghost" onClick={() => setShowSettings(true)}>Ubah URL</button>
        </div>
        {showSettings && (
          <SettingsPanel
            apiUrlInput={apiUrlInput}
            setApiUrlInput={setApiUrlInput}
            onSave={saveApiUrl}
            onClose={() => setShowSettings(false)}
          />
        )}
      </div>
    );
  }

  return (
    <div className="shell">
      <header className="topbar no-print">
        <div className="brand">
          <div className="brand-mark"><ScanLine size={18} strokeWidth={2.2} /></div>
          <div>
            <h1>Rekap Barang Bekas</h1>
            <p>Dinas Perhubungan DIY · Trans Jogja</p>
          </div>
          <button className="icon-btn settings-btn" onClick={() => setShowSettings(true)} title="Pengaturan">
            <Settings size={15} />
          </button>
        </div>
        <nav className="tabs">
          {[
            { id: "input", label: "Input & Data", icon: ClipboardList },
            { id: "lampiran", label: "Lampiran Foto", icon: ImageIcon },
            { id: "io", label: "Impor & Ekspor", icon: FileSpreadsheet },
          ].map(({ id, label, icon: Icon }) => (
            <button key={id} className={`tab ${tab === id ? "active" : ""}`} onClick={() => setTab(id)}>
              <Icon size={15} />{label}
            </button>
          ))}
        </nav>
      </header>

      {tab === "input" && (
        <main className="page no-print">
          <section className="card form-card">
            <h2>Tambah data harian</h2>
            <div className="form-grid">
              <label><span>Tanggal</span>
                <input type="date" value={draft.tanggal} onChange={(e) => setDraft((d) => ({ ...d, tanggal: e.target.value }))} />
              </label>
              <label><span>LB (no. lambang)</span>
                <input type="text" placeholder="mis. 47" value={draft.lb} onChange={(e) => setDraft((d) => ({ ...d, lb: e.target.value }))} />
              </label>
              <label className="col-2"><span>Nama spare part</span>
                <input type="text" placeholder="mis. LINER NQR 71" value={draft.namaPart} onChange={(e) => setDraft((d) => ({ ...d, namaPart: e.target.value }))} />
              </label>
              <label><span>Kode barang</span>
                <input type="text" placeholder="mis. NQR0090" value={draft.kodeBarang} onChange={(e) => setDraft((d) => ({ ...d, kodeBarang: e.target.value }))} />
              </label>
              <label><span>Jumlah</span>
                <input type="text" inputMode="decimal" placeholder="mis. 4" value={draft.jumlah} onChange={(e) => setDraft((d) => ({ ...d, jumlah: e.target.value }))} />
              </label>
              <label><span>Satuan</span>
                <select value={draft.satuan} onChange={(e) => setDraft((d) => ({ ...d, satuan: e.target.value }))}>
                  {SATUAN_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </label>
              <label><span>Mekanik</span>
                <input type="text" placeholder="mis. ARIF TRI" value={draft.mekanik} onChange={(e) => setDraft((d) => ({ ...d, mekanik: e.target.value }))} />
              </label>
            </div>
            <button className="btn primary" onClick={addEntry}><Plus size={16} />Tambah ke rekap</button>
          </section>

          <section className="card">
            <div className="card-head">
              <h2>Rekap tersimpan</h2>
              <span className="count-pill">{entries.length} baris</span>
            </div>
            {!entries.length ? (
              <EmptyState text="Belum ada data. Tambahkan lewat form di atas atau impor file rekap." />
            ) : (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Tanggal</th><th>LB</th><th>Nama spare part</th><th>Kode</th>
                      <th>Jml</th><th>Satuan</th><th>Mekanik</th><th>Foto</th><th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedEntries.map((e) => (
                      <tr key={e.id}>
                        <td>
                          <input type="date" value={e.tanggal}
                            onChange={(ev) => updateFieldLocal(e.id, "tanggal", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "tanggal", ev.target.value)} />
                        </td>
                        <td className="mono">
                          <input className="cell-input narrow" value={e.lb}
                            onChange={(ev) => updateFieldLocal(e.id, "lb", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "lb", ev.target.value)} />
                        </td>
                        <td>
                          <input className="cell-input" value={e.namaPart}
                            onChange={(ev) => updateFieldLocal(e.id, "namaPart", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "namaPart", ev.target.value)} />
                        </td>
                        <td className="mono">
                          <input className="cell-input" value={e.kodeBarang}
                            onChange={(ev) => updateFieldLocal(e.id, "kodeBarang", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "kodeBarang", ev.target.value)} />
                        </td>
                        <td>
                          <input className="cell-input narrow" value={e.jumlah}
                            onChange={(ev) => updateFieldLocal(e.id, "jumlah", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "jumlah", ev.target.value)} />
                        </td>
                        <td>
                          <select value={e.satuan}
                            onChange={(ev) => { updateFieldLocal(e.id, "satuan", ev.target.value); commitField(e.id, "satuan", ev.target.value); }}>
                            {SATUAN_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                          </select>
                        </td>
                        <td>
                          <input className="cell-input" value={e.mekanik}
                            onChange={(ev) => updateFieldLocal(e.id, "mekanik", ev.target.value)}
                            onBlur={(ev) => commitField(e.id, "mekanik", ev.target.value)} />
                        </td>
                        <td>
                          <div className="photo-slots">
                            <PhotoSlot label="Kondisi" value={e.fotoKondisi} onPick={(f) => attachPhoto(e.id, "kondisi", f)} onRemove={() => removePhoto(e.id, "kondisi")} />
                            <PhotoSlot label="Pasang" value={e.fotoPasang} onPick={(f) => attachPhoto(e.id, "pasang", f)} onRemove={() => removePhoto(e.id, "pasang")} />
                          </div>
                        </td>
                        <td>
                          <button className="icon-btn danger" onClick={() => deleteEntry(e.id)} title="Hapus baris"><Trash2 size={15} /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </main>
      )}

      {tab === "lampiran" && (
        <main className="page">
          <section className="card no-print">
            <div className="card-head wrap">
              <h2>Lampiran foto per tanggal</h2>
              <div className="lampiran-controls">
                <div className="select-wrap">
                  <select value={lampiranDate} onChange={(e) => setLampiranDate(e.target.value)}>
                    {tanggalList.length === 0 && <option value="">Belum ada tanggal</option>}
                    {tanggalList.map((t) => <option key={t} value={t}>{formatTanggalID(t)}</option>)}
                  </select>
                  <ChevronDown size={14} className="select-caret" />
                </div>
                <button className="btn primary" onClick={triggerPrint}><Printer size={16} />Cetak / simpan PDF</button>
              </div>
            </div>
            <p className="hint">
              Setiap baris rekap otomatis jadi satu blok foto — keterangan LB, nama part, dan kode mengikuti data
              di Google Sheet, jadi fotonya selalu menempel ke barang yang benar.
            </p>
          </section>

          {!lampiranEntries.length ? (
            <div className="no-print"><EmptyState text="Tidak ada data foto untuk tanggal ini. Pilih tanggal lain atau tambah data dulu." /></div>
          ) : (
            <div className="print-area">
              <div className="sheet-head">
                <h3>LAMPIRAN FOTO SPARE PART BEKAS</h3>
                <p>{formatTanggalID(lampiranDate)}</p>
              </div>
              <div className="photo-grid">
                {lampiranEntries.map((e) => (
                  <Fragment key={e.id}>
                    <PhotoCard entry={e} slot="kondisi" title="(BARU DAN BEKAS)" onPick={(f) => attachPhoto(e.id, "kondisi", f)} />
                    <PhotoCard entry={e} slot="pasang" title="(PENGGANTIAN)" onPick={(f) => attachPhoto(e.id, "pasang", f)} />
                  </Fragment>
                ))}
              </div>
            </div>
          )}
        </main>
      )}

      {tab === "io" && (
        <main className="page no-print">
          <section className="card">
            <h2>Impor data rekap</h2>
            <p className="hint">
              Terima file .csv atau .xlsx dengan kolom seperti rekap harian (LB/Hari, Tanggal, Nama Spare Part,
              Kode Barang, Jumlah, Satuan, Mekanik). Nama kolom tidak harus persis sama. Setiap baris langsung
              ditulis ke Google Sheet.
            </p>
            <input ref={importRef} type="file" accept=".csv,.xlsx,.xls" className="file-input" id="import-file"
              onChange={(e) => handleImport(e.target.files?.[0])} />
            <label htmlFor="import-file" className="btn ghost"><Upload size={16} />Pilih file rekap</label>
          </section>

          <section className="card">
            <h2>Ekspor</h2>
            <p className="hint">
              Unduh seluruh rekap sebagai spreadsheet (data selalu sinkron dengan Google Sheet sumbernya), atau
              buka tab Lampiran Foto dan cetak tanggal yang diinginkan sebagai PDF.
            </p>
            <div className="btn-row">
              <button className="btn primary" onClick={exportData}><Download size={16} />Ekspor rekap (.xlsx)</button>
              <button className="btn ghost" onClick={() => setTab("lampiran")}><Printer size={16} />Buka lampiran foto</button>
            </div>
          </section>

          <section className="card danger-zone">
            <h2>Reset data</h2>
            <p className="hint">Menghapus seluruh baris rekap di Google Sheet (foto di Drive tidak otomatis terhapus). Tidak bisa dibatalkan.</p>
            {!confirmReset ? (
              <button className="btn ghost danger" onClick={() => setConfirmReset(true)}><RotateCcw size={16} />Reset semua data</button>
            ) : (
              <div className="btn-row">
                <button className="btn danger" onClick={resetAll}>Ya, hapus semuanya</button>
                <button className="btn ghost" onClick={() => setConfirmReset(false)}>Batal</button>
              </div>
            )}
          </section>
        </main>
      )}

      {busy && <div className="busy-overlay no-print"><Loader2 className="spin" size={20} /></div>}
      {toast && (
        <div className={`toast ${toast.tone} no-print`}>
          {toast.msg}
          <button onClick={() => setToast(null)}><X size={13} /></button>
        </div>
      )}
      {showSettings && (
        <SettingsPanel
          apiUrlInput={apiUrlInput}
          setApiUrlInput={setApiUrlInput}
          onSave={saveApiUrl}
          onClose={() => setShowSettings(false)}
        />
      )}
    </div>
  );
}

function SettingsPanel({ apiUrlInput, setApiUrlInput, onSave, onClose }) {
  return (
    <div className="modal-backdrop no-print" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <h2>URL Apps Script</h2>
        <input type="text" value={apiUrlInput} onChange={(e) => setApiUrlInput(e.target.value)}
          placeholder="https://script.google.com/macros/s/.../exec" />
        <div className="btn-row">
          <button className="btn primary" onClick={onSave}><CheckCircle2 size={16} />Simpan</button>
          <button className="btn ghost" onClick={onClose}>Tutup</button>
        </div>
      </div>
    </div>
  );
}

function EmptyState({ text }) {
  return <div className="empty-state"><Camera size={22} /><p>{text}</p></div>;
}

function PhotoSlot({ label, value, onPick, onRemove }) {
  const inputId = useRef(`ph-${uid()}`).current;
  return (
    <div className="photo-slot">
      {value ? (
        <div className="thumb">
          <img src={value} alt={label} referrerPolicy="no-referrer" />
          <button className="thumb-remove" onClick={onRemove} title="Hapus foto"><X size={11} /></button>
        </div>
      ) : (
        <>
          <input type="file" accept="image/*" id={inputId} className="file-input" onChange={(e) => onPick(e.target.files?.[0])} />
          <label htmlFor={inputId} className="thumb-empty" title={`Unggah foto ${label}`}><Camera size={13} /></label>
        </>
      )}
      <span className="slot-label">{label}</span>
    </div>
  );
}

function PhotoCard({ entry, slot, title, onPick }) {
  const value = slot === "kondisi" ? entry.fotoKondisi : entry.fotoPasang;
  const inputId = useRef(`pc-${uid()}`).current;
  return (
    <figure className="photo-card">
      <div className="photo-frame">
        {value ? (
          <img src={value} alt={entry.namaPart} referrerPolicy="no-referrer" />
        ) : (
          <div className="photo-missing no-print">
            <input type="file" accept="image/*" id={inputId} className="file-input" onChange={(e) => onPick(e.target.files?.[0])} />
            <label htmlFor={inputId}><Camera size={18} /><span>Unggah foto</span></label>
          </div>
        )}
        <span className="photo-tag">LB {entry.lb || "–"} · {entry.namaPart || "-"}</span>
      </div>
      <figcaption><strong>{entry.namaPart || "-"}</strong><span>{title}</span></figcaption>
    </figure>
  );
}
