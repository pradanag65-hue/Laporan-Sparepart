import { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  Plus, Trash2, Upload, Download, Printer, Image as ImageIcon,
  FileSpreadsheet, Loader2, Camera, X, ClipboardList, Leaf,
  RotateCcw, ChevronDown, Settings, Link2, CheckCircle2, Images,
  Bell, Calendar, Tag, Wrench, Barcode, Hash, Box, User, Bus,
  LogOut, Lock, Eye, Search, BarChart3, History, Users, Mail,
  UserPlus, ShieldCheck, KeyRound,
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { api, getApiUrl, setApiUrl, getAuthUser, setAuthUser, clearAuthUser } from "./api.js";
import "./styles.css";

const SATUAN_OPTIONS = ["PC", "LITER", "SET", "UNIT", "METER", "BUAH"];
const ROLE_LABEL = { admin: "Admin", input: "Input", guest: "Guest" };
const ROLE_OPTIONS = ["admin", "input", "guest"];
const MONTH_LABEL = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

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
function formatDateTimeID(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString("id-ID", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
}
function isBanEntry(e) {
  return /\bban\b/i.test(e.namaPart || "");
}
// Jumlah=2 -> ban depan (8 foto: kanan/kiri x baru/bekas x foto/kode).
// Jumlah=4 -> ban belakang, dual/kembar (14 foto: 4 posisi luar-dalam kanan-kiri).
// Selain itu (nama mengandung "ban" tapi jumlah lain) dianggap depan sebagai default aman.
function getBanKind(e) {
  if (!isBanEntry(e)) return null;
  const n = parseFloat(String(e.jumlah).replace(",", "."));
  if (n === 4) return "belakang";
  return "depan";
}

const BAN_FRONT_SLOTS = [
  { slot: "banBaruKanan", label: "Ban Kanan (Baru)" },
  { slot: "banBaruKiri", label: "Ban Kiri (Baru)" },
  { slot: "banKodeBaruKanan", label: "Kode Ban Kanan (Baru)" },
  { slot: "banKodeBaruKiri", label: "Kode Ban Kiri (Baru)" },
  { slot: "banBekasKanan", label: "Ban Kanan (Bekas)" },
  { slot: "banBekasKiri", label: "Ban Kiri (Bekas)" },
  { slot: "banKodeBekasKanan", label: "Kode Ban Kanan (Bekas)" },
  { slot: "banKodeBekasKiri", label: "Kode Ban Kiri (Bekas)" },
];

const BAN_REAR_SLOTS = [
  { slot: "banRearBaruKananLuar", label: "Ban Kanan Luar (Baru)" },
  { slot: "banRearBaruKananDalam", label: "Ban Kanan Dalam (Baru)" },
  { slot: "banRearBaruKiriLuar", label: "Ban Kiri Luar (Baru)" },
  { slot: "banRearBaruKiriDalam", label: "Ban Kiri Dalam (Baru)" },
  { slot: "banRearKodeBaruKananLuar", label: "Kode Ban Kanan Luar (Baru)" },
  { slot: "banRearKodeBaruKananDalam", label: "Kode Ban Kanan Dalam (Baru)" },
  { slot: "banRearKodeBaruKiriLuar", label: "Kode Ban Kiri Luar (Baru)" },
  { slot: "banRearKodeBaruKiriDalam", label: "Kode Ban Kiri Dalam (Baru)" },
  { slot: "banRearBekasKanan", label: "Ban Belakang Kanan (Bekas)" },
  { slot: "banRearBekasKiri", label: "Ban Belakang Kiri (Bekas)" },
  { slot: "banRearKodeBekasKananLuar", label: "Kode Ban Kanan Luar (Bekas)" },
  { slot: "banRearKodeBekasKananDalam", label: "Kode Ban Kanan Dalam (Bekas)" },
  { slot: "banRearKodeBekasKiriLuar", label: "Kode Ban Kiri Luar (Bekas)" },
  { slot: "banRearKodeBekasKiriDalam", label: "Kode Ban Kiri Dalam (Bekas)" },
];

function banSlotsFor(kind) {
  return kind === "belakang" ? BAN_REAR_SLOTS : BAN_FRONT_SLOTS;
}

// slot -> [namaField di entry, key di dalam field itu] — dipakai untuk baca/tulis state lokal.
const BAN_FIELD_MAP = {
  banBaruKanan: ["fotoBan", "baruKanan"], banBaruKiri: ["fotoBan", "baruKiri"],
  banKodeBaruKanan: ["fotoBan", "kodeBaruKanan"], banKodeBaruKiri: ["fotoBan", "kodeBaruKiri"],
  banBekasKanan: ["fotoBan", "bekasKanan"], banBekasKiri: ["fotoBan", "bekasKiri"],
  banKodeBekasKanan: ["fotoBan", "kodeBekasKanan"], banKodeBekasKiri: ["fotoBan", "kodeBekasKiri"],
  banRearBaruKananLuar: ["fotoBanRear", "baruKananLuar"], banRearBaruKananDalam: ["fotoBanRear", "baruKananDalam"],
  banRearBaruKiriLuar: ["fotoBanRear", "baruKiriLuar"], banRearBaruKiriDalam: ["fotoBanRear", "baruKiriDalam"],
  banRearKodeBaruKananLuar: ["fotoBanRear", "kodeBaruKananLuar"], banRearKodeBaruKananDalam: ["fotoBanRear", "kodeBaruKananDalam"],
  banRearKodeBaruKiriLuar: ["fotoBanRear", "kodeBaruKiriLuar"], banRearKodeBaruKiriDalam: ["fotoBanRear", "kodeBaruKiriDalam"],
  banRearBekasKanan: ["fotoBanRear", "bekasKanan"], banRearBekasKiri: ["fotoBanRear", "bekasKiri"],
  banRearKodeBekasKananLuar: ["fotoBanRear", "kodeBekasKananLuar"], banRearKodeBekasKananDalam: ["fotoBanRear", "kodeBekasKananDalam"],
  banRearKodeBekasKiriLuar: ["fotoBanRear", "kodeBekasKiriLuar"], banRearKodeBekasKiriDalam: ["fotoBanRear", "kodeBekasKiriDalam"],
};
function getBanPhotoValue(entry, slot) {
  const mapping = BAN_FIELD_MAP[slot];
  if (!mapping) return null;
  const [obj, field] = mapping;
  return entry[obj] ? entry[obj][field] : null;
}
function countBanPhotos(entry, kind) {
  return banSlotsFor(kind).filter(({ slot }) => getBanPhotoValue(entry, slot)).length;
}
// Pecah daftar slot jadi kelompok maksimal 8 foto per halaman (ukuran Folio).
function chunkSlots(slots, size = 8) {
  const chunks = [];
  for (let i = 0; i < slots.length; i += size) chunks.push(slots.slice(i, i + size));
  return chunks;
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
  const [apiUrlInput, setApiUrlInput] = useState(getApiUrl().includes("PASTE_URL") ? "" : getApiUrl());
  const [connected, setConnected] = useState(!!getApiUrl() && !getApiUrl().includes("PASTE_URL"));
  const [authUser, setAuthUserState] = useState(getAuthUser());
  const [loginForm, setLoginForm] = useState({ username: "", password: "" });
  const [loginError, setLoginError] = useState("");
  const [loginBusy, setLoginBusy] = useState(false);
  const [entries, setEntries] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [tab, setTab] = useState("input");
  const [draft, setDraft] = useState(emptyDraft());
  const [lampiranSearch, setLampiranSearch] = useState("");
  const [lampiranDateFrom, setLampiranDateFrom] = useState("");
  const [lampiranDateTo, setLampiranDateTo] = useState("");
  const [toast, setToast] = useState(null);
  const [busy, setBusy] = useState(false);
  const [confirmReset, setConfirmReset] = useState(false);
  const [resetPassword, setResetPassword] = useState("");
  const [resetError, setResetError] = useState("");
  const [backupBusy, setBackupBusy] = useState(false);
  const [importProgress, setImportProgress] = useState(null);
  const [search, setSearch] = useState("");
  const [colWidths, setColWidths] = useState({
    tanggal: 122, lb: 56, namaPart: 220, kode: 128, jumlah: 56, satuan: 90, mekanik: 130, foto: 150,
  });
  const colResizeRef = useRef(null);
  const [banModalEntry, setBanModalEntry] = useState(null);
  const [lightbox, setLightbox] = useState(null);
  const [uploadingKeys, setUploadingKeys] = useState(() => new Set());
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const importRef = useRef(null);

  // admin-only data
  const [activityLog, setActivityLog] = useState([]);
  const [activityLoaded, setActivityLoaded] = useState(false);
  const [users, setUsers] = useState([]);
  const [usersLoaded, setUsersLoaded] = useState(false);
  const [newUser, setNewUser] = useState({ username: "", password: "", role: "input" });
  const [userError, setUserError] = useState("");
  const [notifyEmail, setNotifyEmail] = useState("");
  const [configLoaded, setConfigLoaded] = useState(false);
  const [configSaving, setConfigSaving] = useState(false);

  const notify = useCallback((msg, tone = "ok") => {
    setToast({ msg, tone });
    window.clearTimeout(notify._t);
    notify._t = window.setTimeout(() => setToast(null), 3200);
  }, []);

  const handleApiError = useCallback((err, prefix) => {
    const msg = err.message || String(err);
    notify(prefix ? `${prefix}: ${msg}` : msg, "err");
    if (/sesi tidak valid/i.test(msg)) {
      clearAuthUser();
      setAuthUserState(null);
    }
  }, [notify]);

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

  async function handleLogin() {
    if (!loginForm.username.trim() || !loginForm.password) {
      setLoginError("Isi username dan password.");
      return;
    }
    setLoginBusy(true);
    setLoginError("");
    try {
      const result = await api.login(loginForm.username.trim(), loginForm.password);
      const user = { username: result.username, role: result.role, token: result.token };
      setAuthUser(user);
      setAuthUserState(user);
    } catch (err) {
      setLoginError(err.message || "Login gagal.");
    } finally {
      setLoginBusy(false);
    }
  }

  function handleLogout() {
    if (!window.confirm("Keluar dari akun ini?")) return;
    clearAuthUser();
    setAuthUserState(null);
    setLoginForm({ username: "", password: "" });
  }

  const role = authUser ? authUser.role : null;
  const canCreate = role === "admin" || role === "input";
  const canEdit = role === "admin";

  const navItems = useMemo(() => {
    const base = [
      { id: "input", label: "Input & Data", icon: ClipboardList },
      { id: "lampiran", label: "Lampiran Foto", icon: ImageIcon },
      { id: "dashboard", label: "Dashboard", icon: BarChart3 },
      { id: "io", label: "Impor & Ekspor", icon: FileSpreadsheet },
    ];
    if (role === "admin") {
      base.push({ id: "activity", label: "Log Aktivitas", icon: History });
      base.push({ id: "users", label: "Kelola Akun", icon: Users });
    }
    return base;
  }, [role]);

  // ---- load admin-only data lazily when tab opened ----
  useEffect(() => {
    if (tab === "activity" && role === "admin" && !activityLoaded) {
      api.listActivity().then((logs) => { setActivityLog(logs); setActivityLoaded(true); }).catch((err) => handleApiError(err, "Gagal memuat log"));
    }
    if (tab === "users" && role === "admin") {
      if (!usersLoaded) {
        api.listUsers().then((list) => { setUsers(list); setUsersLoaded(true); }).catch((err) => handleApiError(err, "Gagal memuat akun"));
      }
      if (!configLoaded) {
        api.getConfig().then((cfg) => { setNotifyEmail(cfg.notifyEmail || ""); setConfigLoaded(true); }).catch(() => setConfigLoaded(true));
      }
    }
  }, [tab, role, activityLoaded, usersLoaded, configLoaded, handleApiError]);

  const tanggalList = useMemo(() => {
    const set = new Set(entries.map((e) => e.tanggal));
    return Array.from(set).sort((a, b) => (a < b ? 1 : -1));
  }, [entries]);

  const totalFoto = useMemo(
    () => entries.reduce((n, e) => n + (e.fotoKondisi ? 1 : 0) + (e.fotoPasang ? 1 : 0), 0),
    [entries]
  );

  useEffect(() => {
    if (!lampiranDateFrom && !lampiranDateTo && tanggalList.length) {
      setLampiranDateFrom(tanggalList[0]);
      setLampiranDateTo(tanggalList[0]);
    }
  }, [tanggalList, lampiranDateFrom, lampiranDateTo]);

  const sortedEntries = useMemo(
    () =>
      [...entries].sort((a, b) => {
        if (a.tanggal !== b.tanggal) return a.tanggal < b.tanggal ? 1 : -1;
        return (a.lb || "").localeCompare(b.lb || "", undefined, { numeric: true });
      }),
    [entries]
  );

  const lampiranFilteredEntries = useMemo(() => {
    const q = lampiranSearch.trim().toLowerCase();
    return sortedEntries.filter((e) => {
      if (lampiranDateFrom && e.tanggal < lampiranDateFrom) return false;
      if (lampiranDateTo && e.tanggal > lampiranDateTo) return false;
      if (!q) return true;
      return [e.namaPart, e.kodeBarang, e.lb, e.mekanik, e.satuan, formatTanggalID(e.tanggal)]
        .join(" ").toLowerCase().includes(q);
    });
  }, [sortedEntries, lampiranSearch, lampiranDateFrom, lampiranDateTo]);

  // Susun jadi daftar "blok" cetak: satu blok untuk semua part biasa per tanggal,
  // dan satu blok per halaman untuk tiap entri ban (dipecah maks. 8 foto/halaman).
  const lampiranBlocks = useMemo(() => {
    const blocks = [];
    const dates = Array.from(new Set(lampiranFilteredEntries.map((e) => e.tanggal))).sort();
    dates.forEach((date) => {
      const dayEntries = lampiranFilteredEntries.filter((e) => e.tanggal === date);
      const normal = dayEntries.filter((e) => !isBanEntry(e));
      const bans = dayEntries.filter((e) => isBanEntry(e));
      if (normal.length) blocks.push({ type: "normal", date, entries: normal });
      bans.forEach((entry) => {
        const kind = getBanKind(entry);
        const chunks = chunkSlots(banSlotsFor(kind), 8);
        chunks.forEach((chunk, ci) => {
          blocks.push({ type: "ban", date, entry, kind, chunk, page: ci + 1, totalPages: chunks.length });
        });
      });
    });
    return blocks;
  }, [lampiranFilteredEntries]);

  const filteredEntries = useMemo(() => {
    const q = search.trim().toLowerCase();
    return sortedEntries.filter((e) => {
      if (dateFrom && e.tanggal < dateFrom) return false;
      if (dateTo && e.tanggal > dateTo) return false;
      if (!q) return true;
      return [e.namaPart, e.kodeBarang, e.lb, e.mekanik, e.satuan, formatTanggalID(e.tanggal)]
        .join(" ").toLowerCase().includes(q);
    });
  }, [sortedEntries, search, dateFrom, dateTo]);

  const namaPartOptions = useMemo(() => Array.from(new Set(entries.map((e) => e.namaPart).filter(Boolean))).sort(), [entries]);
  const kodeBarangOptions = useMemo(() => Array.from(new Set(entries.map((e) => e.kodeBarang).filter(Boolean))).sort(), [entries]);
  const mekanikOptions = useMemo(() => Array.from(new Set(entries.map((e) => e.mekanik).filter(Boolean))).sort(), [entries]);

  // ---- dashboard aggregates ----
  const chartByMonth = useMemo(() => {
    const map = {};
    entries.forEach((e) => {
      if (!e.tanggal) return;
      const key = e.tanggal.slice(0, 7); // YYYY-MM
      map[key] = (map[key] || 0) + 1;
    });
    return Object.keys(map).sort().slice(-6).map((key) => {
      const [, m] = key.split("-");
      return { name: MONTH_LABEL[parseInt(m, 10) - 1] || key, jumlah: map[key] };
    });
  }, [entries]);

  const topMechanics = useMemo(() => {
    const map = {};
    entries.forEach((e) => { if (e.mekanik) map[e.mekanik] = (map[e.mekanik] || 0) + 1; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, jumlah]) => ({ name, jumlah }));
  }, [entries]);

  const topParts = useMemo(() => {
    const map = {};
    entries.forEach((e) => { if (e.namaPart) map[e.namaPart] = (map[e.namaPart] || 0) + 1; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, jumlah]) => ({ name, jumlah }));
  }, [entries]);

  // ---- entry mutations ----
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
      handleApiError(err, "Gagal menyimpan");
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
      handleApiError(err, "Gagal menghapus");
    } finally {
      setBusy(false);
    }
  }

  function confirmDeleteEntry(id, namaPart) {
    const label = namaPart ? `"${namaPart}"` : "baris ini";
    const yakin = window.confirm(`Hapus data ${label}? Foto yang sudah diunggah untuk baris ini juga akan ikut terhapus. Tindakan ini tidak bisa dibatalkan.`);
    if (yakin) deleteEntry(id);
  }

  function updateFieldLocal(id, field, value) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, [field]: value } : e)));
  }

  async function commitField(id, field, value) {
    try {
      await api.updateEntry(id, { [field]: value });
    } catch (err) {
      handleApiError(err, "Gagal menyimpan perubahan");
    }
  }

  async function attachPhoto(id, slot, file) {
    if (!file) return;
    const uploadKey = `${id}:${slot}`;
    setUploadingKeys((prev) => new Set(prev).add(uploadKey));
    setBusy(true);
    try {
      const dataUrl = await compressImage(file);
      const url = await api.uploadPhoto(id, slot, dataUrl);
      setEntries((prev) =>
        prev.map((e) => {
          if (e.id !== id) return e;
          if (slot === "kondisi") return { ...e, fotoKondisi: url };
          if (slot === "pasang") return { ...e, fotoPasang: url };
          const mapping = BAN_FIELD_MAP[slot];
          if (mapping) { const [obj, field] = mapping; return { ...e, [obj]: { ...(e[obj] || {}), [field]: url } }; }
          return e;
        })
      );
      notify("Foto tersimpan ke Google Drive.");
      setLightbox((lb) => (lb && lb.id === id && lb.slot === slot ? { ...lb, url } : lb));
    } catch (err) {
      console.error(err);
      handleApiError(err, "Gagal mengunggah foto");
    } finally {
      setBusy(false);
      setUploadingKeys((prev) => { const next = new Set(prev); next.delete(uploadKey); return next; });
    }
  }

  async function removePhoto(id, slot) {
    try {
      await api.removePhoto(id, slot);
      setEntries((prev) =>
        prev.map((e) => {
          if (e.id !== id) return e;
          if (slot === "kondisi") return { ...e, fotoKondisi: null };
          if (slot === "pasang") return { ...e, fotoPasang: null };
          const mapping = BAN_FIELD_MAP[slot];
          if (mapping) { const [obj, field] = mapping; return { ...e, [obj]: { ...(e[obj] || {}), [field]: null } }; }
          return e;
        })
      );
      setLightbox((lb) => (lb && lb.id === id && lb.slot === slot ? null : lb));
    } catch (err) {
      handleApiError(err, "Gagal menghapus foto");
    }
  }

  const onColResizeMove = useCallback((e) => {
    const r = colResizeRef.current;
    if (!r) return;
    const delta = e.clientX - r.startX;
    setColWidths((prev) => ({ ...prev, [r.key]: Math.max(44, r.startWidth + delta) }));
  }, []);
  const onColResizeUp = useCallback(() => {
    colResizeRef.current = null;
    window.removeEventListener("mousemove", onColResizeMove);
    window.removeEventListener("mouseup", onColResizeUp);
  }, [onColResizeMove]);
  function startColResize(e, key) {
    e.preventDefault();
    colResizeRef.current = { key, startX: e.clientX, startWidth: colWidths[key] };
    window.addEventListener("mousemove", onColResizeMove);
    window.addEventListener("mouseup", onColResizeUp);
  }

  function isUploading(id, slot) {    return uploadingKeys.has(`${id}:${slot}`);
  }

  function openLightbox(id, slot, url, label, allowRemove) {
    setLightbox({
      id, slot, url, label,
      onReplace: (file) => attachPhoto(id, slot, file),
      onRemove: allowRemove ? () => removePhoto(id, slot) : null,
    });
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

      // Kirim per kelompok (bukan satu per satu) supaya tidak butuh ratusan
      // request terpisah yang rawan putus di tengah jalan.
      const CHUNK_SIZE = 150;
      const chunks = [];
      for (let i = 0; i < fieldsList.length; i += CHUNK_SIZE) chunks.push(fieldsList.slice(i, i + CHUNK_SIZE));

      let insertedTotal = 0;
      setImportProgress({ current: 0, total: fieldsList.length });
      try {
        for (const chunk of chunks) {
          const ids = await api.addEntriesBatch(chunk);
          const newEntries = chunk.map((fields, idx) => ({ id: ids[idx], ...fields, fotoKondisi: null, fotoPasang: null }));
          setEntries((prev) => [...newEntries, ...prev]);
          insertedTotal += chunk.length;
          setImportProgress({ current: insertedTotal, total: fieldsList.length });
        }
        notify(`${insertedTotal} baris berhasil diimpor ke Google Sheet.`);
      } catch (err) {
        console.error(err);
        handleApiError(
          err,
          `Impor terhenti setelah ${insertedTotal} dari ${fieldsList.length} baris (baris ke-${insertedTotal + 1} dan seterusnya belum masuk)`
        );
      }
    } catch (err) {
      console.error(err);
      handleApiError(err, "Gagal mengimpor");
    } finally {
      setBusy(false);
      setImportProgress(null);
      if (importRef.current) importRef.current.value = "";
    }
  }

  function exportData() {
    if (!filteredEntries.length) {
      notify("Tidak ada data pada filter saat ini untuk diekspor.", "err");
      return;
    }
    const rows = filteredEntries.map((e) => ({
      TANGGAL: formatTanggalID(e.tanggal),
      LB: e.lb,
      "NAMA SPARE PART": e.namaPart,
      "KODE BARANG": e.kodeBarang,
      "JUMLAH SPARE PART BEKAS": e.jumlah,
      SATUAN: e.satuan,
      MEKANIK: e.mekanik,
      "FOTO KONDISI": e.fotoKondisi || "-",
      "FOTO PEMASANGAN": e.fotoPasang || "-",
      "FOTO BAN KANAN (BARU)": e.fotoBan?.baruKanan || "-",
      "FOTO BAN KIRI (BARU)": e.fotoBan?.baruKiri || "-",
      "KODE BAN KANAN (BARU)": e.fotoBan?.kodeBaruKanan || "-",
      "KODE BAN KIRI (BARU)": e.fotoBan?.kodeBaruKiri || "-",
      "FOTO BAN KANAN (BEKAS)": e.fotoBan?.bekasKanan || "-",
      "FOTO BAN KIRI (BEKAS)": e.fotoBan?.bekasKiri || "-",
      "KODE BAN KANAN (BEKAS)": e.fotoBan?.kodeBekasKanan || "-",
      "KODE BAN KIRI (BEKAS)": e.fotoBan?.kodeBekasKiri || "-",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 16 }, { wch: 6 }, { wch: 34 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 40 }, { wch: 40 },
      { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Barang Bekas");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const rangeTag = dateFrom || dateTo ? `_${dateFrom || "awal"}_sd_${dateTo || "akhir"}` : "";
    a.download = `rekap-barang-bekas${rangeTag}-${todayISO()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    notify("Rekap data diunduh.");
  }

  async function resetAll() {
    if (!resetPassword) {
      setResetError("Masukkan ulang password admin untuk konfirmasi.");
      return;
    }
    setBusy(true);
    setResetError("");
    try {
      const backupSheet = await api.resetAll(resetPassword);
      setEntries([]);
      notify(`Semua data direset. Cadangan tersimpan di sheet "${backupSheet}".`);
      setConfirmReset(false);
      setResetPassword("");
    } catch (err) {
      setResetError(err.message || "Gagal mereset data.");
    } finally {
      setBusy(false);
    }
  }

  async function backupNow() {
    setBackupBusy(true);
    try {
      const backupSheet = await api.backupNow();
      notify(`Cadangan dibuat: sheet "${backupSheet}".`);
    } catch (err) {
      handleApiError(err, "Gagal membuat cadangan");
    } finally {
      setBackupBusy(false);
    }
  }

  function triggerPrint() {
    if (!lampiranFilteredEntries.length) {
      notify("Tidak ada data pada filter ini untuk dicetak.", "err");
      return;
    }
    window.print();
  }

  // ---- user management ----
  async function addUser() {
    setUserError("");
    if (!newUser.username.trim() || !newUser.password) {
      setUserError("Username dan password wajib diisi.");
      return;
    }
    setBusy(true);
    try {
      await api.addUser(newUser.username.trim(), newUser.password, newUser.role);
      setUsers((prev) => [...prev, { username: newUser.username.trim(), role: newUser.role }]);
      setNewUser({ username: "", password: "", role: "input" });
      notify("Akun baru dibuat.");
    } catch (err) {
      setUserError(err.message || "Gagal membuat akun.");
    } finally {
      setBusy(false);
    }
  }

  async function changeUserRole(username, newRole) {
    try {
      await api.updateUser(username, { role: newRole });
      setUsers((prev) => prev.map((u) => (u.username === username ? { ...u, role: newRole } : u)));
      notify(`Role ${username} diubah jadi ${newRole}.`);
    } catch (err) {
      handleApiError(err, "Gagal mengubah role");
      setUsersLoaded(false); // paksa refresh biar tidak nyasar di UI
    }
  }

  async function resetUserPassword(username) {
    const pw = window.prompt(`Password baru untuk "${username}":`);
    if (!pw) return;
    try {
      await api.updateUser(username, { password: pw });
      notify(`Password ${username} diperbarui.`);
    } catch (err) {
      handleApiError(err, "Gagal mengubah password");
    }
  }

  async function removeUser(username) {
    if (!window.confirm(`Hapus akun "${username}"? Tindakan ini tidak bisa dibatalkan.`)) return;
    try {
      await api.deleteUser(username);
      setUsers((prev) => prev.filter((u) => u.username !== username));
      notify("Akun dihapus.");
    } catch (err) {
      handleApiError(err, "Gagal menghapus akun");
    }
  }

  async function saveNotifyEmail() {
    setConfigSaving(true);
    try {
      await api.setConfig(notifyEmail.trim());
      notify("Pengaturan notifikasi disimpan.");
    } catch (err) {
      handleApiError(err, "Gagal menyimpan pengaturan");
    } finally {
      setConfigSaving(false);
    }
  }

  // ---- onboarding: belum ada URL Apps Script ----
  if (!connected) {
    return (
      <div className="onboard">
        <div className="onboard-card">
          <div className="brand-mark"><Leaf size={18} /></div>
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
          <SettingsPanel apiUrlInput={apiUrlInput} setApiUrlInput={setApiUrlInput} onSave={saveApiUrl} onClose={() => setShowSettings(false)} />
        )}
      </div>
    );
  }

  // ---- login ----
  if (!authUser) {
    return (
      <div className="onboard">
        <div className="onboard-card">
          <div className="brand-mark"><Lock size={18} /></div>
          <h1>Masuk ke Rekap Barang Bekas</h1>
          <p>Gunakan akun yang sudah didaftarkan admin di sheet "Users".</p>
          <input type="text" placeholder="Username" value={loginForm.username}
            onChange={(e) => setLoginForm((f) => ({ ...f, username: e.target.value }))}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
          <input type="password" placeholder="Password" value={loginForm.password}
            onChange={(e) => setLoginForm((f) => ({ ...f, password: e.target.value }))}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
          {loginError && <p className="login-error">{loginError}</p>}
          <button className="btn primary" onClick={handleLogin} disabled={loginBusy}>
            {loginBusy ? <Loader2 size={16} className="spin" /> : <Lock size={16} />}
            Masuk
          </button>
          <button className="link-btn" onClick={() => setShowSettings(true)}>Ubah URL Apps Script</button>
        </div>
        {showSettings && (
          <SettingsPanel apiUrlInput={apiUrlInput} setApiUrlInput={setApiUrlInput} onSave={saveApiUrl} onClose={() => setShowSettings(false)} />
        )}
      </div>
    );
  }

  const activeNav = navItems.find((n) => n.id === tab) || navItems[0];

  return (
    <div className="app-shell">
      <aside className="sidebar no-print">
        <div className="sidebar-brand">
          <div className="brand-mark"><Leaf size={19} strokeWidth={2.2} /></div>
          <div>
            <h1>Rekap Barang Bekas</h1>
            <p>Dinas Perhubungan DIY - Trans Jogja</p>
          </div>
        </div>

        <nav className="side-nav">
          {navItems.map(({ id, label, icon: Icon }) => (
            <button key={id} className={`side-nav-item ${tab === id ? "active" : ""}`} onClick={() => setTab(id)}>
              <Icon size={17} />{label}
            </button>
          ))}
        </nav>

        <div className="side-divider" />

        <div className="side-stats">
          <div className="side-stat">
            <div className="side-stat-icon"><FileSpreadsheet size={15} /></div>
            <div><span>Total Data</span><strong>{entries.length}</strong></div>
          </div>
          <div className="side-stat">
            <div className="side-stat-icon"><Calendar size={15} /></div>
            <div><span>Tanggal Tercatat</span><strong>{tanggalList.length}</strong></div>
          </div>
          <div className="side-stat">
            <div className="side-stat-icon"><ImageIcon size={15} /></div>
            <div><span>Foto Terlampir</span><strong>{totalFoto}</strong></div>
          </div>
        </div>

        <button className="side-settings-btn" onClick={() => setShowSettings(true)}>
          <Settings size={15} />Pengaturan
        </button>

        <div className="side-user">
          <div className="side-user-avatar">{authUser.username.slice(0, 2).toUpperCase()}</div>
          <div className="side-user-info">
            <strong>{authUser.username}</strong>
            <span>{ROLE_LABEL[role] || role}</span>
          </div>
          <button className="side-logout-btn" onClick={handleLogout} title="Keluar"><LogOut size={15} /></button>
        </div>

        <div className="sidebar-footer">
          <Bus size={54} strokeWidth={1.3} />
          <p>Bersama<br />Untuk Transportasi<br />Yang Lebih Baik</p>
        </div>
      </aside>

      <div className="main-area">
        <header className="main-topbar no-print">
          <div className="spacer" />
          <button className="bell-btn" title="Notifikasi"><Bell size={17} /><span className="dot" /></button>
          <button className="admin-chip" onClick={handleLogout} title="Klik untuk keluar">
            <span className="avatar">{authUser.username.slice(0, 2).toUpperCase()}</span>
            {authUser.username}<ChevronDown size={14} />
          </button>
        </header>

        <div className="hero-banner no-print">
          <div className="hero-blobs" aria-hidden="true"><span className="blob b1" /><span className="blob b2" /><span className="blob b3" /></div>
          <p className="hero-kicker">Selamat Datang</p>
          <h2>{activeNav.label}</h2>
          <p className="hero-sub">Dinas Perhubungan DIY - Trans Jogja</p>
        </div>

        {tab === "input" && (
          <main className="page no-print">
            {canCreate ? (
            <section className="card form-card">
              <div className="card-title-row">
                <div className="card-icon gold"><ClipboardList size={16} /></div>
                <div>
                  <h2>Tambah data harian</h2>
                  <p className="card-desc">Silakan lengkapi data berikut untuk menambah rekapan barang bekas.</p>
                </div>
              </div>
              <div className="form-grid">
                <label><span>Tanggal</span>
                  <div className="input-wrap"><Calendar size={15} className="input-icon" />
                    <input type="date" value={draft.tanggal} onChange={(e) => setDraft((d) => ({ ...d, tanggal: e.target.value }))} />
                  </div>
                </label>
                <label><span>LB (no. lambang)</span>
                  <div className="input-wrap"><Tag size={15} className="input-icon" />
                    <input type="text" placeholder="mis. 47" value={draft.lb} onChange={(e) => setDraft((d) => ({ ...d, lb: e.target.value }))} />
                  </div>
                </label>
                <label className="col-2"><span>Nama spare part</span>
                  <div className="input-wrap"><Wrench size={15} className="input-icon" />
                    <input type="text" list="namaPart-options" placeholder="mis. LINER NQR 71" value={draft.namaPart} onChange={(e) => setDraft((d) => ({ ...d, namaPart: e.target.value }))} />
                  </div>
                </label>
                <label><span>Kode barang</span>
                  <div className="input-wrap"><Barcode size={15} className="input-icon" />
                    <input type="text" list="kodeBarang-options" placeholder="mis. NQR0090" value={draft.kodeBarang} onChange={(e) => setDraft((d) => ({ ...d, kodeBarang: e.target.value }))} />
                  </div>
                </label>
                <label><span>Jumlah</span>
                  <div className="input-wrap"><Hash size={15} className="input-icon" />
                    <input type="text" inputMode="decimal" placeholder="mis. 4" value={draft.jumlah} onChange={(e) => setDraft((d) => ({ ...d, jumlah: e.target.value }))} />
                  </div>
                </label>
                <label><span>Satuan</span>
                  <div className="input-wrap"><Box size={15} className="input-icon" />
                    <select value={draft.satuan} onChange={(e) => setDraft((d) => ({ ...d, satuan: e.target.value }))}>
                      {SATUAN_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </label>
                <label><span>Mekanik</span>
                  <div className="input-wrap"><User size={15} className="input-icon" />
                    <input type="text" list="mekanik-options" placeholder="mis. ARIF TRI" value={draft.mekanik} onChange={(e) => setDraft((d) => ({ ...d, mekanik: e.target.value }))} />
                  </div>
                </label>
              </div>
              <datalist id="namaPart-options">{namaPartOptions.map((v) => <option key={v} value={v} />)}</datalist>
              <datalist id="kodeBarang-options">{kodeBarangOptions.map((v) => <option key={v} value={v} />)}</datalist>
              <datalist id="mekanik-options">{mekanikOptions.map((v) => <option key={v} value={v} />)}</datalist>
              <button className="btn primary" onClick={addEntry}><Plus size={16} />Tambah ke rekap</button>
            </section>
            ) : (
            <section className="card role-notice">
              <div className="card-title-row">
                <div className="card-icon"><Eye size={16} /></div>
                <div>
                  <h2>Mode lihat saja</h2>
                  <p className="card-desc">Akun guest hanya bisa melihat rekap, tidak bisa menambah, mengedit, atau menghapus data.</p>
                </div>
              </div>
            </section>
            )}

            <section className="card">
              <div className="card-head wrap">
                <div className="card-title-row">
                  <div className="card-icon"><ClipboardList size={16} /></div>
                  <div>
                    <h2>Rekap tersimpan</h2>
                    <p className="card-desc">Data yang sudah tersimpan akan muncul di sini. Seret garis di sisi kanan judul kolom untuk melebar/mengecilkan lebarnya.</p>
                  </div>
                </div>
                <div className="table-toolbar">
                  <div className="search-wrap">
                    <Search size={14} className="search-icon" />
                    <input type="text" placeholder="Cari nama part, kode, LB, atau mekanik…" value={search} onChange={(e) => setSearch(e.target.value)} />
                    {search && <button className="search-clear" onClick={() => setSearch("")} title="Bersihkan pencarian"><X size={13} /></button>}
                  </div>
                  <div className="date-range">
                    <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} title="Dari tanggal" />
                    <span>–</span>
                    <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} title="Sampai tanggal" />
                    {(dateFrom || dateTo) && <button className="search-clear" onClick={() => { setDateFrom(""); setDateTo(""); }} title="Bersihkan filter tanggal"><X size={13} /></button>}
                  </div>
                  <span className="count-pill">{filteredEntries.length} / {entries.length} baris</span>
                  <button className="btn ghost small" onClick={exportData} title="Ekspor data sesuai filter saat ini">
                    <Download size={14} />Ekspor
                  </button>
                </div>
              </div>
              {!entries.length ? (
                <EmptyState text="Belum ada data. Tambahkan lewat form di atas atau impor file rekap." />
              ) : !filteredEntries.length ? (
                <EmptyState text={`Tidak ada data yang cocok dengan filter saat ini.`} />
              ) : (
                <div className="table-wrap">
                  <table className="resizable-table" style={{ tableLayout: "fixed" }}>
                    <colgroup>
                      <col style={{ width: colWidths.tanggal }} />
                      <col style={{ width: colWidths.lb }} />
                      <col style={{ width: colWidths.namaPart }} />
                      <col style={{ width: colWidths.kode }} />
                      <col style={{ width: colWidths.jumlah }} />
                      <col style={{ width: colWidths.satuan }} />
                      <col style={{ width: colWidths.mekanik }} />
                      <col style={{ width: colWidths.foto }} />
                      {canEdit && <col style={{ width: 50 }} />}
                    </colgroup>
                    <thead>
                      <tr>
                        <th>Tanggal<span className="col-resizer" onMouseDown={(e) => startColResize(e, "tanggal")} /></th>
                        <th>LB<span className="col-resizer" onMouseDown={(e) => startColResize(e, "lb")} /></th>
                        <th>Nama spare part<span className="col-resizer" onMouseDown={(e) => startColResize(e, "namaPart")} /></th>
                        <th>Kode<span className="col-resizer" onMouseDown={(e) => startColResize(e, "kode")} /></th>
                        <th>Jml<span className="col-resizer" onMouseDown={(e) => startColResize(e, "jumlah")} /></th>
                        <th>Satuan<span className="col-resizer" onMouseDown={(e) => startColResize(e, "satuan")} /></th>
                        <th>Mekanik<span className="col-resizer" onMouseDown={(e) => startColResize(e, "mekanik")} /></th>
                        <th>Foto<span className="col-resizer" onMouseDown={(e) => startColResize(e, "foto")} /></th>
                        {canEdit && <th></th>}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredEntries.map((e) => (
                        <tr key={e.id}>
                          <td>
                            {canEdit ? (
                              <input type="date" value={e.tanggal}
                                onChange={(ev) => updateFieldLocal(e.id, "tanggal", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "tanggal", ev.target.value)} />
                            ) : formatTanggalID(e.tanggal)}
                          </td>
                          <td className="mono">
                            {canEdit ? (
                              <input className="cell-input narrow" value={e.lb}
                                onChange={(ev) => updateFieldLocal(e.id, "lb", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "lb", ev.target.value)} />
                            ) : e.lb}
                          </td>
                          <td className="truncate-cell">
                            {canEdit ? (
                              <input className="cell-input" title={e.namaPart} value={e.namaPart}
                                onChange={(ev) => updateFieldLocal(e.id, "namaPart", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "namaPart", ev.target.value)} />
                            ) : <span title={e.namaPart}>{e.namaPart}</span>}
                          </td>
                          <td className="mono">
                            {canEdit ? (
                              <input className="cell-input" value={e.kodeBarang}
                                onChange={(ev) => updateFieldLocal(e.id, "kodeBarang", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "kodeBarang", ev.target.value)} />
                            ) : e.kodeBarang}
                          </td>
                          <td>
                            {canEdit ? (
                              <input className="cell-input narrow" value={e.jumlah}
                                onChange={(ev) => updateFieldLocal(e.id, "jumlah", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "jumlah", ev.target.value)} />
                            ) : e.jumlah}
                          </td>
                          <td>
                            {canEdit ? (
                              <select value={e.satuan}
                                onChange={(ev) => { updateFieldLocal(e.id, "satuan", ev.target.value); commitField(e.id, "satuan", ev.target.value); }}>
                                {SATUAN_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                              </select>
                            ) : e.satuan}
                          </td>
                          <td>
                            {canEdit ? (
                              <input className="cell-input" value={e.mekanik}
                                onChange={(ev) => updateFieldLocal(e.id, "mekanik", ev.target.value)}
                                onBlur={(ev) => commitField(e.id, "mekanik", ev.target.value)} />
                            ) : e.mekanik}
                          </td>
                          <td>
                            {isBanEntry(e) ? (
                              <button className="btn ghost tiny" onClick={() => setBanModalEntry(e)}>
                                <ImageIcon size={13} />
                                {countBanPhotos(e, getBanKind(e))}/{banSlotsFor(getBanKind(e)).length} foto
                              </button>
                            ) : (
                              <div className="photo-slots">
                                <PhotoSlot label="Kondisi" value={e.fotoKondisi} readOnly={!canCreate}
                                  uploading={isUploading(e.id, "kondisi")}
                                  onPreview={() => openLightbox(e.id, "kondisi", e.fotoKondisi, `${e.namaPart} — Kondisi`, canEdit)}
                                  onPick={(f) => attachPhoto(e.id, "kondisi", f)}
                                  onRemove={canEdit ? () => removePhoto(e.id, "kondisi") : undefined} />
                                <PhotoSlot label="Pasang" value={e.fotoPasang} readOnly={!canCreate}
                                  uploading={isUploading(e.id, "pasang")}
                                  onPreview={() => openLightbox(e.id, "pasang", e.fotoPasang, `${e.namaPart} — Pemasangan`, canEdit)}
                                  onPick={(f) => attachPhoto(e.id, "pasang", f)}
                                  onRemove={canEdit ? () => removePhoto(e.id, "pasang") : undefined} />
                              </div>
                            )}
                          </td>
                          {canEdit && (
                            <td>
                              <button className="icon-btn danger" onClick={() => confirmDeleteEntry(e.id, e.namaPart)} title="Hapus baris"><Trash2 size={15} /></button>
                            </td>
                          )}
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
              <p className="card-kicker">Dokumentasi</p>
              <div className="card-head wrap">
                <div className="card-title-row">
                  <div className="card-icon rust"><Images size={16} /></div>
                  <div>
                    <h2>Lampiran foto</h2>
                    <p className="card-desc">Setiap baris rekap otomatis jadi satu blok foto. Part "ban" dengan jumlah=2 dapat 8 foto (depan), jumlah=4 dapat 14 foto (belakang, otomatis terbagi jadi beberapa halaman maksimal 8 foto/halaman).</p>
                  </div>
                </div>
                <div className="lampiran-controls">
                  <div className="search-wrap">
                    <Search size={14} className="search-icon" />
                    <input type="text" placeholder="Cari nama part, kode, LB, mekanik…" value={lampiranSearch} onChange={(e) => setLampiranSearch(e.target.value)} />
                    {lampiranSearch && <button className="search-clear" onClick={() => setLampiranSearch("")} title="Bersihkan pencarian"><X size={13} /></button>}
                  </div>
                  <div className="date-range">
                    <input type="date" value={lampiranDateFrom} onChange={(e) => setLampiranDateFrom(e.target.value)} title="Dari tanggal" />
                    <span>–</span>
                    <input type="date" value={lampiranDateTo} onChange={(e) => setLampiranDateTo(e.target.value)} title="Sampai tanggal" />
                  </div>
                  <button className="btn primary" onClick={triggerPrint}><Printer size={16} />Cetak / simpan PDF</button>
                </div>
              </div>
              <span className="count-pill">{lampiranFilteredEntries.length} data cocok filter</span>
            </section>

            {!lampiranFilteredEntries.length ? (
              <div className="no-print"><EmptyState text="Tidak ada data untuk filter ini. Ubah rentang tanggal atau kata kunci pencarian." /></div>
            ) : (
              lampiranBlocks.map((b, i) =>
                b.type === "normal" ? (
                  <div className={`print-area ${i > 0 ? "force-break" : ""}`} key={`normal-${b.date}`}>
                    <div className="sheet-head">
                      <h3>LAMPIRAN FOTO SPARE PART BEKAS</h3>
                      <p>{formatTanggalID(b.date)}</p>
                    </div>
                    <div className="grid-columns-head">
                      <span>Barang (kondisi)</span>
                      <span>Lampiran (pemasangan)</span>
                    </div>
                    <div className="photo-rows">
                      {b.entries.map((e) => (
                        <div className={`photo-row ${!e.fotoKondisi && !e.fotoPasang ? "print-hide-row" : ""}`} key={e.id}>
                          <PhotoCard entry={e} slot="kondisi" title="(BARU DAN BEKAS)" readOnly={!canCreate}
                            uploading={isUploading(e.id, "kondisi")}
                            onPreview={() => openLightbox(e.id, "kondisi", e.fotoKondisi, `${e.namaPart} — Kondisi`, canEdit)}
                            onPick={(f) => attachPhoto(e.id, "kondisi", f)} />
                          <PhotoCard entry={e} slot="pasang" title="(PENGGANTIAN)" readOnly={!canCreate}
                            uploading={isUploading(e.id, "pasang")}
                            onPreview={() => openLightbox(e.id, "pasang", e.fotoPasang, `${e.namaPart} — Pemasangan`, canEdit)}
                            onPick={(f) => attachPhoto(e.id, "pasang", f)} />
                        </div>
                      ))}
                    </div>
                  </div>
                ) : (
                  <div className={`print-area ban-page ${i > 0 ? "force-break" : ""}`} key={`${b.entry.id}-${b.page}`}>
                    <div className="sheet-head">
                      <h3>
                        PENGGANTIAN {(b.entry.namaPart || "BAN").toUpperCase()}{" "}
                        {b.kind === "belakang" ? "BELAKANG KANAN – BELAKANG KIRI" : "DEPAN KANAN – DEPAN KIRI"}
                      </h3>
                      <p>
                        BUS LAMBUNG {b.entry.lb || "-"} · {formatTanggalID(b.date)}
                        {b.totalPages > 1 ? ` · Halaman ${b.page}/${b.totalPages}` : ""}
                      </p>
                    </div>
                    <div className="ban-page-grid">
                      {b.chunk.map(({ slot, label }) => (
                        <BanPageCell key={slot} label={label} tag={`LB ${b.entry.lb || "-"} · ${label}`}
                          value={getBanPhotoValue(b.entry, slot)} readOnly={!canCreate}
                          uploading={isUploading(b.entry.id, slot)}
                          onPreview={() => openLightbox(b.entry.id, slot, getBanPhotoValue(b.entry, slot), label, canEdit)}
                          onPick={(f) => attachPhoto(b.entry.id, slot, f)} />
                      ))}
                    </div>
                  </div>
                )
              )
            )}
          </main>
        )}

        {tab === "dashboard" && (
          <main className="page no-print">
            <div className="dashboard-grid">
              <section className="card">
                <div className="card-title-row">
                  <div className="card-icon"><BarChart3 size={16} /></div>
                  <div>
                    <h2>Data per bulan</h2>
                    <p className="card-desc">Jumlah spare part bekas tercatat, 6 bulan terakhir.</p>
                  </div>
                </div>
                {chartByMonth.length ? (
                  <div className="chart-box">
                    <ResponsiveContainer width="100%" height={220}>
                      <BarChart data={chartByMonth}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#e1e7f5" vertical={false} />
                        <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#64708c" }} axisLine={{ stroke: "#e1e7f5" }} tickLine={false} />
                        <YAxis allowDecimals={false} tick={{ fontSize: 12, fill: "#64708c" }} axisLine={false} tickLine={false} width={28} />
                        <Tooltip contentStyle={{ borderRadius: 10, border: "1px solid #e1e7f5", fontSize: 12.5 }} cursor={{ fill: "#eef1fb" }} />
                        <Bar dataKey="jumlah" fill="#2f6fed" radius={[6, 6, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                ) : <EmptyState text="Belum ada data untuk ditampilkan." />}
              </section>

              <section className="card">
                <div className="card-title-row">
                  <div className="card-icon gold"><User size={16} /></div>
                  <div>
                    <h2>Mekanik teraktif</h2>
                    <p className="card-desc">5 mekanik dengan jumlah pengerjaan terbanyak.</p>
                  </div>
                </div>
                {topMechanics.length ? (
                  <ul className="leaderboard">
                    {topMechanics.map((m, i) => (
                      <li key={m.name}><span className="rank">{i + 1}</span><span className="name">{m.name}</span><span className="count">{m.jumlah}</span></li>
                    ))}
                  </ul>
                ) : <EmptyState text="Belum ada data untuk ditampilkan." />}
              </section>

              <section className="card">
                <div className="card-title-row">
                  <div className="card-icon rust"><Wrench size={16} /></div>
                  <div>
                    <h2>Spare part paling sering diganti</h2>
                    <p className="card-desc">5 nama part dengan frekuensi penggantian terbanyak.</p>
                  </div>
                </div>
                {topParts.length ? (
                  <ul className="leaderboard">
                    {topParts.map((m, i) => (
                      <li key={m.name}><span className="rank">{i + 1}</span><span className="name">{m.name}</span><span className="count">{m.jumlah}</span></li>
                    ))}
                  </ul>
                ) : <EmptyState text="Belum ada data untuk ditampilkan." />}
              </section>
            </div>
          </main>
        )}

        {tab === "io" && (
          <main className="page no-print">
            {canCreate && (
            <section className="card">
              <div className="card-title-row">
                <div className="card-icon gold"><Upload size={16} /></div>
                <div>
                  <h2>Impor data rekap</h2>
                  <p className="card-desc">Terima file .csv atau .xlsx dengan kolom seperti rekap harian. Setiap baris langsung ditulis ke Google Sheet.</p>
                </div>
              </div>
              <input ref={importRef} type="file" accept=".csv,.xlsx,.xls" className="file-input" id="import-file"
                onChange={(e) => handleImport(e.target.files?.[0])} />
              <label htmlFor="import-file" className="btn ghost"><Upload size={16} />Pilih file rekap</label>
              {importProgress && (
                <div className="import-progress">
                  <div className="import-progress-bar">
                    <div className="import-progress-fill" style={{ width: `${Math.round((importProgress.current / importProgress.total) * 100)}%` }} />
                  </div>
                  <span>Mengimpor {importProgress.current} / {importProgress.total} baris…</span>
                </div>
              )}
            </section>
            )}

            <section className="card">
              <div className="card-title-row">
                <div className="card-icon"><Download size={16} /></div>
                <div>
                  <h2>Ekspor</h2>
                  <p className="card-desc">Mengunduh data sesuai pencarian/rentang tanggal yang sedang aktif di tab Input &amp; Data (kalau tidak ada filter, semua data terunduh). Atau cetak lampiran foto sebagai PDF.</p>
                </div>
              </div>
              <div className="btn-row">
                <button className="btn primary" onClick={exportData}><Download size={16} />Ekspor rekap (.xlsx)</button>
                <button className="btn ghost" onClick={() => setTab("lampiran")}><Printer size={16} />Buka lampiran foto</button>
              </div>
            </section>

            {role === "admin" && (
            <section className="card">
              <div className="card-title-row">
                <div className="card-icon"><Box size={16} /></div>
                <div>
                  <h2>Cadangan data</h2>
                  <p className="card-desc">Menyalin seluruh sheet "Entries" apa adanya ke sheet baru bertanda waktu, tanpa mengubah data yang aktif.</p>
                </div>
              </div>
              <button className="btn ghost" onClick={backupNow} disabled={backupBusy}>
                {backupBusy ? <Loader2 size={16} className="spin" /> : <Download size={16} />}
                Buat cadangan sekarang
              </button>
            </section>
            )}

            {role === "admin" && (
            <section className="card danger-zone">
              <div className="card-title-row">
                <div className="card-icon rust"><RotateCcw size={16} /></div>
                <div>
                  <h2>Reset data</h2>
                  <p className="card-desc">Menghapus seluruh baris rekap di Google Sheet (cadangan otomatis dibuat dulu). Butuh konfirmasi ulang password, tidak bisa dibatalkan setelah dijalankan.</p>
                </div>
              </div>
              {!confirmReset ? (
                <button className="btn ghost danger" onClick={() => setConfirmReset(true)}><RotateCcw size={16} />Reset semua data</button>
              ) : (
                <div className="reset-confirm">
                  <label>
                    <span>Masukkan ulang password admin untuk konfirmasi</span>
                    <input type="password" value={resetPassword}
                      onChange={(e) => { setResetPassword(e.target.value); setResetError(""); }}
                      onKeyDown={(e) => e.key === "Enter" && resetAll()}
                      placeholder="Password" autoFocus />
                  </label>
                  {resetError && <p className="login-error">{resetError}</p>}
                  <div className="btn-row">
                    <button className="btn danger" onClick={resetAll} disabled={!resetPassword}>Ya, hapus semuanya</button>
                    <button className="btn ghost" onClick={() => { setConfirmReset(false); setResetPassword(""); setResetError(""); }}>Batal</button>
                  </div>
                </div>
              )}
            </section>
            )}
          </main>
        )}

        {tab === "activity" && role === "admin" && (
          <main className="page no-print">
            <section className="card">
              <div className="card-head">
                <div className="card-title-row">
                  <div className="card-icon"><History size={16} /></div>
                  <div>
                    <h2>Log aktivitas</h2>
                    <p className="card-desc">Riwayat siapa menambah, mengedit, menghapus, atau login — 500 catatan terakhir.</p>
                  </div>
                </div>
                <button className="btn ghost" onClick={() => { setActivityLoaded(false); }}><RotateCcw size={14} />Muat ulang</button>
              </div>
              {!activityLoaded ? (
                <div className="empty-state"><Loader2 className="spin" size={20} /><p>Memuat log…</p></div>
              ) : !activityLog.length ? (
                <EmptyState text="Belum ada aktivitas tercatat." />
              ) : (
                <div className="table-wrap">
                  <table>
                    <thead><tr><th>Waktu</th><th>Akun</th><th>Role</th><th>Aksi</th><th>Detail</th></tr></thead>
                    <tbody>
                      {activityLog.map((log, i) => (
                        <tr key={i}>
                          <td className="mono">{formatDateTimeID(log.timestamp)}</td>
                          <td>{log.username}</td>
                          <td><span className={`role-badge ${log.role}`}>{ROLE_LABEL[log.role] || log.role}</span></td>
                          <td className="mono">{log.action}</td>
                          <td>{log.detail}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </main>
        )}

        {tab === "users" && role === "admin" && (
          <main className="page no-print">
            <section className="card">
              <div className="card-title-row">
                <div className="card-icon gold"><UserPlus size={16} /></div>
                <div>
                  <h2>Tambah akun</h2>
                  <p className="card-desc">Buat akun baru untuk anggota tim, tentukan perannya.</p>
                </div>
              </div>
              <div className="form-grid user-form-grid">
                <label><span>Username</span>
                  <div className="input-wrap"><User size={15} className="input-icon" />
                    <input type="text" value={newUser.username} onChange={(e) => setNewUser((u) => ({ ...u, username: e.target.value }))} placeholder="mis. budi" />
                  </div>
                </label>
                <label><span>Password</span>
                  <div className="input-wrap"><KeyRound size={15} className="input-icon" />
                    <input type="text" value={newUser.password} onChange={(e) => setNewUser((u) => ({ ...u, password: e.target.value }))} placeholder="Password awal" />
                  </div>
                </label>
                <label><span>Role</span>
                  <div className="input-wrap"><ShieldCheck size={15} className="input-icon" />
                    <select value={newUser.role} onChange={(e) => setNewUser((u) => ({ ...u, role: e.target.value }))}>
                      {ROLE_OPTIONS.map((r) => <option key={r} value={r}>{ROLE_LABEL[r]}</option>)}
                    </select>
                  </div>
                </label>
              </div>
              {userError && <p className="login-error">{userError}</p>}
              <button className="btn primary" onClick={addUser}><UserPlus size={16} />Buat akun</button>
            </section>

            <section className="card">
              <div className="card-head">
                <div className="card-title-row">
                  <div className="card-icon"><Users size={16} /></div>
                  <div>
                    <h2>Daftar akun</h2>
                    <p className="card-desc">Ubah peran, atur ulang password, atau hapus akun.</p>
                  </div>
                </div>
                <span className="count-pill">{users.length} akun</span>
              </div>
              {!usersLoaded ? (
                <div className="empty-state"><Loader2 className="spin" size={20} /><p>Memuat akun…</p></div>
              ) : (
                <div className="table-wrap">
                  <table>
                    <thead><tr><th>Username</th><th>Role</th><th></th></tr></thead>
                    <tbody>
                      {users.map((u) => (
                        <tr key={u.username}>
                          <td>{u.username}{u.username === authUser.username && <span className="you-tag">kamu</span>}</td>
                          <td>
                            <select value={u.role} onChange={(e) => changeUserRole(u.username, e.target.value)}>
                              {ROLE_OPTIONS.map((r) => <option key={r} value={r}>{ROLE_LABEL[r]}</option>)}
                            </select>
                          </td>
                          <td>
                            <div className="btn-row">
                              <button className="icon-btn" onClick={() => resetUserPassword(u.username)} title="Atur ulang password"><KeyRound size={14} /></button>
                              <button className="icon-btn danger" onClick={() => removeUser(u.username)} title="Hapus akun" disabled={u.username === authUser.username}><Trash2 size={14} /></button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>

            <section className="card">
              <div className="card-title-row">
                <div className="card-icon rust"><Mail size={16} /></div>
                <div>
                  <h2>Notifikasi email</h2>
                  <p className="card-desc">Alamat email ini akan menerima pemberitahuan otomatis setiap kali ada yang mereset data.</p>
                </div>
              </div>
              <div className="reset-confirm">
                <label>
                  <span>Alamat email</span>
                  <input type="email" value={notifyEmail} onChange={(e) => setNotifyEmail(e.target.value)} placeholder="admin@contoh.com" />
                </label>
                <button className="btn ghost" onClick={saveNotifyEmail} disabled={configSaving} style={{ alignSelf: "flex-start" }}>
                  {configSaving ? <Loader2 size={15} className="spin" /> : <CheckCircle2 size={15} />}
                  Simpan
                </button>
              </div>
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
          <SettingsPanel apiUrlInput={apiUrlInput} setApiUrlInput={setApiUrlInput} onSave={saveApiUrl} onClose={() => setShowSettings(false)} />
        )}
        {banModalEntry && (
          <BanPhotoModal
            entry={entries.find((en) => en.id === banModalEntry.id) || banModalEntry}
            kind={getBanKind(entries.find((en) => en.id === banModalEntry.id) || banModalEntry)}
            canCreate={canCreate}
            canEdit={canEdit}
            isUploading={(slot) => isUploading(banModalEntry.id, slot)}
            onPreview={(slot, label) => {
              const liveEntry = entries.find((en) => en.id === banModalEntry.id) || banModalEntry;
              openLightbox(banModalEntry.id, slot, getBanPhotoValue(liveEntry, slot), label, canEdit);
            }}
            onPick={(slot, file) => attachPhoto(banModalEntry.id, slot, file)}
            onRemove={(slot) => removePhoto(banModalEntry.id, slot)}
            onClose={() => setBanModalEntry(null)}
          />
        )}
        {lightbox && (
          <PhotoLightbox item={lightbox} canCreate={canCreate} canEdit={canEdit} onClose={() => setLightbox(null)} />
        )}
      </div>
    </div>
  );
}

const BAN_KIND_LABEL = { depan: "Ban Depan (8 foto)", belakang: "Ban Belakang (14 foto)" };

function BanPhotoModal({ entry, kind, canCreate, canEdit, onPick, onRemove, onClose, isUploading, onPreview }) {
  const slots = banSlotsFor(kind);
  return (
    <div className="modal-backdrop no-print" onClick={onClose}>
      <div className="modal-card ban-modal" onClick={(e) => e.stopPropagation()}>
        <h2>Foto ban — LB {entry.lb || "-"} · {entry.namaPart}</h2>
        <p className="card-desc">{BAN_KIND_LABEL[kind] || "Ban"} — terdeteksi dari jumlah = {entry.jumlah || "-"}.</p>
        <div className="ban-modal-grid">
          {slots.map(({ slot, label }) => {
            const value = getBanPhotoValue(entry, slot);
            const inputId = `banmodal-${entry.id}-${slot}`;
            const uploading = isUploading(slot);
            return (
              <div className="ban-modal-cell" key={slot}>
                {uploading ? (
                  <div className="thumb-empty ban-thumb uploading"><Loader2 size={15} className="spin" /></div>
                ) : value ? (
                  <div className="thumb ban-thumb">
                    <img src={value} alt={label} referrerPolicy="no-referrer" onClick={() => onPreview(slot, label)} className="clickable" />
                    {canEdit && <button className="thumb-remove" onClick={() => onRemove(slot)} title="Hapus foto"><X size={11} /></button>}
                  </div>
                ) : canCreate ? (
                  <>
                    <input type="file" accept="image/*" id={inputId} className="file-input" onChange={(e) => onPick(slot, e.target.files?.[0])} />
                    <label htmlFor={inputId} className="thumb-empty ban-thumb" title={`Unggah ${label}`}><Camera size={15} /></label>
                  </>
                ) : (
                  <div className="thumb-empty ban-thumb readonly"><Camera size={15} /></div>
                )}
                <span className="slot-label">{label}</span>
              </div>
            );
          })}
        </div>
        <button className="btn ghost" onClick={onClose}>Tutup</button>
      </div>
    </div>
  );
}

function PhotoLightbox({ item, onClose, canCreate, canEdit }) {
  const inputId = useRef(`lightbox-${uid()}`).current;
  if (!item) return null;
  return (
    <div className="modal-backdrop no-print" onClick={onClose}>
      <div className="lightbox-card" onClick={(e) => e.stopPropagation()}>
        <div className="lightbox-head">
          <span>{item.label}</span>
          <button className="icon-btn" onClick={onClose} title="Tutup"><X size={15} /></button>
        </div>
        <img src={item.url} alt={item.label} referrerPolicy="no-referrer" className="lightbox-img" />
        <div className="btn-row lightbox-actions">
          {canCreate && (
            <>
              <input type="file" accept="image/*" id={inputId} className="file-input"
                onChange={(e) => { const f = e.target.files?.[0]; if (f) item.onReplace(f); e.target.value = ""; }} />
              <label htmlFor={inputId} className="btn ghost"><Upload size={15} />Ganti foto</label>
            </>
          )}
          {canEdit && item.onRemove && (
            <button className="btn ghost danger" onClick={() => { item.onRemove(); onClose(); }}><Trash2 size={15} />Hapus foto</button>
          )}
        </div>
      </div>
    </div>
  );
}

function SettingsPanel({ apiUrlInput, setApiUrlInput, onSave, onClose }) {  return (
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

function PhotoSlot({ label, value, onPick, onRemove, readOnly, uploading, onPreview }) {
  const inputId = useRef(`ph-${uid()}`).current;
  return (
    <div className="photo-slot">
      {uploading ? (
        <div className="thumb-empty uploading"><Loader2 size={13} className="spin" /></div>
      ) : value ? (
        <div className="thumb">
          <img src={value} alt={label} referrerPolicy="no-referrer" onClick={onPreview} className={onPreview ? "clickable" : ""} />
          {onRemove && <button className="thumb-remove" onClick={onRemove} title="Hapus foto"><X size={11} /></button>}
        </div>
      ) : readOnly ? (
        <div className="thumb-empty readonly" title="Belum ada foto"><Camera size={13} /></div>
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

function BanPageCell({ label, tag, value, onPick, readOnly, uploading, onPreview }) {
  const inputId = useRef(`banpg-${uid()}`).current;
  return (
    <figure className={`photo-card ban-cell ${!value ? "is-empty" : ""}`}>
      <div className="photo-frame">
        {uploading ? (
          <div className="photo-missing no-print"><div className="photo-missing-label"><Loader2 size={18} className="spin" /><span>Mengunggah…</span></div></div>
        ) : value ? (
          <img src={value} alt={label} referrerPolicy="no-referrer" onClick={onPreview} className={onPreview ? "clickable" : ""} />
        ) : readOnly ? (
          <div className="photo-missing no-print">
            <div className="photo-missing-label"><Camera size={18} /><span>Belum ada foto</span></div>
          </div>
        ) : (
          <div className="photo-missing no-print">
            <input type="file" accept="image/*" id={inputId} className="file-input" onChange={(e) => onPick(e.target.files?.[0])} />
            <label htmlFor={inputId}><Camera size={18} /><span>Unggah foto</span></label>
          </div>
        )}
        {tag && <span className="photo-tag">{tag}</span>}
      </div>
      <figcaption><strong>{label}</strong></figcaption>
    </figure>
  );
}

function PhotoCard({ entry, slot, title, onPick, readOnly, uploading, onPreview }) {
  const value = slot === "kondisi" ? entry.fotoKondisi : entry.fotoPasang;
  const inputId = useRef(`pc-${uid()}`).current;
  return (
    <figure className={`photo-card ${!value ? "is-empty" : ""}`}>
      <div className="photo-frame">
        {uploading ? (
          <div className="photo-missing no-print"><div className="photo-missing-label"><Loader2 size={18} className="spin" /><span>Mengunggah…</span></div></div>
        ) : value ? (
          <img src={value} alt={entry.namaPart} referrerPolicy="no-referrer" onClick={onPreview} className={onPreview ? "clickable" : ""} />
        ) : readOnly ? (
          <div className="photo-missing no-print">
            <div className="photo-missing-label"><Camera size={18} /><span>Belum ada foto</span></div>
          </div>
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
