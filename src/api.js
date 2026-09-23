// Client tipis untuk Apps Script Web App yang jadi backend + database (Google Sheet).
// URL Web App disimpan di localStorage supaya bisa diganti dari UI tanpa rebuild.

const STORAGE_KEY = "rekap-apps-script-url";
const AUTH_KEY = "rekap-auth-user";

// Tempel URL Web App Apps Script kamu (yang diakhiri /exec) di bawah ini.
// Setelah diisi dan di-deploy, semua orang yang membuka situs ini otomatis
// tersambung ke Google Sheet — tidak perlu isi URL manual lagi.
const DEFAULT_API_URL = "PASTE_URL_APPS_SCRIPT_DI_SINI";

export function getApiUrl() {
  return localStorage.getItem(STORAGE_KEY) || DEFAULT_API_URL;
}

export function setApiUrl(url) {
  localStorage.setItem(STORAGE_KEY, url.trim());
}

export function getAuthUser() {
  try {
    const raw = localStorage.getItem(AUTH_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {
    return null;
  }
}

export function setAuthUser(user) {
  localStorage.setItem(AUTH_KEY, JSON.stringify(user));
}

export function clearAuthUser() {
  localStorage.removeItem(AUTH_KEY);
}

async function get(action, params = {}) {
  const url = new URL(getApiUrl());
  url.searchParams.set("action", action);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Permintaan gagal");
  return data;
}

async function post(body) {
  const auth = getAuthUser();
  const payload = auth && body.action !== "login" ? { ...body, token: auth.token } : body;
  const res = await fetch(getApiUrl(), {
    method: "POST",
    // text/plain menghindari CORS preflight yang tidak didukung Apps Script
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Permintaan gagal");
  return data;
}

export const api = {
  login: async (username, password) => post({ action: "login", username, password }),
  listEntries: async () => (await get("list")).entries,
  addEntry: async (fields) => (await post({ action: "add", ...fields })).id,
  addEntriesBatch: async (rows) => (await post({ action: "addBatch", rows })).ids,
  updateEntry: async (id, fields) => post({ action: "update", id, fields }),
  deleteEntry: async (id) => post({ action: "delete", id }),
  uploadPhoto: async (id, slot, dataUrl) =>
    (await post({ action: "uploadPhoto", id, slot, dataUrl })).url,
  removePhoto: async (id, slot) => post({ action: "removePhoto", id, slot }),
  backupNow: async () => (await post({ action: "backup" })).backupSheet,
  resetAll: async (password) => (await post({ action: "resetAll", password })).backupSheet,
  listActivity: async () => (await post({ action: "listActivity" })).logs,
  listUsers: async () => (await post({ action: "listUsers" })).users,
  addUser: async (username, password, role) => post({ action: "addUser", username, password, role }),
  updateUser: async (username, fields) => post({ action: "updateUser", username, fields }),
  deleteUser: async (username) => post({ action: "deleteUser", username }),
  getConfig: async () => (await post({ action: "getConfig" })).config,
  setConfig: async (notifyEmail) => post({ action: "setConfig", notifyEmail }),
};
