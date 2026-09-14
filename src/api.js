// Client tipis untuk Apps Script Web App yang jadi backend + database (Google Sheet).
// URL Web App disimpan di localStorage supaya bisa diganti dari UI tanpa rebuild.

const STORAGE_KEY = "rekap-apps-script-url";

export function getApiUrl() {
  return localStorage.getItem(STORAGE_KEY) || "";
}

export function setApiUrl(url) {
  localStorage.setItem(STORAGE_KEY, url.trim());
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
  const res = await fetch(getApiUrl(), {
    method: "POST",
    // text/plain menghindari CORS preflight yang tidak didukung Apps Script
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Permintaan gagal");
  return data;
}

export const api = {
  listEntries: async () => (await get("list")).entries,
  addEntry: async (fields) => (await post({ action: "add", ...fields })).id,
  updateEntry: async (id, fields) => post({ action: "update", id, fields }),
  deleteEntry: async (id) => post({ action: "delete", id }),
  uploadPhoto: async (id, slot, dataUrl) =>
    (await post({ action: "uploadPhoto", id, slot, dataUrl })).url,
  removePhoto: async (id, slot) => post({ action: "removePhoto", id, slot }),
};
