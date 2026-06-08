// api/catalogo.js — Vercel Function
// Lee productos y stock del Excel via Microsoft Graph
// Si falla, devuelve la última copia guardada en Blob (fallback)

import { put, list } from '@vercel/blob';

const CLIENT_ID = process.env.MS_CLIENT_ID;
const INITIAL_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;
const FILENAME = "tiendapino_excel.xlsx";
const TOKEN_BLOB = 'tiendapino-ms-token.json';
const CACHE_BLOB = 'tiendapino-catalog-cache.json';

// Precios de venta hardcodeados
const PRECIOS_VENTA = {
  "musculosa a": 25000,
  "musculosa b": 25000,
  "gorras c": 17000,
  "gorras n": 17000,
  "stickers": 500,
};

// ── Token management ──
async function getCurrentRefreshToken() {
  try {
    const { blobs } = await list({ prefix: TOKEN_BLOB });
    const blob = blobs.find(b => b.pathname === TOKEN_BLOB);
    if (blob) {
      const r = await fetch(blob.downloadUrl || blob.url);
      if (r.ok) {
        const data = await r.json();
        if (data.refresh_token) return data.refresh_token;
      }
    }
  } catch (e) { console.warn("No hay token en Blob:", e.message); }
  return INITIAL_REFRESH_TOKEN;
}

async function saveRefreshToken(token) {
  await put(TOKEN_BLOB, JSON.stringify({ refresh_token: token, updated: new Date().toISOString() }), {
    access: 'public',
    contentType: 'application/json',
    allowOverwrite: true,
    addRandomSuffix: false,
  });
}

async function getAccessToken() {
  const refreshToken = await getCurrentRefreshToken();
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "Files.Read User.Read offline_access",
  });
  const res = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Origin": "https://catalogotiendapino.vercel.app",
    },
    body,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("No se pudo obtener access token: " + JSON.stringify(data));
  if (data.refresh_token) {
    try { await saveRefreshToken(data.refresh_token); }
    catch(e) { console.warn("No se pudo guardar token:", e.message); }
  }
  return data.access_token;
}

// ── Excel reading ──
async function findFile(token) {
  const hdrs = { Authorization: `Bearer ${token}` };
  const res = await fetch("https://graph.microsoft.com/v1.0/me/drive/root/children", { headers: hdrs });
  const data = await res.json();
  const f = data.value?.find(i => i.name === FILENAME);
  if (f) return f.id;
  const shared = await fetch("https://graph.microsoft.com/v1.0/me/drive/sharedWithMe", { headers: hdrs });
  const sData = await shared.json();
  const sf = sData.value?.find(i => i.name === FILENAME);
  if (sf) return sf.remoteItem?.id
    ? `drives/${sf.remoteItem.parentReference?.driveId}/items/${sf.remoteItem.id}`
    : sf.id;
  return null;
}

async function readSheet(base, sheetName, hdrs) {
  const res = await fetch(`${base}/${encodeURIComponent(sheetName)}/usedRange`, { headers: hdrs });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values || null;
}

// ── Catalog cache (fallback si falla la conexión con Excel) ──
async function saveCatalogCache(data) {
  try {
    await put(CACHE_BLOB, JSON.stringify({ ...data, updated: new Date().toISOString() }), {
      access: 'public',
      contentType: 'application/json',
      allowOverwrite: true,
      addRandomSuffix: false,
    });
  } catch(e) { console.warn("No se pudo guardar cache:", e.message); }
}

async function loadCatalogCache() {
  try {
    const { blobs } = await list({ prefix: CACHE_BLOB });
    const blob = blobs.find(b => b.pathname === CACHE_BLOB);
    if (!blob) return null;
    const r = await fetch(blob.downloadUrl || blob.url);
    if (!r.ok) return null;
    return await r.json();
  } catch(e) { return null; }
}

// ── Main handler ──
async function fetchFromExcel() {
  const token = await getAccessToken();
  const hdrs = { Authorization: `Bearer ${token}` };

  const fileId = await findFile(token);
  if (!fileId) throw new Error("Archivo " + FILENAME + " no encontrado");

  const base = fileId.includes("/drives/")
    ? `https://graph.microsoft.com/v1.0/${fileId}/workbook/worksheets`
    : `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`;

  // STOCK
  const stockRows = await readSheet(base, "STOCK", hdrs);
  const stock = [];
  const stockNamesOrdered = [];
  const seenNames = new Set();

  if (stockRows) {
    let hIdx = -1, col = 0;
    for (let i = 0; i < stockRows.length; i++) {
      const ci = stockRows[i].findIndex(c => String(c).toLowerCase() === "producto");
      if (ci >= 0) { hIdx = i; col = ci; break; }
    }
    const dataRows = hIdx >= 0 ? stockRows.slice(hIdx + 1) : stockRows;
    dataRows.forEach(r => {
      const producto = String(r[col]     || "").trim();
      const talle    = String(r[col + 1] || "").trim();
      const si       = Number(r[col + 2]) || 0;
      const vendidos = Number(r[col + 3]) || 0;
      if (!producto || !talle) return;
      stock.push({ producto, talle, stock_inicio: si, vendidos });
      const lower = producto.toLowerCase();
      if (!seenNames.has(lower)) {
        seenNames.add(lower);
        stockNamesOrdered.push({ name: producto, lower });
      }
    });
  }

  // CAPITAL
  const capitalRows = await readSheet(base, "CAPITAL", hdrs);
  const costMap = {};
  if (capitalRows) {
    let hIdx = -1, col = 0;
    for (let i = 0; i < capitalRows.length; i++) {
      const ci = capitalRows[i].findIndex(c => String(c).toLowerCase() === "fecha");
      if (ci >= 0) { hIdx = i; col = ci; break; }
    }
    const dataRows = hIdx >= 0 ? capitalRows.slice(hIdx + 1) : capitalRows;
    dataRows.forEach(r => {
      const nombre = String(r[col + 2] || "").trim();
      const costo  = Number(r[col + 3]) || 0;
      if (!nombre) return;
      const key = nombre.toLowerCase();
      if (costo && !costMap[key]) costMap[key] = costo;
    });
  }

  const products = stockNamesOrdered.map(({ name, lower }) => {
    let cost = costMap[lower] || 0;
    if (!cost) {
      for (const k of Object.keys(costMap)) {
        if (k.includes(lower) || lower.includes(k)) { cost = costMap[k]; break; }
      }
    }
    let price = PRECIOS_VENTA[lower] || 0;
    if (!price) {
      for (const k of Object.keys(PRECIOS_VENTA)) {
        if (k.includes(lower) || lower.includes(k)) { price = PRECIOS_VENTA[k]; break; }
      }
    }
    return { name, cost, price };
  });

  return { products, stock };
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=120");

  try {
    // Intento leer del Excel
    const data = await fetchFromExcel();
    // Si funcionó, guardo copia en cache
    await saveCatalogCache(data);
    res.status(200).json({ ...data, ok: true, source: "excel" });
  } catch (excelErr) {
    console.error("Excel falló, intentando cache:", excelErr.message);
    // Si falla, intento devolver la última copia guardada
    const cached = await loadCatalogCache();
    if (cached) {
      res.status(200).json({
        products: cached.products || [],
        stock: cached.stock || [],
        ok: true,
        source: "cache",
        cachedAt: cached.updated,
      });
      return;
    }
    // Si no hay cache, devuelvo error
    res.status(500).json({ ok: false, error: excelErr.message });
  }
}
