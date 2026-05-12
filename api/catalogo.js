// api/catalogo.js — Vercel Function
// Lee las hojas CAPITAL y STOCK del Excel de OneDrive via Microsoft Graph
// Usa refresh_token guardado en env vars

const CLIENT_ID = process.env.MS_CLIENT_ID;
const REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;
const FILENAME = "tiendapino_excel.xlsx";

async function getAccessToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: REFRESH_TOKEN,
    scope: "Files.Read User.Read offline_access",
  });
  const res = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Origin": "https://tienda-pino-catalogo.vercel.app",
    },
    body,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("No se pudo obtener access token: " + JSON.stringify(data));
  return data.access_token;
}

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

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=120");

  try {
    const token = await getAccessToken();
    const hdrs = { Authorization: `Bearer ${token}` };

    const fileId = await findFile(token);
    if (!fileId) throw new Error("Archivo " + FILENAME + " no encontrado");

    const base = fileId.includes("/drives/")
      ? `https://graph.microsoft.com/v1.0/${fileId}/workbook/worksheets`
      : `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`;

    // ── CAPITAL (productos) ──
    const capitalRows = await readSheet(base, "CAPITAL", hdrs);
    const products = [];
    if (capitalRows) {
      let hIdx = -1, col = 0;
      for (let i = 0; i < capitalRows.length; i++) {
        const ci = capitalRows[i].findIndex(c => String(c).toLowerCase() === "fecha");
        if (ci >= 0) { hIdx = i; col = ci; break; }
      }
      const dataRows = hIdx >= 0 ? capitalRows.slice(hIdx + 1) : capitalRows;
      const seen = new Set();
      dataRows.forEach(r => {
        const nombre = String(r[col + 2] || "").trim();
        const costo  = Number(r[col + 3]) || 0;
        const precio = Number(r[col + 4]) || 0;
        if (!nombre || seen.has(nombre.toLowerCase())) return;
        seen.add(nombre.toLowerCase());
        products.push({ name: nombre, cost: costo, price: precio });
      });
    }

    // ── STOCK ──
    const stockRows = await readSheet(base, "STOCK", hdrs);
    const stock = [];
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
      });
    }

    res.status(200).json({ products, stock, ok: true });

  } catch (e) {
    console.error("Error API catalogo:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
}
