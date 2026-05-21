// api/refresh-token.js — Vercel Cron Job
// Se ejecuta automáticamente cada 3 días para refrescar el token de Microsoft
// Guarda el nuevo refresh_token en Vercel Blob (porque Microsoft lo rota cada uso)

import { put, list } from '@vercel/blob';

const CLIENT_ID = process.env.MS_CLIENT_ID;
const INITIAL_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;
const TOKEN_BLOB = 'tiendapino-ms-token.json';

async function getCurrentRefreshToken() {
  // Buscar el token guardado en Blob (es el más reciente)
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
  } catch (e) { console.warn("No hay token en Blob, uso el de env:", e.message); }
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

export default async function handler(req, res) {
  const authHeader = req.headers.authorization;
  if (process.env.CRON_SECRET && authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ ok: false, error: 'Unauthorized' });
  }

  try {
    const refreshToken = await getCurrentRefreshToken();
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: "Files.Read User.Read offline_access",
    });

    const r = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": "https://catalogotiendapino.vercel.app",
      },
      body,
    });

    const data = await r.json();

    if (!data.access_token) {
      console.error("❌ Refresh token falló:", data);
      return res.status(500).json({
        ok: false,
        error: data.error_description || "No se pudo refrescar el token",
        timestamp: new Date().toISOString(),
      });
    }

    // ✅ IMPORTANTE: guardar el nuevo refresh_token (rotativo)
    if (data.refresh_token) {
      await saveRefreshToken(data.refresh_token);
      console.log("✓ Token guardado en Blob:", new Date().toISOString());
    }

    return res.status(200).json({
      ok: true,
      message: "Token refrescado y guardado correctamente",
      expires_in: data.expires_in,
      timestamp: new Date().toISOString(),
    });

  } catch (e) {
    console.error("Error en cron:", e);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
