// api/save-token.js
// Guarda un refresh token manualmente en Vercel Blob
// Llamado desde gettoken.html cuando el usuario pega el token a mano

import { put } from '@vercel/blob';

const TOKEN_BLOB = 'tiendapino-ms-token.json';

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");

  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Método no permitido" });
  }

  try {
    const { refresh_token } = req.body;

    if (!refresh_token || typeof refresh_token !== "string" || refresh_token.trim().length < 20) {
      return res.status(400).json({ ok: false, error: "Token inválido o vacío" });
    }

    await put(
      TOKEN_BLOB,
      JSON.stringify({
        refresh_token: refresh_token.trim(),
        updated: new Date().toISOString(),
        source: "manual",
      }),
      {
        access: 'public',
        contentType: 'application/json',
        allowOverwrite: true,
        addRandomSuffix: false,
      }
    );

    return res.status(200).json({ ok: true, updated: new Date().toISOString() });

  } catch (e) {
    console.error("save-token error:", e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
