// api/refresh-token.js — Vercel Cron Job
// Se ejecuta automáticamente cada 3 días para refrescar el token de Microsoft
// y mantenerlo siempre vigente (el token vence a los 90 días sin uso)

const CLIENT_ID = process.env.MS_CLIENT_ID;
const REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;

export default async function handler(req, res) {
  // Solo permitir ejecución desde el cron de Vercel
  // Vercel manda el header authorization con CRON_SECRET
  const authHeader = req.headers.authorization;
  if (process.env.CRON_SECRET && authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ ok: false, error: 'Unauthorized' });
  }

  try {
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      grant_type: "refresh_token",
      refresh_token: REFRESH_TOKEN,
      scope: "Files.Read User.Read offline_access",
    });

    const r = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": "https://tienda-pino-catalogo.vercel.app",
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

    // ✅ Token refrescado correctamente
    console.log("✓ Token refrescado:", new Date().toISOString());

    return res.status(200).json({
      ok: true,
      message: "Token refrescado correctamente",
      expires_in: data.expires_in,
      timestamp: new Date().toISOString(),
    });

  } catch (e) {
    console.error("Error en cron:", e);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
