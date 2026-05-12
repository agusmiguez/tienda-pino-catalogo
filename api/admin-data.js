// api/admin-data.js — Vercel Function
// GET: lee las ediciones del admin desde Vercel Blob
// POST: guarda las ediciones del admin en Vercel Blob (requiere ADMIN_PASS)

import { put, list } from '@vercel/blob';

const BLOB_FILENAME = 'tiendapino-admin-data.json';
const ADMIN_PASS = process.env.ADMIN_PASS || 'pino2026';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Admin-Pass');

  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  try {
    if (req.method === 'GET') {
      const { blobs } = await list({ prefix: BLOB_FILENAME });
      const blob = blobs.find(b => b.pathname === BLOB_FILENAME);
      if (!blob) {
        res.setHeader('Cache-Control', 'no-store');
        res.status(200).json({ ok: true, data: {} });
        return;
      }
      const r = await fetch(blob.downloadUrl || blob.url);
      if (!r.ok) {
        res.setHeader('Cache-Control', 'no-store');
        res.status(200).json({ ok: true, data: {} });
        return;
      }
      const data = await r.json();
      res.setHeader('Cache-Control', 'no-store');
      res.status(200).json({ ok: true, data });
      return;
    }

    if (req.method === 'POST') {
      const pass = req.headers['x-admin-pass'];
      if (pass !== ADMIN_PASS) {
        res.status(401).json({ ok: false, error: 'Contraseña incorrecta' });
        return;
      }
      const body = req.body;
      if (typeof body !== 'object' || body === null) {
        res.status(400).json({ ok: false, error: 'Body inválido' });
        return;
      }
      await put(BLOB_FILENAME, JSON.stringify(body), {
        access: 'public',
        contentType: 'application/json',
        allowOverwrite: true,
        addRandomSuffix: false,
      });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ ok: false, error: 'Método no permitido' });
  } catch (e) {
    console.error('Error admin-data:', e);
    res.status(500).json({ ok: false, error: e.message });
  }
}
