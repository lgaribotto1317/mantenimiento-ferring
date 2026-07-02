// ═══════════════════════════════════════════════════════════════════
// Service Worker — Reporte Diario de Mantenimiento (PR-1)
// ═══════════════════════════════════════════════════════════════════
// Estrategia:
//   - Shell (HTML/JS/CSS del mismo origen): NETWORK-FIRST. Siempre intenta
//     traer la versión nueva; solo cae al cache si NO hay red. Nunca deja
//     bundle viejo pegado → respeta el gate min_version del lado de la app.
//   - Supabase (datos): NETWORK-ONLY. Jamás se cachea un dato del reporte.
//   - Cache: solo fallback offline del shell (se llena on-the-fly).
//
// skipWaiting + clients.claim: el SW nuevo toma control apenas se instala,
// así un deploy nuevo se sirve en la próxima carga/reapertura de la app.

const CACHE = 'rdm-shell-v1';
const SUPABASE_HOST = 'qdcmrirwkfesaqxvhxgh.supabase.co';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Solo GET. POST/PATCH/DELETE (escrituras a Supabase) pasan directo, sin tocar.
  if (req.method !== 'GET') return;

  let url;
  try {
    url = new URL(req.url);
  } catch {
    return;
  }

  // Datos de Supabase → NETWORK-ONLY: no intervenir, no cachear jamás.
  if (url.hostname === SUPABASE_HOST) return;

  // Shell (mismo origen) → NETWORK-FIRST con fallback a cache.
  if (url.origin === self.location.origin) {
    event.respondWith((async () => {
      try {
        const fresh = await fetch(req);
        const cache = await caches.open(CACHE);
        cache.put(req, fresh.clone());
        return fresh;
      } catch (err) {
        const cached = await caches.match(req);
        if (cached) return cached;
        if (req.mode === 'navigate') {
          const root = await caches.match('/');
          if (root) return root;
        }
        throw err;
      }
    })());
  }
});
