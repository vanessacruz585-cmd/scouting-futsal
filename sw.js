// Guarda a aplicação no aparelho para abrir sem internet.
// Com rede: abre sempre a versão mais recente do site (e atualiza a cópia guardada).
// Sem rede (ou rede muito lenta): abre a cópia guardada.
const CACHE = 'scouting-v4';
const FILES = ['./scouting.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png', './xlsx-export.js', './vendor/exceljs.min.js'];
const TIMEOUT_MS = 4000;

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Pede ao servidor sem usar a cache HTTP do browser.
const fresh = req => fetch(req, { cache: 'no-store' });

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin || url.searchParams.has('fresh')) return;
  // Páginas e o módulo do Excel: rede primeiro (versão mais recente); o resto: cópia guardada primeiro.
  const isPage = req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html') || url.pathname.endsWith('/xlsx-export.js');

  if (isPage) {
    // Rede primeiro, com limite de tempo; se falhar, cópia guardada.
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      const network = fresh(req).then(res => {
        if (res && res.ok) cache.put(req, res.clone());
        return res;
      });
      const timeout = new Promise(r => setTimeout(() => r(null), TIMEOUT_MS));
      try {
        const res = await Promise.race([network, timeout]);
        if (res && res.ok) return res;
      } catch (err) { /* sem rede */ }
      const cached = await cache.match(req, { ignoreSearch: true }) || (req.mode === 'navigate' ? await cache.match('./scouting.html') : undefined);
      if (cached) return cached;
      try { return await network; } catch (err) {
        return new Response('Sem ligação à internet.', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
      }
    })());
    return;
  }

  // Ícones e manifesto: cópia guardada primeiro, atualizada em segundo plano.
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(req, { ignoreSearch: true });
    const network = fresh(req).then(res => { if (res && res.ok) cache.put(req, res.clone()); return res; }).catch(() => null);
    if (cached) { e.waitUntil(network); return cached; }
    return (await network) || new Response('', { status: 504 });
  })());
});
