/* Schulplaner – Service Worker: offline verfügbar, Updates automatisch, Benachrichtigungen */
const CACHE = 'schulplaner-v2.1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/badge-96.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // Google-Skript nie anfassen
  if (e.request.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('index.html')) {
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put('./index.html', c)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => {
    const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return res; })));
});

/* ---- kleiner Speicher (gleiche Datenbank wie die App) ---- */
function kv(mode, fn) {
  return new Promise((res, rej) => {
    const r = indexedDB.open('schulplaner', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onerror = () => rej(r.error);
    r.onsuccess = () => { const t = r.result.transaction('kv', mode); const q = fn(t.objectStore('kv')); t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error); };
  });
}
const kvGet = k => kv('readonly', s => s.get(k));
const kvSet = (k, v) => kv('readwrite', s => s.put(v, k));

/* ---- Push: das Google-Skript stupst an, wir holen den Text ab und zeigen ihn ---- */
self.addEventListener('push', e => { e.waitUntil(onPush()); });
async function onPush() {
  const opts = { icon: 'icons/icon-192.png', badge: 'icons/badge-96.png', data: { url: './' } };
  let items = null;
  try {
    const conf = await kvGet('conf');
    if (conf && conf.url) {
      const r = await fetch(conf.url + (conf.url.includes('?') ? '&' : '?') + 'action=pending&token=' + encodeURIComponent(conf.token));
      const j = await r.json(); if (j.ok) items = j.items || [];
    }
  } catch (err) {}
  if (!items) return self.registration.showNotification('Schulplaner', Object.assign({ body: 'Du hast eine Erinnerung – tippe zum Öffnen', tag: 'generic' }, opts));
  const shown = (await kvGet('shown').catch(() => null)) || {};
  const fresh = items.filter(i => !shown[i.id]);
  const limit = Date.now() - 3 * 864e5;
  for (const k in shown) if (shown[k] < limit) delete shown[k];
  for (const i of fresh) {
    shown[i.id] = Date.now();
    await self.registration.showNotification(i.title || 'Schulplaner', Object.assign({ body: i.body || '', tag: i.id, timestamp: i.ts || Date.now() }, opts));
  }
  await kvSet('shown', shown).catch(() => {});
}
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    const c = list.find(w => w.url.includes(self.registration.scope));
    return c ? c.focus() : self.clients.openWindow('./');
  }));
});
