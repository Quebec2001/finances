// Code de l'app gardé sur l'appareil (service worker) — SEULEMENT le code et les icônes de ce site, jamais les données :
// OneDrive, Microsoft et le CDN ne passent pas par ici. L'écran s'affiche tout de suite, même avec un réseau faible.
// Nouvelle version publiée : la page est prévenue (« Nouvelle version · Recharger »).
// Pour tout désactiver d'urgence : publier un sw.js qui ne contient que
//   self.addEventListener("install", () => self.skipWaiting()); self.addEventListener("activate", () => self.registration.unregister());
"use strict";
const VERSION = "2cef903c3d", CACHE = "finances-" + VERSION;
const BASE = new URL("./", self.location).pathname;                        // ex. /finances/
const PAGE = BASE;                                                          // la page (index.html) est gardée sous l'adresse du dossier
const exclu = p => p.startsWith(BASE + "v1/") || p.startsWith(BASE + "v2/") || p.endsWith("/sw.js");

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil((async () => {
  for(const k of await caches.keys()) if(k.startsWith("finances-") && k !== CACHE) await caches.delete(k);
  await self.clients.claim();
})()));

async function previent(){
  for(const c of await self.clients.matchAll({type: "window"})) c.postMessage({type: "nouvelle-version"});
}
// la page : celle gardée tout de suite ; en arrière-plan, la version en ligne est relue et gardée pour la prochaine fois
// MARQUE : la page gardée est plus récente que celle affichée (la page le demande à l'ouverture et au retour dans l'app)
const MARQUE = BASE + "__nouvelle-version";
async function page(req){
  const cache = await caches.open(CACHE), garde = await cache.match(PAGE);
  await cache.delete(MARQUE);                                                // ce qui va être affiché = la page gardée
  const vieuxTexte = garde ? garde.clone().text() : Promise.resolve(null);   // lu AVANT d'envoyer la page (le corps ne se lit qu'une fois)
  const enLigne = fetch(PAGE, {cache: "no-store", credentials: "same-origin"}).then(async r => {
    if(!r || !r.ok || r.type !== "basic") return r;
    const neuf = await r.clone().text(), vieux = await vieuxTexte;
    if(vieux !== neuf){ await cache.put(PAGE, r.clone()); if(vieux !== null){ await cache.put(MARQUE, new Response("1")); await previent(); } }
    return r;
  });
  const suite = enLigne.then(() => {}, () => {});                            // le service worker reste actif jusqu'à la fin de la vérification
  if(garde) return {reponse: garde, suite};
  try { return {reponse: await enLigne, suite}; } catch(e){ return {reponse: new Response("Hors ligne", {status: 503, headers: {"Content-Type": "text/plain; charset=utf-8"}}), suite}; }
}
// le code (adresse avec ?v=…) ne change jamais : gardé une fois ; le reste (icônes, manifeste) : gardé, puis rafraîchi
async function fichier(req, versionne){
  const cache = await caches.open(CACHE), garde = await cache.match(req);
  const enLigne = fetch(req).then(r => { if(r && r.ok && r.type === "basic") cache.put(req, r.clone()); return r; });
  if(garde){ if(!versionne) enLigne.catch(() => {}); return garde; }
  return enLigne;
}
self.addEventListener("fetch", e => {
  const req = e.request, u = new URL(req.url);
  if(req.method !== "GET" || u.origin !== self.location.origin || exclu(u.pathname)) return;      // tout le reste : réseau, sans copie
  if(req.mode === "navigate"){ if(u.pathname === BASE || u.pathname === BASE + "index.html"){ const r = page(req); e.respondWith(r.then(x => x.reponse)); e.waitUntil(r.then(x => x.suite)); } return; }
  if(/\.(js|png|webmanifest|py|sql)$/.test(u.pathname)) e.respondWith(fichier(req, u.searchParams.has("v")));
});
self.addEventListener("message", e => { if(e.data && e.data.type === "version?") e.waitUntil((async () => {
  const c = await caches.open(CACHE); if(await c.match(MARQUE) && e.source) e.source.postMessage({type: "nouvelle-version"});
})()); });
