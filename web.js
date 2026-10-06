// web.js — version en ligne de l'app (GitHub Pages)
// Connexion Microsoft (OAuth PKCE, comptes personnels), fichier Excel dans OneDrive,
// moteur Python exécuté dans le navigateur (Pyodide). Aucune donnée ne passe par un serveur tiers.
(function(){
"use strict";
const VERSION = "86135c4420";
const CLIENT_ID = "c7ebaba5-e820-450d-92c0-7efd4f7517c8";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPES = "Files.ReadWrite offline_access User.Read";
const G = "https://graph.microsoft.com/v1.0";
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const CONFIG = {premier_mois_modifiable: "Octobre 2026", alias_investissements: {"Fonds communs de placement": "CÉLI"}};
const CHEMIN = "/data/finances.xlsx";
const MODULES = ["formules.py", "xlsx.py", "modele.py", "releve.py", "routes.py", "pont.py"];
const LECTURE_SEULE = ["/api/releve/analyse"];   // POST sans écriture dans le fichier
const REDIRECT = location.origin + location.pathname.replace(/[^/]*$/, "");
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const LS = {
  get(k){ try { return localStorage.getItem(k); } catch(e){ return null; } },
  set(k, v){ try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch(e){} }
};
const esc = t => String(t ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
let py = null, fichier = null, compte = null, derniereVerif = 0, file = Promise.resolve();

// ------------------------------------------------------------------ écran de chargement
function ecran(html){
  let e = document.getElementById("webEcran");
  if(!e){ e = document.createElement("div"); e.id = "webEcran"; document.body.appendChild(e); }
  e.innerHTML = `<div class="we-boite"><img src="icone.png" alt="" class="we-logo">${html}</div>`;
  e.style.display = "flex";
  return e;
}
function cacheEcran(){ const e = document.getElementById("webEcran"); if(e) e.style.display = "none"; }
function etape(msg){ ecran(`<div class="we-spin"></div><p>${esc(msg)}</p>`); }
const style = document.createElement("style");
style.textContent = `#webEcran{position:fixed;inset:0;z-index:100;background:var(--bg);display:flex;align-items:center;justify-content:center;padding:24px}
.we-boite{max-width:380px;text-align:center}.we-boite p{color:var(--ink-2);margin:12px 0}.we-boite h1{font-size:22px;margin:8px 0}
.we-logo{width:72px;height:72px;border-radius:18px}.we-spin{width:28px;height:28px;margin:18px auto 0;border:3px solid var(--line);border-top-color:var(--accent);border-radius:50%;animation:we 0.8s linear infinite}
@keyframes we{to{transform:rotate(360deg)}}.we-liste{display:grid;gap:8px;margin-top:12px;text-align:left}.we-liste button{text-align:left}`;
document.head.appendChild(style);

// ------------------------------------------------------------------ connexion Microsoft
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
function aleatoire(n){ const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a); }

async function versConnexion(choisirCompte){
  const verifier = aleatoire(48), state = "web-" + aleatoire(12);
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  LS.set("w_verifier", verifier); LS.set("w_state", state);
  const q = new URLSearchParams({client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT, response_mode: "query",
    scope: SCOPES, state, code_challenge: challenge, code_challenge_method: "S256"});
  if(choisirCompte) q.set("prompt", "select_account");
  else if(LS.get("w_login")) q.set("login_hint", LS.get("w_login"));
  location.href = AUTH + "/authorize?" + q;
  return new Promise(() => {});   // la page est remplacée
}
async function jeton(corps){
  const r = await fetch(AUTH + "/token", {method: "POST", headers: {"Content-Type": "application/x-www-form-urlencoded"}, body: new URLSearchParams(corps)});
  const j = await r.json().catch(() => ({}));
  if(!r.ok) throw new Error((j.error || r.status) + (j.error_description ? " – " + j.error_description.split("\r")[0] : ""));
  return j;
}
function garde(t){
  LS.set("w_at", t.access_token);
  LS.set("w_exp", String(Date.now() + ((t.expires_in || 3600) - 120) * 1000));
  if(t.refresh_token) LS.set("w_rt", t.refresh_token);
}
async function retourConnexion(){
  const p = new URLSearchParams(location.search);
  if(!(p.get("state") || "").startsWith("web-")) return false;
  history.replaceState(null, "", location.pathname);
  if(p.has("error")){
    if(p.get("error") === "access_denied") return "refus";
    throw new Error("Connexion Microsoft refusée : " + (p.get("error_description") || p.get("error")));
  }
  if(p.get("state") !== LS.get("w_state")) throw new Error("Retour de connexion invalide. Réessaie.");
  garde(await jeton({client_id: CLIENT_ID, grant_type: "authorization_code", code: p.get("code"), redirect_uri: REDIRECT,
    code_verifier: LS.get("w_verifier"), scope: SCOPES}));
  LS.set("w_verifier", null); LS.set("w_state", null);
  return true;
}
async function accessToken(){
  const at = LS.get("w_at");
  if(at && Date.now() < +(LS.get("w_exp") || 0)) return at;
  const rt = LS.get("w_rt");
  if(rt){
    try { const t = await jeton({client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: rt, scope: SCOPES}); garde(t); return t.access_token; }
    catch(e){ LS.set("w_rt", null); }
  }
  etape("Reconnexion à Microsoft…");
  return versConnexion(false);   // la connexion Microsoft est en général déjà ouverte : retour automatique
}
async function graph(chemin, opts = {}, essai = 0){
  const tok = await accessToken();
  const r = await fetch(chemin.startsWith("http") ? chemin : G + chemin,
    Object.assign({}, opts, {headers: Object.assign({Authorization: "Bearer " + tok}, opts.headers || {})}));
  if(r.status === 401 && essai === 0){ LS.set("w_at", null); return graph(chemin, opts, 1); }
  if(!r.ok){
    let msg = ""; try { const j = await r.json(); msg = (j.error && (j.error.message || j.error.code)) || ""; } catch(e){}
    const err = new Error("OneDrive " + r.status + (msg ? " : " + msg : "")); err.status = r.status; throw err;
  }
  return r;
}

// ------------------------------------------------------------------ fichier OneDrive
async function choisitFichier(forcer){
  const id = forcer ? null : LS.get("w_fichier");
  if(id){
    try { return await (await graph(`/me/drive/items/${id}?$select=id,name,eTag,size,lastModifiedDateTime,parentReference`)).json(); }
    catch(e){ if(e.status !== 404) throw e; LS.set("w_fichier", null); }
  }
  etape("Recherche de ton fichier Excel dans OneDrive…");
  const xlsx = l => (l || []).filter(i => /\.xlsx$/i.test(i.name) && !/^~\$/.test(i.name) && !(i.parentReference && /Sauvegardes/i.test(i.parentReference.path || "")));
  // racine puis recherche dans tout le OneDrive (le fichier peut être dans un sous-dossier, ex. Documents)
  let items = xlsx((await (await graph("/me/drive/root/children?$top=200&$select=id,name,eTag,size,lastModifiedDateTime,file,parentReference")).json()).value);
  if(!items.some(i => /financ/i.test(i.name))){
    try { items = xlsx((await (await graph("/me/drive/root/search(q='Finances')?$top=50&$select=id,name,eTag,size,lastModifiedDateTime,file,parentReference")).json()).value).concat(items); } catch(e){}
  }
  let cands = items.filter(i => /financ/i.test(i.name));
  if(!cands.length) cands = items;
  if(!cands.length) throw new Error("Aucun fichier Excel « Finances… » trouvé dans ton OneDrive. Vérifie qu'il y est bien, puis réessaie.");
  let choix = cands[0];
  if(cands.length > 1 || forcer){
    choix = await new Promise(res => {
      const e = ecran(`<h1>Quel fichier utiliser ?</h1><p>Fichiers Excel trouvés dans ton OneDrive :</p><div class="we-liste">${cands.map((c, i) =>
        `<button class="btn" data-i="${i}"><b>${esc(c.name)}</b><br><small class="muted">${esc(((c.parentReference && c.parentReference.path) || "").replace(/^.*root:?/, "") || "/")} · modifié le ${new Date(c.lastModifiedDateTime).toLocaleDateString("fr-CA")}</small></button>`).join("")}</div>`);
      e.querySelectorAll("[data-i]").forEach(b => b.onclick = () => res(cands[+b.dataset.i]));
    });
  }
  LS.set("w_fichier", choix.id);
  return choix;
}
async function telecharge(){
  for(let k = 0; k < 3; k++){
    const avant = await (await graph(`/me/drive/items/${fichier.id}?$select=id,name,eTag,parentReference`)).json();
    const buf = new Uint8Array(await (await graph(`/me/drive/items/${fichier.id}/content`)).arrayBuffer());
    const apres = await (await graph(`/me/drive/items/${fichier.id}?$select=id,name,eTag,parentReference`)).json();
    if(avant.eTag === apres.eTag){
      if(!(buf[0] === 0x50 && buf[1] === 0x4B)) throw new Error("Le fichier téléchargé n'est pas un classeur Excel valide.");
      fichier.eTag = apres.eTag; fichier.name = apres.name; fichier.parentReference = apres.parentReference;
      py.FS.writeFile(CHEMIN, buf);
      py.globals.set("_cfg", JSON.stringify(CONFIG));
      py.runPython(`pont.ouvre("${CHEMIN}", _cfg)`);
      derniereVerif = Date.now();
      return;
    }
  }
  throw new Error("Le fichier change pendant le téléchargement (il est peut-être en cours de modification). Réessaie dans un instant.");
}
async function aChange(){
  const m = await (await graph(`/me/drive/items/${fichier.id}?$select=id,eTag`)).json();
  derniereVerif = Date.now();
  return m.eTag !== fichier.eTag;
}
async function envoie(){
  const buf = py.FS.readFile(CHEMIN);
  try {
    const r = await graph(`/me/drive/items/${fichier.id}/content`, {method: "PUT", body: buf,
      headers: {"Content-Type": XLSX_TYPE, "If-Match": fichier.eTag}});
    const j = await r.json(); fichier.eTag = j.eTag;
  } catch(e){
    if(e.status === 412 || e.status === 409){ const c = new Error("conflit"); c.conflit = true; throw c; }
    if(e.status === 423) throw new Error("Le fichier est verrouillé (ouvert dans Excel ?). Ferme-le, puis réessaie.");
    throw e;
  }
}
const date2 = n => String(n).padStart(2, "0");
async function sauvegardeDuJour(){
  const d = new Date(), jour = `${d.getFullYear()}-${date2(d.getMonth() + 1)}-${date2(d.getDate())}`;
  const cle = "w_sauv_" + fichier.id;
  if(LS.get(cle) === jour) return;
  try {
    const base = fichier.name.replace(/\.xlsx$/i, "");
    const nom = `${base} – ${jour} ${date2(d.getHours())}h${date2(d.getMinutes())}.xlsx`;
    // dans le dossier « Sauvegardes app finances » situé à côté du fichier Excel (suit le fichier s'il est déplacé)
    const parent = fichier.parentReference && fichier.parentReference.id;
    const base_ = parent ? `/me/drive/items/${parent}:` : "/me/drive/root:";
    await graph(`${base_}/${encodeURIComponent("Sauvegardes app finances")}/${encodeURIComponent(nom)}:/content`,
      {method: "PUT", body: py.FS.readFile(CHEMIN), headers: {"Content-Type": XLSX_TYPE}});
    LS.set(cle, jour);
  } catch(e){ console.warn("Sauvegarde OneDrive impossible", e); }
}

// ------------------------------------------------------------------ moteur Python
function chargeScript(src){
  return new Promise((ok, ko) => { const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = () => ko(new Error("Chargement impossible : " + src)); document.head.appendChild(s); });
}
async function chargePython(){
  etape("Chargement du moteur de calcul… (la première fois peut prendre une dizaine de secondes)");
  await chargeScript(PYODIDE + "pyodide.js");
  py = await loadPyodide({indexURL: PYODIDE});
  py.FS.mkdirTree("/app"); py.FS.mkdirTree("/data");
  for(const n of MODULES){
    const r = await fetch(n + "?v=" + VERSION, {cache: "no-cache"});
    if(!r.ok) throw new Error("Module manquant : " + n);
    py.FS.writeFile("/app/" + n, await r.text());
  }
  py.runPython("import sys\nsys.path.insert(0, '/app')\nimport pont");
}
function appelPy(methode, chemin, corps){
  py.globals.set("_m", methode); py.globals.set("_p", chemin); py.globals.set("_c", corps ? JSON.stringify(corps) : "");
  return JSON.parse(py.runPython("pont.appel(_m, _p, _c)"));
}

// ------------------------------------------------------------------ API utilisée par l'interface
async function api(chemin, corps){
  if(!corps){
    const j = appelPy("GET", chemin);
    if(j && j.erreur) throw new Error(j.erreur);
    if(chemin === "/api/etat"){ j.excel_ouvert = false; j.web = {compte, fichier: fichier && fichier.name}; }
    return j;
  }
  if(LECTURE_SEULE.includes(chemin)){
    const j = appelPy("POST", chemin, corps);
    if(j && j.erreur) throw new Error(j.erreur);
    return j;
  }
  // écritures : une à la fois
  const tache = file.then(() => ecriture(chemin, corps));
  file = tache.catch(() => {});
  return tache;
}
async function ecriture(chemin, corps){
  if(await aChange()) await telecharge();          // quelqu'un a modifié le fichier : on part de la dernière version
  await sauvegardeDuJour();
  let j = appelPy("POST", chemin, corps);
  if(j.erreur) throw new Error(j.erreur);
  try { await envoie(); }
  catch(e){
    if(e.conflit){                                  // modifié entre-temps : on recommence sur la nouvelle version
      await telecharge();
      j = appelPy("POST", chemin, corps);
      if(j.erreur) throw new Error(j.erreur);
      await envoie();
    } else {
      try { await telecharge(); } catch(x){}
      throw new Error("Non enregistré : " + e.message);
    }
  }
  return j;
}

// ------------------------------------------------------------------ démarrage
async function demarre(){
  etape("Connexion…");
  let r;
  try { r = await retourConnexion(); }
  catch(e){ return erreurDemarrage(e); }
  if(r === "refus" || (!LS.get("w_rt") && !LS.get("w_at"))){
    await new Promise(res => {
      const e = ecran(`<h1>Finances</h1><p>Connecte-toi avec ton compte Microsoft pour ouvrir ton fichier Excel dans OneDrive.</p>
        <button class="btn primary" id="weCnx">Se connecter avec Microsoft</button>`);
      e.querySelector("#weCnx").onclick = () => { versConnexion(true); res(); };
    });
    return new Promise(() => {});
  }
  try {
    const me = await (await graph("/me?$select=displayName,userPrincipalName,mail")).json();
    compte = me.userPrincipalName || me.mail || me.displayName || "";
    if(compte) LS.set("w_login", compte);
    fichier = await choisitFichier(false);
    await chargePython();
    etape("Ouverture de « " + fichier.name + " »…");
    await telecharge();
  } catch(e){ return erreurDemarrage(e); }
  cacheEcran();
  // au retour dans l'app : recharger si le fichier a été modifié ailleurs (ex. sur le Mac)
  document.addEventListener("visibilitychange", async () => {
    if(document.visibilityState !== "visible" || Date.now() - derniereVerif < 20000) return;
    try { if(await aChange()){ await telecharge(); if(typeof rafraichit === "function") await rafraichit(); toast("Données mises à jour depuis OneDrive"); } } catch(e){}
  });
}
function erreurDemarrage(e){
  console.error(e);
  const el = ecran(`<h1>Oups</h1><p>${esc(e.message || e)}</p><button class="btn primary" id="weRe">Réessayer</button>
    <p><button class="btn small" id="weOut">Se déconnecter</button></p>`);
  el.querySelector("#weRe").onclick = () => location.reload();
  el.querySelector("#weOut").onclick = deconnexion;
  return new Promise(() => {});
}
function deconnexion(){
  ["w_at", "w_exp", "w_rt", "w_login", "w_fichier"].forEach(k => LS.set(k, null));
  location.href = REDIRECT;
}
async function changerFichier(){
  try { fichier = await choisitFichier(true); etape("Ouverture de « " + fichier.name + " »…"); await telecharge(); cacheEcran(); await rafraichit(); }
  catch(e){ erreurDemarrage(e); }
}
function carteReglages(){
  return `<div class="card"><h3>Compte et fichier</h3>
    <div class="kv small"><div class="k">Compte Microsoft</div><div style="word-break:break-all">${esc(compte)}</div>
    <div class="k">Fichier OneDrive</div><div style="word-break:break-all">${esc(fichier && fichier.name)}</div></div>
    <p class="hint">Une copie de sécurité est faite chaque jour d'utilisation dans le dossier OneDrive « Sauvegardes app finances », à côté du fichier Excel.</p>
    <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn small" id="weFich">Changer de fichier</button><button class="btn small" id="weDec">Se déconnecter</button></div></div>`;
}
function brancheReglages(el){
  const a = el.querySelector("#weFich"); if(a) a.onclick = changerFichier;
  const b = el.querySelector("#weDec"); if(b) b.onclick = deconnexion;
}
window.WEB = {demarre, api, carteReglages, brancheReglages, version: VERSION};
})();
