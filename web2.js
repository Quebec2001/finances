// web2.js — version en ligne de la v2 (GitHub Pages, /finances/v2/)
// Connexion Microsoft (OAuth PKCE, comptes personnels), base « Finances app.json » dans OneDrive, à côté de l'Excel.
// Le moteur de calcul est en JavaScript (calculs.js, api2.js) : aucun moteur Python au quotidien. Pyodide n'est chargé
// que pour lire l'Excel (essai en mode miroir, import) et pour lire des relevés ou des captures. Aucune donnée ne passe par un serveur tiers.
(function(){
"use strict";
const VERSION = "46cdaddf63";
const CLIENT_ID = "c7ebaba5-e820-450d-92c0-7efd4f7517c8";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPES_BASE = "Files.ReadWrite offline_access User.Read";
// la synchronisation du partage lit un fichier dans le OneDrive de l'autre personne : permission demandée seulement quand on l'active
const SCOPES_SYNC = SCOPES_BASE + " Files.ReadWrite.All";
const scopes = () => LS.get("w_scope") === "all" ? SCOPES_SYNC : SCOPES_BASE;
const G = "https://graph.microsoft.com/v1.0";
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
// empreinte exacte de pyodide.js 0.26.4 : si le fichier servi par le CDN était modifié, le navigateur refuserait de l'exécuter
const PYODIDE_SRI = "sha384-i3R37b3tF+HWudsUf1VSEOY2YxwSNMqY8DQa9Z0O3xh+NkJ9o+yjcGyIi5huj+nB";
const CONFIG = {premier_mois_modifiable: "Octobre 2026", alias_investissements: {"Fonds communs de placement": "CÉLI"}};
const CHEMIN = "/data/finances.xlsx";
// modules Python : ceux de la v1 (« ^ » = dossier racine de l'app, où sont aussi copiés les modules de la v1) pour lire l'Excel et les relevés, et ceux de la v2 (import)
const MODULES = [["^", "formules.py"], ["^", "xlsx.py"], ["^", "modele.py"], ["^", "releve.py"], ["^", "rapprochement.py"],
                 ["", "magasin.py"], ["", "schema.sql"], ["", "importe.py"], ["", "lecture_v2.py"]];
const LECTURE_SEULE = ["/api/releve/analyse", "/api/releve/revenus", "/api/rappro/lire", "/api/rappro/images", "/api/rappro/analyse"];   // POST sans écriture
const BASE_NOM = "Finances app.json";
const JSON_TYPE = "application/json";
// adresse de retour de la connexion Microsoft : la racine de l'app (déjà inscrite chez Microsoft) ; si la v2 est dans v2/, la racine la renvoie ici
const REDIRECT = location.origin + location.pathname.replace(/v2\/[^/]*$/, "").replace(/[^/]*$/, "");
const ICI = location.origin + location.pathname.replace(/[^/]*$/, "");
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const LS = {
  get(k){ try { return localStorage.getItem(k); } catch(e){ return null; } },
  set(k, v){ try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch(e){} }
};
const esc = t => String(t ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
let py = null, fichier = null, compte = null, nomCompte = "", derniereVerif = 0, file = Promise.resolve();
// base de la v2 : D (tables), A (API v2), base = élément OneDrive (id, eTag) ; mode « miroir » (essai, lecture de l'Excel) ou « reel »
let D = null, A = null, base = null, mode = null, rapportImport = null;

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
#webVerrou{position:fixed;inset:0;z-index:300;background:var(--bg);display:none;align-items:center;justify-content:center;padding:24px}@keyframes we{to{transform:rotate(360deg)}}.we-liste{display:grid;gap:8px;margin-top:12px;text-align:left}.we-liste button{text-align:left}`;
document.head.appendChild(style);

// ------------------------------------------------------------------ connexion Microsoft
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
function aleatoire(n){ const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a); }

async function versConnexion(choisirCompte){
  const verifier = aleatoire(48), state = "web2-" + aleatoire(12);
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  LS.set("w_verifier", verifier); LS.set("w_state", state);
  const q = new URLSearchParams({client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT, response_mode: "query",
    scope: scopes(), state, code_challenge: challenge, code_challenge_method: "S256"});
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
  if(!(p.get("state") || "").startsWith("web2-")) return false;
  history.replaceState(null, "", location.pathname);
  if(p.has("error")){
    if(p.get("error") === "access_denied"){
      if(LS.get("w_apres")){ LS.set("w_scope", null); LS.set("w_apres", null); LS.set("w_apres_refus", "1"); return LS.get("w_rt") ? true : "refus"; }
      return "refus";
    }
    throw new Error("Connexion Microsoft refusée : " + (p.get("error_description") || p.get("error")));
  }
  if(p.get("state") !== LS.get("w_state")) throw new Error("Retour de connexion invalide. Réessaie.");
  garde(await jeton({client_id: CLIENT_ID, grant_type: "authorization_code", code: p.get("code"), redirect_uri: REDIRECT,
    code_verifier: LS.get("w_verifier"), scope: scopes()}));
  LS.set("w_verifier", null); LS.set("w_state", null);
  return true;
}
async function accessToken(){
  const at = LS.get("w_at");
  if(at && Date.now() < +(LS.get("w_exp") || 0)) return at;
  const rt = LS.get("w_rt");
  if(rt){
    try { const t = await jeton({client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: rt, scope: scopes()}); garde(t); return t.access_token; }
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
  // la recherche OneDrive peut être lente ou incomplète (index) : on regarde aussi directement le dossier Documents
  if(!items.some(i => /financ/i.test(i.name))){
    try { items = xlsx((await (await graph("/me/drive/special/documents/children?$top=200&$select=id,name,eTag,size,lastModifiedDateTime,file,parentReference")).json()).value).concat(items); } catch(e){}
  }
  let cands = items.filter(i => /financ/i.test(i.name));
  if(!cands.length) cands = items;
  if(!cands.length) throw new Error("Aucun fichier Excel « Finances… » trouvé dans le OneDrive du compte " + (compte || "connecté") + ". Si ce n'est pas le bon compte Microsoft, touche « Se déconnecter » et reconnecte-toi avec le bon.");
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
async function telechargeExcel(){
  for(let k = 0; k < 3; k++){
    const avant = await (await graph(`/me/drive/items/${fichier.id}?$select=id,name,eTag,parentReference`)).json();
    const buf = new Uint8Array(await (await graph(`/me/drive/items/${fichier.id}/content`)).arrayBuffer());
    const apres = await (await graph(`/me/drive/items/${fichier.id}?$select=id,name,eTag,parentReference`)).json();
    if(avant.eTag === apres.eTag){
      if(!(buf[0] === 0x50 && buf[1] === 0x4B)) throw new Error("Le fichier téléchargé n'est pas un classeur Excel valide.");
      fichier.eTag = apres.eTag; fichier.name = apres.name; fichier.parentReference = apres.parentReference;
      return buf;
    }
  }
  throw new Error("Le fichier change pendant le téléchargement (il est peut-être en cours de modification). Réessaie dans un instant.");
}
// ------------------------------------------------------------------ base de la v2 dans OneDrive
const dossierExcel = () => { const p = fichier && fichier.parentReference && fichier.parentReference.id; return p ? `/me/drive/items/${p}:` : "/me/drive/root:"; };
async function trouveBase(){
  const id = LS.get("w2_base");
  if(id){ try { return await (await graph(`/me/drive/items/${id}?$select=id,name,eTag,parentReference`)).json(); } catch(e){ if(e.status !== 404) throw e; LS.set("w2_base", null); } }
  try { const b = await (await graph(`${dossierExcel()}/${encodeURIComponent(BASE_NOM)}?$select=id,name,eTag,parentReference`)).json(); LS.set("w2_base", b.id); return b; }
  catch(e){ if(e.status === 404) return null; throw e; }
}
async function chargeBase(){
  for(let k = 0; k < 3; k++){
    const a = await (await graph(`/me/drive/items/${base.id}?$select=id,eTag`)).json();
    const txt = await (await graph(`/me/drive/items/${base.id}/content`, {cache: "no-store"})).text();
    const b = await (await graph(`/me/drive/items/${base.id}?$select=id,eTag`)).json();
    if(a.eTag === b.eTag){ D = JSON.parse(txt); base.eTag = b.eTag; derniereVerif = Date.now(); nouvelleApi(); return; }
  }
  throw new Error("La base change pendant le téléchargement. Réessaie dans un instant.");
}
async function baseAChange(){
  const m = await (await graph(`/me/drive/items/${base.id}?$select=id,eTag`)).json();
  derniereVerif = Date.now();
  return m.eTag !== base.eTag;
}
async function envoieBase(){
  try {
    const r = await graph(`/me/drive/items/${base.id}/content`, {method: "PUT", body: JSON.stringify(D), headers: {"Content-Type": JSON_TYPE, "If-Match": base.eTag}});
    base.eTag = (await r.json()).eTag;
  } catch(e){
    if(e.status === 412 || e.status === 409){ const c = new Error("conflit"); c.conflit = true; throw c; }
    throw e;
  }
}
async function creeBase(){
  const r = await graph(`${dossierExcel()}/${encodeURIComponent(BASE_NOM)}:/content?@microsoft.graph.conflictBehavior=fail`,
    {method: "PUT", body: JSON.stringify(D), headers: {"Content-Type": JSON_TYPE}});
  base = await r.json(); LS.set("w2_base", base.id);
}
const date2 = n => String(n).padStart(2, "0");
const jourISO = (d = new Date()) => `${d.getFullYear()}-${date2(d.getMonth() + 1)}-${date2(d.getDate())}`;
// Sauvegarde du jour (au premier changement de la journée) dans « Sauvegardes app finances », puis ménage :
// on garde les 30 dernières copies quotidiennes et une par mois (la plus ancienne du mois) au-delà.
async function sauvegardeDuJour(){
  const jour = jourISO(), cle = "w2_sauv_" + base.id;
  if(LS.get(cle) === jour) return;
  const dossier = `${dossierExcel()}/${encodeURIComponent("Sauvegardes app finances")}`;
  try {
    await graph(`${dossier}/${encodeURIComponent(`Finances app – ${jour}.json`)}:/content`, {method: "PUT", body: JSON.stringify(D), headers: {"Content-Type": JSON_TYPE}});
    LS.set(cle, jour);
    const l = ((await (await graph(`${dossier}:/children?$top=999&$select=id,name`)).json()).value || [])
      .map(x => ({id: x.id, jour: (/^Finances app – (\d{4}-\d{2}-\d{2})\.json$/.exec(x.name) || [])[1]})).filter(x => x.jour).sort((a, b) => b.jour.localeCompare(a.jour));
    for(const x of aSupprimer(l, jour)) { try { await graph(`/me/drive/items/${x.id}`, {method: "DELETE"}); } catch(e){} }
  } catch(e){ console.warn("Sauvegarde OneDrive impossible", e); }
}
function aSupprimer(liste, aujourdhui){
  // liste triée du plus récent au plus ancien : 30 dernières gardées ; au-delà, une par mois (la plus ancienne du mois)
  const garde = new Set(liste.slice(0, 30).map(x => x.id)), parMois = {};
  for(const x of liste.slice(30)) parMois[x.jour.slice(0, 7)] = x;          // la dernière vue = la plus ancienne du mois
  Object.values(parMois).forEach(x => garde.add(x.id));
  return liste.filter(x => !garde.has(x.id));
}

// ------------------------------------------------------------------ moteur Python
function chargeScript(src, integrite){
  return new Promise((ok, ko) => { const s = document.createElement("script"); s.src = src; if(integrite){ s.integrity = integrite; s.crossOrigin = "anonymous"; } s.onload = ok; s.onerror = () => ko(new Error("Chargement impossible : " + src)); document.head.appendChild(s); });
}
let pyPret = null;
function chargePython(msg){
  if(!pyPret) pyPret = (async () => {
    if(msg) etape(msg);
    await chargeScript(PYODIDE + "pyodide.js", PYODIDE_SRI);
    py = await loadPyodide({indexURL: PYODIDE});
    py.FS.mkdirTree("/app"); py.FS.mkdirTree("/data");
    for(const [dossier, n] of MODULES){
      const r = await fetch((dossier === "^" ? REDIRECT : dossier) + n + "?v=" + VERSION, {cache: "no-cache"});
      if(!r.ok) throw new Error("Module manquant : " + n);
      py.FS.writeFile("/app/" + n, await r.text());
    }
    py.runPython("import sys\nsys.path.insert(0, '/app')\nimport importe, lecture_v2, json");
  })().catch(e => { pyPret = null; throw e; });
  return pyPret;
}
// lecture de l'Excel → tables de la v2 (même code que les contrôles de l'import)
function importeExcel(buf){
  py.FS.writeFile(CHEMIN, buf);
  py.globals.set("_cfg", JSON.stringify(CONFIG));
  const r = JSON.parse(py.runPython(`json.dumps(importe.importe("${CHEMIN}", None, json.loads(_cfg)), ensure_ascii=False, default=str)`));
  return r;
}
// résumé de la v1 (lecture de l'Excel) pour l'écran Concordance : totaux de chaque mois, solde avec le partenaire
function resumeV1(){
  return JSON.parse(py.runPython(`
from modele import Application
_a = Application(dict(json.loads(_cfg), fichier="${CHEMIN}", dossier_sauvegardes="/data/s")); _a.charge(); _e = _a.etat(); _out = {"mois": {}, "pourboires": bool(_e["reglages"].get("pourboires"))}
for _i in _e["mois"]:
    _m = _a.modele(_i["feuille"]) if _i["canvas"] else _a.historique(_i["feuille"])
    _out["mois"][_i["cle"]] = {"nom": _i["nom"], "depenses": _m["total_depenses"], "revenus": _m["total_revenus"], "debut": _m.get("solde_debut"), "fin": _m.get("solde_fin"),
                               "du_net": _m.get("du_net") if _i["canvas"] else None}
json.dumps(_out, default=str)`));
}

// ------------------------------------------------------------------ API utilisée par l'interface (moteur v2 en JavaScript)
function nouvelleApi(){
  A = ApiV2.API(D, {fichier: mode === "miroir" ? (fichier && fichier.name) + " (essai : lecture seule)" : BASE_NOM, lectureSeule: route => LECTURE_SEULE.includes(route) ? lecture(route) : null});
}
const lecture = route => async corps => {
  await chargePython();
  const etat = A.etat(), mois = {};
  const i = etat.mois.findIndex(m => m.feuille === (corps.feuille || "").trim());
  for(const m of [etat.mois[i], etat.mois[i - 1]]) if(m && m.canvas) mois[m.feuille] = A.mois(m.feuille);
  py.globals.set("_in", JSON.stringify({route, corps, etat, mois}));
  const r = JSON.parse(py.runPython("lecture_v2.appel_json(_in)"));
  if(r.erreur) throw new Error(r.erreur);
  return r.resultat;
};
async function api(chemin, corps){
  if(!corps){
    const j = await A.api(chemin);
    if(chemin === "/api/etat"){ j.web = {compte, nom: nomCompte, fichier: fichier && fichier.name}; j.mode_v2 = mode; j.concordance_v2 = !!rapportImport;
      if(mode !== "reel") j.mois.forEach(m => { m.modifiable = false; });      // essai : lecture seule
    }
    if(mode !== "reel" && chemin.startsWith("/api/mois") && j) j.modifiable = false;
    return j;
  }
  if(LECTURE_SEULE.includes(chemin)) return A.api(chemin, corps);
  if(mode !== "reel") throw new Error("Mode essai : la v2 lit ton Excel sans rien modifier. Fais tes changements dans l'app actuelle.");
  const tache = file.then(() => ecriture(chemin, corps));     // écritures : une à la fois
  file = tache.catch(() => {});
  return tache;
}
async function ecriture(chemin, corps){
  if(await baseAChange()) await chargeBase();                 // modifiée ailleurs (autre appareil) : on part de la dernière version
  await sauvegardeDuJour();
  let j;
  try { j = await A.api(chemin, corps); }
  catch(e){ await chargeBase(); throw e; }                    // rien n'est envoyé : on revient à la version enregistrée
  try { await envoieBase(); }
  catch(e){
    if(e.conflit){ await chargeBase(); j = await A.api(chemin, corps); await envoieBase(); }
    else { try { await chargeBase(); } catch(x){} throw new Error("Non enregistré : " + e.message); }
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
  if(LS.get("w_verrou")) await verrouille(true);       // Face ID / code de l'appareil avant d'ouvrir quoi que ce soit
  try {
    const me = await (await graph("/me?$select=displayName,userPrincipalName,mail")).json();
    compte = me.userPrincipalName || me.mail || me.displayName || "";
    nomCompte = me.displayName || "";
    if(compte) LS.set("w_login", compte);
    fichier = await choisitFichier(false);
    etape("Recherche de la base de la v2…");
    base = await trouveBase();
    if(base){
      mode = "reel"; etape("Ouverture de tes finances…"); await chargeBase();
      try { await chargeLien(); } catch(e){}
    } else {
      mode = "miroir"; await essaiMiroir();
    }
  } catch(e){ return erreurDemarrage(e); }
  cacheEcran();
  // au retour dans l'app : recharger si la base (ou, pendant l'essai, l'Excel) a été modifiée ailleurs
  document.addEventListener("visibilitychange", async () => {
    if(document.visibilityState !== "visible" || Date.now() - derniereVerif < 20000) return;
    try {
      if(mode === "reel" && await baseAChange()){ await chargeBase(); if(typeof rafraichit === "function") await rafraichit(); toast("Données mises à jour depuis OneDrive"); }
      if(mode === "miroir" && await excelAChange()){ await essaiMiroir(true); if(typeof rafraichit === "function") await rafraichit(); toast("Excel relu"); }
    } catch(e){}
  });
}
// ------------------------------------------------------------------ essai (mode miroir) et passage à la v2
// Pendant l'essai, la v2 relit l'Excel à chaque ouverture, sans rien écrire. L'écran Concordance compare chaque mois.
async function excelAChange(){
  const m = await (await graph(`/me/drive/items/${fichier.id}?$select=id,eTag`)).json(); derniereVerif = Date.now(); return m.eTag !== fichier.eTag;
}
async function essaiMiroir(silencieux){
  await chargePython(silencieux ? null : "Essai de la v2 : lecture de ton Excel… (environ 15 secondes)");
  if(!silencieux) etape("Lecture de « " + fichier.name + " »…");
  const buf = await telechargeExcel();
  const r = importeExcel(buf);
  D = r.donnees; delete r.donnees;
  rapportImport = {rapport: r, v1: resumeV1(), lu: new Date().toISOString()};
  derniereVerif = Date.now(); nouvelleApi();
}
// Concordance v1 / v2 : pour chaque mois, les totaux lus dans l'Excel (v1) et ceux calculés par la v2 sur les mêmes données.
// La comparaison se fait avec les règles de la v1 ; l'écart dû à la nouvelle règle (part payée par le partenaire
// comptée au règlement) est affiché à part.
function concordance(){
  if(!rapportImport) return null;
  const v1 = rapportImport.v1, logique = v1.pourboires ? "v1-zach" : "v2";
  const M1 = CalculsV2.Moteur(D), ch = M1.chaine(logique), ch2 = M1.chaine("v2");
  const eg = (a, b) => Math.abs((a || 0) - (b || 0)) < 0.005;
  const lignes = Object.entries(v1.mois).map(([cle, x]) => {
    const y = ch[cle] || {}; const c = [["Dépenses", x.depenses, y.totalDepenses], ["Revenus", x.revenus, y.revenus]];
    if(x.debut != null) c.push(["Solde de début", x.debut, y.debut], ["Solde de fin", x.fin, y.fin]);
    return {cle, nom: x.nom, controles: c.map(([lib, a, b]) => ({lib, v1: a, v2: b, ok: eg(a, b)})), nouvelle_regle: (ch2[cle] || {}).totalDepenses - (y.totalDepenses || 0)};
  });
  const der = Object.entries(v1.mois).filter(([, x]) => x.du_net != null).pop();
  const part = der ? {v1: der[1].du_net, v2: M1.soldePartenaire(), ok: eg(der[1].du_net, M1.soldePartenaire())} : null;
  const ok = lignes.every(l => l.controles.every(c => c.ok)) && (!part || part.ok) && !rapportImport.rapport.avertissements.some(a => /non reconnu|introuvable/.test(a));
  return {ok, lignes, partenaire: part, avertissements: rapportImport.rapport.avertissements, lu: rapportImport.lu, fichier: fichier && fichier.name};
}
async function bascule(){
  const c = concordance(); if(!c || !c.ok) throw new Error("La concordance n'est pas complète : on ne peut pas passer à la v2.");
  if(await excelAChange()) throw new Error("Ton Excel a changé depuis la lecture. Rouvre l'app, vérifie la concordance, puis recommence.");
  etape("Création de ta base « " + BASE_NOM + " » dans OneDrive…");
  const r = D.reglages.find(x => x.cle === "bascule"); const v = JSON.stringify({date: new Date().toISOString(), excel: fichier.name});
  if(r) r.valeur = v; else D.reglages.push({cle: "bascule", valeur: v});
  await creeBase();
  mode = "reel"; rapportImport = null;
  location.reload();
}

// si la politique de sécurité bloque une adresse imprévue (ex. un nouveau serveur de téléchargement OneDrive),
// on l'affiche dans le message d'erreur pour pouvoir l'ajouter à la liste
const bloques = [];
document.addEventListener("securitypolicyviolation", ev => { try { const h = new URL(ev.blockedURI).host; if(h && !bloques.includes(h)) bloques.push(h); } catch(x){} });
function erreurDemarrage(e){
  console.error(e);
  if(bloques.length) e = new Error((e.message || e) + " (adresse bloquée par la sécurité : " + bloques.join(", ") + ")");
  const el = ecran(`<h1>Oups</h1><p>${esc(e.message || e)}</p><button class="btn primary" id="weRe">Réessayer</button>
    <p><button class="btn small" id="weOut">Se déconnecter</button></p>`);
  el.querySelector("#weRe").onclick = () => location.reload();
  el.querySelector("#weOut").onclick = deconnexion;
  return new Promise(() => {});
}
function deconnexion(){
  ["w_at", "w_exp", "w_rt", "w_login", "w_fichier", "w_lien", "w_scope", "w_sync_off", "w_verrou", "w2_base"].forEach(k => LS.set(k, null));
  location.href = REDIRECT;
}
async function changerFichier(){
  try { fichier = await choisitFichier(true); LS.set("w2_base", null); location.reload(); }
  catch(e){ erreurDemarrage(e); }
}
function carteReglages(){
  return `<div class="card"><h3>Compte et fichier</h3>
    <div class="kv small"><div class="k">Compte Microsoft</div><div style="word-break:break-all">${esc(compte)}</div>
    <div class="k">${mode === "reel" ? "Base de l'app" : "Excel (essai)"}</div><div style="word-break:break-all">${esc(mode === "reel" ? BASE_NOM : fichier && fichier.name)}</div></div>
    <p class="hint">${mode === "reel" ? "Une copie de sécurité est faite chaque jour d'utilisation dans le dossier OneDrive « Sauvegardes app finances », à côté de ton Excel (les 30 derniers jours, puis une par mois)." : "Essai de la v2 : ton Excel est seulement lu, rien n'est modifié."}</p>
    <div class="we-verrou" style="display:flex;align-items:center;gap:10px;margin:10px 0"><div class="grow" style="flex:1"><b>Verrouillage Face ID / code</b><div class="hint" style="margin:2px 0 0">${LS.get("w_verrou") ? "Activé : demandé à l'ouverture et après 5 minutes hors de l'app." : "Demande Face ID ou le code de l'appareil pour ouvrir l'app."}</div></div><button class="btn small" id="weVerrou">${LS.get("w_verrou") ? "Désactiver" : "Activer"}</button></div>
    <div style="display:flex;gap:8px;flex-wrap:wrap">${mode === "reel" ? "" : `<button class="btn small" id="weFich">Changer de fichier</button>`}<button class="btn small" id="weDec">Se déconnecter</button>${typeof window.dlgSync === "function" ? `<button class="btn small" id="weSync">Synchronisation du partage</button>` : ""}</div></div>`;
}
function brancheReglages(el){
  const a = el.querySelector("#weFich"); if(a) a.onclick = changerFichier;
  const b = el.querySelector("#weDec"); if(b) b.onclick = deconnexion;
  const v = el.querySelector("#weVerrou");
  if(v) v.onclick = async () => {
    try {
      if(LS.get("w_verrou")){ await verifieIdentite(); LS.set("w_verrou", null); toast("Verrouillage désactivé"); }
      else { await activeVerrou(); toast("Verrouillage activé"); }
    } catch(e){ if(!(e && e.name === "NotAllowedError")) toast(e.message || String(e)); return; }
    const c = v.closest(".card"); if(c){ const par = c.parentNode; c.outerHTML = carteReglages(); brancheReglages(par); }
  };
  const c = el.querySelector("#weSync"); if(c) c.onclick = () => window.dlgSync();
}

// ------------------------------------------------------------------ verrouillage (Face ID / Touch ID / code de l'appareil)
// WebAuthn avec l'authentificateur intégré de l'appareil : l'app demande au téléphone de confirmer l'identité de la
// personne (Face ID, empreinte ou code). Aucune donnée biométrique n'est vue par l'app ni envoyée nulle part.
// C'est un verrou d'accès à l'app : il empêche quelqu'un qui prend le téléphone déverrouillé d'ouvrir les finances.
const VERROU_DELAI = 5 * 60 * 1000;            // re-verrouiller après 5 minutes hors de l'app
const deB64 = t => Uint8Array.from(atob(t.replace(/-/g, "+").replace(/_/g, "/")), c => c.charCodeAt(0));
const octets = n => crypto.getRandomValues(new Uint8Array(n));
let verrouOuvert = null, quitteLe = 0;
async function verrouDispo(){
  try { return !!(window.PublicKeyCredential && await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable()); } catch(e){ return false; }
}
async function verifieIdentite(){
  const id = LS.get("w_verrou"); if(!id) return true;
  await navigator.credentials.get({publicKey: {challenge: octets(32), rpId: location.hostname, timeout: 60000, userVerification: "required",
    allowCredentials: [{type: "public-key", id: deB64(id), transports: ["internal", "hybrid"]}]}});
  return true;
}
function verrouille(auto){
  if(verrouOuvert) return verrouOuvert;
  let el = document.getElementById("webVerrou");
  if(!el){ el = document.createElement("div"); el.id = "webVerrou"; document.body.appendChild(el); }
  el.innerHTML = `<div class="we-boite"><img src="icone.png" alt="" class="we-logo"><h1>Finances verrouillées</h1>
    <p>Confirme ton identité avec Face ID ou le code de ton appareil.</p><p id="wvMsg" class="small"></p>
    <button class="btn primary" id="wvOk">Déverrouiller</button>
    <p><button class="btn small" id="wvOut">Se déconnecter</button></p></div>`;
  el.style.display = "flex";
  verrouOuvert = new Promise(ok => {
    const essai = async () => {
      try { await verifieIdentite(); el.style.display = "none"; el.innerHTML = ""; verrouOuvert = null; ok(); }
      catch(e){ el.querySelector("#wvMsg").textContent = e && e.name === "NotAllowedError" ? "" : "Vérification impossible : " + (e.message || e); }
    };
    el.querySelector("#wvOk").onclick = essai;
    el.querySelector("#wvOut").onclick = deconnexion;
    if(auto) essai();
  });
  return verrouOuvert;
}
async function activeVerrou(){
  if(!await verrouDispo()) throw new Error("Cet appareil ou ce navigateur n'offre pas Face ID / code pour les sites web.");
  const c = await navigator.credentials.create({publicKey: {
    challenge: octets(32), rp: {name: "Finances", id: location.hostname}, timeout: 60000, attestation: "none",
    user: {id: octets(16), name: compte || "finances", displayName: nomCompte || compte || "Finances"},
    pubKeyCredParams: [{type: "public-key", alg: -7}, {type: "public-key", alg: -257}],
    authenticatorSelection: {authenticatorAttachment: "platform", userVerification: "required", residentKey: "discouraged"}}});
  LS.set("w_verrou", b64url(c.rawId));
}
document.addEventListener("visibilitychange", () => {
  if(!LS.get("w_verrou")) return;
  if(document.visibilityState === "hidden") quitteLe = Date.now();
  else if(quitteLe && Date.now() - quitteLe > VERROU_DELAI) verrouille(true);
});

// ------------------------------------------------------------------ fichier partagé du couple (synchronisation)
// « Partage couple.json » est créé dans le OneDrive de la personne qui lie les comptes, à côté de son Excel,
// puis partagé (lecture-écriture) avec l'autre personne seulement. Chaque app y lit et y écrit avec If-Match (eTag).
const PARTAGE_NOM = "Partage couple.json";
const LIEN_APPROOT = "/me/drive/special/approot:/partage-couple-lien.json";
const lien = () => { try { return JSON.parse(LS.get("w_lien") || "null"); } catch(e){ return null; } };
function activeScope(){ if(LS.get("w_scope") !== "all"){ LS.set("w_scope", "all"); LS.set("w_at", null); } }
async function assureScope(apres){
  if(LS.get("w_scope") === "all") return true;
  LS.set("w_apres", apres); LS.set("w_scope", "all"); LS.set("w_at", null);
  etape("Autorisation OneDrive pour la synchronisation…");
  return versConnexion(false);            // Microsoft demande le consentement une seule fois, puis revient ici
}
function apres(){ const a = LS.get("w_apres"), r = LS.get("w_apres_refus"); LS.set("w_apres", null); LS.set("w_apres_refus", null); return r ? "refus" : a; }
async function gardeLien(l){
  LS.set("w_lien", JSON.stringify(l)); activeScope();
  try { await graph(LIEN_APPROOT + ":/content", {method: "PUT", body: JSON.stringify(l), headers: {"Content-Type": "application/json"}}); } catch(e){ console.warn("lien non mémorisé dans OneDrive", e); }
}
async function chargeLien(){
  if(lien()) return lien();
  try { if(sessionStorage.getItem("w_lien_vu")) return null; } catch(e){}
  try {
    // liste du dossier de l'app (évite une erreur 404 quand rien n'est encore lié)
    const f = ((await (await graph("/me/drive/special/approot/children?$select=id,name")).json()).value || []).find(x => x.name === "partage-couple-lien.json");
    if(!f){ try { sessionStorage.setItem("w_lien_vu", "1"); } catch(e){} return null; }
    const l = await (await graph(LIEN_APPROOT + ":/content")).json();
    if(l && l.d && l.i){ LS.set("w_lien", JSON.stringify(l)); activeScope(); return l; }
  } catch(e){}
  try { sessionStorage.setItem("w_lien_vu", "1"); } catch(e){}
  return null;
}
async function creePartage(email, initial, message){
  const parent = fichier.parentReference && fichier.parentReference.id;
  const base = parent ? `/me/drive/items/${parent}:` : "/me/drive/root:";
  let it;
  try { it = await (await graph(`${base}/${encodeURIComponent(PARTAGE_NOM)}?$select=id,eTag,parentReference`)).json(); }
  catch(e){
    if(e.status !== 404) throw e;
    it = await (await graph(`${base}/${encodeURIComponent(PARTAGE_NOM)}:/content`, {method: "PUT", body: JSON.stringify(initial), headers: {"Content-Type": "application/json"}})).json();
  }
  const l = {d: it.parentReference.driveId, i: it.id, proprio: initial.proprio};
  if(email) await graph(`/drives/${l.d}/items/${l.i}/invite`, {method: "POST", headers: {"Content-Type": "application/json"},
    body: JSON.stringify({recipients: [{email}], requireSignIn: true, sendInvitation: true, roles: ["write"], message})});
  await gardeLien(l);
  return l;
}
async function rejoint(l){
  await graph(`/drives/${l.d}/items/${l.i}?$select=id,eTag`);       // 403/404 si l'invitation n'est pas encore acceptée
  await gardeLien(l);
  return l;
}
async function litPartage(){
  const l = lien(); if(!l) throw new Error("Comptes non liés");
  const u = `/drives/${l.d}/items/${l.i}`;
  for(let k = 0; k < 3; k++){
    const a = await (await graph(u + "?$select=id,eTag")).json();
    const txt = await (await graph(u + "/content", {cache: "no-store"})).text();
    const b = await (await graph(u + "?$select=id,eTag")).json();
    if(a.eTag === b.eTag){ let j = {}; try { j = JSON.parse(txt || "{}"); } catch(e){ j = {}; } return {j, etag: b.eTag}; }
  }
  throw new Error("Le fichier partagé change en ce moment. Réessaie.");
}
async function ecritPartage(j, etag){
  const l = lien();
  try {
    const r = await graph(`/drives/${l.d}/items/${l.i}/content`, {method: "PUT", body: JSON.stringify(j), headers: {"Content-Type": "application/json", "If-Match": etag}});
    return (await r.json()).eTag;
  } catch(e){
    if(e.status === 412 || e.status === 409){ const c = new Error("conflit"); c.conflit = true; throw c; }
    throw e;
  }
}
async function delie(){
  LS.set("w_lien", null);
  try { await graph(LIEN_APPROOT, {method: "DELETE"}); } catch(e){}
}
const partage = {lien: () => mode === "reel" ? lien() : null, chargeLien, assureScope, apres, creePartage, rejoint, lit: litPartage, ecrit: ecritPartage, delie,
  pause: () => LS.get("w_sync_off") === "1", metPause: b => LS.set("w_sync_off", b ? "1" : null), compte: () => compte};
window.WEB = {demarre, api, carteReglages, brancheReglages, partage, version: VERSION, v2: true,
  mode: () => mode, concordance, bascule, aSupprimer};
// Mise à jour automatique : l'iPhone garde parfois l'ancienne page (icône de l'écran d'accueil). Au lancement et au
// retour dans l'app, on compare la version publiée sur GitHub à celle-ci ; si elle a changé, on recharge (jamais
// pendant une écriture ni avec une fenêtre ouverte).
async function verifieMiseAJour(){
  if(VERSION.indexOf("__") === 0) return;
  try {
    const t = await (await fetch(location.pathname + "?maj=" + Date.now(), {cache: "no-store"})).text();
    const v = (t.match(/web\.js\?v=(\w+)/) || [])[1];
    if(!v || v === VERSION) return;
    const dl = document.querySelector("#dlg");
    if((dl && dl.open) || document.body.classList.contains("loading")) return;
    let deja = null; try { deja = sessionStorage.getItem("maj_" + v); sessionStorage.setItem("maj_" + v, "1"); } catch(e){}
    if(deja) return;                                   // déjà essayé : pas de boucle
    location.replace(location.pathname + "?v=" + v + location.hash);
  } catch(e){}
}
window.WEB.verifieMiseAJour = verifieMiseAJour;
setTimeout(verifieMiseAJour, 4000);
document.addEventListener("visibilitychange", () => { if(document.visibilityState === "visible") verifieMiseAJour(); });
})();
