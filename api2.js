// API de la v2 : mêmes routes et mêmes formes de réponse que la v1 (/api/etat, /api/mois, /api/tendances, écritures),
// mais sur la base de la v2 (tables en mémoire, enregistrées en JSON) et le moteur calculs.js.
// L'interface existante (index.html) fonctionne ainsi sans changement de fond ; les écrans propres à Excel disparaissent.
"use strict";
(function(racine){
  const {Moteur, moisPrec} = (typeof module !== "undefined" && module.exports) ? require("./calculs.js") : racine.CalculsV2;
  const MOIS_FR = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
  const RX_A_REMB = /\s*\(carte ([^()]+?) – à rembourser\)\s*$/;
  const RX_CARTE = /\s*\(carte ([^()]+?)\)\s*$/;
  const norm = t => String(t || "").trim().toLowerCase().normalize("NFC");
  const r2 = x => Math.round((x || 0) * 100) / 100;
  const nomMois = cle => { const [a, m] = cle.split("-").map(Number); return MOIS_FR[m - 1] + " " + a; };
  const cleDe = nom => { const m = /^\s*([^\s]+)\s+(\d{4})\s*$/.exec(nom || ""); if(!m) return null;
    const i = MOIS_FR.findIndex(x => norm(x) === norm(m[1])); return i < 0 ? null : m[2] + "-" + String(i + 1).padStart(2, "0"); };
  const cents = v => Math.round(v * 100);
  class Erreur extends Error {}

  // montant saisi : « 12,50 », « 120+15,50 », « =33,34+15,53 » (petite calculatrice, sans eval)
  function calcule(txt){
    let t = String(txt ?? "").replace(/\s|\$/g, "").replace(/^=/, "").replace(/,/g, ".");
    if(!t) return null;
    if(!/^[\d.+\-*/()]+$/.test(t)) throw new Erreur("Montant invalide : " + txt);
    let i = 0;
    const nb = () => { const m = /^(\d+\.?\d*|\.\d+)/.exec(t.slice(i)); if(!m) throw new Erreur("Montant invalide : " + txt); i += m[0].length; return parseFloat(m[0]); };
    const fact = () => { const c = t[i]; if(c === "+"){ i++; return fact(); } if(c === "-"){ i++; return -fact(); }
      if(c === "("){ i++; const v = somme(); if(t[i] !== ")") throw new Erreur("Montant invalide : " + txt); i++; return v; } return nb(); };
    const prod = () => { let v = fact(); while(t[i] === "*" || t[i] === "/"){ const o = t[i++], w = fact(); v = o === "*" ? v * w : v / w; } return v; };
    const somme = () => { let v = prod(); while(t[i] === "+" || t[i] === "-"){ const o = t[i++], w = prod(); v = o === "+" ? v + w : v - w; } return v; };
    const v = somme(); if(i !== t.length || !isFinite(v)) throw new Erreur("Montant invalide : " + txt);
    return {valeur: v, calcul: /[+\-*/()]/.test(t.replace(/^-/, "")) ? t.replace(/\./g, ",") : null};
  }

  const TABLES = ["reglages","categories","mots_cles","reglements","depenses","autres_montants","pourboires","revenus","revenus_pourboires",
    "repartition","comptes_placement","placements_mois","versements_epargne","soldes_mois","ajustements_solde","releves_banque","budgets",
    "exclusions_rappro","journal","financements","financements_recus","financements_soldes","fixes","prevus"];
  function API(D, opts = {}){
    for(const t of TABLES) if(!Array.isArray(D[t])) D[t] = [];
    const auj = () => opts.aujourdhui ? opts.aujourdhui() : new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
    const REG = () => { const o = {}; for(const r of D.reglages) o[r.cle] = JSON.parse(r.valeur); return o; };
    const setReg = (k, v) => { const r = D.reglages.find(x => x.cle === k); if(r) r.valeur = JSON.stringify(v); else D.reglages.push({cle: k, valeur: JSON.stringify(v)}); };
    const partenaire = () => REG().partenaire || "partenaire";
    const TYPES = () => [["P", "Personnel"], ["50", "SPLIT 50/50"], ["SP", "SPLIT Proportionnel"], ["M", "100% " + partenaire()]];
    const libType = c => c === "R" ? "Remboursement" : (TYPES().find(x => x[0] === c) || [c, c])[1];
    const codeType = lib => { const n = norm(lib);
      if(n.startsWith("personnel")) return "P"; if(n.startsWith("split 50")) return "50"; if(n.startsWith("split prop")) return "SP";
      if(n.startsWith("rembo")) return "R"; if(n.startsWith("100")) return "M"; return null; };
    const nouvelId = t => D[t].reduce((a, x) => Math.max(a, x.id || 0), 0) + 1;
    const ordreSuivant = () => Math.max(0, ...D.depenses.map(d => d.ordre || 0), ...D.autres_montants.map(d => d.ordre || 0)) + 1;
    const journal = (table, id, action, avant, apres) => D.journal.push({id: nouvelId("journal"), quand: new Date().toISOString(), table_nom: table,
      ligne_id: id, action, avant: avant ? JSON.stringify(avant) : null, apres: apres ? JSON.stringify(apres) : null});
    // le mois suivant est toujours ouvert (dépenses et revenus prévus, ex. le loyer du 1er) : plus de « Créer le mois suivant »
    const moisSuivant = mo => { let [y, n] = mo.split("-").map(Number); n++; if(n > 12){ n = 1; y++; } return y + "-" + String(n).padStart(2, "0"); };
    const M = () => Moteur(D, {moisMax: moisSuivant(auj().slice(0, 7))});
    const catId = nom => { if(!nom) return null; const c = D.categories.find(x => norm(x.nom) === norm(nom)); return c ? c.id : null; };
    const catNom = id => (D.categories.find(x => x.id === id) || {}).nom;

    // ------------------------------------------------------------------ mois : liste et verrou
    const verrou = () => REG().verrou_avant || "0000-00";
    const historique = (mo) => { const d = D.depenses.filter(x => x.mois === mo); return d.length > 0 && d.every(x => x.historique); };
    function infos(){
      return M().listeMois().map(cle => {
        const nom = nomMois(cle), hist = cle < verrou() && historique(cle);
        return {feuille: nom, nom, cle, annee: +cle.slice(0, 4), mois: +cle.slice(5), groupe: cle >= verrou() ? "final" : (hist ? "historique" : "reference"),
                canvas: !hist, modifiable: cle >= verrou()};
      });
    }
    const info = f => { const c = cleDe(f); const i = infos().find(x => x.cle === c); if(!i) throw new Erreur("Mois introuvable : " + f); return i; };
    const modifiable = f => { const i = info(f); if(!i.modifiable) throw new Erreur(`${i.nom} est en lecture seule (mois d'avant ${nomMois(verrou())}).`); return i; };
    const moisDeDate = (date, i) => date && /^\d{4}-\d{2}-\d{2}$/.test(date) ? date.slice(0, 7) : i.cle;

    // ------------------------------------------------------------------ lecture : état
    function etat(){
      const reg = REG(), m = M();
      const reRev = new RegExp(reg.mots_revenu_principal || "^$", "i"), freq = {};
      for(const r of D.revenus) if(!r.historique && r.description && !reRev.test(r.description)) freq[r.description] = (freq[r.description] || 0) + 1;
      const bud = {}; for(const b of D.budgets) if(b.mois === "") bud[catNom(b.categorie_id)] = b.montant_c / 100;
      return {mois: infos(), categories: D.categories.filter(c => c.active).sort((a, b) => a.ordre - b.ordre).map(c => c.nom),
        types: TYPES().map(x => x[1]), type_cent_a_faire: false, correction_part_a_faire: false,
        reglages: Object.assign({prenom: "Zach", partenaire: "Marika", pourboires: true, revenu_principal: "Revenu Fiorellino",
          descriptions_revenu_principal: [], mots_revenu_principal: "", configuration_a_faire: false}, reg),
        budgets: bud, date_demo: opts.date_demo || null, ui_defaut: opts.ui_defaut || "pro",
        types_invest: D.comptes_placement.slice().sort((a, b) => a.ordre - b.ordre).map(c => c.nom),
        motscles: D.mots_cles.slice().sort((a, b) => a.ordre - b.ordre).map(k => ({mot: k.mot, categorie: catNom(k.categorie_id)})),
        fichier: opts.fichier || "Finances.json", excel_ouvert: false, prochain_mois: null, v2: true,
        descriptions_autres_revenus: Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 8).map(x => x[0])};
    }

    // ------------------------------------------------------------------ lecture : un mois
    function moisHistorique(i){
      const m = M(), ch = m.chaine()[i.cle];
      const lignes = D.depenses.filter(d => d.mois === i.cle).sort((a, b) => a.ordre - b.ordre)
        .map(d => ({date: d.date, description: d.description, montant: d.part_hist, categorie: m.categorie(d)}));
      const pl = m.placements(i.cle);
      return {feuille: i.feuille, nom: i.nom, groupe: i.groupe, lignes, total_depenses: ch.totalDepenses, somme_lignes: ch.totalDepenses,
        par_categorie: Object.entries(ch.parCategorie).map(([categorie, montant]) => ({categorie, montant})),
        total_revenus: ch.revenus, revenu_net: ch.revenuNet, solde_debut: ch.debut, retrait: ch.epargne, solde_fin: ch.fin,
        investissements: pl.map(x => ({type: x.compte, investi: x.investi, jv: x.valeur})),
        total_investi: pl.reduce((a, x) => a + x.investi, 0), total_jv: pl.reduce((a, x) => a + x.valeur, 0)};
    }
    function mois(f){
      const i = info(f); if(!i.canvas) return moisHistorique(i);
      const m = M(), reg = REG(), P = partenaire(), ch = m.chaine()[i.cle], mo = i.cle;
      const rg = new Map(D.reglements.map(r => [r.id, r]));
      const rep = D.repartition.find(r => r.mois === mo) || {};
      const deps = D.depenses.filter(d => d.mois === mo).sort((a, b) => a.ordre - b.ordre);
      const reglsMois = D.reglements.filter(r => r.date.slice(0, 7) === mo || deps.some(d => d.ordre === r.ordre)).sort((a, b) => a.ordre - b.ordre);
      const marq = reglsMois.length ? reglsMois[reglsMois.length - 1] : null;
      const depMarq = marq ? (deps.filter(d => d.ordre <= marq.ordre).pop() || null) : null;
      const transactions = deps.map(d => {
        const part = d.paye_par === "partenaire", ouvert = !d.reglement_id;
        const desc = part ? d.description + (ouvert ? ` (carte ${P} – à rembourser)` : ` (carte ${P})`) : d.description;
        const auto = m.categorie(Object.assign({}, d, {categorie_id: null}));
        const prop = part ? (ouvert ? null : 1) : {P: 1, "50": 0.5, SP: 1 - m.pct(mo), R: 1 - m.pct(mo), M: 0}[d.type];
        return {ligne: d.id, utilise: true, date: d.date, date_brute: null, description: desc,
          type: part ? (ouvert ? "" : "Personnel") : libType(d.type), type_exact: true,
          deboursement: (d.montant_c || 0) / 100, deb_formule: d.calcul ? "=" + d.calcul : null, proportion: prop,
          depense: part && ouvert ? null : m.maPart(d),
          categorie: m.categorie(d), cat_auto: auto, cat_manuelle: !!d.categorie_id && catNom(d.categorie_id) !== auto,
          marika: part ? 0 : m.partPartenaire(d), marqueur: !!(depMarq && depMarq.id === d.id), erreur: false,
          a_rembourser: part && ouvert, regle: !!d.reglement_id, paye_par: d.paye_par, ordre: d.ordre, fixe: !!d.fixe_id, estime: !!d.estime};
      });
      // solde avec le partenaire, présenté comme dans la v1 : ce mois-ci + reports des mois précédents
      const ouvertes = x => !x.reglement_id && !x.historique;
      const avant = D.depenses.filter(d => d.mois < mo && ouvertes(d));
      const autresAvant = D.autres_montants.filter(x => x.date.slice(0, 7) < mo && ouvertes(x));
      const reportAuto = avant.filter(d => d.paye_par !== "partenaire").reduce((a, d) => a + m.partPartenaire(d), 0) + autresAvant.reduce((a, x) => a + x.montant_c / 100, 0);
      const reports = [];
      // mois précédents réglés par un règlement fait ce mois-ci (pour « Règlements du mois »)
      const reglesIci = new Set(D.reglements.filter(r => deps.some(d => d.ordre === r.ordre)).map(r => r.id));
      const regleAvant = D.depenses.filter(d => d.mois < mo && !d.historique && d.paye_par !== "partenaire" && reglesIci.has(d.reglement_id)).reduce((a, d) => a + m.partPartenaire(d), 0)
        + D.autres_montants.filter(x => x.date.slice(0, 7) < mo && reglesIci.has(x.reglement_id)).reduce((a, x) => a + x.montant_c / 100, 0);
      if(Math.abs(regleAvant) > 0.004) reports.push({ligne: null, libelle: "Portion " + nomMois(moisPrec(mo)), montant: 0, auto: true, formule: null, regle: true, montant_origine: regleAvant});
      if(Math.abs(reportAuto) > 0.004) reports.push({ligne: null, libelle: "Portion " + nomMois(moisPrec(mo)), montant: reportAuto, auto: true, formule: null, regle: false, montant_origine: null});
      for(const x of D.autres_montants.filter(x => x.date.slice(0, 7) === mo).sort((a, b) => a.ordre - b.ordre))
        reports.push({ligne: x.id, libelle: x.libelle, montant: x.reglement_id ? 0 : x.montant_c / 100, auto: false, formule: null,
                      regle: !!x.reglement_id, montant_origine: x.reglement_id ? x.montant_c / 100 : null});
      const duDepuis = deps.filter(d => ouvertes(d) && d.paye_par !== "partenaire").reduce((a, d) => a + m.partPartenaire(d), 0);
      const duTotal = duDepuis + reports.reduce((a, x) => a + (x.montant || 0), 0);
      const aRembMois = deps.filter(d => d.paye_par === "partenaire" && ouvertes(d)).reduce((a, d) => a + d.montant_c / 100, 0);
      const aRembAvant = avant.filter(d => d.paye_par === "partenaire").reduce((a, d) => a + d.montant_c / 100, 0);
      // pourboires et revenus
      const pbs = D.pourboires.filter(p => p.mois === mo).sort((a, b) => (a.date || "").localeCompare(b.date || "") || a.id - b.id);
      const pourboires = pbs.map(p => ({ligne: p.id, utilise: true, date: p.date, recu: p.especes_c == null ? null : p.especes_c / 100,
        a_recevoir: p.carte_c == null ? null : p.carte_c / 100, heures: p.heures, recu_txt: null, a_recevoir_txt: null,
        statut: p.recu === 1 ? "Oui" : p.recu === 0 ? "Non" : ""}));
      const reRev = new RegExp(reg.mots_revenu_principal || "^$", "i");
      const revs = D.revenus.filter(r => r.mois === mo).sort((a, b) => (a.date || "").localeCompare(b.date || "") || a.id - b.id);
      const revenus = revs.map(r => {
        const l = D.revenus_pourboires.filter(x => x.revenu_id === r.id).map(x => x.pourboire_id);
        return {ligne: r.id, utilise: true, date: r.date, montant: r.montant_c / 100, montant_txt: r.calcul ? "=" + r.calcul : null, description: r.description,
                cases: r.cases_colonne ? {colonne: r.cases_colonne === "especes" ? "recu" : "arec", lignes: l, ajout: (r.cases_ajout_c || 0) / 100} : null,
                fixe: !!r.fixe_id, estime: !!r.estime};
      });
      // placements et épargne
      const pl = m.placements(mo);
      const investissements = pl.map(x => { const c = D.comptes_placement.find(y => y.id === x.id);
        return {ligne: x.id, type: x.compte, investi: x.investi, jv: x.valeur, variation: x.variation, investi_auto: !x.manuel,
                valeur_fixe: !!c.valeur_fixe, investi_auto_possible: x.auto != null}; });
      const retraits = D.comptes_placement.slice().sort((a, b) => a.ordre - b.ordre).map(c => ({libelle: `(-) Épargne (${c.nom})`, compte: c.nom, adr: String(c.id),
        montant: D.versements_epargne.filter(v => v.mois === mo && v.compte_id === c.id).reduce((a, v) => a + v.montant_c / 100, 0)}));
      // budgets
      const bm = m.budgets(mo); const de = m.budgetsDe(mo);
      const rel = D.releves_banque.filter(r => r.date.slice(0, 7) === mo).sort((a, b) => a.date.localeCompare(b.date)).pop();
      const totRev = revs.reduce((a, r) => a + r.montant_c / 100, 0);
      const out = {feuille: i.feuille, nom: i.nom, groupe: i.groupe, modifiable: i.modifiable,
        pct_marika: m.pct(mo), pct_manuel: !!rep.pct_manuel, revenu_marika: rep.revenu_partenaire_c == null ? null : rep.revenu_partenaire_c / 100,
        revenu_zach: rep.revenu_moi_c == null ? null : rep.revenu_moi_c / 100, pct_repartition: m.pct(mo),
        transactions, total_depenses: ch.totalDepenses, total_split: transactions.reduce((a, t) => a + (t.marika || 0), 0),
        du_depuis_reglement: duDepuis, reports, du_total: duTotal, total_a_rembourser: r2(aRembMois + aRembAvant), a_rembourser_reporte: r2(aRembAvant),
        du_net: r2(duTotal - aRembMois - aRembAvant), ligne_dernier_reglement: depMarq ? depMarq.id : null,
        par_categorie: Object.entries(ch.parCategorie).map(([categorie, montant]) => ({categorie, montant})), hors_resume: {},
        pourboires, total_pourboires: {recu: ch.pourboires.especes, a_recevoir: ch.pourboires.carte, heures: ch.pourboires.heures}, colonne_heures: true,
        prevus: D.prevus.filter(x => x.mois === mo).sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id).map(prevuVue),
        revenus, total_revenus: totRev, total_fiorellino: revs.filter(r => reRev.test(r.description || "")).reduce((a, r) => a + r.montant_c / 100, 0),
        investissements, total_investi: pl.reduce((a, x) => a + x.investi, 0), total_jv: pl.reduce((a, x) => a + x.valeur, 0),
        budgets: bm, budgets_de: de ? nomMois(de) : null,
        solde_debut: ch.debut, retrait: ch.epargne, retraits, revenu_net: ch.revenuNet, solde_fin: ch.fin, solde_fin_standard: ch.fin,
        solde_reel: rel ? rel.solde_c / 100 : null, controles: [], resume_controles: "", nb_a_verifier: 0, ajustements: [],
        mois_param: i.nom, lignes_libres: 9999};
      out.financements = m.financements(mo).map(f => ({ligne: f.id, nom: f.nom, recu: f.recu, recu_mois: f.recuMois, solde_du: f.soldeDu, solde_du_date: f.soldeDuDate,
        recus: f.recus.map(x => ({ligne: x.id, date: x.date, montant: x.montant_c / 100, note: x.note}))}));
      out.financement_recu_mois = ch.financement || 0;
      if(mo === auj().slice(0, 7)) out.soldes_jour = soldesJour(m);
      return out;
    }
    // ------------------------------------------------------------------ soldes du jour, concordance avec la banque, pistes d'écart
    function soldesJour(m){
      m = m || M(); const a = auj(), s = m.soldes(a);
      const rel = D.releves_banque.filter(r => r.origine === "app" && r.date <= a).sort((x, y) => x.date.localeCompare(y.date) || x.id - y.id).pop() || null;
      const conc = rel ? m.concordance(rel) : null;
      const precOk = D.releves_banque.filter(r => r.origine === "app" && rel && r.date < rel.date).sort((x, y) => y.date.localeCompare(x.date))
        .find(r => Math.abs(m.concordance(r).ecart) < 0.005) || null;
      return {date: a, app: s.app, au_compte: s.auCompte, attendu: s.attendu,
        revenus_a_venir: s.revFut.map(r => ({ligne: r.id, date: r.date, description: r.description, montant: r.montant_c / 100})), total_revenus_a_venir: s.sR,
        depenses_a_venir: s.depFut.map(d => ({ligne: d.id, date: d.date, description: d.description, montant: d.montant_c / 100})), total_depenses_a_venir: s.sD,
        partenaire: s.partenaire, pourboires_attente: s.pbAttente.map(p => ({ligne: p.id, date: p.date, montant: (p.carte_c || 0) / 100})), total_pourboires_attente: s.sPb,
        releve: rel && {ligne: rel.id, date: rel.date, solde: rel.solde_c / 100},
        concordance: conc && Object.assign(conc, {derniere_concordance: precOk ? precOk.date : null, pistes: Math.abs(conc.ecart) >= 0.005 ? pistes(conc, precOk, m) : []})};
    }
    function pistes(conc, precOk, m){
      // pistes pour expliquer un écart entre le solde de la banque et le solde au compte calculé
      const e = Math.abs(conc.ecart), out = [], eg = (x, y) => Math.abs(x - y) < 0.005;
      const depuis = precOk ? precOk.date : (() => { const d = new Date(conc.date); d.setDate(d.getDate() - 45); return d.toISOString().slice(0, 10); })();
      const periode = x => x.date && x.date > depuis && x.date <= conc.date;
      const P = partenaire(), dl = x => x.description + " (" + x.date + ")";
      const deps = D.depenses.filter(d => !d.historique && d.paye_par === "moi" && periode(d));
      // banque plus basse que l'app : une dépense manque peut-être dans l'app (ou un revenu est en trop) ; et inversement
      for(const d of deps){
        const mt = (d.montant_c || 0) / 100;
        if(eg(mt, e)) out.push({ligne: d.id, texte: conc.ecart > 0 ? `Même montant que « ${dl(d)} » : saisie en double ?` : `Même montant que « ${dl(d)} » : une autre dépense du même montant manque peut-être.`});
        if(eg(mt * 2, e)) out.push({ligne: d.id, texte: `Le double de « ${dl(d)} » : montant inscrit avec le mauvais signe ?`});
        if(d.type !== "P" && eg(mt - m.maPart(d), e)) out.push({ligne: d.id, texte: `La part de ${P} de « ${dl(d)} » : type de partage à vérifier.`});
      }
      for(const r of D.revenus.filter(r => !r.historique && periode(r))) if(eg(r.montant_c / 100, e))
        out.push({ligne: r.id, texte: conc.ecart < 0 ? `Même montant que le revenu « ${dl(r)} » : pas encore arrivé à la banque ?` : `Même montant que le revenu « ${dl(r)} » : inscrit deux fois ?`});
      for(const p of D.pourboires.filter(p => periode(p))) if(eg((p.carte_c || 0) / 100, e))
        out.push({ligne: p.id, texte: `Pourboires carte du ${p.date} : statut « reçu » à vérifier.`});
      for(const d of D.depenses.filter(d => !d.historique && d.paye_par === "moi" && d.date > conc.date)) if(eg((d.montant_c || 0) / 100, e))
        out.push({ligne: d.id, texte: `« ${dl(d)} » est datée après le relevé : la date est-elle juste ?`});
      const net = m.detailPartenaire(conc.date).net;
      if(Math.abs(net) >= 0.005 && eg(Math.abs(net), e)) out.push({texte: `Égal au solde avec ${P} : le virement du règlement est-il déjà passé ?`});
      out.push({texte: "Le rapprochement bancaire (Plus › Rapprochement) compare ton relevé ligne par ligne avec l'app."});
      return out;
    }
    function tendances(){
      const m = M(), ch = m.chaine();
      return infos().map(i => { const x = ch[i.cle], pl = m.placements(i.cle);
        return {feuille: i.feuille, nom: i.nom, cle: i.cle, groupe: i.groupe, total_depenses: x.totalDepenses, par_categorie: x.parCategorie,
                revenus: x.revenus, revenu_net: x.revenuNet, solde_fin: x.fin, investi: pl.reduce((a, y) => a + y.investi, 0),
                jv: pl.reduce((a, y) => a + y.valeur, 0), epargne_placee: x.epargne}; });
    }

    // ------------------------------------------------------------------ écritures
    function ecritDepense(f, data){
      const i = modifiable(f);
      let desc = String(data.description || "").trim(); if(!desc) throw new Erreur("La description est obligatoire.");
      let paye = "moi"; const mr = RX_A_REMB.exec(desc);
      if(mr){ paye = "partenaire"; desc = desc.replace(RX_A_REMB, "").trim(); }
      const tl = String(data.type || "").trim();
      let typ = codeType(tl); if(!typ){ if(paye === "partenaire" || !tl) typ = "P"; else throw new Erreur("Type inconnu : " + tl); }
      const mt = calcule(data.montant); if(mt == null) throw new Erreur("Le montant est obligatoire.");
      const date = data.date || null;
      const catTxt = String(data.categorie || "").trim();
      const m = M(), auto = m.categorie({description: desc, categorie_id: null});
      if(catTxt && !catId(catTxt)) throw new Erreur("Catégorie inconnue : " + catTxt);
      if(catTxt && norm(catTxt) !== norm(auto) && REG().pourboires) apprendMotCle(desc, catTxt);
      const auto2 = M().categorie({description: desc, categorie_id: null});
      const champs = {mois: moisDeDate(date, i) < verrou() ? i.cle : moisDeDate(date, i), date, description: desc, montant_c: cents(mt.valeur), calcul: mt.calcul,
                      type: typ, categorie_id: catTxt && norm(catTxt) !== norm(auto2) ? catId(catTxt) : null, paye_par: paye};
      if(champs.mois < verrou()) throw new Erreur("Date dans un mois en lecture seule.");
      if(data.ligne != null){
        const d = D.depenses.find(x => x.id === +data.ligne); if(!d) throw new Erreur("Dépense introuvable");
        const avant = Object.assign({}, d); Object.assign(d, champs); if(d.estime) d.estime = 0;     // enregistrée = montant confirmé
        journal("depenses", d.id, "modification", avant, d); ajusteEstimation(d.fixe_id, D.depenses); return d.id;
      }
      const d = Object.assign({id: nouvelId("depenses"), reglement_id: null, id_partage: null, historique: 0, part_hist: null,
                               ordre: ordreSuivant(), origine: "app"}, champs);
      D.depenses.push(d); journal("depenses", d.id, "ajout", null, d); return d.id;
    }
    function apprendMotCle(desc, categorie){
      const mot = desc.replace(RX_A_REMB, "").trim();
      if(mot.length < 3 || !/[A-Za-zÀ-ÿ]/.test(mot) || !catId(categorie)) return;
      const k = D.mots_cles.find(x => norm(x.mot) === norm(mot));
      if(k){ if(k.categorie_id !== catId(categorie)){ const av = Object.assign({}, k); k.categorie_id = catId(categorie); journal("mots_cles", k.id, "modification", av, k); } return; }
      insereMotCle(mot, categorie);
    }
    function insereMotCle(mot, categorie){
      // les 3 derniers mots-clés (qui l'emportent) restent en bas : le nouveau prend la place juste avant eux
      const l = D.mots_cles.slice().sort((a, b) => a.ordre - b.ordre);
      const pos = Math.max(0, l.length - 3);
      l.forEach((x, j) => { x.ordre = j + 1 + (j >= pos ? 1 : 0); });
      const k = {id: nouvelId("mots_cles"), mot, categorie_id: catId(categorie), ordre: pos + 1};
      D.mots_cles.push(k); journal("mots_cles", k.id, "ajout", null, k); return k.id;
    }
    function supprime(table, id, controle){
      const j = D[table].findIndex(x => x.id === +id); if(j < 0) throw new Erreur("Ligne introuvable");
      if(controle) controle(D[table][j]);
      const x = D[table][j]; D[table].splice(j, 1);
      // un revenu calculé à partir de jours de pourboires : on garde les liens pour pouvoir annuler ; un jour de pourboires
      // supprimé est retiré des revenus qui le comptaient (leur montant est recalculé)
      const liens = table === "revenus" ? D.revenus_pourboires.filter(y => y.revenu_id === x.id).map(y => y.pourboire_id)
                  : table === "pourboires" ? D.revenus_pourboires.filter(y => y.pourboire_id === x.id).map(y => y.revenu_id) : null;
      journal(table, x.id, "suppression", liens ? Object.assign({_liens: liens}, x) : x, null);
      if(table === "revenus") D.revenus_pourboires = D.revenus_pourboires.filter(y => y.revenu_id !== x.id);
      if(table === "pourboires") D.revenus_pourboires = D.revenus_pourboires.filter(y => y.pourboire_id !== x.id);
      return x.id;
    }
    function reglement(f, ligne){
      const i = modifiable(f);
      // on déplace le règlement du mois s'il y en a déjà un (comme le marqueur « Solde réglé » de la v1)
      const deps = D.depenses.filter(d => d.mois === i.cle);
      const dernier = D.reglements.slice().sort((a, b) => b.ordre - a.ordre)[0];
      if(dernier && deps.some(d => d.ordre === dernier.ordre) && dernier.source !== "import"){
        for(const t of ["depenses", "autres_montants"]) for(const x of D[t]) if(x.reglement_id === dernier.id) x.reglement_id = null;
        D.reglements.splice(D.reglements.indexOf(dernier), 1); journal("reglements", dernier.id, "suppression", dernier, null);
      }
      if(ligne == null) return null;
      const d = D.depenses.find(x => x.id === +ligne); if(!d) throw new Erreur("Dépense introuvable");
      const r = {id: nouvelId("reglements"), date: auj(), ordre: d.ordre, confirme: 1, source: "app"};
      D.reglements.push(r);
      for(const t of ["depenses", "autres_montants"])
        // dépenses partagées jusqu'à la ligne choisie ; tous les autres montants ouverts (comme dans la v1 : « Autres montants → 0 »)
        for(const x of D[t]) if(!x.reglement_id && !x.historique && (t === "autres_montants" || (x.ordre <= d.ordre && (x.type !== "P" || x.paye_par === "partenaire")))) x.reglement_id = r.id;
      journal("reglements", r.id, "ajout", null, r); return r.id;
    }
    function recalculeRevenusCases(){
      for(const r of D.revenus) if(r.cases_colonne){
        const ids = new Set(D.revenus_pourboires.filter(x => x.revenu_id === r.id).map(x => x.pourboire_id));
        const s = D.pourboires.filter(p => ids.has(p.id)).reduce((a, p) => a + ((r.cases_colonne === "especes" ? p.especes_c : p.carte_c) || 0), 0);
        r.montant_c = s + (r.cases_ajout_c || 0);
      }
    }
    function ecritRevenu(f, data){
      const i = modifiable(f);
      const champs = {mois: i.cle, date: data.date || null, description: String(data.description || "").trim()};
      let liens = null;
      const exist = data.ligne != null ? D.revenus.find(x => x.id === +data.ligne) : null;
      if(exist && exist.cases_colonne && !data.cases && (data.montant == null || data.montant === "" || /^=?SUM\(/i.test(String(data.montant)))){
        // revenu calculé à partir des jours de pourboires : seules la date et la description changent
        const av = Object.assign({}, exist); Object.assign(exist, champs); journal("revenus", exist.id, "modification", av, exist); recalculeRevenusCases(); return exist.id;
      }
      if(data.cases){
        const col = data.colonne === "recu" ? "especes" : "carte";
        const pbs = data.cases.map(Number); if(!pbs.length) throw new Erreur("Choisis au moins un jour.");
        for(const id of pbs) if(!D.pourboires.some(p => p.id === id && p.mois === i.cle)) throw new Erreur("Jour de pourboires introuvable dans ce mois");
        const aj = data.ajout ? calcule(data.ajout) : null;
        Object.assign(champs, {cases_colonne: col, cases_ajout_c: aj ? cents(aj.valeur) : 0, calcul: null, montant_c: 0}); liens = pbs;
      } else {
        const mt = calcule(data.montant); if(mt == null) throw new Erreur("Le montant est obligatoire.");
        Object.assign(champs, {montant_c: cents(mt.valeur), calcul: mt.calcul, cases_colonne: null, cases_ajout_c: 0});
      }
      let r;
      if(data.ligne != null){ r = D.revenus.find(x => x.id === +data.ligne); if(!r) throw new Erreur("Revenu introuvable");
        const av = Object.assign({}, r); Object.assign(r, champs); if(r.estime) r.estime = 0; journal("revenus", r.id, "modification", av, r); ajusteEstimation(r.fixe_id, D.revenus); }
      else { r = Object.assign({id: nouvelId("revenus"), historique: 0, origine: "app"}, champs); D.revenus.push(r); journal("revenus", r.id, "ajout", null, r); }
      if(liens){ D.revenus_pourboires = D.revenus_pourboires.filter(x => x.revenu_id !== r.id).concat(liens.map(p => ({revenu_id: r.id, pourboire_id: p}))); }
      else D.revenus_pourboires = D.revenus_pourboires.filter(x => x.revenu_id !== r.id);
      recalculeRevenusCases(); return r.id;
    }
    function ecritPourboire(f, data){
      const i = modifiable(f);
      const v = x => { const c = calcule(x); return c == null ? null : cents(c.valeur); };
      const st = norm(data.statut);
      const champs = {mois: i.cle, date: data.date || null, especes_c: v(data.recu), carte_c: v(data.a_recevoir),
                      heures: data.heures === "" || data.heures == null ? null : parseFloat(String(data.heures).replace(",", ".")),
                      recu: st === "oui" ? 1 : st === "non" ? 0 : null};
      let p;
      if(data.ligne != null){ p = D.pourboires.find(x => x.id === +data.ligne); if(!p) throw new Erreur("Ligne introuvable");
        const av = Object.assign({}, p); Object.assign(p, champs); journal("pourboires", p.id, "modification", av, p); }
      else { p = Object.assign({id: nouvelId("pourboires"), origine: "app"}, champs); D.pourboires.push(p); journal("pourboires", p.id, "ajout", null, p); }
      recalculeRevenusCases(); return p.id;
    }
    function parametres(f, data){
      const i = modifiable(f), mo = i.cle;
      let rep = D.repartition.find(r => r.mois === mo);
      if(!rep){ const pr = D.repartition.filter(r => r.mois < mo).sort((a, b) => b.mois.localeCompare(a.mois))[0] || {pct_partenaire: 0};
        rep = {mois: mo, revenu_moi_c: null, revenu_partenaire_c: null, pct_partenaire: pr.pct_partenaire, pct_manuel: pr.pct_manuel || 0}; D.repartition.push(rep); }
      const av = Object.assign({}, rep);
      if(data.revenu_marika != null && data.revenu_marika !== "") rep.revenu_partenaire_c = cents(calcule(data.revenu_marika).valeur);
      if(data.revenu_zach != null && data.revenu_zach !== "") rep.revenu_moi_c = cents(calcule(data.revenu_zach).valeur);
      if((data.revenu_marika != null || data.revenu_zach != null) && rep.revenu_moi_c != null && rep.revenu_partenaire_c != null){
        const t = rep.revenu_moi_c + rep.revenu_partenaire_c; rep.pct_partenaire = t ? rep.revenu_partenaire_c / t : 0; rep.pct_manuel = 0; }
      if(data.revenu_marika != null || data.revenu_zach != null) journal("repartition", null, "modification", av, rep);
      for(const [adr, val] of Object.entries(data.retraits || {})){
        const cid = +adr; D.versements_epargne = D.versements_epargne.filter(v => !(v.mois === mo && v.compte_id === cid));
        const c = calcule(val); if(c && c.valeur) D.versements_epargne.push({id: nouvelId("versements_epargne"), mois: mo, date: finMois(mo), compte_id: cid, montant_c: cents(c.valeur), origine: "app"});
      }
      if(data.solde_reel != null && data.solde_reel !== ""){
        const c = calcule(data.solde_reel); const date = mo === auj().slice(0, 7) ? auj() : finMois(mo);
        D.releves_banque = D.releves_banque.filter(r => r.date !== date);
        const r = {id: nouvelId("releves_banque"), date, solde_c: cents(c.valeur), origine: "app"};
        D.releves_banque.push(r); journal("releves_banque", r.id, "ajout", null, r);
      }
      return true;
    }
    const finMois = mo => { const [a, m] = mo.split("-").map(Number); return mo + "-" + String(new Date(a, m, 0).getDate()).padStart(2, "0"); };
    function investissement(f, data){
      const i = modifiable(f), cid = +data.ligne;
      let x = D.placements_mois.find(y => y.compte_id === cid && y.mois === i.cle);
      if(!x){ x = {compte_id: cid, mois: i.cle, valeur_c: null, investi_c: null, investi_manuel: 0}; D.placements_mois.push(x); }
      const av = Object.assign({}, x);
      if(data.jv != null && data.jv !== "") x.valeur_c = cents(calcule(data.jv).valeur);
      if(data.investi_auto) { x.investi_manuel = 0; x.investi_c = null; }
      else if(data.investi != null && data.investi !== ""){ x.investi_manuel = 1; x.investi_c = cents(calcule(data.investi).valeur); }
      journal("placements_mois", cid, "modification", av, x); return true;
    }

    // ------------------------------------------------------------------ dépenses et revenus fixes (essai, app de Zach)
    // Une ligne fixe = un montant qui revient chaque mois (loyer, cellulaire, paie…). « Inscrire chaque mois » (actif) :
    // l'app ajoute la ligne du mois à sa date (datée d'avance → « à venir » jusqu'au jour J). Montant variable : la ligne
    // est inscrite avec un montant estimé (« estimé ») jusqu'à ce que tu inscrives le vrai montant.
    const cleFixe = t => String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()
      .replace(/\(carte [^)]*\)/g, "").replace(/[^a-z]+/g, " ").trim();
    const finDuMois = mo => { const [a, m] = mo.split("-").map(Number); return new Date(a, m, 0).getDate(); };
    const dateFixe = (mo, jour) => mo + "-" + String(Math.min(Math.max(1, jour || 1), finDuMois(mo))).padStart(2, "0");
    const partDe = (typ, pct) => ({P: 1, "50": 0.5, SP: 1 - pct, M: 0}[typ] ?? 1);
    function fixeVue(x, pct){
      return {id: x.id, genre: x.genre, description: x.description, montant: x.montant_c / 100, variable: !!x.variable, jour: x.jour,
              categorie: x.genre === "depense" ? (catNom(x.categorie_id) || M().categorie({description: x.description, categorie_id: null})) : null,
              type: x.genre === "depense" ? libType(x.type || "P") : null, type_code: x.type || "P", actif: !!x.actif,
              ma_part: x.genre === "depense" ? x.montant_c / 100 * partDe(x.type || "P", pct) : x.montant_c / 100, dernier: x.dernier || null};
    }
    function suggestionsFixes(mo){
      const ign = new Set(REG().fixes_ignores || []), deja = new Set(D.fixes.map(x => x.genre + ":" + cleFixe(x.description)));
      const derniers = []; for(let k = 1, m = mo; k <= 4; k++){ m = (() => { let [a, n] = m.split("-").map(Number); n--; if(n < 1){ n = 12; a--; } return a + "-" + String(n).padStart(2, "0"); })(); derniers.push(m); }
      const out = [];
      const analyse = (genre, lignes) => {
        const g = new Map();
        for(const l of lignes){ const c = cleFixe(l.description); if(c.length < 3) continue; if(!g.has(c)) g.set(c, []); g.get(c).push(l); }
        for(const [c, l] of g){
          const parMois = new Map(); for(const x of l){ if(!parMois.has(x.mois)) parMois.set(x.mois, []); parMois.get(x.mois).push(x); }
          if(parMois.size < 3 || [...parMois.values()].some(v => v.length > 2)) continue;
          if(deja.has(genre + ":" + c) || ign.has(genre + ":" + c)) continue;
          const mts = [...parMois.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([, v]) => v.reduce((a, x) => a + x.mt, 0));
          const moy = mts.reduce((a, x) => a + x, 0) / mts.length, ecart = Math.max(...mts) - Math.min(...mts);
          // fixe si les 2 derniers mois ont le même montant (un loyer augmenté en juillet reste un fixe)
          const variable = Math.abs(mts[mts.length - 1] - mts[mts.length - 2]) > Math.max(0.5, moy * 0.01);
          if(variable && ecart > moy * 0.5) continue;        // trop irrégulier pour être un fixe (restaurant, magasinage…)
          const der = l.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).pop();
          const derPaye = l.filter(x => x.paye != null).sort((a, b) => (a.date || "").localeCompare(b.date || "")).pop();
          const jours = l.map(x => +(x.date || "").slice(8, 10)).filter(Boolean).sort((a, b) => a - b);
          // la série compare « ma part » (seule connue pour l'historique) ; le montant proposé = montant payé de la dernière
          // ligne détaillée (avec son type de partage), mis à l'échelle si le montant varie
          let est = variable ? mts.slice(-3).reduce((a, x) => a + x, 0) / Math.min(3, mts.length) : mts[mts.length - 1];
          let typ = "P";
          if(derPaye && derPaye.mt > 0.004){ est = est * derPaye.paye / derPaye.mt; typ = derPaye.type || "P"; }
          out.push({cle: genre + ":" + c, genre, description: der.description, montant: Math.round(est * 100) / 100, variable,
                    jour: jours.length ? jours[Math.floor(jours.length / 2)] : 1, mois: parMois.size,
                    categorie: genre === "depense" ? M().categorie(der) : null, type: genre === "depense" ? libType(typ) : null});
        }
      };
      const m = M();
      analyse("depense", D.depenses.filter(d => derniers.includes(d.mois) && d.paye_par !== "partenaire" && !d.fixe_id)
        .map(d => Object.assign({}, d, {mt: m.maPart(d), paye: d.historique ? null : (d.montant_c || 0) / 100})));
      analyse("revenu", D.revenus.filter(r => derniers.includes(r.mois) && !r.fixe_id && !r.historique && !/pourboire/i.test(r.description || "")
        && !D.revenus_pourboires.some(x => x.revenu_id === r.id)).map(r => Object.assign({}, r, {mt: (r.montant_c || 0) / 100, paye: (r.montant_c || 0) / 100})));
      return out.sort((a, b) => b.montant - a.montant).slice(0, 10);
    }
    // montant variable : l'estimation suit la moyenne des 3 derniers vrais montants (les prochains mois sont mieux estimés)
    function ajusteEstimation(fid, table){
      const x = fid && D.fixes.find(y => y.id === fid); if(!x || !x.variable) return;
      const l = table.filter(y => y.fixe_id === fid && !y.estime && y.montant_c > 0).sort((a, b) => (a.date || "").localeCompare(b.date || "")).slice(-3);
      if(l.length) x.montant_c = Math.round(l.reduce((a, y) => a + y.montant_c, 0) / l.length);
    }
    function aGenerer(mo){ return D.fixes.filter(x => x.actif && (!x.dernier || x.dernier < mo) && (!x.depuis || x.depuis <= mo)); }
    function fixes(){
      const mo = auj().slice(0, 7), pct = M().pct(mo);
      const l = D.fixes.slice().sort((a, b) => (a.jour || 0) - (b.jour || 0) || a.id - b.id).map(x => fixeVue(x, pct));
      const dep = l.filter(x => x.genre === "depense"), rev = l.filter(x => x.genre === "revenu");
      return {fixes: l, suggestions: suggestionsFixes(mo), a_generer: aGenerer(mo).length, mois: mo,
              total_depenses: dep.reduce((a, x) => a + x.ma_part, 0), total_revenus: rev.reduce((a, x) => a + x.montant, 0)};
    }
    function ecritFixe(data){
      const genre = data.genre === "revenu" ? "revenu" : "depense";
      const desc = String(data.description || "").trim(); if(!desc) throw new Erreur("La description est obligatoire.");
      const mt = calcule(data.montant); if(!mt || !(mt.valeur > 0)) throw new Erreur(data.variable ? "Indique un montant estimé." : "Indique le montant.");
      const jour = Math.min(31, Math.max(1, parseInt(data.jour, 10) || 1));
      const champs = {genre, description: desc, montant_c: cents(mt.valeur), variable: data.variable ? 1 : 0, jour, actif: data.actif ? 1 : 0};
      if(genre === "depense"){
        const catTxt = String(data.categorie || "").trim();
        champs.categorie_id = catTxt && norm(catTxt) !== norm(M().categorie({description: desc, categorie_id: null})) ? catId(catTxt) : null;
        champs.type = codeType(String(data.type || "")) || "P";
      }
      const mo = auj().slice(0, 7);
      if(data.id != null){
        const x = D.fixes.find(y => y.id === +data.id); if(!x) throw new Erreur("Ligne fixe introuvable");
        const av = Object.assign({}, x); Object.assign(x, champs); majPrevus(x, !x.actif);
        if(x.actif && !av.actif) x.depuis = +jour >= +auj().slice(8, 10) ? mo : moisSuivant(mo);   // réactivée : à partir de la prochaine date
        journal("fixes", x.id, "modification", av, x); return x.id;
      }
      // nouvelle ligne : inscrite dès ce mois si sa date n'est pas passée ; sinon à partir du mois prochain
      const x = Object.assign({id: nouvelId("fixes"), depuis: jour >= +auj().slice(8, 10) ? mo : moisSuivant(mo), dernier: null}, champs);
      if(data.cle) setReg("fixes_ignores", (REG().fixes_ignores || []).filter(k => k !== data.cle));
      D.fixes.push(x); journal("fixes", x.id, "ajout", null, x); return x.id;
    }
    // chaque mois, une prévision par ligne fixe active (« Dépenses / Revenus à venir ») : elle ne compte nulle part tant que
    // tu ne l'as pas cochée « passée » ; cochée, elle devient une vraie dépense (ou un vrai revenu), date et montant modifiables.
    function genereFixes(){
      const mo = auj().slice(0, 7); let n = 0;
      if(mo < verrou()) return 0;
      for(const x of aGenerer(mo)){
        const pv = {id: nouvelId("prevus"), fixe_id: x.id, genre: x.genre, mois: mo, date: dateFixe(mo, x.jour), description: x.description,
                    montant_c: x.montant_c, estime: x.variable ? 1 : 0, categorie_id: x.categorie_id || null, type: x.type || null};
        D.prevus.push(pv); journal("prevus", pv.id, "ajout", null, pv);
        x.dernier = mo; n++;
      }
      return n;
    }
    function prevuVue(pv){
      const pct = M().pct(pv.mois);
      return {id: pv.id, fixe_id: pv.fixe_id, genre: pv.genre, date: pv.date, description: pv.description, montant: pv.montant_c / 100, estime: !!pv.estime,
              categorie: pv.genre === "depense" ? (catNom(pv.categorie_id) || M().categorie({description: pv.description, categorie_id: null})) : null,
              type: pv.genre === "depense" ? libType(pv.type || "P") : null, type_code: pv.type || "P",
              ma_part: pv.genre === "depense" ? pv.montant_c / 100 * partDe(pv.type || "P", pct) : pv.montant_c / 100};
    }
    const champsPrevu = data => {
      const c = {};
      if(data.description != null && String(data.description).trim()) c.description = String(data.description).trim();
      if(data.montant != null && String(data.montant).trim() !== ""){ const mt = calcule(data.montant); if(!mt || !(mt.valeur > 0)) throw new Erreur("Montant invalide."); c.montant_c = cents(mt.valeur); }
      if(data.date){ if(!/^\d{4}-\d{2}-\d{2}$/.test(data.date)) throw new Erreur("Date invalide."); c.date = data.date; }
      return c;
    };
    function modifiePrevu(data){
      const pv = D.prevus.find(x => x.id === +data.id); if(!pv) throw new Erreur("Prévision introuvable");
      const av = Object.assign({}, pv), c = champsPrevu(data); Object.assign(pv, c); if(c.montant_c != null && c.montant_c !== av.montant_c) pv.estime = 0;
      journal("prevus", pv.id, "modification", av, pv); return pv.id;
    }
    // « C'est passé » : la prévision devient une vraie dépense / un vrai revenu (dans le mois de sa date)
    function passePrevu(data){
      const pv = D.prevus.find(x => x.id === +data.id); if(!pv) throw new Erreur("Prévision introuvable");
      const c = Object.assign({date: pv.date, montant_c: pv.montant_c, description: pv.description}, champsPrevu(data));
      const mo = c.date.slice(0, 7); if(mo < verrou()) throw new Erreur("Date dans un mois en lecture seule.");
      let id;
      if(pv.genre === "depense"){
        const d = {id: nouvelId("depenses"), mois: mo, date: c.date, description: c.description, montant_c: c.montant_c, calcul: null, type: pv.type || "P",
                   categorie_id: pv.categorie_id || null, paye_par: "moi", reglement_id: null, id_partage: null, historique: 0, part_hist: null,
                   ordre: ordreSuivant(), origine: "app", fixe_id: pv.fixe_id};
        D.depenses.push(d); journal("depenses", d.id, "ajout", null, d); id = d.id; ajusteEstimation(pv.fixe_id, D.depenses);
      } else {
        const r = {id: nouvelId("revenus"), mois: mo, date: c.date, description: c.description, montant_c: c.montant_c, calcul: null, cases_colonne: null,
                   cases_ajout_c: 0, historique: 0, origine: "app", fixe_id: pv.fixe_id};
        D.revenus.push(r); journal("revenus", r.id, "ajout", null, r); id = r.id; ajusteEstimation(pv.fixe_id, D.revenus);
      }
      supprime("prevus", pv.id);
      return id;
    }
    // une ligne fixe modifiée / désactivée / supprimée : ses prévisions pas encore cochées suivent
    function majPrevus(x, retirer){
      for(const pv of D.prevus.filter(y => y.fixe_id === x.id)){
        if(retirer){ D.prevus = D.prevus.filter(y => y !== pv); continue; }
        Object.assign(pv, {description: x.description, montant_c: x.montant_c, date: dateFixe(pv.mois, x.jour),
                           categorie_id: x.categorie_id || null, type: x.type || null, estime: x.variable ? 1 : 0});
      }
    }

    // versement vers un placement (Placements › compte › Verser) : l'argent sort du compte (épargne du mois) ;
    // au besoin, la valeur du placement augmente du même montant (annuler le versement retire aussi cet ajout)
    function versement(f, data){
      const i = modifiable(f), cid = +data.compte, c = D.comptes_placement.find(x => x.id === cid);
      if(!c) throw new Erreur("Compte de placement introuvable");
      const mt = calcule(data.montant); if(!mt || !mt.valeur || mt.valeur <= 0) throw new Erreur("Indique le montant versé.");
      const date = data.date || auj(), mo = moisDeDate(date, i);
      if(mo < verrou()) throw new Erreur("Date dans un mois en lecture seule.");
      const ajoute = !!data.ajoute_valeur && !c.valeur_fixe;
      const r = {id: nouvelId("versements_epargne"), mois: mo, date, compte_id: cid, montant_c: cents(mt.valeur), origine: "app",
                 libelle: "Versement " + c.nom, valeur_ajoutee: ajoute ? 1 : 0};
      if(ajoute){
        const pl = M().placements(mo).find(y => y.id === cid);
        let x = D.placements_mois.find(y => y.compte_id === cid && y.mois === mo);
        if(!x){ x = {compte_id: cid, mois: mo, valeur_c: null, investi_c: null, investi_manuel: 0}; D.placements_mois.push(x); }
        x.valeur_c = (x.valeur_c != null ? x.valeur_c : Math.round(((pl && pl.valeur) || 0) * 100)) + r.montant_c;
      }
      D.versements_epargne.push(r); journal("versements_epargne", r.id, "ajout", null, r); return r.id;
    }
    function budgets(f, data){
      const mo = data.par_mois && f ? info(f).cle : "";
      // budgets d'un mois : la liste envoyée remplace celle du mois (comme une colonne de mois dans la v1)
      if(mo){ const prec = M().budgets(mo); if(!D.budgets.some(b => b.mois === mo)) for(const [n, v] of Object.entries(prec)) if(!(n in (data.budgets || {}))) data.budgets[n] = ""; D.budgets = D.budgets.filter(b => b.mois !== mo); }
      for(const [n, v] of Object.entries(data.budgets || {})){
        const id = catId(n); if(!id) continue;
        D.budgets = D.budgets.filter(b => !(b.categorie_id === id && b.mois === mo));
        const c = v === "" || v == null ? null : calcule(v);
        if(c) D.budgets.push({categorie_id: id, mois: mo, montant_c: cents(c.valeur)});
      }
      journal("budgets", null, "modification", null, data.budgets); return true;
    }
    function autreMontant(f, data){
      const i = modifiable(f); const c = calcule(data.montant); if(!c) throw new Erreur("Le montant est obligatoire.");
      const montant_c = cents(Math.abs(c.valeur)) * (data.sens === "je_dois" ? -1 : 1);
      if(data.ligne != null){ const x = D.autres_montants.find(y => y.id === +data.ligne); if(!x) throw new Erreur("Ligne introuvable");
        const av = Object.assign({}, x); Object.assign(x, {libelle: data.libelle || "", montant_c}); journal("autres_montants", x.id, "modification", av, x); return x.id; }
      const x = {id: nouvelId("autres_montants"), date: i.cle === auj().slice(0, 7) ? auj() : finMois(i.cle), libelle: data.libelle || "", montant_c,
                 reglement_id: null, ordre: ordreSuivant(), origine: "app"};
      D.autres_montants.push(x); journal("autres_montants", x.id, "ajout", null, x); return x.id;
    }
    function configuration(f, data){
      const i = modifiable(f);
      const sd = calcule(data.solde_debut); if(sd == null) throw new Erreur("Le solde de début est obligatoire.");
      let s = D.soldes_mois.find(x => x.mois === i.cle); if(!s){ s = {mois: i.cle, debut_c: null, debut_raison: null, ajustement_c: 0, ajustement_raison: null}; D.soldes_mois.push(s); }
      s.debut_c = cents(sd.valeur); s.debut_raison = "Solde de départ";
      for(const [nom, v] of Object.entries(data.placements || {})){
        const c = D.comptes_placement.find(x => x.nom === nom || String(x.id) === String(nom)); if(!c) continue;
        investissement(f, {ligne: c.id, investi: v.investi ?? v.investi_c, jv: v.jv ?? v.valeur});
      }
      if(data.pct_partenaire != null && data.pct_partenaire !== ""){
        let rep = D.repartition.find(r => r.mois === i.cle); if(!rep){ rep = {mois: i.cle, revenu_moi_c: null, revenu_partenaire_c: null, pct_partenaire: 0, pct_manuel: 1}; D.repartition.push(rep); }
        rep.pct_partenaire = calcule(data.pct_partenaire).valeur / 100; rep.pct_manuel = 1;
      }
      setReg("configuration_a_faire", false); return true;
    }

    // ------------------------------------------------------------------ historique des modifications et « Annuler »
    const TABLE_FR = {depenses: "Dépense", revenus: "Revenu", pourboires: "Pourboires", autres_montants: "Autre montant", reglements: "Règlement",
      mots_cles: "Mot-clé", categories: "Catégorie", placements_mois: "Placement", repartition: "Répartition", budgets: "Budgets",
      ajustements_solde: "Ajustement de solde", comptes_placement: "Compte de placement", financements: "Financement", financements_recus: "Financement reçu", financements_soldes: "Solde dû d'un financement", releves_banque: "Solde de la banque", versements_epargne: "Épargne", fixes: "Dépense ou revenu fixe", prevus: "Prévision (fixe)"};
    const ANNULABLES = new Set(["prevus", "fixes", "versements_epargne", "financements", "financements_recus", "financements_soldes", "releves_banque", "depenses", "revenus", "pourboires", "autres_montants", "reglements", "mots_cles", "ajustements_solde", "placements_mois", "categories"]);
    function resume(j){
      const o = JSON.parse(j.apres || j.avant || "null") || {};
      const mt = o.montant_c != null ? " · " + (o.montant_c / 100).toFixed(2).replace(".", ",") + " $" : "";
      const sol = o.solde_c != null ? " · " + (o.solde_c / 100).toFixed(2).replace(".", ",") + " $" : "";
      return (o.description || o.libelle || o.mot || o.note || o.nom || (j.table_nom === "reglements" ? "Solde réglé" : "")) + (o.date ? " (" + o.date + ")" : "") + mt + sol;
    }
    function historiqueModifs(n){
      const annulees = new Set(D.journal.filter(j => j.action === "annulation").map(j => j.ligne_id));
      return D.journal.slice().sort((a, b) => b.id - a.id).filter(j => j.action !== "annulation").slice(0, n).map(j => ({id: j.id, quand: j.quand,
        quoi: TABLE_FR[j.table_nom] || j.table_nom, action: j.action, resume: resume(j), annulee: annulees.has(j.id),
        annulable: ANNULABLES.has(j.table_nom) && !annulees.has(j.id) && !!(j.avant || j.apres)}));
    }
    function annuler(id){
      const j = D.journal.find(x => x.id === id); if(!j) throw new Erreur("Modification introuvable");
      if(D.journal.some(x => x.action === "annulation" && x.ligne_id === id)) throw new Erreur("Déjà annulée.");
      if(!ANNULABLES.has(j.table_nom)) throw new Erreur("Cette modification ne peut pas être annulée ici.");
      const t = D[j.table_nom], avant = j.avant && JSON.parse(j.avant), apres = j.apres && JSON.parse(j.apres);
      const cle = j.table_nom === "placements_mois" ? (x, o) => x.compte_id === o.compte_id && x.mois === o.mois : (x, o) => x.id === o.id;
      // on n'annule que si l'élément n'a pas été modifié depuis (sinon : annuler d'abord les modifications plus récentes)
      const plusRecente = D.journal.find(x => x.id > id && x.table_nom === j.table_nom && x.ligne_id === j.ligne_id && x.action !== "annulation"
        && !D.journal.some(y => y.action === "annulation" && y.ligne_id === x.id));
      if(plusRecente) throw new Erreur("Cet élément a été modifié depuis : annule d'abord la modification plus récente.");
      if(j.action === "ajout"){
        const k = t.findIndex(x => cle(x, apres)); if(k >= 0) t.splice(k, 1);
        if(j.table_nom === "reglements") for(const tt of ["depenses", "autres_montants"]) for(const x of D[tt]) if(x.reglement_id === apres.id) x.reglement_id = null;
        if(j.table_nom === "revenus") D.revenus_pourboires = D.revenus_pourboires.filter(x => x.revenu_id !== apres.id);
        if(j.table_nom === "versements_epargne" && apres.valeur_ajoutee){
          const x = D.placements_mois.find(y => y.compte_id === apres.compte_id && y.mois === apres.mois);
          if(x && x.valeur_c != null) x.valeur_c -= apres.montant_c;
        }
      } else if(j.action === "suppression"){
        if(j.table_nom === "reglements") throw new Erreur("Pour rétablir un règlement, marque de nouveau le solde réglé.");
        const liens = avant._liens; delete avant._liens; t.push(avant);
        if(liens) for(const p of liens) D.revenus_pourboires.push(j.table_nom === "pourboires" ? {revenu_id: p, pourboire_id: avant.id} : {revenu_id: avant.id, pourboire_id: p});
      } else if(j.action === "modification"){
        const k = t.findIndex(x => cle(x, avant)); if(k >= 0) t[k] = avant; else t.push(avant);
      }
      if(["revenus", "pourboires"].includes(j.table_nom)) recalculeRevenusCases();
      D.journal.push({id: nouvelId("journal"), quand: new Date().toISOString(), table_nom: j.table_nom, ligne_id: id, action: "annulation", avant: null, apres: null});
      return true;
    }

    // ------------------------------------------------------------------ routes
    async function api(chemin, corps){
      const u = new URL(chemin, "http://x"), p = u.pathname, f = corps ? corps.feuille : null;
      if(!corps){
        if(p === "/api/etat") return etat();
        if(p === "/api/journal") return historiqueModifs(+(u.searchParams.get("n") || 60));
        if(p === "/api/fixes") return fixes();
        if(p === "/api/export"){
          const X = (typeof module !== "undefined" && module.exports) ? require("./export_xlsx.js") : racine.ExportXlsx;
          const octets = X.exporte(null, D, M(), nomMois);
          let b64; if(typeof Buffer !== "undefined") b64 = Buffer.from(octets).toString("base64");
          else { let t = ""; for(let i = 0; i < octets.length; i += 0x8000) t += String.fromCharCode.apply(null, octets.subarray(i, i + 0x8000)); b64 = btoa(t); }
          return {nom: `Finances – export ${auj()}.xlsx`, b64};
        }
        if(p === "/api/mois") return mois(u.searchParams.get("f"));
        if(p === "/api/tendances") return tendances();
        throw new Erreur("Route inconnue : " + p);
      }
      if(opts.lectureSeule && opts.lectureSeule(p)) return {ok: true, resultat: await opts.lectureSeule(p)(corps)};
      let res;
      switch(p){
        case "/api/transaction": res = ecritDepense(f, corps); break;
        case "/api/transactions/lot": res = (corps.transactions || []).map(t => ecritDepense(f, t)); break;
        case "/api/transaction/supprimer": res = supprime("depenses", corps.ligne, d => { modifiable(nomMois(d.mois)); if(d.historique) throw new Erreur("Ligne d'historique : lecture seule."); }); break;
        case "/api/reglement": res = reglement(f, corps.ligne); break;
        case "/api/marika": res = autreMontant(f, corps); break;
        case "/api/marika/supprimer": res = supprime("autres_montants", corps.ligne); break;
        case "/api/revenu": res = ecritRevenu(f, corps); break;
        case "/api/revenus/lot": res = (corps.revenus || []).map(r => ecritRevenu(f, r)); break;
        case "/api/revenu/supprimer": res = supprime("revenus", corps.ligne); break;
        case "/api/pourboire": res = ecritPourboire(f, corps); break;
        case "/api/pourboire/supprimer": res = supprime("pourboires", corps.ligne); recalculeRevenusCases(); break;
        case "/api/parametres": res = parametres(f, corps); break;
        case "/api/investissement": res = investissement(f, corps); break;
        case "/api/versement": res = versement(f, corps); break;
        case "/api/fixe": res = ecritFixe(corps); break;
        case "/api/fixe/supprimer": res = supprime("fixes", corps.id, x => majPrevus(x, true)); break;
        case "/api/prevu": res = modifiePrevu(corps); break;
        case "/api/prevu/passe": res = passePrevu(corps); break;
        case "/api/prevu/retirer": res = supprime("prevus", corps.id); break;
        case "/api/fixe/ignorer": setReg("fixes_ignores", [...new Set([...(REG().fixes_ignores || []), String(corps.cle || "")])]); res = true; break;
        case "/api/fixes/generer": res = genereFixes(); break;
        case "/api/budgets": res = budgets(f, corps); break;
        case "/api/configuration": res = configuration(f, corps); break;
        case "/api/ajustement": {
          const i = modifiable(f || nomMois(auj().slice(0, 7))); const c = calcule(corps.montant); if(!c || !c.valeur) throw new Erreur("Montant obligatoire");
          const note = String(corps.note || "").trim(); if(!note) throw new Erreur("Une note est obligatoire (ex. « écart inexpliqué au 14 octobre »).");
          const x = {id: nouvelId("ajustements_solde"), mois: i.cle, date: auj(), montant_c: cents(c.valeur), note};
          D.ajustements_solde.push(x); journal("ajustements_solde", x.id, "ajout", null, x); res = x.id; break; }
        case "/api/ajustement/supprimer": res = supprime("ajustements_solde", corps.ligne); break;
        case "/api/annuler": res = annuler(+corps.id); break;
        // ---- financement (prêts) : argent reçu (entre au compte, pas un revenu) et solde dû inscrit à la main
        case "/api/financement": {
          const n = String(corps.nom || "").trim(); if(!n) throw new Erreur("Nom obligatoire (ex. « Prêt étudiant »).");
          const x = {id: nouvelId("financements"), nom: n, ordre: D.financements.length + 1, cree: f ? info(f).cle : auj().slice(0, 7)};
          D.financements.push(x); journal("financements", x.id, "ajout", null, x); res = x.id; break; }
        case "/api/financement/recu": {
          const i = modifiable(f); const fid = +corps.financement; if(!D.financements.some(x => x.id === fid)) throw new Erreur("Financement introuvable");
          const c = calcule(corps.montant); if(!c || !c.valeur) throw new Erreur("Le montant est obligatoire.");
          const date = corps.date || auj(); const mois_ = date.slice(0, 7) >= verrou() ? date.slice(0, 7) : i.cle;
          const champs = {financement_id: fid, mois: mois_, date, montant_c: cents(c.valeur), note: String(corps.note || "").trim() || null};
          if(corps.ligne != null){ const x = D.financements_recus.find(y => y.id === +corps.ligne); if(!x) throw new Erreur("Ligne introuvable");
            const av = Object.assign({}, x); Object.assign(x, champs); journal("financements_recus", x.id, "modification", av, x); res = x.id; break; }
          const x = Object.assign({id: nouvelId("financements_recus")}, champs); D.financements_recus.push(x); journal("financements_recus", x.id, "ajout", null, x); res = x.id; break; }
        case "/api/financement/recu/supprimer": res = supprime("financements_recus", corps.ligne); break;
        case "/api/financement/solde": {
          const fid = +corps.financement; if(!D.financements.some(x => x.id === fid)) throw new Erreur("Financement introuvable");
          const c = calcule(corps.solde_du); if(!c) throw new Erreur("Le solde dû est obligatoire.");
          const x = {id: nouvelId("financements_soldes"), financement_id: fid, date: corps.date || auj(), solde_du_c: cents(c.valeur)};
          D.financements_soldes.push(x); journal("financements_soldes", x.id, "ajout", null, x); res = x.id; break; }
        case "/api/listes/categorie": {
          const n = String(corps.nom || "").trim(); if(!n) throw new Erreur("Nom obligatoire"); if(catId(n)) throw new Erreur("Cette catégorie existe déjà.");
          const c = {id: nouvelId("categories"), nom: n, ordre: D.categories.length + 1, couleur: null, active: 1}; D.categories.push(c); journal("categories", c.id, "ajout", null, c); res = c.id; break; }
        case "/api/listes/mot-cle": {
          const mot = String(corps.mot || "").trim(); if(!mot) throw new Erreur("Mot-clé obligatoire"); if(!catId(corps.categorie)) throw new Erreur("Catégorie inconnue");
          res = insereMotCle(mot, corps.categorie); break; }
        case "/api/compte-epargne": {
          const n = String(corps.nom || "").trim(); if(!n) throw new Erreur("Nom obligatoire"); if(D.comptes_placement.some(c => norm(c.nom) === norm(n))) throw new Erreur("Ce compte existe déjà.");
          const c = {id: nouvelId("comptes_placement"), nom: n, valeur_fixe: /urgence|épargne|epargne/i.test(n) ? 1 : 0, ordre: D.comptes_placement.length + 1, cree: (f ? info(f).cle : auj().slice(0, 7))};
          D.comptes_placement.push(c); journal("comptes_placement", c.id, "ajout", null, c); res = c.id; break; }
        default: throw new Erreur("Fonction retirée dans la v2 : " + p);
      }
      if(opts.apresEcriture) await opts.apresEcriture(p);
      return {ok: true, resultat: res};
    }
    return {api, etat, mois, tendances, donnees: () => D, Erreur, calcule};
  }
  const exp = {API, nomMois, cleDe, calcule};
  if(typeof module !== "undefined" && module.exports) module.exports = exp; else racine.ApiV2 = exp;
})(typeof globalThis !== "undefined" ? globalThis : this);
