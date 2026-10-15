// Moteur de calcul de la v2 (navigateur et Node). Entrée : les tables de la base sous forme d'objets simples
// (voir donnees() dans le test, ou sql.js dans l'app). Aucun total n'est enregistré : tout est calculé ici.
// Montants de la base en cents entiers ; les résultats sont en dollars NON arrondis (arrondir à l'affichage).
"use strict";
(function(racine){
  const norm = t => String(t || "").trim().toLowerCase().normalize("NFC");
  const cle = d => (d || "").slice(0, 7);
  const moisPrec = m => { let [a, n] = m.split("-").map(Number); n--; if(n < 1){ n = 12; a--; } return a + "-" + String(n).padStart(2, "0"); };
  const somme = (l, f) => l.reduce((a, x) => a + (f(x) || 0), 0);

  const moisSuiv = m => { let [a, n] = m.split("-").map(Number); n++; if(n > 12){ n = 1; a++; } return a + "-" + String(n).padStart(2, "0"); };
  function Moteur(D, opts = {}){
    const cats = new Map(D.categories.map(c => [c.id, c.nom]));
    const mots = D.mots_cles.slice().sort((a, b) => a.ordre - b.ordre).map(k => [norm(k.mot), cats.get(k.categorie_id)]);
    const rep = new Map(D.repartition.map(r => [r.mois, r]));
    const soldes = new Map(D.soldes_mois.map(s => [s.mois, s]));
    const dep = D.depenses.slice().sort((a, b) => a.ordre - b.ordre);
    const pbLies = new Set(D.revenus_pourboires.map(x => x.pourboire_id));
    const M = {};

    // ------------------------------------------------------------ dépenses
    M.categorie = d => {
      if(d.categorie_id) return cats.get(d.categorie_id);
      const t = norm(d.description); let r = "Autres";
      for(const [mot, cat] of mots) if(mot && t.includes(mot)) r = cat;     // le dernier mot-clé trouvé l'emporte
      return r;
    };
    M.pct = mois => (rep.get(mois) || {}).pct_partenaire || 0;
    // logique « v2 » (décision 4) : une part payée par le partenaire ne compte qu'une fois réglée ;
    // « v1-zach » : elle compte tout de suite (ancien comportement de l'app de Zach), pour les comparaisons.
    M.maPart = (d, logique = "v2") => {
      if(d.historique) return d.part_hist || 0;
      const m = (d.montant_c || 0) / 100;
      if(d.paye_par === "partenaire") return (logique === "v2" && !d.reglement_id) ? 0 : m;
      const p = M.pct(d.mois);
      return {P: m, "50": m * 0.5, SP: m * (1 - p), R: m * (1 - p), M: 0}[d.type] ?? 0;
    };
    M.partPartenaire = d => {          // > 0 : le partenaire me doit ; < 0 : je lui dois
      if(d.historique) return 0;
      const m = (d.montant_c || 0) / 100;
      if(d.paye_par === "partenaire") return -m;
      const p = M.pct(d.mois);
      return {P: 0, "50": m * 0.5, SP: m * p, R: -m * (1 - p), M: m}[d.type] ?? 0;
    };

    // ------------------------------------------------------------ un mois
    // tous les mois, sans trou, du premier mois de données jusqu'au dernier (ou jusqu'à opts.moisMax, ex. le mois du jour)
    let _liste = null;
    M.listeMois = () => {
      if(_liste) return _liste;
      const vus = new Set([...soldes.keys(), ...dep.map(d => d.mois), ...D.revenus.map(r => r.mois), ...D.pourboires.map(p => p.mois),
                           ...D.placements_mois.map(p => p.mois)].filter(Boolean));
      if(opts.moisMax) vus.add(opts.moisMax);
      const l = [...vus].sort(); if(!l.length) return (_liste = []);
      const out = []; for(let m = l[0]; m <= l[l.length - 1]; m = moisSuiv(m)) out.push(m);
      return (_liste = out);
    };
    M.mois = (mo, logique = "v2") => {
      const deps = dep.filter(d => d.mois === mo);
      const parCategorie = {};
      for(const d of deps){ const k = M.categorie(d); parCategorie[k] = (parCategorie[k] || 0) + M.maPart(d, logique); }
      const total = Object.values(parCategorie).reduce((a, x) => a + x, 0);
      const revenus = somme(D.revenus.filter(r => r.mois === mo), r => r.montant_c) / 100;
      const epargne = somme(D.versements_epargne.filter(v => v.mois === mo), v => v.montant_c) / 100;
      const pbs = D.pourboires.filter(p => p.mois === mo);
      return {mois: mo, lignes: deps.length, totalDepenses: total, parCategorie, revenus, revenuNet: revenus - total, epargne,
              pourboires: {especes: somme(pbs, p => p.especes_c) / 100, carte: somme(pbs, p => p.carte_c) / 100, heures: somme(pbs, p => p.heures)}};
    };
    // début = point de départ explicite (premier mois, ou départ repris de l'Excel), sinon fin du mois précédent
    M.chaine = (logique = "v2") => {
      const out = {}; let prec = null;
      for(const mo of M.listeMois()){
        const s = soldes.get(mo) || {debut_c: null, ajustement_c: 0}, m = M.mois(mo, logique);
        const debut = s.debut_c != null ? s.debut_c / 100 : (prec ?? 0);
        const ajust = (s.ajustement_c || 0) / 100 + somme((D.ajustements_solde || []).filter(a => a.mois === mo), a => a.montant_c) / 100;
        // financement reçu (prêt) : l'argent entre au compte, sans être un revenu (le revenu net n'en tient pas compte)
        const finan = somme((D.financements_recus || []).filter(f => f.mois === mo), f => f.montant_c) / 100;
        const fin = debut + m.revenuNet - m.epargne + ajust + finan;
        out[mo] = Object.assign({debut, fin, ajustement: ajust, financement: finan}, m); prec = fin;
      }
      return out;
    };

    // ------------------------------------------------------------ partenaire
    const regl = new Map((D.reglements || []).map(r => [r.id, r]));
    // ouvert = pas encore réglé (à une date donnée : réglé seulement si le règlement a eu lieu ce jour-là ou avant)
    const ouvert = (x, quand) => !x.historique &&
      (!x.reglement_id || (quand && regl.get(x.reglement_id) && regl.get(x.reglement_id).date > quand));
    M.soldePartenaire = quand => somme(dep.filter(d => ouvert(d, quand)), M.partPartenaire)
                                + somme(D.autres_montants.filter(x => ouvert(x, quand)), x => x.montant_c) / 100;
    // détail du solde avec le partenaire, tel qu'il était à une date (règlements faits ce jour-là ou avant),
    // pour les mois jusqu'à moisMax
    M.detailPartenaire = (quand, moisMax) => {
      const o = x => ouvert(x, quand) && (!moisMax || (x.mois || cle(x.date)) <= moisMax);
      const mesParts = somme(dep.filter(d => o(d) && d.paye_par !== "partenaire"), M.partPartenaire);
      const autres = somme(D.autres_montants.filter(o), x => x.montant_c) / 100;
      const sesParts = somme(dep.filter(d => o(d) && d.paye_par === "partenaire"), d => (d.montant_c || 0) / 100);
      return {mesParts, autres, sesParts, net: mesParts + autres - sesParts};
    };

    // ------------------------------------------------------------ soldes du jour
    // Solde de l'app (fin du mois du jour) → Solde au compte (la banque aujourd'hui) et Solde réel attendu.
    M.soldes = (auj, logique = "v2") => {
      const ch = M.chaine(logique), mo = cle(auj);
      const app = (ch[mo] || ch[M.listeMois().filter(x => x <= mo).pop()] || {fin: 0}).fin;
      const jusque = x => x.mois <= mo;
      const revFut = D.revenus.filter(r => jusque(r) && r.date && r.date > auj && r.montant_c > 0);
      const depFut = dep.filter(d => jusque(d) && !d.historique && d.paye_par !== "partenaire" && d.date && d.date > auj && d.montant_c > 0);
      const p = M.detailPartenaire(auj, mo);
      const nonRegl = p.sesParts;
      const finFut = (D.financements_recus || []).filter(f => f.mois <= mo && f.date > auj);
      const sR = (somme(revFut, r => r.montant_c) + somme(finFut, f => f.montant_c)) / 100, sD = somme(depFut, d => d.montant_c) / 100;
      // en logique v1 (ancienne app de Zach), le solde de l'app comptait déjà ses parts payées par le partenaire
      const aRembDansApp = logique === "v1-zach" ? nonRegl : 0;
      const auCompte = app + aRembDansApp - sR + sD - (p.mesParts + p.autres);
      const pbAttente = D.pourboires.filter(x => (x.recu === 0 || (x.recu == null && x.carte_c > 0)) && !pbLies.has(x.id) && jusque(x));
      const sPb = somme(pbAttente, x => x.carte_c) / 100;
      const attendu = auCompte + sR - sD + p.net + sPb;
      const releve = D.releves_banque.filter(r => r.date <= auj).sort((a, b) => a.date.localeCompare(b.date)).pop() || null;
      return {app, auCompte, attendu, revFut: revFut.concat(finFut.map(f => ({id: "F" + f.id, date: f.date, description: "Financement reçu", montant_c: f.montant_c}))), depFut, sR, sD, partenaire: p, pbAttente, sPb,
              releve: releve && {date: releve.date, solde: releve.solde_c / 100}};
    };
    // contrôle gardé : le solde au compte calculé à la date d'un relevé de banque, comparé à ce relevé
    M.concordance = (releve, logique = "v2") => {
      const s = M.soldes(releve.date, logique);
      return {date: releve.date, banque: releve.solde_c / 100, calcule: s.auCompte, ecart: releve.solde_c / 100 - s.auCompte};
    };

    // ------------------------------------------------------------ placements
    // capital investi : saisi à la main, sinon capital du mois précédent + épargne versée le mois précédent ;
    // un mois sans saisie reprend la valeur marchande du mois précédent
    const _pl = new Map();
    M.placements = mo => {
      if(_pl.has(mo)) return _pl.get(mo);
      const premier = M.listeMois()[0];
      const out = D.comptes_placement.slice().sort((a, b) => a.ordre - b.ordre).map(c => {
        const x = D.placements_mois.find(y => y.compte_id === c.id && y.mois === mo);
        const pr = mo > premier ? (M.placements(moisPrec(mo)).find(y => y.id === c.id) || null) : null;
        if(!x && !pr && !(c.cree && mo >= c.cree)) return null;
        const vers = somme(D.versements_epargne.filter(v => v.compte_id === c.id && v.mois === moisPrec(mo)), v => v.montant_c) / 100;
        const auto = pr ? pr.investi + vers : null;
        const investi = x && (x.investi_manuel || auto == null) ? (x.investi_c || 0) / 100 : auto;
        const valeur = c.valeur_fixe ? investi : (x && x.valeur_c != null ? x.valeur_c / 100 : (pr ? pr.valeur : 0));
        return {id: c.id, compte: c.nom, investi, valeur, variation: valeur - investi, manuel: !!(x && x.investi_manuel), auto, saisi: !!x};
      }).filter(Boolean);
      _pl.set(mo, out); return out;
    };

    // ------------------------------------------------------------ financement (prêts) : reçu à ce jour et solde dû
    M.financements = mo => (D.financements || []).slice().sort((a, b) => a.ordre - b.ordre).filter(f => !f.cree || f.cree <= mo).map(f => {
      const recus = (D.financements_recus || []).filter(x => x.financement_id === f.id && x.mois <= mo);
      const sd = (D.financements_soldes || []).filter(x => x.financement_id === f.id && x.date.slice(0, 7) <= mo).sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id).pop();
      const recu = somme(recus, x => x.montant_c) / 100;
      return {id: f.id, nom: f.nom, recu, recuMois: somme(recus.filter(x => x.mois === mo), x => x.montant_c) / 100,
              soldeDu: sd ? sd.solde_du_c / 100 : recu, soldeDuDate: sd ? sd.date : null, recus};
    });

    // ------------------------------------------------------------ budgets
    // budgets d'un mois (comme la v1) : la liste du dernier mois fixé (ce mois-ci ou avant) — une catégorie absente de
    // cette liste n'a pas de budget ce mois-là — sinon les budgets par défaut
    M.budgetsDe = mo => D.budgets.filter(b => b.mois && b.mois <= mo).map(b => b.mois).sort().pop() || null;
    M.budgets = mo => {
      const de = M.budgetsDe(mo), parCat = {};
      for(const b of D.budgets.filter(b => b.mois === (de || ""))) if(b.montant_c != null) parCat[cats.get(b.categorie_id)] = b.montant_c / 100;
      return parCat;
    };
    M.propositionBudgets = mo => {
      const avant = M.listeMois().filter(x => x < mo).slice(-12), tot = {};
      for(const x of avant) for(const [k, v] of Object.entries(M.mois(x).parCategorie)) tot[k] = (tot[k] || 0) + v;
      const n = avant.length || 1; const out = {};
      for(const k in tot) out[k] = tot[k] / n;
      return out;
    };

    // ------------------------------------------------------------ tendances (graphiques)
    M.tendances = (logique = "v2") => {
      const ch = M.chaine(logique);
      return M.listeMois().map(mo => {
        const pl = M.placements(mo);
        return {mois: mo, totalDepenses: ch[mo].totalDepenses, parCategorie: ch[mo].parCategorie, revenus: ch[mo].revenus,
                revenuNet: ch[mo].revenuNet, soldeFin: ch[mo].fin, investi: somme(pl, x => x.investi), valeur: somme(pl, x => x.valeur),
                epargne: ch[mo].epargne};
      });
    };
    return M;
  }
  const exp = {Moteur, moisPrec};
  if(typeof module !== "undefined" && module.exports) module.exports = exp; else racine.CalculsV2 = exp;
})(typeof globalThis !== "undefined" ? globalThis : this);
