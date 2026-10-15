// Export Excel de la v2 : classeur simple et lisible (aucune formule), généré sans bibliothèque externe.
// Onglets : Résumé (par mois), Dépenses, Revenus, Pourboires, Placements, Budgets, Partage (autres montants).
"use strict";
(function(racine){
  // ------------------------------------------------------------------ ZIP minimal (fichiers « stockés », sans compression)
  const CRC = (() => { const t = new Uint32Array(256); for(let n = 0; n < 256; n++){ let c = n; for(let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for(let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const utf8 = s => new TextEncoder().encode(s);
  function zip(fichiers){
    const parts = [], central = []; let pos = 0;
    const u16 = v => [v & 0xFF, (v >>> 8) & 0xFF], u32 = v => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];
    for(const [nom, contenu] of fichiers){
      const n = utf8(nom), d = typeof contenu === "string" ? utf8(contenu) : contenu, c = crc32(d);
      const tete = [...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(d.length), ...u32(d.length), ...u16(n.length), ...u16(0)];
      parts.push(new Uint8Array(tete), n, d);
      central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(d.length), ...u32(d.length),
        ...u16(n.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(pos)]), n);
      pos += tete.length + n.length + d.length;
    }
    const tailleC = central.reduce((a, x) => a + x.length, 0);
    const fin = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(fichiers.length), ...u16(fichiers.length), ...u32(tailleC), ...u32(pos), ...u16(0)]);
    const tout = [...parts, ...central, fin], out = new Uint8Array(tout.reduce((a, x) => a + x.length, 0)); let o = 0;
    for(const x of tout){ out.set(x, o); o += x.length; }
    return out;
  }
  // ------------------------------------------------------------------ feuilles
  const xe = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const col = i => { let s = ""; i++; while(i){ const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
  const serie = iso => { const [a, m, j] = iso.split("-").map(Number); return Math.round((Date.UTC(a, m - 1, j) - Date.UTC(1899, 11, 30)) / 864e5); };
  // styles : 0 normal, 1 en-tête (gras), 2 montant (# ##0,00 $), 3 date (aaaa-mm-jj), 4 pourcentage
  function feuille(lignes, largeurs){
    const rows = lignes.map((l, r) => `<row r="${r + 1}">` + l.map((v, c) => {
      const ref = col(c) + (r + 1);
      if(v == null || v === "") return "";
      if(typeof v === "object" && v.t === "d") return `<c r="${ref}" s="3"><v>${serie(v.v)}</v></c>`;
      if(typeof v === "object" && v.t === "%") return `<c r="${ref}" s="4"><v>${v.v}</v></c>`;
      if(typeof v === "number") return `<c r="${ref}" s="${r === 0 ? 1 : 2}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr" s="${r === 0 ? 1 : 0}"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
    }).join("") + `</row>`).join("");
    const cols = largeurs ? `<cols>${largeurs.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${rows}</sheetData></worksheet>`;
  }
  function classeur(onglets){
    const f = [["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${onglets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`],
      ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${onglets.map((o, i) => `<sheet name="${xe(o.nom)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`],
      ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${onglets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${onglets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="# ##0.00 &quot;$&quot;"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="5"><xf/><xf fontId="1" applyFont="1"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="165" applyNumberFormat="1"/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`]];
    onglets.forEach((o, i) => f.push([`xl/worksheets/sheet${i + 1}.xml`, feuille(o.lignes, o.largeurs)]));
    return zip(f);
  }
  // ------------------------------------------------------------------ contenu à partir de l'API v2
  function exporte(A, D, M, nomMoisFn){
    const P = (D.reglages.find(r => r.cle === "partenaire") || {valeur: '"partenaire"'}).valeur.replace(/"/g, "");
    const cat = new Map(D.categories.map(c => [c.id, c.nom])), cpt = new Map(D.comptes_placement.map(c => [c.id, c.nom]));
    const TY = {P: "Personnel", "50": "50/50", SP: "Proportionnel", M: "100 % " + P, R: "Remboursement"};
    const d = x => x ? {t: "d", v: x} : "";
    const ch = M.chaine();
    const resume = [["Mois", "Solde de début", "Revenus", "Dépenses (ma part)", "Revenu net", "Épargne placée", "Ajustement", "Solde de fin"]]
      .concat(M.listeMois().map(m => { const x = ch[m]; return [nomMoisFn(m), x.debut, x.revenus, x.totalDepenses, x.revenuNet, x.epargne, x.ajustement || 0, x.fin]; }));
    const dep = [["Mois", "Date", "Description", "Montant payé", "Type", "Payé par", "Ma part", "Part de " + P, "Catégorie", "Réglé", "Calcul saisi"]]
      .concat(D.depenses.slice().sort((a, b) => (a.mois + (a.date || "")).localeCompare(b.mois + (b.date || "")) || a.ordre - b.ordre).map(x => [nomMoisFn(x.mois), d(x.date), x.description,
        x.historique ? "" : (x.montant_c || 0) / 100, x.historique ? "Historique" : TY[x.type] || x.type, x.paye_par === "partenaire" ? P : "Moi",
        M.maPart(x), x.historique ? "" : M.partPartenaire(x), M.categorie(x), x.historique ? "" : x.reglement_id ? "Oui" : "Non", x.calcul || ""]));
    const pbLien = new Map(D.revenus_pourboires.map(l => [l.revenu_id, (D.revenus_pourboires.filter(y => y.revenu_id === l.revenu_id).length)]));
    const rev = [["Mois", "Date", "Description", "Montant", "Jours de pourboires inclus"]]
      .concat(D.revenus.slice().sort((a, b) => (a.mois + (a.date || "")).localeCompare(b.mois + (b.date || ""))).map(x => [nomMoisFn(x.mois), d(x.date), x.description, x.montant_c / 100, pbLien.get(x.id) || ""]));
    const pb = [["Mois", "Date", "Espèces", "Carte", "Heures", "Reçu"]]
      .concat(D.pourboires.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(x => [nomMoisFn(x.mois), d(x.date), x.especes_c == null ? "" : x.especes_c / 100,
        x.carte_c == null ? "" : x.carte_c / 100, x.heures ?? "", x.recu === 1 ? "Oui" : x.recu === 0 ? "Non" : ""]));
    const pl = [["Mois", "Compte", "Capital investi", "Valeur marchande", "Variation", "Épargne versée ce mois"]];
    for(const m of M.listeMois()) for(const x of M.placements(m))
      pl.push([nomMoisFn(m), x.compte, x.investi, x.valeur, x.variation, D.versements_epargne.filter(v => v.mois === m && v.compte_id === x.id).reduce((a, v) => a + v.montant_c / 100, 0)]);
    const bud = [["Catégorie", "Mois (vide = par défaut)", "Budget"]].concat(D.budgets.map(b => [cat.get(b.categorie_id), b.mois ? nomMoisFn(b.mois) : "", b.montant_c / 100]));
    const aut = [["Date", "Libellé", "Montant (+ " + P + " me doit, − je dois)", "Réglé"]].concat(D.autres_montants.map(x => [d(x.date), x.libelle, x.montant_c / 100, x.reglement_id ? "Oui" : "Non"]));
    const rep = [["Mois", "Mon revenu", "Revenu de " + P, "% de " + P]].concat(D.repartition.slice().sort((a, b) => a.mois.localeCompare(b.mois)).map(r =>
      [nomMoisFn(r.mois), r.revenu_moi_c == null ? "" : r.revenu_moi_c / 100, r.revenu_partenaire_c == null ? "" : r.revenu_partenaire_c / 100, {t: "%", v: r.pct_partenaire}]));
    return classeur([{nom: "Résumé", lignes: resume, largeurs: [16, 15, 13, 17, 13, 15, 12, 14]},
      {nom: "Dépenses", lignes: dep, largeurs: [16, 11, 34, 13, 14, 10, 11, 13, 16, 7, 16]}, {nom: "Revenus", lignes: rev, largeurs: [16, 11, 34, 12, 12]},
      {nom: "Pourboires", lignes: pb, largeurs: [16, 11, 10, 10, 8, 7]}, {nom: "Placements", lignes: pl, largeurs: [16, 14, 15, 16, 12, 16]},
      {nom: "Budgets", lignes: bud, largeurs: [18, 20, 12]}, {nom: "Partage", lignes: aut, largeurs: [11, 30, 26, 7]},
      {nom: "Répartition", lignes: rep, largeurs: [16, 13, 15, 11]}]);
  }
  const exp = {exporte, classeur, zip};
  if(typeof module !== "undefined" && module.exports) module.exports = exp; else racine.ExportXlsx = exp;
})(typeof globalThis !== "undefined" ? globalThis : this);
