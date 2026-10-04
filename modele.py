"""Logique financière : lecture des onglets, calculs du Canvas, opérations d'écriture."""
import datetime
import os
import re
import threading

from formules import (Evaluateur, NonSupporte, ErreurExcel, col_en_num, num_en_col,
                      prefixe_feuille, decoupe_adresse)
from xlsx import Classeur, normalise, serie_en_date, copie_sauvegarde

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]
MOIS_AFF = [m.capitalize() for m in MOIS]


def parse_mois(nom):
    m = re.match(r"^\s*([A-Za-zÀ-ÿ]+)\s+(\d{4})\s*$", nom or "")
    if not m or m.group(1).lower() not in MOIS:
        return None
    return int(m.group(2)), MOIS.index(m.group(1).lower()) + 1


def nom_mois(annee, mois):
    return "%s %d" % (MOIS_AFF[mois - 1], annee)


def num(v):
    if isinstance(v, bool) or v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def r2(x):
    if x is None:
        return None
    v = round(x, 2)
    return 0.0 if v == 0 else v


class Erreur(Exception):
    pass


# ---------------------------------------------------------------------------
# Disposition d'un onglet au format Canvas (tout est repéré par les titres)
# ---------------------------------------------------------------------------

class Disposition:
    def __init__(self, f):
        self.f = f
        self.fusions = {}
        for a, b in re.findall(r'<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"', f.xml):
            self.fusions[a] = b
        d = f.cherche("Date", cols=["A"])
        if d is None:
            raise Erreur("en-tête « Date » introuvable dans " + f.nom)
        self.lig_entete = h = d.lig
        entetes = sorted((c for c in f.cellules.values() if c.lig == h and isinstance(c.valeur, str)),
                         key=lambda c: col_en_num(c.col))
        vus = {}
        for c in entetes:
            k = normalise(c.valeur)
            vus.setdefault(k, []).append(c.col)
        def col(nom, k=0):
            l = vus.get(nom) or []
            return l[k] if len(l) > k else None
        self.c_date = col("date")
        self.c_desc = col("description")
        self.c_type = col("type de dépenses")
        self.c_deb = col("déboursé")
        self.c_prop = col("proportion")
        self.c_dep = col("dépense")
        self.c_cat = col("catégorie")
        if not (self.c_deb and self.c_type and self.c_cat):
            raise Erreur("format Canvas non reconnu")
        self.c_marq = num_en_col(col_en_num(self.c_cat) + 1)
        self.c_mtype = col("type de dépense")
        self.c_mprop = col("proportion", 1)
        self.c_mmont = col("montant")
        t = f.cherche("TOTAL DÉPENSES")
        self.lig_total = t.lig
        self.lignes = list(range(h + 1, t.lig))
        # Remboursement
        r = f.cherche("Remboursement", debut=h - 2, fin=h)
        self.c_rlib = r.col if r else None
        self.c_rval = num_en_col(col_en_num(r.col) + 1) if r else None
        self.lig_tableau = None
        self.lig_rtotal = None
        if r:
            tot = [c for c in f.cellules.values() if c.col == r.col and c.lig > h and
                   isinstance(c.valeur, str) and normalise(c.valeur) == "total"]
            tot.sort(key=lambda c: c.lig)
            self.lig_tableau = h + 1
            self.lig_rtotal = tot[0].lig if tot else None
        # Répartition
        self.a_marika = self._a_droite("Revenu Marika", "N")
        self.a_zach = self._a_droite("Revenu Zach", "N")
        pm = f.cherche("% Marika")
        # fichiers d'autres personnes : 1er « Revenu … » sous « Répartition » = partenaire, 2e = moi, 1er « % … » = % partenaire
        rep = f.cherche("Répartition")
        if rep:
            sous = sorted((c for c in f.cellules.values() if c.col == rep.col and rep.lig < c.lig <= rep.lig + 10
                           and isinstance(c.valeur, str) and c.valeur.strip()), key=lambda c: c.lig)
            revs = [c for c in sous if normalise(c.valeur).startswith("revenu")]
            pcts = [c for c in sous if c.valeur.strip().startswith("%")]
            droite = lambda c: num_en_col(col_en_num(c.col) + 1) + str(c.lig)
            if len(revs) >= 2:
                self.a_marika, self.a_zach = droite(revs[0]), droite(revs[1])
            if pcts:
                pm = pcts[0]
        self.a_pct = num_en_col(col_en_num(pm.col) + 1) + str(pm.lig) if pm else None
        self.a_pct_utilise = self.c_mprop + str(self.lig_total) if self.c_mprop else None
        # Pourboires
        p = f.cherche("Revenus - Pourboires", cols=["A"])
        self.pourboires = []
        self.pb = {"date": "A", "recu": "B", "arec": "C", "heures": None, "statut": "D"}
        if p:
            ent = p.lig + 1
            r_ = ent + 1
            while r_ < ent + 200:
                c = f.cell("B" + str(r_))
                if c is not None and c.formule and c.formule.upper().startswith("SUM("):
                    break
                r_ += 1
            self.pourboires = list(range(ent + 1, r_))
            self.lig_pourb_total = r_
            self.lig_pourb_entete = ent
            cols = {}
            for c in f.cellules.values():
                if c.lig == ent and isinstance(c.valeur, str):
                    n = normalise(c.valeur)
                    if n == "date": cols["date"] = c.col
                    elif n.startswith("montant reçu"): cols["recu"] = c.col
                    elif n.startswith("montant à recevoir"): cols["arec"] = c.col
                    elif n.startswith("heure"): cols["heures"] = c.col
                    elif n.startswith("reçu"): cols["statut"] = c.col
            self.pb = {"date": cols.get("date", "A"), "recu": cols.get("recu", "B"), "arec": cols.get("arec", "C"),
                       "heures": cols.get("heures"), "statut": cols.get("statut", "D")}
        # Revenus globaux
        g = f.cherche("Revenus globaux", cols=["A"])
        tr = f.cherche("TOTAL REVENUS", cols=["A"])
        self.revenus = list(range(g.lig + 2, tr.lig)) if g and tr else []
        self.a_total_rev = "B" + str(tr.lig) if tr else None
        # Investissements
        i = f.cherche("Investissements", cols=["A"], debut=h + 1)
        ti = f.cherche("Total investissements", cols=["A"])
        self.invest = list(range(i.lig + 2, ti.lig)) if i and ti else []
        self.lig_invest_entete = i.lig + 1 if i else None
        # Résumé global
        rg = f.cherche("Résumé global", cols=["A"])
        deb = rg.lig if rg else h
        self.a_rev = self._b("Revenus", deb)
        self.a_dep = self._b("(-) Dépenses", deb)
        self.a_net = self._b("Revenus net", deb)
        self.a_debut = self._b("Solde de compte - DÉBUT", deb)
        # retraits d'épargne : une ou plusieurs lignes « (-) Épargne … » (ex. « (-) Épargne (CÉLI) »)
        self.retraits = []
        fin_res = f.cherche("Solde de compte - FIN", cols=["A"], debut=deb)
        for c in sorted((c for c in f.cellules.values() if c.col == "A" and isinstance(c.valeur, str)
                         and deb <= c.lig <= (fin_res.lig if fin_res else deb + 15)
                         and normalise(c.valeur).startswith("(-) épargne")), key=lambda c: c.lig):
            mc = re.search(r"\(([^()]*)\)\s*$", c.valeur.strip())
            compte = mc.group(1).strip() if mc and not normalise(c.valeur).endswith("période") else None
            self.retraits.append({"libelle": c.valeur.strip(), "adr": "B" + str(c.lig), "compte": compte})
        self.a_retrait = self.retraits[0]["adr"] if self.retraits else None
        self.a_fin = self._b("Solde de compte - FIN", deb)
        # Résumé par catégorie
        rc = f.cherche("Résumé des dépenses par catégorie", cols=["A"])
        self.cats = []
        if rc:
            r_ = rc.lig + 1
            while isinstance(f.val("A" + str(r_)), str) and f.val("A" + str(r_)).strip():
                self.cats.append(r_)
                r_ += 1
        # Contrôles
        ce = f.cherche("Contrôle", cols=None)
        self.controles = []
        if ce:
            ent = {normalise(c.valeur): c.col for c in f.cellules.values()
                   if c.lig == ce.lig and isinstance(c.valeur, str)}
            self.c_ctl_lib = ce.col
            self.c_ctl_ecart = ent.get("écart / nombre")
            self.c_ctl_statut = ent.get("statut")
            self.c_ctl_forcer = ent.get("forcer ok")
            self.c_ctl_just = ent.get("justification")
            r_ = ce.lig + 1
            while isinstance(f.val(ce.col + str(r_)), str) and f.val(ce.col + str(r_)).strip():
                self.controles.append(r_)
                r_ += 1
        # Paramètres
        self.a_mois = self._param("Mois du tableau")
        self.a_reel = self._param("Solde réel au compte (relevé)")
        tl = f.cherche("Tableau de finances personnelles", exact=False)
        self.a_titre = tl.adr if tl else None

    def _a_droite(self, lib, col_pref=None):
        c = self.f.cherche(lib)
        return num_en_col(col_en_num(c.col) + 1) + str(c.lig) if c else None

    def _b(self, lib, debut):
        c = self.f.cherche(lib, cols=["A"], debut=debut)
        return "B" + str(c.lig) if c else None

    def _param(self, lib):
        c = self.f.cherche(lib)
        if not c:
            return None
        fin = self.fusions.get(c.adr)
        col = decoupe_adresse(fin)[0] if fin else c.col
        return num_en_col(col_en_num(col) + 1) + str(c.lig)


def sem_type(t):
    n = normalise(t or "")
    if n.startswith("personnel"):
        return "P"
    if n.startswith("split 50"):
        return "50"
    if n.startswith("split prop"):
        return "SP"
    if n.startswith("rembo"):
        return "R"
    return None


def date_de(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
        return serie_en_date(v)
    return None


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class Application:
    def __init__(self, config):
        self.config = config
        self.verrou = threading.RLock()
        self.mtime = None
        self.wb = None
        self.charge()

    @property
    def chemin(self):
        return self.config["fichier"]

    def charge(self, force=False):
        with self.verrou:
            mt = os.path.getmtime(self.chemin)
            if force or self.wb is None or mt != self.mtime:
                self.wb = Classeur(self.chemin)
                self.mtime = mt
                self._reinit()

    def _reinit(self):
        self.modeles = {}
        self.en_cours = set()
        self._partiels = {}
        self.memo = {}
        self.dispos = {}
        self._listes()
        self._classe()

    # --- listes de référence ---
    def _listes(self):
        f = self.wb.feuille("Listes")
        h = f.cherche("Mot-clé (description)")
        self.l_entete = h.lig
        def colonne(lib):
            c = f.cherche(lib, debut=h.lig, fin=h.lig)
            vals, r = [], h.lig + 1
            last = h.lig
            while r <= f.max_lig():
                v = f.val(c.col + str(r))
                if isinstance(v, str) and v.strip():
                    vals.append(v)
                    last = r
                r += 1
            return c.col, vals, last
        self.l_cat_col, self.categories_brutes, self.l_cat_der = colonne("Catégorie")
        self.categories = [v.strip() for v in self.categories_brutes]
        self.l_type_col, self.types_bruts, self.l_type_der = colonne("Type de dépenses")
        self.types = [v.strip() for v in self.types_bruts]
        self.l_inv_col, inv, self.l_inv_der = colonne("Type d'investissements")
        self.types_invest = [v.strip() for v in inv]
        kc = h.col
        vc = num_en_col(col_en_num(kc) + 1)
        self.l_mot_col, self.l_motcat_col = kc, vc
        self.motscles = []
        r = h.lig + 1
        self.l_mot_der = h.lig
        self.l_mot_debut = h.lig + 1
        while r <= f.max_lig():
            k, v = f.val(kc + str(r)), f.val(vc + str(r))
            if isinstance(k, str) and k != "":
                self.motscles.append((k, (v or "").strip() if isinstance(v, str) else ""))
                self.l_mot_der = r
            r += 1

    def categorie_auto(self, desc):
        if not desc or not desc.strip():
            return ""
        d = desc.lower()
        res = "Autres"
        for k, v in self.motscles:
            if k.lower() in d:
                res = v
        return res

    def type_exact(self, t):
        """Valeur exacte (avec espaces) de la liste pour un type saisi."""
        for b in self.types_bruts:
            if normalise(b) == normalise(t):
                return b
        raise Erreur("Type de dépense inconnu : %s" % t)

    # --- classement des onglets ---
    def _classe(self):
        premier = parse_mois(self.config.get("premier_mois_modifiable", ""))
        self.mois = []
        for f in self.wb.feuilles:
            pm = parse_mois(f.nom)
            if not pm:
                continue
            canvas = f.cherche("Déboursé") is not None
            if canvas:
                groupe = "final" if (premier and pm >= premier) else "reference"
            else:
                groupe = "historique"
            self.mois.append({"feuille": f.nom, "nom": f.nom.strip(), "cle": "%04d-%02d" % pm,
                              "annee": pm[0], "mois": pm[1], "groupe": groupe,
                              "canvas": canvas, "modifiable": groupe == "final"})
        self.mois.sort(key=lambda m: m["cle"])

    def info_mois(self, feuille):
        for m in self.mois:
            if m["feuille"] == feuille:
                return m
        raise Erreur("Mois introuvable : %s" % feuille)

    def dispo(self, nom):
        if nom not in self.dispos:
            self.dispos[nom] = Disposition(self.wb.feuille(nom))
        return self.dispos[nom]

    # --- évaluation ---
    def lire(self, feuille, adr):
        cle = (feuille, adr)
        if cle in self.memo:
            return self.memo[cle]
        if feuille not in self.wb.par_nom:
            raise ErreurExcel("#REF!")
        info = next((m for m in self.mois if m["feuille"] == feuille), None)
        if info and info["canvas"]:
            mod = self.modeles.get(feuille)
            if mod is None and feuille not in self.en_cours:
                mod = self.modele(feuille)
            if mod is None:
                mod = self._partiels.get(feuille)
            if mod is not None and adr in mod["_calc"]:
                return mod["_calc"][adr]
        f = self.wb.feuille(feuille)
        c = f.cell(adr)
        if c is None:
            return None
        if c.formule:
            try:
                v = Evaluateur(self.lire, feuille).evalue(c.formule)
            except NonSupporte:
                v = c.valeur
            self.memo[cle] = v
            return v
        if isinstance(c.valeur, tuple):
            raise ErreurExcel(c.valeur[1])
        return c.valeur

    def eval_cell(self, feuille, adr):
        try:
            return self.lire(feuille, adr)
        except ErreurExcel as e:
            return "#" + e.code.strip("#")

    # --- modèle d'un mois au format Canvas ---
    def modele(self, feuille):
        if feuille in self.modeles:
            return self.modeles[feuille]
        if not hasattr(self, "_partiels"):
            self._partiels = {}
        self.en_cours.add(feuille)
        try:
            m = self._calcule(feuille)
        finally:
            self.en_cours.discard(feuille)
            self._partiels.pop(feuille, None)
        self.modeles[feuille] = m
        return m

    def _calcule(self, nom):
        f = self.wb.feuille(nom)
        d = self.dispo(nom)
        calc = {}
        m = {"_calc": calc}
        self._partiels[nom] = m
        info = self.info_mois(nom)
        ev = lambda adr: self.eval_cell(nom, adr) if adr else None

        # % Marika utilisé
        pct_cell = f.cell(d.a_pct_utilise) if d.a_pct_utilise else None
        pct_manuel = pct_cell is not None and not pct_cell.formule and num(pct_cell.valeur) is not None
        pct = num(ev(d.a_pct_utilise))
        marika_rev = num(ev(d.a_marika)) if d.a_marika else None
        zach_rev = num(ev(d.a_zach)) if d.a_zach else None
        pct_repart = None
        if d.a_pct:
            pr = ev(d.a_pct)
            pct_repart = num(pr)

        trans = []
        der_marq = None
        erreurs = 0
        for r in d.lignes:
            a = lambda c: (c + str(r)) if c else None
            date_v = f.val(a(d.c_date))
            desc = f.texte(a(d.c_desc))
            typ = f.texte(a(d.c_type))
            cd = f.cell(a(d.c_deb))
            deb_brut = cd.valeur if cd else None
            deb = num(ev(a(d.c_deb))) if cd is not None and (cd.formule or cd.valeur not in (None, "")) else None
            cg = f.cell(a(d.c_cat))
            override = cg is not None and not cg.formule and isinstance(cg.valeur, str) and cg.valeur.strip() != ""
            marq = "solde réglé" in normalise(f.texte(a(d.c_marq)))
            utilise = bool(date_v not in (None, "") or desc.strip() or typ.strip() or deb_brut not in (None, "")
                           or override or marq)
            st = sem_type(typ)
            exact = typ in self.types_bruts or not typ
            prop = {"P": 1.0, "50": 0.5, "SP": (1 - pct) if pct is not None else None,
                    "R": (1 - pct) if pct is not None else None}.get(st) if exact else None
            erreur = False
            dep = None
            if deb is not None and deb > 0:
                if prop is None:
                    erreur = True
                else:
                    dep = deb * prop
            elif deb_brut not in (None, "") and deb is None:
                erreur = True  # texte dans Déboursé
            cat_auto = self.categorie_auto(desc)
            cat = cg.valeur.strip() if override else cat_auto
            km = None
            if exact and st == "P":
                km = 0.0
            elif exact and st == "SP":
                if dep is None or pct is None:
                    erreur = erreur or bool(typ)
                else:
                    km = dep / (1 - pct) * pct
            elif exact and st == "50":
                km = dep
            elif exact and st == "R":
                if dep is None:
                    erreur = erreur or bool(typ)
                else:
                    km = -dep
            if erreur:
                erreurs += 1
            if marq:
                der_marq = r
            calc[d.c_dep + str(r)] = dep if dep is not None else ""
            calc[d.c_cat + str(r)] = cat
            if d.c_prop:
                calc[d.c_prop + str(r)] = prop if prop is not None else ""
            if d.c_mmont:
                calc[d.c_mmont + str(r)] = km if km is not None else ""
            trans.append({"ligne": r, "utilise": utilise, "date": date_de(date_v).isoformat() if date_de(date_v) else None,
                          "date_brute": None if date_de(date_v) or date_v is None else str(date_v),
                          "description": desc.strip(), "type": typ.strip(), "type_exact": exact,
                          "deboursement": deb, "deb_formule": ("=" + cd.formule) if cd is not None and cd.formule else None,
                          "proportion": prop, "depense": dep, "categorie": cat, "cat_auto": cat_auto,
                          "cat_manuelle": override, "marika": km, "marqueur": marq, "erreur": erreur})
        total = sum(t["depense"] or 0 for t in trans)
        k_total = sum(t["marika"] or 0 for t in trans)
        calc[d.c_dep + str(d.lig_total)] = total
        if d.c_mmont:
            calc[d.c_mmont + str(d.lig_total)] = k_total
        du_depuis = sum((t["marika"] or 0) for t in trans if der_marq is None or t["ligne"] > der_marq)
        if d.lig_tableau:
            calc[d.c_rval + str(d.lig_tableau)] = du_depuis
        # catégories
        par_cat = {}
        for t in trans:
            if t["depense"]:
                par_cat[t["categorie"]] = par_cat.get(t["categorie"], 0) + t["depense"]
        resume_cats = []
        for r in d.cats:
            c = f.val("A" + str(r)).strip()
            v = sum(val for k, val in par_cat.items() if k.lower() == c.lower())
            calc["B" + str(r)] = v
            resume_cats.append({"categorie": c, "montant": v})
        hors_resume = {k: v for k, v in par_cat.items()
                       if k.lower() not in [x["categorie"].lower() for x in resume_cats]}
        # lignes « report » du tableau Remboursement
        reports = []
        if d.lig_tableau and d.lig_rtotal:
            for r in range(d.lig_tableau + 1, d.lig_rtotal):
                lib = f.texte(d.c_rlib + str(r)).strip()
                c = f.cell(d.c_rval + str(r))
                if lib or (c is not None and (c.formule or c.valeur not in (None, ""))):
                    fx = c.formule if c is not None and c.formule else None
                    origine = None
                    mo = re.match(r'^IF\(COUNTIF\([^)]*\)>0,0,(.*)\)$', fx or "", re.S)
                    if mo:
                        try:
                            origine = num(Evaluateur(self.lire, nom).evalue(mo.group(1)))
                        except (NonSupporte, ErreurExcel):
                            origine = None
                    reports.append({"ligne": r, "libelle": lib, "montant": num(ev(d.c_rval + str(r))) or 0.0,
                                    "auto": bool(fx and "!" in fx), "formule": ("=" + fx) if fx else None,
                                    "regle": bool(fx and "COUNTIF" in fx.upper() and der_marq is not None),
                                    "montant_origine": origine, "expr_origine": ("=" + mo.group(1)) if mo else None})
        du_total = num(ev(d.c_rval + str(d.lig_rtotal))) if d.lig_rtotal else du_depuis
        if du_total is None:
            du_total = du_depuis + sum(x["montant"] for x in reports)
        # pourboires
        pourb = []
        pb = d.pb
        for r in d.pourboires:
            dv = f.val(pb["date"] + str(r))
            b, c = f.cell(pb["recu"] + str(r)), f.cell(pb["arec"] + str(r))
            h = f.cell(pb["heures"] + str(r)) if pb["heures"] else None
            vide = lambda x: x is None or (x.valeur in (None, "") and not x.formule)
            if dv in (None, "") and vide(b) and vide(c) and vide(h):
                pourb.append({"ligne": r, "utilise": False})
                continue
            pourb.append({"ligne": r, "utilise": True,
                          "date": date_de(dv).isoformat() if date_de(dv) else None,
                          "recu": num(ev(pb["recu"] + str(r))), "a_recevoir": num(ev(pb["arec"] + str(r))),
                          "heures": num(ev(pb["heures"] + str(r))) if pb["heures"] else None,
                          "recu_txt": ("=" + b.formule) if b is not None and b.formule else None,
                          "a_recevoir_txt": ("=" + c.formule) if c is not None and c.formule else None,
                          "statut": f.texte(pb["statut"] + str(r)).strip() if pb["statut"] else ""})
        # revenus globaux
        revs = []
        for r in d.revenus:
            dv = f.val("A" + str(r))
            b = f.cell("B" + str(r))
            desc = f.texte("C" + str(r)).strip()
            if dv in (None, "") and not desc and (b is None or (b.valeur in (None, "") and not b.formule)):
                revs.append({"ligne": r, "utilise": False})
                continue
            revs.append({"ligne": r, "utilise": True, "date": date_de(dv).isoformat() if date_de(dv) else None,
                         "montant": num(ev("B" + str(r))),
                         "montant_txt": ("=" + b.formule) if b is not None and b.formule else None,
                         "description": desc,
                         "cases": self._cases_formule(b.formule if b is not None else None, d)})
        total_rev = sum(x.get("montant") or 0 for x in revs if x["utilise"])
        total_fio = sum(x.get("montant") or 0 for x in revs if x["utilise"] and self.re_revenu_principal().search(x.get("description") or ""))
        # investissements
        invs = []
        for r in d.invest:
            t = f.texte("A" + str(r)).strip()
            if not t:
                continue
            bi, cj = num(ev("B" + str(r))), num(ev("C" + str(r)))
            invs.append({"ligne": r, "type": t, "investi": bi, "jv": cj,
                         "variation": (cj - bi) if bi is not None and cj is not None else None})
        # résumé
        debut = num(ev(d.a_debut))
        retraits = [{"libelle": x["libelle"], "compte": x["compte"], "adr": x["adr"],
                     "montant": num(ev(x["adr"])) or 0.0} for x in d.retraits]
        retrait = sum(x["montant"] for x in retraits)
        net = total_rev - total
        fin_std = (debut or 0.0) + net - retrait
        fin_excel = num(ev(d.a_fin))
        fc = f.cell(d.a_fin) if d.a_fin else None
        canvas_fin = self.wb.feuille("Canvas").cell(d.a_fin) if "Canvas" in self.wb.par_nom else None
        fin_formule = fc.formule if fc is not None else None
        reel = num(ev(d.a_reel)) if d.a_reel else None
        # contrôles
        ctls = []
        pm = parse_mois(f.texte(d.a_mois)) if d.a_mois else None
        for r in d.controles:
            lib = f.texte(d.c_ctl_lib + str(r)).strip()
            n = normalise(lib)
            forcer = normalise(f.texte(d.c_ctl_forcer + str(r))) == "oui" if d.c_ctl_forcer else False
            just = f.texte(d.c_ctl_just + str(r)).strip() if d.c_ctl_just else ""
            ecart, source = None, "app"
            if "catégories" in n:
                ecart = r2(sum(x["montant"] for x in resume_cats) - total)
            elif "dates hors" in n:
                if pm:
                    ecart = 0
                    for t in trans:
                        dd = date_de(f.val(d.c_date + str(t["ligne"])))
                        if t["date"] is None and t["date_brute"] is None:
                            continue
                        if dd is None or (dd.year, dd.month) != pm:
                            ecart += 1
            elif "erreurs" in n:
                ecart = erreurs
            elif "solde calculé" in n:
                ecart = r2((fin_excel or 0) - reel) if reel is not None else None
            elif "recalculée" in n:
                h = 0.0
                for t in trans:
                    st, dv = sem_type(t["type"]), t["deboursement"] or 0
                    if st == "SP":
                        h += dv * (pct or 0)
                    elif st == "50":
                        h += dv * 0.5
                    elif st == "R":
                        h -= dv * (1 - (pct or 0))
                ecart = r2(h - k_total)
            elif "dernier solde réglé" in n:
                h = 0.0
                for t in trans:
                    if der_marq is not None and t["ligne"] <= der_marq:
                        continue
                    st, dv = sem_type(t["type"]), t["deboursement"] or 0
                    if st == "SP":
                        h += dv * (pct or 0)
                    elif st == "50":
                        h += dv * 0.5
                    elif st == "R":
                        h -= dv * (1 - (pct or 0))
                ecart = r2(h - du_depuis)
            elif "proportion" in n:
                ecart = (round((pct or 0) - pct_repart, 4) + 0.0 or 0.0) if pct_repart is not None else None
            else:
                source = "excel"
                ev_ = f.val(d.c_ctl_ecart + str(r))
                ecart = num(ev_)
            if forcer:
                statut = "OK (forcé)"
            elif ecart is None:
                statut = "À compléter"
            elif abs(ecart) < 0.005:
                statut = "OK"
            else:
                statut = "À vérifier"
            ctls.append({"ligne": r, "libelle": lib, "ecart": ecart, "statut": statut, "forcer": forcer,
                         "justification": just, "source": source})
        nb_verif = sum(1 for c in ctls if "vérifier" in c["statut"])
        # ajustements manuels détectés
        ajust = []
        if pct_manuel:
            ajust.append("%% Marika saisi à la main (%s) au lieu de la répartition des revenus" % (
                "{:.0%}".format(pct) if pct is not None else "?"))
        if fin_excel is not None and abs(fin_excel - fin_std) > 0.005:
            ajust.append("Solde de fin : la formule du fichier (=%s) donne %.2f $ au lieu de %.2f $ (écart %.2f $)"
                         % (fin_formule, fin_excel, fin_std, fin_excel - fin_std))
        for t in trans:
            if t["cat_manuelle"] and t["categorie"] != t["cat_auto"]:
                ajust.append("Catégorie corrigée à la main ligne %d : « %s » → %s (auto : %s)"
                             % (t["ligne"], t["description"], t["categorie"], t["cat_auto"]))
        m.update({
            "feuille": nom, "nom": nom.strip(), "groupe": info["groupe"], "modifiable": info["modifiable"],
            "pct_marika": pct, "pct_manuel": pct_manuel, "revenu_marika": marika_rev, "revenu_zach": zach_rev,
            "pct_repartition": pct_repart, "transactions": trans, "total_depenses": total,
            "total_split": k_total, "du_depuis_reglement": du_depuis, "reports": reports, "du_total": du_total,
            "ligne_dernier_reglement": der_marq, "par_categorie": resume_cats, "hors_resume": hors_resume,
            "pourboires": pourb,
            "total_pourboires": {"recu": sum(p.get("recu") or 0 for p in pourb if p["utilise"]),
                                 "a_recevoir": sum(p.get("a_recevoir") or 0 for p in pourb if p["utilise"]),
                                 "heures": sum(p.get("heures") or 0 for p in pourb if p["utilise"])},
            "colonne_heures": bool(d.pb["heures"]),
            "revenus": revs, "total_revenus": total_rev, "total_fiorellino": total_fio, "investissements": invs,
            "total_investi": sum(i["investi"] or 0 for i in invs), "total_jv": sum(i["jv"] or 0 for i in invs),
            "solde_debut": debut, "retrait": retrait, "retraits": retraits, "revenu_net": net, "solde_fin": fin_excel if fin_excel is not None else fin_std,
            "solde_fin_standard": fin_std, "solde_reel": reel, "controles": ctls,
            "resume_controles": "Tout est OK" if nb_verif == 0 else "%d contrôle(s) à vérifier" % nb_verif,
            "nb_a_verifier": nb_verif, "ajustements": ajust, "mois_param": f.texte(d.a_mois).strip() if d.a_mois else "",
            "lignes_libres": sum(1 for t in trans if not t["utilise"]),
        })
        return m

    # --- historique ---
    def historique(self, nom):
        cle = ("hist", nom)
        if cle in self.memo:
            return self.memo[cle]
        f = self.wb.feuille(nom)
        h = f.cherche("Date", cols=["A"])
        ent = {normalise(c.valeur): c.col for c in f.cellules.values() if c.lig == h.lig and isinstance(c.valeur, str)}
        c_mont = ent.get("montant")
        c_cat = ent.get("catégorie")
        c_desc = ent.get("description")
        tot = f.cherche("Total dépenses")
        total_c = num_en_col(col_en_num(tot.col) + 1) + str(tot.lig)
        total = num(f.val(total_c))
        fc = f.cell(total_c)
        lig_deb, lig_fin = h.lig + 1, tot.lig - 1
        if fc is not None and fc.formule:
            mm = re.match(r"SUM\(\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)\)", fc.formule.replace(" ", ""), re.I)
            if mm:
                lig_deb, lig_fin = int(mm.group(1)), int(mm.group(2))
        lignes, par_cat = [], {}
        for r in range(lig_deb, lig_fin + 1):
            v = num(f.val(c_mont + str(r)))
            if v is None:
                continue
            cat = f.texte(c_cat + str(r)).strip() or "(sans catégorie)"
            par_cat[cat] = par_cat.get(cat, 0) + v
            dd = date_de(f.val("A" + str(r)))
            lignes.append({"date": dd.isoformat() if dd else None, "description": f.texte(c_desc + str(r)).strip(),
                           "montant": v, "categorie": cat})
        rg = f.cherche("Résumé global", cols=["A"])
        def lab(prefixe):
            for r in range(rg.lig + 1, rg.lig + 12):
                t = normalise(f.texte("A" + str(r)))
                if t.startswith(prefixe):
                    return num(f.val("B" + str(r)))
            return None
        revenus = lab("revenus")
        alias = {normalise(k): v for k, v in self.config.get("alias_investissements", {}).items()}
        i = f.cherche("Investissements", cols=["A"])
        ti = f.cherche("Total investissements", cols=["A"])
        invs = []
        if i and ti:
            for r in range(i.lig + 2, ti.lig):
                t = f.texte("A" + str(r)).strip()
                if t:
                    invs.append({"type": alias.get(normalise(t), t), "investi": num(f.val("B" + str(r))),
                                 "jv": num(f.val("C" + str(r)))})
        res = {"feuille": nom, "nom": nom.strip(), "groupe": "historique", "lignes": lignes,
               "total_depenses": total, "somme_lignes": sum(l["montant"] for l in lignes),
               "par_categorie": [{"categorie": k, "montant": v} for k, v in par_cat.items()],
               "total_revenus": revenus, "revenu_net": lab("net"),
               "solde_debut": lab("solde au compte en début"), "retrait": lab("épargne"),
               "solde_fin": lab("solde au compte en fin"), "investissements": invs,
               "total_investi": sum(x["investi"] or 0 for x in invs), "total_jv": sum(x["jv"] or 0 for x in invs)}
        self.memo[cle] = res
        return res

    def tendances(self):
        out = []
        for m in self.mois:
            if m["canvas"]:
                d = self.modele(m["feuille"])
                cats = {c["categorie"]: c["montant"] for c in d["par_categorie"]}
                for k, v in d["hors_resume"].items():
                    cats[k] = cats.get(k, 0) + v
            else:
                d = self.historique(m["feuille"])
                cats = {c["categorie"]: c["montant"] for c in d["par_categorie"]}
            out.append({"feuille": m["feuille"], "nom": m["nom"], "cle": m["cle"], "groupe": m["groupe"],
                        "total_depenses": d["total_depenses"], "par_categorie": cats,
                        "revenus": d["total_revenus"], "revenu_net": d["revenu_net"], "solde_fin": d["solde_fin"],
                        "investi": d["total_investi"], "jv": d["total_jv"],
                        "epargne_placee": sum(x.get("montant") or 0 for x in d.get("retraits") or [])})
        return out

    # --- réglages propres à chaque fichier (bloc « Réglages de l'app » de l'onglet Listes) ---
    REGLAGES_DEFAUT = {"prenom": "Zach", "partenaire": "Marika", "pourboires": True,
                       "revenu_principal": "Revenu Fiorellino",
                       "descriptions_revenu_principal": ["Dépôts d'espèces", "Salaire horaire", "Pourboire"],
                       "mots_revenu_principal": "fiorellino|esp[eè]ces|salaire horaire|pourboire",
                       "configuration_a_faire": False}
    CLES_REGLAGES = {"prénom": "prenom", "partenaire": "partenaire", "pourboires": "pourboires",
                     "revenu principal": "revenu_principal",
                     "descriptions du revenu principal": "descriptions_revenu_principal",
                     "mots du revenu principal": "mots_revenu_principal",
                     "configuration de départ": "configuration_a_faire"}

    def _bloc_reglages(self):
        f = self.wb.feuille("Listes")
        h = f.cherche("Réglages de l'app")
        if not h:
            return None, {}
        vc = num_en_col(col_en_num(h.col) + 1)
        lignes = {}
        r = h.lig + 1
        while f.texte(h.col + str(r)).strip():
            lignes[normalise(f.texte(h.col + str(r)))] = (vc + str(r), f.texte(vc + str(r)).strip())
            r += 1
        return h, lignes

    def reglages(self):
        if "reglages" in self.memo:
            return self.memo["reglages"]
        reg = dict(self.REGLAGES_DEFAUT)
        _, lignes = self._bloc_reglages()
        for k, (adr, v) in lignes.items():
            cle = self.CLES_REGLAGES.get(k)
            if not cle:
                continue
            if cle == "pourboires":
                reg[cle] = normalise(v) in ("oui", "o", "vrai")
            elif cle == "configuration_a_faire":
                reg[cle] = normalise(v).startswith("à faire") or normalise(v).startswith("a faire")
            elif cle == "descriptions_revenu_principal":
                reg[cle] = [x.strip() for x in v.split(";") if x.strip()]
            elif v:
                reg[cle] = v
        self.memo["reglages"] = reg
        return reg

    def re_revenu_principal(self):
        try:
            return re.compile(self.reglages()["mots_revenu_principal"], re.I)
        except re.error:
            return self.RE_FIORELLINO

    def etat(self):
        return {"mois": self.mois, "categories": self.categories, "types": self.types, "reglages": self.reglages(),
                "types_invest": self.types_invest,
                "motscles": [{"mot": k, "categorie": v} for k, v in self.motscles],
                "fichier": self.chemin, "excel_ouvert": self.excel_ouvert(),
                "prochain_mois": self.prochain_mois(),
                "descriptions_autres_revenus": self.descriptions_autres_revenus()}

    RE_FIORELLINO = re.compile(r"fiorellino|esp[eè]ces|salaire horaire|pourboire", re.I)

    def descriptions_autres_revenus(self, n=3):
        """Les n descriptions de revenus (hors Fiorellino) les plus fréquentes de l'historique."""
        cle_memo = ("desc_rev", n)
        if cle_memo in self.memo:
            return self.memo[cle_memo]
        import unicodedata
        from collections import Counter, defaultdict

        def cle(t):
            t = unicodedata.normalize("NFKD", t.lower())
            t = "".join(ch for ch in t if not unicodedata.combining(ch))
            mots = [m.rstrip("s") for m in re.findall(r"[a-z0-9]+", t) if m not in ("sur", "de", "du", "d", "l", "la", "le")]
            return " ".join(mots)
        freq, variantes = Counter(), defaultdict(Counter)
        for f in self.wb.feuilles:
            if not parse_mois(f.nom):
                continue
            # sections « Date | Montant | Description » (revenus)
            for c in f.cellules.values():
                if c.col != "A" or normalise(f.texte(c.adr)) != "date":
                    continue
                r = c.lig
                if normalise(f.texte("B%d" % r)) != "montant" or normalise(f.texte("C%d" % r)) != "description":
                    continue
                r += 1
                while r <= f.max_lig() and "total" not in normalise(f.texte("A%d" % r)):
                    d = re.sub(r"\s+", " ", f.texte("C%d" % r)).strip()
                    if d and num(f.val("B%d" % r)) is not None and not self.re_revenu_principal().search(d):
                        k = cle(d)
                        freq[k] += 1
                        variantes[k][d] += 1
                    r += 1
        res = [variantes[k].most_common(1)[0][0] for k, _ in freq.most_common(n)]
        self.memo[cle_memo] = res
        return res

    def prochain_mois(self):
        if not self.mois:
            return None
        a, mm = self.mois[-1]["annee"], self.mois[-1]["mois"] + 1
        if mm > 12:
            a, mm = a + 1, 1
        return nom_mois(a, mm)

    # =====================================================================
    # Écriture
    # =====================================================================
    def excel_ouvert(self):
        d, b = os.path.split(os.path.abspath(self.chemin))
        return os.path.exists(os.path.join(d, "~$" + b))

    def _ecriture(self, fn):
        with self.verrou:
            if self.excel_ouvert():
                raise Erreur("Le fichier est ouvert dans Excel. Ferme-le (Cmd+Q dans Excel), puis réessaie.")
            self.charge()
            dossier = self.config.get("dossier_sauvegardes") or os.path.join(
                os.path.dirname(os.path.abspath(self.chemin)), "Sauvegardes")
            copie_sauvegarde(self.chemin, dossier)
            try:
                res = fn()
                self.wb.enregistre()
            finally:
                self.charge(force=True)
            return res

    def _verifie_modifiable(self, feuille):
        info = self.info_mois(feuille)
        if not info["modifiable"]:
            raise Erreur("%s est en lecture seule dans l'application." % info["nom"])
        return self.wb.feuille(feuille), self.dispo(feuille)

    @staticmethod
    def montant_saisi(txt):
        """« 12,50 » → 12.5 ; « 33,34+15,53 » → formule. Renvoie (valeur, formule)."""
        if txt is None:
            return None, None
        if isinstance(txt, (int, float)):
            return float(txt), None
        t = str(txt).strip().replace(" ", "").replace(" ", "").replace("$", "")
        if t == "":
            return None, None
        if t.startswith("="):
            t = t[1:]
        t = t.replace(",", ".")
        if re.fullmatch(r"-?\d+(\.\d+)?", t):
            return float(t), None
        if not re.fullmatch(r"[\d.+\-*/() ]+", t):
            raise Erreur("Montant invalide : %s" % txt)
        try:
            v = Evaluateur(lambda f, a: None, "").evalue(t)
        except Exception:
            raise Erreur("Montant invalide : %s" % txt)
        return float(v), "=" + t

    def _ecrit_montant(self, f, adr, txt):
        v, formule = self.montant_saisi(txt)
        if formule:
            f.ecrit(adr, valeur=v, formule=formule)
        else:
            f.ecrit(adr, valeur=v)

    @staticmethod
    def _date(txt):
        if not txt:
            return None
        if isinstance(txt, datetime.date):
            return txt
        t = str(txt).strip()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(t, fmt).date()
            except ValueError:
                pass
        raise Erreur("Date invalide : %s" % txt)

    # --- transactions ---
    def _lignes_utilisees(self, nom):
        m = self.modele(nom)
        return [t["ligne"] for t in m["transactions"] if t["utilise"]]

    def _restaure_categorie(self, f, d, r):
        c = f.cell(d.c_cat + str(r))
        if c is not None and c.formule:
            return
        xml = f.formule_modele(d.c_cat, r, d.lignes[0], d.lignes[-1])
        if xml is None:
            can = self.wb.feuille("Canvas")
            cd = self.dispo("Canvas")
            xml = can.formule_modele(cd.c_cat, r, cd.lignes[0], cd.lignes[-1])
        f.ecrit_xml_brut(d.c_cat + str(r), xml)

    def _lit_entrees(self, f, d, r):
        def brut(col):
            c = f.cell(col + str(r))
            if c is None:
                return None
            if c.formule:
                return ("f", c.formule, c.valeur)
            return ("v", c.valeur)
        cg = f.cell(d.c_cat + str(r))
        cat = cg.valeur if (cg is not None and not cg.formule and isinstance(cg.valeur, str) and cg.valeur.strip()) else None
        return {"A": brut(d.c_date), "B": brut(d.c_desc), "C": brut(d.c_type), "D": brut(d.c_deb),
                "cat": cat, "H": brut(d.c_marq)}

    def _ecrit_entrees(self, f, d, r, e):
        for cle, col in (("A", d.c_date), ("B", d.c_desc), ("C", d.c_type), ("D", d.c_deb), ("H", d.c_marq)):
            v = e.get(cle) if e else None
            adr = col + str(r)
            if v is None:
                f.ecrit(adr, None)
            elif v[0] == "f":
                f.ecrit(adr, valeur=num(v[2]), formule=v[1])
            else:
                f.ecrit(adr, valeur=serie_en_date(v[1]) if cle == "A" and isinstance(v[1], float) else v[1])
        if e and e.get("cat"):
            f.ecrit(d.c_cat + str(r), e["cat"])
        else:
            self._restaure_categorie(f, d, r)

    def _assure_lignes(self, nom):
        """Garantit au moins une ligne de saisie libre après la dernière utilisée."""
        f, d = self.wb.feuille(nom), self.dispo(nom)
        utilisees = self._lignes_utilisees(nom)
        der = max(utilisees) if utilisees else d.lignes[0] - 1
        if der < d.lignes[-1]:
            return der + 1
        # insérer 10 lignes en recopiant les formules (comme Insérer dans Excel)
        n = 10
        a_partir = d.lignes[-1]
        entrees = self._lit_entrees(f, d, a_partir)
        cols = [num_en_col(k) for k in range(col_en_num(d.c_date), col_en_num(d.c_mmont or d.c_marq) + 2)]
        self.wb.insere_lignes(nom, a_partir, n, lig_modele=a_partir - 1, cols_modele=cols)
        self.dispos.pop(nom, None)
        self.modeles.clear(); self.memo.clear()
        d = self.dispo(nom)
        f = self.wb.feuille(nom)
        # la dernière ligne utilisée a été poussée en bas : on la remonte
        self._ecrit_entrees(f, d, a_partir, entrees)
        self._ecrit_entrees(f, d, a_partir + n, None)
        self.modeles.clear(); self.memo.clear()
        return a_partir + 1

    def _ecrit_transaction(self, feuille, data):
        f, d = self._verifie_modifiable(feuille)
        r = data.get("ligne")
        if r is None:
            r = self._assure_lignes(feuille)
            f, d = self.wb.feuille(feuille), self.dispo(feuille)
        elif r not in d.lignes:
            raise Erreur("Ligne invalide")
        date = self._date(data.get("date"))
        desc = (data.get("description") or "").strip()
        if not desc:
            raise Erreur("La description est obligatoire.")
        typ = self.type_exact(data.get("type") or "")
        f.ecrit(d.c_date + str(r), date)
        f.ecrit(d.c_desc + str(r), desc)
        f.ecrit(d.c_type + str(r), typ)
        self._ecrit_montant(f, d.c_deb + str(r), data.get("montant"))
        cat = (data.get("categorie") or "").strip()
        if cat and cat != self.categorie_auto(desc):
            f.ecrit(d.c_cat + str(r), cat)
        else:
            self._restaure_categorie(f, d, r)
        return r

    def enregistre_transaction(self, feuille, data):
        return self._ecriture(lambda: self._ecrit_transaction(feuille, data))

    def enregistre_transactions_lot(self, feuille, liste):
        """Ajoute plusieurs dépenses en une seule écriture (import de relevé)."""
        if not liste:
            raise Erreur("Aucune dépense à ajouter.")
        info = self.info_mois(feuille)
        for k, t in enumerate(liste, 1):
            if not (t.get("description") or "").strip():
                raise Erreur("Ligne %d : description manquante." % k)
            v, _ = self.montant_saisi(t.get("montant"))
            if v is None:
                raise Erreur("Ligne %d : montant invalide." % k)
            dt = self._date(t.get("date"))
            if dt is None or dt.strftime("%Y-%m") != info["cle"]:
                raise Erreur("Ligne %d : la date doit être dans %s." % (k, info["nom"]))

        def op():
            lignes = []
            for t in liste:
                t = dict(t); t.pop("ligne", None)
                lignes.append(self._ecrit_transaction(feuille, t))
                self.modeles.clear(); self.memo.clear()
            return lignes
        return self._ecriture(op)

    def analyse_releve(self, feuille, images):
        import releve
        info = self.info_mois(feuille)
        existantes = []
        if info["canvas"]:
            m = self.modele(feuille)
            for t in m["transactions"]:
                if t["utilise"]:
                    existantes.append({"date": t.get("date"), "montant": t.get("deboursement"), "libelle": t.get("description")})
            # montants manuels du tableau Remboursement (ex. achats faits pour Marika)
            for r in m.get("reports") or []:
                if r.get("montant") and not r.get("auto"):
                    existantes.append({"date": None, "montant": abs(r["montant"]), "libelle": r.get("libelle")})
        revenus = []
        if info["canvas"]:
            for r in self.modele(feuille).get("revenus") or []:
                if r.get("utilise") and r.get("montant") is not None:
                    # « =1402.63-200 » : le dépôt de 1 402,63 $ est aussi reconnu
                    montants = [r["montant"]] + [float(x) for x in re.findall(r"(?<![\d.])\d+\.\d{2}(?![\d])", str(r.get("montant_txt") or ""))[:1]]
                    revenus.append({"date": r.get("date"), "montants": montants, "libelle": r.get("description")})
        lignes = releve.analyse(images or [], info["cle"], existantes, self.categorie_auto, revenus)
        return {"mois": info["nom"], "cle": info["cle"], "modifiable": info["modifiable"], "lignes": lignes}

    def supprime_transaction(self, feuille, ligne):
        def op():
            f, d = self._verifie_modifiable(feuille)
            utilisees = self._lignes_utilisees(feuille)
            if ligne not in utilisees:
                raise Erreur("Ligne vide")
            der = max(utilisees)
            for r in range(ligne, der):
                self._ecrit_entrees(f, d, r, self._lit_entrees(f, d, r + 1))
            self._ecrit_entrees(f, d, der, None)
        return self._ecriture(op)

    def marque_reglement(self, feuille, ligne):
        def op():
            f, d = self._verifie_modifiable(feuille)
            m = re.search(r'<dataValidation [^>]*sqref="[^"]*%s\d+[^"]*"[^>]*>\s*<formula1>"([^"]*)"</formula1>' % d.c_marq, f.xml)
            marque = m.group(1) if m and "réglé" in m.group(1) else "── Solde réglé ──"
            for t in self.modele(feuille)["transactions"]:
                if t["marqueur"]:
                    f.ecrit(d.c_marq + str(t["ligne"]), None)
            if ligne:
                f.ecrit(d.c_marq + str(ligne), marque)
                self._regle_autres_montants(f, d)
        return self._ecriture(op)

    def _condition_reglement(self, d):
        plage = "%s%d:%s%d" % (d.c_marq, d.lignes[0], d.c_marq, d.lignes[-1])
        return 'COUNTIF(%s,"*Solde réglé*")>0' % plage

    def _regle_autres_montants(self, f, d):
        """Les « autres montants » (report du mois précédent, ajouts manuels) existants
        sont considérés comme réglés : ils valent 0 tant qu'un « Solde réglé » est présent."""
        if not (d.lig_tableau and d.lig_rtotal):
            return
        cond = self._condition_reglement(d)
        for r in range(d.lig_tableau + 1, d.lig_rtotal):
            adr = d.c_rval + str(r)
            c = f.cell(adr)
            if c is None:
                continue
            if c.formule:
                if "COUNTIF" in c.formule.upper() and "SOLDE RÉGLÉ" in c.formule.upper():
                    continue
                expr = c.formule
            elif num(c.valeur) is not None and num(c.valeur) != 0:
                expr = repr(num(c.valeur))
            else:
                continue
            f.ecrit(adr, valeur=0.0, formule="IF(%s,0,%s)" % (cond, expr))

    # --- revenus calculés à partir des jours du calendrier des pourboires ---
    RX_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")

    def _cases_formule(self, formule, d):
        """« SUM(C68:C71,C73)+0.06 » → {"colonne": "arec", "lignes": [68,69,70,71,73], "ajout": 0.06} si la
        formule n'additionne que des cases d'une même colonne du calendrier des pourboires (plus une constante)."""
        if not formule or not d.pourboires:
            return None
        cols = {d.pb["recu"]: "recu", d.pb["arec"]: "arec"}
        t = formule.replace(" ", "").upper()
        lignes, col = [], None
        for m in self.RX_REF.finditer(t):
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3) or m.group(1), int(m.group(4) or m.group(2))
            if c1 != c2 or c1 not in cols or (col and c1 != col):
                return None
            col = c1
            lignes += list(range(min(r1, r2), max(r1, r2) + 1))
        if not lignes or any(r not in d.pourboires for r in lignes):
            return None
        reste = self.RX_REF.sub("", t)
        reste = re.sub(r"SUM\((,)*\)", "", reste)
        reste = reste.replace(",", "")
        ajout = 0.0
        if reste:
            if not re.fullmatch(r"(\+\d+(\.\d+)?)+", reste):
                return None
            ajout = sum(float(x) for x in re.findall(r"\d+(?:\.\d+)?", reste))
        return {"colonne": cols[col], "lignes": sorted(set(lignes)), "ajout": round(ajout, 2)}

    @staticmethod
    def _formule_cases(col, lignes, ajout=0.0):
        """[68,69,70,71,73] → « SUM(C68:C71,C73) » (+ constante éventuelle)."""
        lignes = sorted(set(lignes))
        plages, debut = [], None
        for i, r in enumerate(lignes):
            if debut is None:
                debut = r
            if i == len(lignes) - 1 or lignes[i + 1] != r + 1:
                plages.append("%s%d" % (col, debut) if debut == r else "%s%d:%s%d" % (col, debut, col, r))
                debut = None
        f = "SUM(%s)" % ",".join(plages)
        if ajout:
            f += "+" + ("%.2f" % ajout).rstrip("0").rstrip(".")
        return f

    def _ecrit_revenu_cases(self, f, d, r, cases, colonne, ajout=0.0):
        if colonne not in ("recu", "arec"):
            raise Erreur("Colonne invalide")
        lignes = sorted({int(x) for x in cases})
        if not lignes:
            raise Erreur("Sélectionne au moins un jour")
        utilisees = {x["ligne"] for x in self.modele(f.nom)["pourboires"] if x["utilise"]}
        if any(x not in d.pourboires or x not in utilisees for x in lignes):
            raise Erreur("Jour du calendrier invalide")
        col = d.pb[colonne]
        cle = "recu" if colonne == "recu" else "a_recevoir"
        valeurs = {x["ligne"]: x.get(cle) or 0 for x in self.modele(f.nom)["pourboires"] if x["utilise"]}
        total = round(sum(valeurs[x] for x in lignes) + (ajout or 0), 2)
        f.ecrit("B" + str(r), valeur=total, formule=self._formule_cases(col, lignes, ajout))

    def _recale_revenus_apres_suppression(self, f, d, ligne_suppr, der):
        """Une ligne du calendrier des pourboires a été retirée (les suivantes remontent d'une ligne) :
        les revenus qui additionnaient des jours du calendrier suivent leurs jours."""
        for r in d.revenus:
            c = f.cell("B" + str(r))
            cases = self._cases_formule(c.formule if c is not None else None, d)
            if not cases:
                continue
            nouvelles = [x - 1 if ligne_suppr < x <= der else x for x in cases["lignes"] if x != ligne_suppr]
            if nouvelles == cases["lignes"]:
                continue
            col = d.pb[cases["colonne"]]
            if not nouvelles:
                f.ecrit("B" + str(r), valeur=cases["ajout"] or None)
                continue
            total = sum(num(f.val(col + str(x))) or 0 for x in nouvelles) + cases["ajout"]
            f.ecrit("B" + str(r), valeur=round(total, 2), formule=self._formule_cases(col, nouvelles, cases["ajout"]))

    # --- revenus / pourboires (lignes compactées vers le haut) ---
    def _bloc(self, d, bloc):
        return {"revenus": d.revenus, "pourboires": d.pourboires}[bloc]

    def enregistre_ligne_bloc(self, feuille, bloc, data):
        def op():
            f, d = self._verifie_modifiable(feuille)
            lignes = self._bloc(d, bloc)
            mod = self.modele(feuille)[bloc]
            r = data.get("ligne")
            if r is None:
                r = self._assure_ligne_bloc(feuille, bloc)
                f, d = self.wb.feuille(feuille), self.dispo(feuille)
                lignes = self._bloc(d, bloc)
            elif r not in lignes:
                raise Erreur("Ligne invalide")
            f.ecrit("A" + str(r), self._date(data.get("date")))
            if bloc == "revenus":
                actuelle = f.cell("B" + str(r))
                if data.get("cases") is not None:
                    self._ecrit_revenu_cases(f, d, r, data["cases"], data.get("colonne"), num(data.get("ajout")) or 0)
                elif actuelle is not None and actuelle.formule and \
                        str(data.get("montant") or "").replace(" ", "") == ("=" + actuelle.formule).replace(" ", ""):
                    pass                                        # formule inchangée (ex. =SUM(C68:C71))
                else:
                    self._ecrit_montant(f, "B" + str(r), data.get("montant"))
                f.ecrit("C" + str(r), (data.get("description") or "").strip() or None)
            else:
                pb = d.pb
                self._ecrit_montant(f, pb["recu"] + str(r), data.get("recu"))
                self._ecrit_montant(f, pb["arec"] + str(r), data.get("a_recevoir"))
                if pb["heures"] and "heures" in data:
                    self._ecrit_montant(f, pb["heures"] + str(r), data.get("heures"))
                if pb["statut"]:
                    f.ecrit(pb["statut"] + str(r), data.get("statut") or None)
            return r
        return self._ecriture(op)

    def _assure_ligne_bloc(self, nom, bloc):
        """Première ligne libre après la dernière utilisée d'une section (revenus, pourboires) ;
        ajoute 5 lignes (formules et totaux ajustés comme dans Excel) si la section est pleine."""
        d = self.dispo(nom)
        lignes = self._bloc(d, bloc)
        mod = self.modele(nom)[bloc]
        utilisees = [x["ligne"] for x in mod if x["utilise"]]
        der = max(utilisees) if utilisees else lignes[0] - 1
        if der < lignes[-1]:
            return der + 1
        n = 5
        a_partir = lignes[-1]
        f = self.wb.feuille(nom)
        cols = ["A", "B", "C"] if bloc == "revenus" else sorted({x for x in d.pb.values() if x}, key=col_en_num)
        sauve = {}
        for col in cols:
            c = f.cell(col + str(a_partir))
            if c is not None and (c.formule or c.valeur not in (None, "")):
                sauve[col] = ("f", c.formule, c.valeur) if c.formule else ("v", c.valeur)
        self.wb.insere_lignes(nom, a_partir, n, lig_modele=a_partir - 1, cols_modele=cols)
        self.dispos.pop(nom, None)
        self.modeles.clear(); self.memo.clear()
        f = self.wb.feuille(nom)
        for col in cols:
            v = sauve.get(col)
            if v is None:
                f.ecrit(col + str(a_partir), None)
            elif v[0] == "f":
                f.ecrit(col + str(a_partir), valeur=num(v[2]), formule=v[1])
            else:
                f.ecrit(col + str(a_partir), serie_en_date(v[1]) if col == "A" and isinstance(v[1], float) else v[1])
            f.ecrit(col + str(a_partir + n), None)
        self.modeles.clear(); self.memo.clear()
        return a_partir + 1

    def supprime_ligne_bloc(self, feuille, bloc, ligne):
        def op():
            f, d = self._verifie_modifiable(feuille)
            mod = self.modele(feuille)[bloc]
            utilisees = [x["ligne"] for x in mod if x["utilise"]]
            if ligne not in utilisees:
                raise Erreur("Ligne vide")
            der = max(utilisees)
            cols = ["A", "B", "C"] if bloc == "revenus" else [x for x in d.pb.values() if x]
            for r in range(ligne, der + 1):
                for col in cols:
                    src = f.cell(col + str(r + 1)) if r < der else None
                    if src is None or (src.valeur in (None, "") and not src.formule):
                        f.ecrit(col + str(r), None)
                    elif src.formule:
                        f.ecrit(col + str(r), valeur=num(src.valeur), formule=src.formule)
                    else:
                        v = src.valeur
                        if col == "A" and isinstance(v, float):
                            v = serie_en_date(v)
                        f.ecrit(col + str(r), v)
            if bloc == "pourboires":
                self._recale_revenus_apres_suppression(f, d, ligne, der)
        return self._ecriture(op)

    # --- paramètres du mois ---
    def enregistre_parametres(self, feuille, data):
        def op():
            f, d = self._verifie_modifiable(feuille)
            cibles = {"revenu_marika": d.a_marika, "revenu_zach": d.a_zach, "solde_reel": d.a_reel,
                      "retrait": d.a_retrait, "solde_debut": d.a_debut}
            for k, adr in cibles.items():
                if k in data and adr:
                    self._ecrit_montant(f, adr, data[k])
            if ("revenu_marika" in data or "revenu_zach" in data) and d.a_pct:
                self._retablit_formule_pct(f, d)
            # retraits par compte : {"retraits": {"B159": "1250", ...}}
            adrs = {x["adr"] for x in d.retraits}
            for adr, v in (data.get("retraits") or {}).items():
                if adr in adrs:
                    self._ecrit_montant(f, adr, v)
        return self._ecriture(op)

    def _retablit_formule_pct(self, f, d):
        """% du partenaire saisi à la main (configuration de départ) → formule du Canvas (répartition des revenus)."""
        c = f.cell(d.a_pct)
        if c is not None and c.formule:
            return
        modele_c = self.wb.feuille("Canvas").cell(d.a_pct) if "Canvas" in self.wb.par_nom else None
        if modele_c is not None and modele_c.formule:
            f.ecrit(d.a_pct, formule=modele_c.formule)

    def configuration_depart(self, feuille, data):
        """Premier mois d'un fichier vierge : solde de départ, placements et % du partenaire."""
        def op():
            f, d = self._verifie_modifiable(feuille)
            if data.get("solde_debut") in (None, ""):
                raise Erreur("Indique le solde du compte au début du mois.")
            self._ecrit_montant(f, d.a_debut, data["solde_debut"])
            for r in d.invest:
                x = (data.get("placements") or {}).get(str(r)) or {}
                if x.get("investi") not in (None, ""):
                    self._ecrit_montant(f, "B" + str(r), x["investi"])
                if x.get("jv") not in (None, ""):
                    self._ecrit_montant(f, "C" + str(r), x["jv"])
            if data.get("pct_partenaire") not in (None, "") and d.a_pct:
                v, _ = self.montant_saisi(data["pct_partenaire"])
                if v is None or not 0 <= v <= 100:
                    raise Erreur("Pourcentage invalide")
                f.ecrit(d.a_pct, round(v / 100.0, 4))
            h, lignes = self._bloc_reglages()
            if "configuration de départ" in lignes:
                self.wb.feuille("Listes").ecrit(lignes["configuration de départ"][0], "Faite")
        return self._ecriture(op)

    def enregistre_investissement(self, feuille, data):
        def op():
            f, d = self._verifie_modifiable(feuille)
            r = int(data["ligne"])
            if r not in d.invest:
                raise Erreur("Ligne invalide")
            if "investi" in data:
                self._ecrit_montant(f, "B" + str(r), data["investi"])
            if "jv" in data:
                self._ecrit_montant(f, "C" + str(r), data["jv"])
        return self._ecriture(op)

    # --- montants manuels du tableau Remboursement (Split Marika) ---
    def _lignes_ajust(self, d):
        if not (d.lig_tableau and d.lig_rtotal):
            raise Erreur("Tableau Remboursement introuvable dans ce mois.")
        return list(range(d.lig_tableau + 1, d.lig_rtotal))

    def enregistre_ajust_marika(self, feuille, data):
        def op():
            f, d = self._verifie_modifiable(feuille)
            lignes = self._lignes_ajust(d)
            lib = (data.get("libelle") or "").strip()
            if not lib:
                raise Erreur("La description est obligatoire.")
            v, formule = self.montant_saisi(data.get("montant"))
            if v is None:
                raise Erreur("Montant invalide.")
            if data.get("sens") == "je_dois":
                v = -abs(v)
                formule = "=-(%s)" % formule[1:] if formule else None
            elif data.get("sens") == "marika_doit":
                if v < 0:
                    v = -v
                    formule = "=-(%s)" % formule[1:] if formule else None
            r = data.get("ligne")
            if r is None:
                libres = [x for x in lignes if not f.texte(d.c_rlib + str(x)).strip()
                          and (f.cell(d.c_rval + str(x)) is None
                               or (f.cell(d.c_rval + str(x)).valeur in (None, "") and not f.cell(d.c_rval + str(x)).formule))]
                if not libres:
                    raise Erreur("Plus de ligne libre dans le tableau Remboursement (%d lignes)." % len(lignes))
                r = libres[0]
            elif r not in lignes:
                raise Erreur("Ligne invalide")
            else:
                c = f.cell(d.c_rval + str(r))
                if c is not None and c.formule and "!" in c.formule:
                    raise Erreur("Cette ligne est le report automatique du mois précédent : elle ne se modifie pas ici.")
            f.ecrit(d.c_rlib + str(r), lib)
            anc = f.cell(d.c_rval + str(r))
            if data.get("ligne") is not None and anc is not None and anc.formule and "COUNTIF" in anc.formule.upper():
                expr = formule[1:] if formule else repr(v)
                f.ecrit(d.c_rval + str(r), valeur=0.0, formule="IF(%s,0,%s)" % (self._condition_reglement(d), expr))
                return r
            if formule:
                f.ecrit(d.c_rval + str(r), valeur=v, formule=formule)
            else:
                f.ecrit(d.c_rval + str(r), v)
            return r
        return self._ecriture(op)

    def supprime_ajust_marika(self, feuille, ligne):
        def op():
            f, d = self._verifie_modifiable(feuille)
            if ligne not in self._lignes_ajust(d):
                raise Erreur("Ligne invalide")
            c = f.cell(d.c_rval + str(ligne))
            if c is not None and c.formule and "!" in c.formule:
                raise Erreur("Cette ligne est le report automatique du mois précédent : elle ne se supprime pas ici.")
            f.ecrit(d.c_rlib + str(ligne), None)
            f.ecrit(d.c_rval + str(ligne), None)
        return self._ecriture(op)

    def force_controle(self, feuille, ligne, forcer, justification):
        def op():
            f, d = self._verifie_modifiable(feuille)
            if ligne not in d.controles:
                raise Erreur("Contrôle invalide")
            if forcer and not (justification or "").strip():
                raise Erreur("Une justification est obligatoire pour forcer un contrôle.")
            f.ecrit(d.c_ctl_forcer + str(ligne), "Oui" if forcer else None)
            f.ecrit(d.c_ctl_just + str(ligne), (justification or "").strip() or None)
        return self._ecriture(op)

    # --- listes ---
    def ajoute_mot_cle(self, mot, categorie):
        mot = (mot or "").strip()
        if not mot:
            raise Erreur("Mot-clé vide")
        if categorie not in self.categories:
            raise Erreur("Catégorie inconnue")
        def op():
            f = self.wb.feuille("Listes")
            r = self.l_mot_der + 1
            if r > 250:
                raise Erreur("La table des mots-clés est pleine (ligne 250, limite des formules du Canvas).")
            # Les 3 derniers mots-clés (les plus fréquents, ils l'emportent) restent en bas :
            # on les descend d'une ligne et le nouveau mot-clé prend la place libérée.
            premier = max(self.l_mot_debut, self.l_mot_der - 2) if self.l_mot_der >= self.l_mot_debut else r
            for c in (self.l_mot_col, self.l_motcat_col):
                ref = f.cell(c + str(self.l_mot_der))
                st = ref.style if ref is not None else None
                for x in range(self.l_mot_der, premier - 1, -1):
                    src = f.cell(c + str(x))
                    f.ecrit(c + str(x + 1), src.valeur if src is not None else None, style=st)
            f.ecrit(self.l_mot_col + str(premier), mot)
            f.ecrit(self.l_motcat_col + str(premier), categorie)
            return premier
        return self._ecriture(op)

    def _etend_validations(self, col, nouvelle_der):
        """Étend les listes déroulantes qui pointent vers Listes!$col$3:$col$N."""
        motif = re.compile(r"(Listes!\$%s\$(\d+):\$%s\$)(\d+)" % (col, col))
        # seulement Canvas, Listes et les mois modifiables : jamais l'historique ni le mois de référence
        permis = {"Canvas", "Listes"} | {m["feuille"] for m in self.mois if m["modifiable"]}
        for g in self.wb.feuilles:
            if g.nom not in permis:
                continue
            x = g.xml
            x2 = motif.sub(lambda m: m.group(1) + str(max(int(m.group(3)), nouvelle_der)), x)
            if x2 != x:
                g.xml = x2
                g.analyse()

    def ajoute_categorie(self, nom):
        nom = (nom or "").strip()
        if not nom or nom.lower() in [c.lower() for c in self.categories]:
            raise Erreur("Catégorie vide ou déjà existante")
        def op():
            f = self.wb.feuille("Listes")
            r = self.l_cat_der + 1
            f.ecrit(self.l_cat_col + str(r), nom)
            self._etend_validations(self.l_cat_col, r)
            # ajouter la catégorie au résumé du Canvas et des mois modifiables (avant « Autres »)
            for m in [{"feuille": "Canvas"}] + [x for x in self.mois if x["modifiable"]]:
                self._ajoute_cat_resume(m["feuille"], nom)
        return self._ecriture(op)

    def _ajoute_cat_resume(self, feuille, nom):
        d = Disposition(self.wb.feuille(feuille))
        if not d.cats:
            return
        noms = [normalise(self.wb.feuille(feuille).texte("A" + str(r))) for r in d.cats]
        if normalise(nom) in noms:
            return
        r_autres = d.cats[noms.index("autres")] if "autres" in noms else d.cats[-1]
        self.wb.insere_lignes(feuille, r_autres, 1, lig_modele=r_autres)
        f = self.wb.feuille(feuille)
        # la ligne modèle est maintenant r_autres+1 ; la nouvelle ligne est r_autres
        src = f.cell("B" + str(r_autres + 1))
        f.ecrit("A" + str(r_autres), nom)
        if src is not None and src.formule:
            from formules import copie_formule
            f.ecrit("B" + str(r_autres), formule=copie_formule(src.formule, -1))
        self.dispos.pop(feuille, None)

    def ajoute_type(self, nom):
        nom = (nom or "").strip()
        if not nom:
            raise Erreur("Type vide")
        def op():
            f = self.wb.feuille("Listes")
            r = self.l_type_der + 1
            f.ecrit(self.l_type_col + str(r), nom + " ")
            self._etend_validations(self.l_type_col, r)
        return self._ecriture(op)

    # --- nouveau mois ---
    def nouveau_mois(self, data):
        def op():
            nom = self.prochain_mois()
            if nom in self.wb.par_nom:
                raise Erreur("L'onglet %s existe déjà." % nom)
            prec = self.mois[-1]
            # --- validation : tous les montants de clôture doivent être inscrits ---
            def requis(cle, libelle, valeur=None):
                v = data.get(cle) if valeur is None else valeur
                if v is None or str(v).strip() == "":
                    raise Erreur("Montant manquant : %s." % libelle)
                val, _ = self.montant_saisi(v)
                if val is None:
                    raise Erreur("Montant invalide : %s." % libelle)
                return v
            a_ecrire = []          # (feuille, adresse, texte) écrits dans le mois qui se termine
            if prec["canvas"] and prec["modifiable"]:
                dp0 = self.dispo(prec["feuille"])
                fp0 = self.wb.feuille(prec["feuille"])
                saisis = data.get("retraits") or {}
                for x in dp0.retraits:
                    lib = "Épargne " + (x["compte"] or x["libelle"])
                    a_ecrire.append((fp0, x["adr"], requis(None, lib, saisis.get(x["adr"], ""))))
                if dp0.a_reel:
                    a_ecrire.append((fp0, dp0.a_reel, requis("solde_reel", "Solde réel au compte")))
            reg = self.reglages()
            requis("revenu_marika", "Revenus " + reg["partenaire"])
            requis("revenu_zach", "Revenus " + reg["prenom"])
            for fp0, adr, v in a_ecrire:
                self._ecrit_montant(fp0, adr, v)
            f = self.wb.copie_feuille("Canvas", nom)
            # textes « XX 2026 » du gabarit
            for c in list(f.cellules.values()):
                if isinstance(c.valeur, str) and not c.formule and re.search(r"\bXX\b", c.valeur):
                    f.ecrit(c.adr, re.sub(r"\bXX(\s+\d{4})?", nom, c.valeur))
            d = Disposition(f)
            self.dispos[nom] = d
            if d.a_mois:
                f.ecrit(d.a_mois, nom)
            p = prefixe_feuille(prec["feuille"])
            if prec["canvas"]:
                dp = self.dispo(prec["feuille"])
                if d.a_debut:
                    fin_prec = p + dp.a_fin
                    if dp.a_reel:
                        f.ecrit(d.a_debut, formule="IF(ISNUMBER(%s%s),%s%s,%s)" % (p, dp.a_reel, p, dp.a_reel, fin_prec))
                    else:
                        f.ecrit(d.a_debut, formule=fin_prec)
                # report du solde Marika non réglé
                if d.lig_tableau and dp.lig_rtotal:
                    r = d.lig_tableau + 1
                    plage = "%s%d:%s%d" % (d.c_marq, d.lignes[0], d.c_marq, d.lignes[-1])
                    f.ecrit(d.c_rlib + str(r), "Portion " + prec["nom"])
                    f.ecrit(d.c_rval + str(r), formule='IF(COUNTIF(%s,"*Solde réglé*")>0,0,%s%s%d)'
                            % (plage, p, dp.c_rval, dp.lig_rtotal))
                # investissements : montant investi = investi du mois précédent
                #                     + épargne retirée vers ce compte le mois précédent
                alloc = data.get("allocations") or {}
                prec_inv = {normalise(self.wb.feuille(prec["feuille"]).texte("A" + str(r))): r for r in dp.invest}
                prec_ret = {normalise(x["compte"]): x["adr"] for x in dp.retraits if x["compte"]}
                for r in d.invest:
                    t = f.texte("A" + str(r))
                    rp = prec_inv.get(normalise(t))
                    ret = prec_ret.get(normalise(t))
                    a, _ = self.montant_saisi(alloc.get(t.strip()) or 0)
                    termes = []
                    if rp:
                        termes.append("%sB%d" % (p, rp))
                    if ret:
                        termes.append(p + ret)
                    if a:
                        termes.append(repr(a).rstrip("0").rstrip("."))
                    if termes:
                        f.ecrit("B" + str(r), formule="+".join(termes))
            else:
                h = self.historique(prec["feuille"])
                if d.a_debut and h["solde_fin"] is not None:
                    f.ecrit(d.a_debut, h["solde_fin"])
            if data.get("revenu_marika") not in (None, "") and d.a_marika:
                self._ecrit_montant(f, d.a_marika, data["revenu_marika"])
            if data.get("revenu_zach") not in (None, "") and d.a_zach:
                self._ecrit_montant(f, d.a_zach, data["revenu_zach"])
            return nom
        return self._ecriture(op)
