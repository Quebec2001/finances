"""Import d'un classeur Excel (format de l'app v1) dans une base v2 (SQLite).

Lecture seule de l'Excel : on passe par le modèle de la v1 (modele.Application), qui sait déjà lire chaque bloc.
Usage : importe(chemin_xlsx, chemin_db, config) -> rapport (dict)."""
import json, os, re, sys
ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ICI, "..", "app"))
from modele import Application, sem_type, parse_mois, RX_A_REMBOURSER, RX_FIXE   # noqa: E402
from magasin import Magasin   # noqa: E402

VERSION_BASE = 1


def c(x):
    """Dollars -> cents entiers (None si absent)."""
    if x is None or x == "" or isinstance(x, str):
        return None
    return int(round(float(x) * 100))


def fin_de_mois(cle):
    a, m = map(int, cle.split("-"))
    import calendar
    return "%04d-%02d-%02d" % (a, m, calendar.monthrange(a, m)[1])


def importe(xlsx, db, config=None):
    cfg = {"premier_mois_modifiable": "Octobre 2026", "dossier_sauvegardes": os.path.join(os.path.dirname(db or "") or ".", ".sauv-import")}
    cfg.update(config or {})
    cfg["fichier"] = xlsx
    a = Application(cfg)
    a.charge()
    e = a.etat()
    con = Magasin()
    q = con.execute
    rap = {"mois": [], "avertissements": []}

    # ---------------------------------------------------------------- réglages
    reg = dict(e["reglages"])
    reg["theme"] = "pro" if reg.get("pourboires") else "rose"
    reg["verrou_avant"] = "-".join(map(str, (lambda p: (p[0], "%02d" % p[1]))(parse_mois(cfg["premier_mois_modifiable"]))))
    reg["version_base"] = VERSION_BASE
    for k, v in reg.items():
        q("INSERT INTO reglages VALUES (?,?)", (k, json.dumps(v, ensure_ascii=False)))

    # ---------------------------------------------------------------- catégories et mots-clés
    cats = {}
    dir_locals = {}
    def cat_id(nom, active=1):
        if not nom:
            return None
        k = nom.strip().lower()
        if k not in cats:
            if "actives" in dir_locals and k not in dir_locals["actives"]:
                active = 0
            cur = q("INSERT INTO categories(nom, ordre, active) VALUES (?,?,?)", (nom.strip(), len(cats) + 1, active))
            cats[k] = cur.lastrowid
        return cats[k]
    for n in e["categories"]:
        cat_id(n)
    dir_locals["actives"] = set(cats)   # catégories de l'onglet Listes ; celles trouvées seulement dans l'historique : inactives
    for i, k in enumerate(e["motscles"]):
        q("INSERT INTO mots_cles(mot, categorie_id, ordre) VALUES (?,?,?)", (k["mot"], cat_id(k["categorie"]), i + 1))

    # ---------------------------------------------------------------- budgets
    defaut, par_mois, _ = a._lit_budgets()
    for n, v in defaut.items():
        if v is not None:
            q("INSERT INTO budgets VALUES (?,?,?)", (cat_id(n), "", c(v)))
    for nm, d in par_mois.items():
        p = parse_mois(nm)
        if not p:
            continue
        for n, v in d.items():
            if v is not None:
                q("INSERT OR REPLACE INTO budgets VALUES (?,?,?)", (cat_id(n), "%04d-%02d" % p, c(v)))

    # ---------------------------------------------------------------- comptes de placement
    comptes = {}
    alias = cfg.get("alias_investissements", {"Fonds communs de placement": "CÉLI"})
    def compte_id(nom, fixe=None):
        if not nom:
            return None
        nom = alias.get(nom, nom)
        if nom not in comptes:
            cur = q("INSERT INTO comptes_placement(nom, valeur_fixe, ordre) VALUES (?,?,?)",
                    (nom, 1 if (fixe if fixe is not None else RX_FIXE.search(nom)) else 0, len(comptes) + 1))
            comptes[nom] = cur.lastrowid
        return comptes[nom]

    # ---------------------------------------------------------------- mois par mois
    ordre = 0
    a_regler = []          # (table, id) des éléments partagés pas encore couverts par un règlement
    premier = True
    reel_prec = None
    for info in e["mois"]:
        f, cle = info["feuille"], info["cle"]
        fin = fin_de_mois(cle)
        if info["canvas"]:
            m = a.modele(f)
        else:
            m = a.historique(f)
        r = {"mois": cle, "feuille": f.strip(), "groupe": info["groupe"]}
        if m.get("solde_debut") is None and premier:
            rap["avertissements"].append(f"{f.strip()} : solde de début vide (fichier pas encore configuré)")
        sd = m.get("solde_debut")
        std = (sd or 0) + (m.get("revenu_net") or 0) - (m.get("retrait") or 0)
        aj = (m.get("solde_fin") or 0) - std if m.get("solde_fin") is not None else 0
        # point de départ explicite : le tout premier mois, ou quand la v1 a repris le solde réel de la banque
        # inscrit à la fin du mois précédent (règle de la v1 : début = solde réel du relevé s'il existe)
        ancre = premier or (reel_prec is not None and sd is not None)
        q("INSERT INTO soldes_mois VALUES (?,?,?,?,?)", (cle, c(sd) if ancre else None,
          None if not ancre else ("Solde de départ" if premier else "Solde réel de la banque inscrit à la fin du mois précédent"), c(aj) or 0,
          ("Écart repris de l'Excel : solde de fin du fichier différent du calcul standard" if abs(aj) >= 0.005 else None)))
        reel_prec = m.get("solde_reel")
        if premier and sd is not None:
            q("INSERT OR REPLACE INTO reglages VALUES ('solde_depart', ?)", (json.dumps({"mois": cle, "montant_c": c(sd)}),))
        premier = False

        if not info["canvas"]:
            # ---- historique : lignes (ma part exacte), un revenu global, l'épargne globale, les placements
            for i, l in enumerate(m["lignes"]):
                ordre += 1
                q("""INSERT INTO depenses(mois, date, description, montant_c, type, categorie_id, historique, part_hist, ordre, origine)
                     VALUES (?,?,?,?,?,?,?,?,?,?)""", (cle, l.get("date"), l.get("description") or "", None, "P", cat_id(l.get("categorie")),
                     1, l.get("montant") or 0, ordre, f"{f.strip()} ligne {i + 1}"))
            if m.get("total_revenus"):
                q("INSERT INTO revenus(mois, date, description, montant_c, historique, origine) VALUES (?,?,?,?,?,?)",
                  (cle, fin, "Revenus du mois (historique)", c(m["total_revenus"]), 1, f.strip()))
            invs = m.get("investissements") or []
            cpts = {compte_id(x["type"]) for x in invs}
            if m.get("retrait"):     # épargne globale : attribuée au placement s'il n'y en a qu'un ce mois-là
                q("INSERT INTO versements_epargne(mois, date, compte_id, montant_c, origine) VALUES (?,?,?,?,?)",
                  (cle, fin, next(iter(cpts)) if len(cpts) == 1 else None, c(m["retrait"]), f.strip()))
            tot = {}
            for x in invs:           # plusieurs lignes vers le même compte (alias) : additionnées
                k = compte_id(x["type"]); a_, b_ = tot.get(k, (0, 0))
                tot[k] = (a_ + (c(x.get("jv")) or 0), b_ + (c(x.get("investi")) or 0))
            for k, (jv, inv) in tot.items():
                q("INSERT INTO placements_mois VALUES (?,?,?,?,?)", (k, cle, jv, inv, 1))
            r.update(lignes=len(m["lignes"]))
            rap["mois"].append(r)
            continue

        # ---- mois au format de l'app
        q("INSERT INTO repartition VALUES (?,?,?,?,?)", (cle, c(m.get("revenu_zach")), c(m.get("revenu_marika")),
          m.get("pct_marika") or 0, 1 if m.get("pct_manuel") else 0))
        n_dep = 0
        for t in m["transactions"]:
            if not t["utilise"]:
                continue
            n_dep += 1
            ordre += 1
            desc = t.get("description") or ""
            paye_par = "moi"
            typ = sem_type(t.get("type"))
            if t.get("a_rembourser"):
                paye_par = "partenaire"
                desc = RX_A_REMBOURSER.sub("", desc).rstrip()
                typ = "P"
            if typ is None:
                rap["avertissements"].append(f"{f.strip()} ligne {t['ligne']} : type « {t.get('type')} » non reconnu ({desc})")
                typ = "P"
            calc = t.get("deb_formule")
            cur = q("""INSERT INTO depenses(mois, date, description, montant_c, calcul, type, categorie_id, paye_par, ordre, origine)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""", (cle, t.get("date"), desc, c(t.get("deboursement")),
                       calc.lstrip("=") if isinstance(calc, str) else None, typ,
                       cat_id(t.get("categorie")) if t.get("cat_manuelle") else None, paye_par, ordre, f"{f.strip()} ligne {t['ligne']}"))
            if typ != "P" or paye_par == "partenaire":
                a_regler.append(("depenses", cur.lastrowid))
            if t.get("marqueur"):
                # « Solde réglé » sur cette ligne : couvre tout ce qui est partagé jusqu'ici (mois précédents compris)
                rg = q("INSERT INTO reglements(date, ordre, confirme, source) VALUES (?,?,1,'import')", (t.get("date") or fin, ordre)).lastrowid
                for tab, i in a_regler:
                    q(f"UPDATE {tab} SET reglement_id=? WHERE id=?", (rg, i))
                a_regler = []
                r["reglement"] = rg
        # autres montants saisis à la main (les reports automatiques disparaissent : le solde est continu)
        for x in m.get("reports") or []:
            if x.get("auto"):
                continue
            ordre += 1
            mt = x.get("montant_origine") if x.get("regle") and x.get("montant_origine") is not None else x.get("montant")
            cur = q("INSERT INTO autres_montants(date, libelle, montant_c, ordre, origine) VALUES (?,?,?,?,?)",
                    (fin, x.get("libelle") or "", c(mt) or 0, ordre, f"{f.strip()} ligne {x['ligne']}"))
            if x.get("regle") and r.get("reglement"):
                q("UPDATE autres_montants SET reglement_id=? WHERE id=?", (r["reglement"], cur.lastrowid))
            elif not x.get("regle"):
                a_regler.append(("autres_montants", cur.lastrowid))
        # pourboires et revenus
        pb = {}
        for p in m.get("pourboires") or []:
            if not p.get("utilise"):
                continue
            st = (p.get("statut") or "").strip().lower()
            pb[p["ligne"]] = q("INSERT INTO pourboires(mois, date, especes_c, carte_c, heures, recu, origine) VALUES (?,?,?,?,?,?,?)",
                               (cle, p.get("date"), c(p.get("recu")), c(p.get("a_recevoir")), p.get("heures"),
                                1 if st == "oui" else 0 if st == "non" else None, f"{f.strip()} ligne {p['ligne']}")).lastrowid
        n_rev = 0
        for x in m.get("revenus") or []:
            if not x.get("utilise"):
                continue
            n_rev += 1
            cs = x.get("cases")
            rid = q("""INSERT INTO revenus(mois, date, description, montant_c, calcul, cases_colonne, cases_ajout_c, origine)
                       VALUES (?,?,?,?,?,?,?,?)""", (cle, x.get("date"), x.get("description") or "", c(x.get("montant")) or 0,
                       (x.get("montant_txt") or "").lstrip("=") or None,
                       ({"recu": "especes", "arec": "carte"}.get(cs["colonne"]) if cs else None),
                       c(cs.get("ajout")) or 0 if cs else 0, f"{f.strip()} ligne {x['ligne']}")).lastrowid
            for lg in (cs or {}).get("lignes", []):
                if lg in pb:
                    q("INSERT INTO revenus_pourboires VALUES (?,?)", (rid, pb[lg]))
                else:
                    rap["avertissements"].append(f"{f.strip()} : revenu ligne {x['ligne']} additionne la ligne de pourboires {lg}, introuvable")
        # placements, épargne, solde de banque
        for x in m.get("investissements") or []:
            q("INSERT OR REPLACE INTO placements_mois VALUES (?,?,?,?,?)", (compte_id(x["type"], x.get("valeur_fixe")), cle,
              c(x.get("jv")), c(x.get("investi")), 0 if x.get("investi_auto") else 1))
        for x in m.get("retraits") or []:
            if x.get("montant"):
                q("INSERT INTO versements_epargne(mois, date, compte_id, montant_c, origine) VALUES (?,?,?,?,?)",
                  (cle, fin, compte_id(x.get("compte")), c(x["montant"]), f"{f.strip()} {x.get('adr')}"))
        if m.get("solde_reel") is not None:
            q("INSERT INTO releves_banque(date, solde_c, origine) VALUES (?,?,?)", (fin, c(m["solde_reel"]), f.strip()))
        r.update(lignes=n_dep, revenus=n_rev, pourboires=len(pb))
        rap["mois"].append(r)

    if db:
        con.enregistre(db)
    rap["donnees"] = con.json()
    return rap


if __name__ == "__main__":
    r = importe(sys.argv[1], sys.argv[2]); r.pop("donnees"); print(json.dumps(r, ensure_ascii=False, indent=1))
