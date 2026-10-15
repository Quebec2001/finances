"""Lectures de relevés et captures (OCR), rapprochement bancaire : réutilise tel quel le code de la v1
(releve.py, rapprochement.py, méthodes de modele.Application) sur les données de la v2.
L'appelant fournit l'état et les mois au format de l'API (produits par api2.js). Rien n'est écrit."""
import json, os, re, sys
ICI = os.path.dirname(os.path.abspath(__file__))
for p in (os.path.join(ICI, "..", "app"), ICI):
    if p not in sys.path:
        sys.path.insert(0, p)
from modele import Application, Erreur, normalise   # noqa: E402


class AppLecture:
    """Juste ce qu'il faut de modele.Application pour les lectures (mois, catégories automatiques, réglages)."""
    def __init__(self, etat, mois):
        self.etat = etat
        self.mois = etat["mois"]
        self._mois = mois                      # {nom de feuille: modèle du mois (format de l'API)}
        self.motscles = [(k["mot"], k["categorie"]) for k in etat.get("motscles", [])]
        self.memo = {}

    def info_mois(self, feuille):
        for m in self.mois:
            if m["feuille"].strip() == (feuille or "").strip():
                return m
        raise Erreur("Mois introuvable : %s" % feuille)

    def modele(self, feuille):
        f = self.info_mois(feuille)["feuille"]
        if f not in self._mois:
            raise Erreur("Mois non fourni : %s" % f)
        return self._mois[f]

    def reglages(self):
        return self.etat.get("reglages", {})

    def descriptions_autres_revenus(self, n=3):
        return (self.etat.get("descriptions_autres_revenus") or [])[:n]

    def categorie_auto(self, desc):
        d, r = normalise(desc or ""), "Autres"
        for mot, cat in self.motscles:
            if mot and normalise(mot) in d:
                r = cat
        return r

    analyse_releve = Application.analyse_releve
    analyse_releve_revenus = Application.analyse_releve_revenus
    rappro_lire = Application.rappro_lire
    rappro_images = Application.rappro_images
    rappro_analyse = Application.rappro_analyse


def appel(route, corps, etat, mois):
    a = AppLecture(etat, mois)
    f = corps.get("feuille")
    if route == "/api/releve/analyse":
        return a.analyse_releve(f, corps.get("images") or [])
    if route == "/api/releve/revenus":
        return a.analyse_releve_revenus(f, corps.get("images") or [])
    if route == "/api/rappro/lire":
        return a.rappro_lire(f, corps.get("fichier") or {})
    if route == "/api/rappro/images":
        return a.rappro_images(f, corps.get("images") or [], corps.get("noms") or [])
    if route == "/api/rappro/analyse":
        return a.rappro_analyse(f, corps.get("fichiers") or [], corps.get("options") or {})
    raise Erreur("Route inconnue : " + route)


def appel_json(texte):
    """Entrée/sortie en JSON (Pyodide ou ligne de commande)."""
    o = json.loads(texte)
    try:
        return json.dumps({"ok": True, "resultat": appel(o["route"], o["corps"], o["etat"], o["mois"])}, ensure_ascii=False, default=str)
    except Erreur as e:
        return json.dumps({"erreur": str(e)}, ensure_ascii=False)


if __name__ == "__main__":
    sys.stdout.write(appel_json(sys.stdin.read()))
