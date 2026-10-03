"""Pont entre l'interface web (navigateur, Pyodide) et le moteur Python.

Le navigateur télécharge le classeur depuis OneDrive, l'écrit dans le système de
fichiers virtuel, puis appelle `appel()` exactement comme le serveur local.
Les sauvegardes sont faites dans OneDrive par l'interface (pas sur le disque virtuel)."""
import json
import traceback
import urllib.parse

import modele
import routes
from modele import Application, Erreur

modele.copie_sauvegarde = lambda *a, **k: None   # sauvegardes gérées côté OneDrive

APP = None


def ouvre(chemin, config_json):
    """(Re)charge le classeur `chemin` avec la configuration donnée."""
    global APP
    cfg = json.loads(config_json)
    cfg["fichier"] = chemin
    APP = Application(cfg)
    APP.charge(force=True)
    return "ok"


def recharge():
    APP.charge(force=True)
    return "ok"


def appel(methode, chemin_complet, corps_json):
    """Même contrat que l'API HTTP du serveur local : renvoie une chaîne JSON."""
    try:
        url = urllib.parse.urlparse(chemin_complet)
        if methode == "GET":
            res = routes.get(APP, url.path, urllib.parse.parse_qs(url.query))
            if res is None:
                return json.dumps({"erreur": "Route inconnue"})
            return json.dumps(res, ensure_ascii=False, default=str)
        data = json.loads(corps_json or "{}")
        try:
            res = routes.post(APP, url.path, data)
        except KeyError as e:
            if e.args and e.args[0] == url.path:
                return json.dumps({"erreur": "Route inconnue"})
            raise
        return json.dumps({"ok": True, "resultat": res}, ensure_ascii=False, default=str)
    except Erreur as e:
        return json.dumps({"erreur": str(e)}, ensure_ascii=False)
    except Exception as e:  # pragma: no cover
        traceback.print_exc()
        return json.dumps({"erreur": "Erreur interne : %s" % e}, ensure_ascii=False)
