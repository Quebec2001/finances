"""Routes de l'API, partagées par le serveur local (serveur.py) et la version web (pont.py)."""


def get(app, chemin, q):
    """Requête de lecture. `q` : dict {clé: [valeurs]} (comme urllib.parse.parse_qs)."""
    app.charge()
    if chemin == "/api/etat":
        return app.etat()
    if chemin == "/api/mois":
        f = q.get("f", [""])[0]
        info = app.info_mois(f)
        return app.modele(f) if info["canvas"] else app.historique(f)
    if chemin == "/api/tendances":
        return app.tendances()
    return None


def post(app, chemin, data):
    """Requête d'écriture. Renvoie le résultat, ou lève KeyError si la route est inconnue."""
    f = data.get("feuille")
    if chemin == "/api/transaction":
        return app.enregistre_transaction(f, data)
    if chemin == "/api/transaction/supprimer":
        return app.supprime_transaction(f, int(data["ligne"]))
    if chemin == "/api/reglement":
        return app.marque_reglement(f, data.get("ligne"))
    if chemin == "/api/revenu":
        return app.enregistre_ligne_bloc(f, "revenus", data)
    if chemin == "/api/revenu/supprimer":
        return app.supprime_ligne_bloc(f, "revenus", int(data["ligne"]))
    if chemin == "/api/pourboire":
        return app.enregistre_ligne_bloc(f, "pourboires", data)
    if chemin == "/api/pourboire/supprimer":
        return app.supprime_ligne_bloc(f, "pourboires", int(data["ligne"]))
    if chemin == "/api/parametres":
        return app.enregistre_parametres(f, data)
    if chemin == "/api/investissement":
        return app.enregistre_investissement(f, data)
    if chemin == "/api/marika":
        return app.enregistre_ajust_marika(f, data)
    if chemin == "/api/marika/supprimer":
        return app.supprime_ajust_marika(f, int(data["ligne"]))
    if chemin == "/api/controle":
        return app.force_controle(f, int(data["ligne"]), bool(data.get("forcer")), data.get("justification"))
    if chemin == "/api/nouveau-mois":
        return app.nouveau_mois(data)
    if chemin == "/api/listes/mot-cle":
        return app.ajoute_mot_cle(data.get("mot"), data.get("categorie"))
    if chemin == "/api/listes/categorie":
        return app.ajoute_categorie(data.get("nom"))
    if chemin == "/api/listes/type":
        return app.ajoute_type(data.get("nom"))
    raise KeyError(chemin)
