"""Lecture des captures d'écran des apps bancaires (Desjardins compte / carte de crédit, Amex).

L'interface fait la reconnaissance de texte (Tesseract) dans le navigateur et envoie, pour
chaque image, la liste des mots avec leur position. Ce module reconstitue les transactions
(date, description, montant, sens) à partir de la disposition des mots."""
import datetime
import re
import unicodedata

MOIS = {"jan": 1, "janv": 1, "fev": 2, "fevr": 2, "feb": 2, "mar": 3, "mars": 3, "avr": 4, "apr": 4, "mai": 5,
        "may": 5, "jui": 6, "juin": 6, "jun": 6, "juil": 7, "jul": 7, "aou": 8, "aout": 8, "aug": 8,
        "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
MOIS_LONGS = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6, "juillet": 7, "aout": 8,
              "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}
RX_MONTANT = re.compile(r"([+\-−–]?)\s?(\d{1,3}(?:[   ]\d{3})+|\d+)(?:[,.](\d{2}))?\s?[$S§]")
RX_DATE = re.compile(r"^(\d{1,2})\s*([A-Za-zéûÉÛ]{3,5})(\.?)\s*(\d{4})?$")
RX_ENTETE = re.compile(r"(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)\s+(\d{4})")
VILLES = {"PQ", "BRAMPTON", "MISSISSAUGA", "OTTAWA", "MONTREAL", "WESTMOUNT", "LAVAL", "LONGUEUIL", "QC", "QUEBEC", "CA", "CAN", "OUTREMONT", "VERDUN",
          "BROSSARD", "TORONTO", "ON", "SAINT-LAURENT", "MTL"}


def sans_accents(t):
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


ONGLETS = re.compile(r"\b(apercu|adhesion|offres amex|compte|accueil|virements|payer|plus|home|accounts|move money|more)\b", re.I)
RX_DATE_EN = re.compile(r"^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$")      # RBC : « Sep 30, 2026 »
ENTETES = ("amex", "rbc_credit", "rbc_compte")    # dates en en-tête au-dessus des transactions
RX_MONTANT_NU = re.compile(r"([+\-−–]?)\s?(\d{1,3}(?:,\d{3})+|\d+)\.(\d{2})")   # RBC : « 13.30 », « -900.12 »


def _mots(img):
    W = img.get("largeur") or 1000
    H = img.get("hauteur") or 2000
    mots = []
    for m in img.get("mots") or []:
        t = (m.get("t") or "").strip()
        if not t:
            continue
        x0, y0, x1, y1 = m["x0"], m["y0"], m["x1"], m["y1"]
        c = m.get("c", 100)
        if y1 < H * 0.075 or y0 > H * 0.975:          # barre d'état / indicateur du bas
            continue
        if c < 30 and not re.search(r"\d", t):
            continue
        if t in (">", "›", "<", "»", ")", "(", "|") or (len(t) <= 2 and x0 > W * 0.86 and not re.search(r"\d", t)):
            continue
        mots.append({"t": t, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "yc": (y0 + y1) / 2, "h": max(1, y1 - y0)})
    return mots, W, H


def _grouper(mots):
    mots = sorted(mots, key=lambda m: (m["yc"], m["x0"]))
    lignes = []
    for m in mots:
        for L in lignes:
            if abs(L["yc"] - m["yc"]) < max(L["h"], m["h"]) * 0.6:
                L["mots"].append(m)
                L["yc"] = (L["yc"] * (len(L["mots"]) - 1) + m["yc"]) / len(L["mots"])
                L["h"] = max(L["h"], m["h"])
                break
        else:
            lignes.append({"mots": [m], "yc": m["yc"], "h": m["h"]})
    for L in lignes:
        L["mots"].sort(key=lambda m: m["x0"])
        L["texte"] = " ".join(m["t"] for m in L["mots"])
    lignes.sort(key=lambda L: L["yc"])
    return lignes


def _lignes(img, source_prec=None):
    """Lignes de texte (de haut en bas) après nettoyage propre à l'app bancaire reconnue."""
    mots, W, H = _mots(img)
    source = _source(_grouper(list(mots)), H, source_prec)
    if source.startswith("desjardins"):
        # Desjardins : icônes rondes à gauche des descriptions
        mots = [m for m in mots if m["x1"] > W * 0.16]
    lignes = [L for L in _grouper(mots) if not (L["yc"] > H * 0.88 and ONGLETS.search(sans_accents(L["texte"])))]
    return lignes, W, H, source


def _montant(L, W, source=""):
    """Montant tout à droite de la ligne → (signe, valeur, texte à gauche) ou None."""
    ms = L["mots"]
    if source.startswith("rbc"):
        # RBC : montant sans « $ », point décimal (« 1,234.56 », « -900.12 »)
        der = ms[-1]
        k = 2 if len(ms) >= 2 and ms[-2]["t"] in ("-", "−", "–") and ms[-2]["x0"] > W * 0.5 else 1
        txt = "".join(m["t"] for m in ms[-k:])
        m = RX_MONTANT_NU.fullmatch(txt)
        if m and ms[-k]["x0"] > W * 0.5:
            signe = "-" if m.group(1) in ("-", "−", "–") else ("+" if m.group(1) == "+" else "")
            return signe, float(m.group(2).replace(",", "") + "." + m.group(3)), " ".join(x["t"] for x in ms[:-k]).strip()
    for k in range(min(4, len(ms)), 0, -1):          # le plus long d'abord (« +1 402,63 $ »)
        suffixe = ms[-k:]
        if suffixe[0]["x0"] < W * 0.5:
            continue
        mots_s = [m["t"] for m in suffixe]
        if len(mots_s) >= 2 and mots_s[-1] in ("5", "9", "S", "§") and re.fullmatch(r"\(?[+\-−–]?\d[\d ]*[,.]\d{2}", mots_s[-2]):
            mots_s[-1] = "$"                             # « $ » isolé lu comme « 5 » ou « 9 »
        txt = " ".join(mots_s).replace("—", "-").strip()
        txt = re.sub(r"(\d)[:;](\d{2})[:;.]?(?=\s?[$S§]$)", r"\1,\2", txt)     # « 29:72: $ »
        if "%" in txt:
            return None
        paren = txt.startswith("(") and txt.rstrip(".").endswith(")")
        if paren:
            txt = txt.strip("().").strip()
        m = RX_MONTANT.fullmatch(txt)
        if not m:
            continue
        entier = re.sub(r"\D", "", m.group(2))
        if m.group(3) is not None:
            valeur = float(entier + "." + m.group(3))
        else:
            if len(entier) < 3:
                return None
            valeur = float(entier[:-2] + "." + entier[-2:])   # virgule mal lue : 2 derniers chiffres = cents
        signe = "+" if paren else {"+": "+", "-": "-", "−": "-", "–": "-"}.get(m.group(1), "")
        gauche = [x["t"] for x in ms[:-k]]
        if not signe and gauche and gauche[-1] in ("-", "−", "–") and ms[-k - 1]["x0"] > W * 0.5:
            signe = "-"
            gauche = gauche[:-1]
        return signe, round(valeur, 2), " ".join(gauche).strip()
    # « +14,27 » sans « $ » (signe et cents bien lus)
    der = ms[-1]
    m = re.fullmatch(r"([+\-−–])(\d{1,3}(?:[ ]\d{3})*|\d+)[,.](\d{2})", der["t"])
    if m and der["x0"] > W * 0.5:
        return ("+" if m.group(1) == "+" else "-"), float(re.sub(r"\D", "", m.group(2)) + "." + m.group(3)), \
            " ".join(x["t"] for x in ms[:-1]).strip()
    # dernier recours : « $ » collé et lu comme « 5 » ou « S » (ex. « +0135 » pour « +0,13 $ »)
    der = ms[-1]
    if der["x0"] > W * 0.5 and "$" not in L["texte"]:
        m = re.fullmatch(r"([+\-−–])(\d[\d ,.]*\d)[5S]", der["t"])
        if m:
            chiffres = m.group(2)
            if re.search(r"[,.]\d{2}$", chiffres):
                valeur = float(re.sub(r"[^\d]", "", chiffres[:-3]) + "." + chiffres[-2:])
            else:
                c = re.sub(r"\D", "", chiffres)
                if len(c) < 3:
                    return None
                valeur = float(c[:-2] + "." + c[-2:])
            signe = "+" if m.group(1) == "+" else "-"
            return signe, round(valeur, 2), " ".join(x["t"] for x in ms[:-1]).strip()
    return None


RX_DATE_BRUIT = re.compile(r"^([0-9|\]lIZzOo!]{1,2})\s*[.:,;]?\s*([A-Za-z5$][A-Za-z]{2,4})(\.?)\s*(\d{4})?$")


def _date(texte, annee_def):
    t = sans_accents(texte.strip()).rstrip(",")
    me = RX_DATE_EN.match(t)
    if me:
        mois = MOIS.get(me.group(1).lower()) or MOIS.get(me.group(1).lower()[:3])
        if mois and 1 <= int(me.group(2)) <= 31:
            return {"jour": int(me.group(2)), "mois": mois, "annee": int(me.group(3)), "style": "entete"}
    m = RX_DATE.match(t)
    if m:
        jour, mois_txt, point, an = int(m.group(1)), m.group(2).lower(), m.group(3), m.group(4)
        brut_mois = m.group(2)
    else:
        # chiffres mal lus par la reconnaissance : « 2| SEP », « Z| SEP », « 24.SEP », « 25.5eP »
        m = RX_DATE_BRUIT.match(t)
        if not m or not re.search(r"[\d|\]!]", m.group(1) + m.group(2)):
            return None
        j = m.group(1).translate(str.maketrans("|]lI!ZzOo", "111112200"))
        mois_txt = re.sub(r"^[5$]", "s", m.group(2)).lower()
        if not j.isdigit() or mois_txt not in MOIS and mois_txt[:3] not in MOIS:
            return None
        jour, point, an, brut_mois = int(j), m.group(3), m.group(4), m.group(2).upper()
    mois = MOIS.get(mois_txt) or MOIS.get(mois_txt[:3])
    if not mois or not 1 <= jour <= 31:
        return None
    style = "entete" if (point or brut_mois[0].islower() and brut_mois != brut_mois.upper()) else "dessous"
    return {"jour": jour, "mois": mois, "annee": int(an) if an else annee_def, "style": style}


def _source(lignes, H, prec=None):
    """App bancaire d'après le titre en haut de l'écran ; la sorte de compte Desjardins se reporte
    d'une capture à la suivante (l'en-tête « Facturées en … » n'est visible que sur la première)."""
    haut = sans_accents(" ".join(L["texte"] for L in lignes if L["yc"] < H * 0.12)).lower()
    tout = sans_accents(" ".join(L["texte"] for L in lignes)).lower()
    haut_rbc = sans_accents(" ".join(L["texte"] for L in lignes if L["yc"] < H * 0.2)).lower()
    if not re.search(r"american|cobalt|amex", haut_rbc) and (
            re.search(r"\bposted\b|move money|\bas of \w+ \d|comptabilisee", tout) or
            re.search(r"\b(mastercard|visa)\b", haut_rbc) and "transactions" not in haut):
        if re.search(r"\b(mastercard|visa|credit card|carte de credit)\b", haut_rbc):
            return "rbc_credit"
        if re.search(r"chequing|savings|day to day|cheques|epargne|compte", haut_rbc):
            return "rbc_compte"
        return prec if prec in ("rbc_credit", "rbc_compte") else "rbc_credit"
    if re.search(r"american|cobalt|\bcard\b|amex", haut):
        return "amex"
    if "transactions" not in haut and re.search(r"american express|cobalt", tout) and "facturees" not in tout:
        return "amex"
    if "facturees en" in tout or "non facturees" in tout:
        return "desjardins_credit"
    if RX_ENTETE.search(tout) or re.search(r"(^|\s)[-−–]\s?[-−–]?\s?\d[\d ]*[,.]\d{2}", tout):
        return "desjardins_compte"              # mois en en-tête ou montants négatifs : compte courant
    if prec in ("desjardins_credit", "desjardins_compte"):
        return prec
    return "desjardins_compte"


def nettoie_description(desc, source):
    d = re.sub(r"\s+", " ", desc.replace(" / ", " ").replace("/", " ")).strip(" -")
    if source in ENTETES:
        # « CANADIAN TIRE / #435 MONTREAL » → « Canadian Tire » (1re ligne = commerçant)
        num = r"#?\d+" if source == "amex" else r"#|#?\d{3,}"      # RBC : garder « 3 » de « 3 Brasseurs »
        mots = [w for w in d.split() if not re.fullmatch(num, w)]
        while len(mots) > 1 and mots[-1].upper().strip(",.") in VILLES:
            mots.pop()
        d = " ".join(mots)
    if d.isupper():
        d = " ".join(w if (len(w) <= 3 and w.isalpha()) else w.capitalize() for w in d.lower().split())
        d = " ".join(w.upper() if w.upper() in ("IGA", "HEC", "SAQ", "STM", "BMO", "TD", "RBC", "UQAM", "PC") else w
                     for w in d.split())
    return d.strip()


def analyse_image(img, annee_def, date_precedente=None, source_prec=None):
    lignes, W, H, source = _lignes(img, source_prec)
    if source != source_prec:
        date_precedente = None                         # autre carte : la date d'en-tête ne se reporte pas
    texte_complet = " ".join(L["texte"] for L in lignes)
    m_ent = RX_ENTETE.search(sans_accents(texte_complet).lower())
    annee = int(m_ent.group(2)) if m_ent else annee_def
    montants, dates, textes = [], [], []
    y_posted, en_attente_vu = None, False
    for L in lignes:
        tn = sans_accents(L["texte"]).lower()
        if source.startswith("rbc"):
            if L["yc"] < H * 0.19:
                continue                               # en-tête : nom de la carte et solde
            if re.match(r"(posted|comptabilisee)", tn):
                y_posted = L["yc"]
                continue
            if re.match(r"(as of|en date du|pending|en attente)", tn):
                en_attente_vu = True
                continue
        mt = _montant(L, W, source)
        if mt:
            signe, valeur, gauche = mt
            montants.append({"signe": signe, "montant": valeur, "yc": L["yc"], "h": L["h"],
                             "lignes": [(L["yc"], gauche)] if gauche else [], "date": None})
            continue
        if RX_ENTETE.search(tn):
            continue                                   # « Septembre 2026 », « Facturées en … »
        d = _date(L["texte"], annee)
        if d:
            dates.append((L["yc"], d))
            continue
        mp = re.fullmatch(r"([0-9|\]lI]{1,2})(?:[\s.,:;]*\S{2,3})?", L["texte"].strip())
        if source not in ENTETES and mp and L["mots"][0]["x0"] < W * 0.5 and mp.group(1).translate(str.maketrans("|]lI", "1111")).isdigit():
            # « 22.56 » pour « 22 SEP » : jour lisible, mois deviné d'après les dates voisines
            dates.append((L["yc"], {"jour": int(mp.group(1).translate(str.maketrans("|]lI", "1111"))), "mois": None,
                                    "annee": annee, "style": "dessous"}))
            continue
        if re.search(r"en attente|pending", tn):
            if montants and L["yc"] - montants[-1]["yc"] < montants[-1]["h"] * 2:
                montants[-1]["attente"] = True
            continue
        if re.search(r"\d+[,.]\d+ ?%|^\d+ ?%$", L["texte"]) or "32007" in L["texte"] or L["mots"][0]["x0"] > W * 0.6:
            continue
        if source in ENTETES and L["yc"] < H * 0.17:
            continue                                   # titre de la carte
        alnum = sum(ch.isalnum() for ch in L["texte"])
        if alnum < max(3, 0.6 * len(L["texte"].replace(" ", ""))):
            continue                                   # icônes / bruit
        textes.append((L["yc"], L["texte"]))
    if not montants:
        return [], source, date_precedente
    if source.startswith("rbc"):
        for m in montants:
            # transactions en attente : au-dessus de « Posted » (ou toute la capture si seulement « As of … »)
            if (y_posted is not None and m["yc"] < y_posted) or (y_posted is None and en_attente_vu):
                m["attente"] = True
            if source == "rbc_credit":
                m["signe"] = "+" if m["signe"] == "-" else ""   # carte : « -900.12 » = paiement / crédit
    # chaque ligne de texte va au montant le plus proche verticalement (description sur 1 à 3 lignes)
    for yc, t in textes:
        best = min(montants, key=lambda m: abs(m["yc"] - yc))
        dy = yc - best["yc"]
        if source in ENTETES:
            if abs(dy) <= best["h"] * 1.6:
                best["lignes"].append((yc, t))
        elif -best["h"] * 0.8 <= dy <= best["h"] * 4.5:
            best["lignes"].append((yc, t))
    # dates : Amex = en-tête au-dessus d'un groupe ; Desjardins = sous chaque transaction
    if source in ENTETES:
        courante = date_precedente
        evts = sorted([(yc, "d", d) for yc, d in dates] + [(m["yc"], "m", m) for m in montants], key=lambda e: e[0])
        for yc, k, v in evts:
            if k == "d":
                courante = v
            else:
                v["date"] = courante
        date_fin = courante
    else:
        for yc, d in dates:
            avant = [m for m in montants if m["yc"] < yc and m["date"] is None]
            if avant:
                max(avant, key=lambda m: m["yc"])["date"] = dict(d)
        connus = [m["date"]["mois"] for m in montants if m["date"] and m["date"]["mois"]]
        for m in montants:
            if m["date"] and not m["date"]["mois"]:
                if connus and 1 <= m["date"]["jour"] <= 31:
                    m["date"]["mois"] = max(set(connus), key=connus.count)
                else:
                    m["date"] = None
        # la liste va du plus récent au plus ancien : une date qui casse l'ordre est mal lue
        def rang(m):
            return (m["date"]["mois"], m["date"]["jour"]) if m["date"] else None
        for i in range(1, len(montants)):
            x, y = montants[i - 1], montants[i]
            rx, ry = rang(x), rang(y)
            if not rx or not ry or ry <= rx:
                continue
            avant = rang(montants[i - 2]) if i >= 2 else None
            if str(rx[1]) in str(ry[1]) and rx[1] != ry[1] and (avant is None or avant >= ry):
                x["date"] = dict(y["date"])            # chiffre perdu : « 1 SEP » pour « 11 SEP »
            elif i == len(montants) - 1:
                y["date"] = None
            else:
                x["date"] = y["date"] = None
        date_fin = date_precedente
    out = []
    for x in montants:
        lignes_tx = [t for _, t in sorted(x["lignes"])]
        desc = " ".join(lignes_tx).strip()
        if not desc:
            continue                                   # transaction coupée en haut/bas de l'écran
        base = lignes_tx[0] if source in ENTETES and len(lignes_tx[0]) >= 3 else desc
        if re.fullmatch(r"(sous-)?total.*", sans_accents(desc).lower()):
            continue                                   # ligne « Total » de la période
        out.append({"source": source, "signe": x["signe"], "montant": x["montant"],
                    "description_brute": desc, "description": nettoie_description(base, source),
                    "date": x["date"], "attente": bool(x.get("attente")), "yc": x["yc"]})
    return out, source, date_fin


EXCLUSIONS = [
    (r"temp auth|auth hold|autorisation temp", "Autorisation temporaire (sera annulée)"),
    (r"fidelity|wealthsimple|disnat|questrade|placement|celiapp|celi\b|reer\b", "Transfert vers un placement (épargne)"),
    (r"paiement facture|paiement carte|visa desjardins|remises pour etudiants|american express|amex", "Paiement de carte de crédit (déjà compté sur la carte)"),
    (r"^paiement$|^paiement merci|^payment|payment received", "Paiement reçu sur la carte"),
    (r"virement interac a |virement a |transfert", "Virement à une personne : à vérifier"),
    (r"avance de fonds|retrait", "Avance de fonds / retrait d'argent"),
]


def classe(tx):
    """Décide si la ligne est une dépense à proposer (cochée) et pourquoi pas sinon."""
    d = sans_accents(tx["description_brute"]).lower()
    if tx["signe"] == "+":
        return False, ("Entrée d'argent (dépôt, remboursement)" if tx["source"] in ("desjardins_compte", "rbc_compte")
                       else "Crédit sur la carte (paiement ou remboursement)")
    for rx, raison in EXCLUSIONS:
        if re.search(rx, d):
            return False, raison
    if tx["source"] in ("desjardins_compte", "rbc_compte") and tx["signe"] != "-":
        return False, "Sens du montant incertain"
    return True, ""


def _jour(d):
    return datetime.date.fromisoformat(d) if d else None


def _cle_desc(t):
    t = sans_accents(t["description"]).lower()
    return re.sub(r"[^a-z]", "", t)[:8]


def _mots_communs(a, b):
    def mots(t):
        return {w for w in re.findall(r"[a-z]{4,}", sans_accents(t["description_brute"]).lower())
                if w.upper() not in VILLES and w != "montreal"}
    return bool(mots(a) & mots(b))


def analyse(images, cle_mois, existantes, categorie_auto, revenus=None):
    """images : [{largeur, hauteur, mots:[{t,x0,y0,x1,y1,c}]}] dans l'ordre du défilement ;
    cle_mois : 'AAAA-MM' du mois visé ; existantes : [{date:'AAAA-MM-JJ' ou None, montant, libelle}]
    déjà dans le fichier. Renvoie les lignes proposées."""
    an, mo = (int(x) for x in cle_mois.split("-"))
    brut, date_prec, prec_img, src_prec = [], None, [], None
    for k, img in enumerate(images):
        txs, source, date_prec = analyse_image(img, an, date_prec, src_prec)
        src_prec = source
        for tx in txs:
            d = tx["date"]
            date = None
            if d:
                annee = d["annee"]
                if d["mois"] > mo + 1 and annee == an:   # ex. décembre vu en janvier → année précédente
                    annee -= 1
                try:
                    date = datetime.date(annee, d["mois"], d["jour"]).isoformat()
                except ValueError:
                    date = None
            tx["date_iso"] = date
            tx["image"] = k
        # chevauchement avec la capture précédente : les premières lignes identiques aux dernières de la précédente
        if prec_img and txs and txs[0]["source"] == prec_img[0]["source"]:
            # chevauchement : les k premières lignes = les k dernières de la capture précédente, dans le même ordre
            # (vérifier la suite complète évite d'enlever un achat identique répété, ex. 4 × Ggpoker 14,19 $)
            queue = prec_img[-8:]

            def pareil(p, t, premier):
                if p["signe"] != t["signe"]:
                    return False
                ecart = abs(p["montant"] - t["montant"])
                meme_date = p["date_iso"] == t["date_iso"] or not p["date_iso"] or not t["date_iso"]
                if ecart < 0.005 and _cle_desc(p) == _cle_desc(t) and meme_date:
                    return True
                # 1re ligne coupée en haut de l'écran : description illisible, date d'en-tête héritée (Amex)
                return premier and (ecart < 0.005 or ecart < 0.1 and _mots_communs(p, t)) and \
                    (meme_date or t["source"] == "amex")
            k_ok = 0
            for k in range(min(len(queue), len(txs)), 0, -1):
                if all(pareil(queue[-k + j], txs[j], j == 0) for j in range(k)):
                    k_ok = k
                    break
            for j in range(k_ok):
                p, t0 = queue[-k_ok + j], txs[j]
                if not p["date_iso"] and t0["date_iso"]:
                    p["date_iso"] = t0["date_iso"]             # la date était coupée sur la capture précédente
            txs = txs[k_ok:]
        brut.extend(txs)
        prec_img = txs if txs else prec_img

    # correspondance avec ce qui est déjà dans le fichier (1 pour 1, date la plus proche)
    restants = [dict(e, _i=i) for i, e in enumerate(existantes) if e.get("montant") is not None]

    def ecart_j(e, date):
        return abs((_jour(e["date"]) - _jour(date)).days) if e.get("date") and date else 9

    doublons = {}
    # 1) montant exact, ±3 jours ; 2) écart de quelques cents (faute de frappe ou de lecture), ±1 jour
    for tol, jours, probable in ((0.005, 3, False), (0.055, 1, True)):
        for n, tx in enumerate(brut):
            if tx["signe"] == "+" or n in doublons:
                continue
            date = tx["date_iso"]
            cands = [e for e in restants if abs(e["montant"] - tx["montant"]) < tol and
                     (not e.get("date") or not date or ecart_j(e, date) <= jours)]
            if probable:
                cands = [e for e in cands if e.get("date") and date]
            if cands:
                e = min(cands, key=lambda e: (ecart_j(e, date), abs(e["montant"] - tx["montant"])))
                restants.remove(e)
                doublons[n] = (e, probable)
    res = []
    for n, tx in enumerate(brut):
        ok, raison = classe(tx)
        date = tx["date_iso"]
        doublon, probable = doublons.get(n, (None, False))
        autre_mois = bool(date) and date[:7] != cle_mois
        if ok and doublon:
            lib = " : « %s »" % doublon["libelle"] if doublon.get("libelle") else ""
            if probable:
                ok, raison = False, "Probablement déjà dans ton fichier%s à %s $" % (lib, ("%.2f" % doublon["montant"]).replace(".", ","))
            else:
                ok, raison = False, "Déjà dans ton fichier" + lib
        if ok and autre_mois:
            ok, raison = False, "Date hors du mois sélectionné"
        if ok and tx.get("attente"):
            raison = "En attente : le montant peut encore changer"
        if ok and not date:
            raison = "Date non lue : vérifie-la"
        res.append({"image": tx["image"], "source": tx["source"], "date": date, "description": tx["description"],
                    "description_brute": tx["description_brute"], "montant": tx["montant"], "signe": tx["signe"],
                    "categorie": categorie_auto(tx["description"]) or categorie_auto(tx["description_brute"]) or "",
                    "cocher": ok, "raison": raison, "doublon": bool(doublon), "autre_mois": autre_mois,
                    "attente": bool(tx.get("attente"))})

    # achats inscrits en un seul total dans le fichier (ex. 2 × Ggpoker, ou Maxi + Provigo = « Épicerie »)
    import itertools
    for e in sorted(restants, key=lambda e: -e["montant"]):
        if not e.get("date"):
            continue
        pool = [r for r in res if r["cocher"] and r["date"] and abs((_jour(r["date"]) - _jour(e["date"])).days) <= 2]
        trouve = None
        for n in (2, 3, 4):
            for combo in itertools.combinations(pool, n):
                if abs(sum(r["montant"] for r in combo) - e["montant"]) < 0.015:
                    trouve = combo
                    break
            if trouve or len(pool) > 14:
                break
        if trouve:
            for r in trouve:
                r["cocher"] = False
                r["doublon"] = True
                r["raison"] = "Déjà dans ton fichier, regroupé : « %s » %s $" % (e.get("libelle") or "", ("%.2f" % e["montant"]).replace(".", ","))
            e["_regroupe"] = True
    # montant net inscrit dans le fichier (ex. Ggpoker : dépôts − retraits de la semaine)
    for e in restants:
        if e.get("_regroupe") or not e.get("date") or not e.get("libelle"):
            continue
        cle = re.sub(r"[^a-z]", "", sans_accents(e["libelle"]).lower())[:6]
        if len(cle) < 4:
            continue
        pool = [r for r in res if (r["cocher"] or r["signe"] == "+" and not r["doublon"]) and
                re.sub(r"[^a-z]", "", sans_accents(r["description"]).lower())[:6] == cle and
                (not r["date"] or abs((_jour(r["date"]) - _jour(e["date"])).days) <= 3)]
        if not any(r["signe"] == "+" for r in pool) or len(pool) > 12:
            continue
        meilleur = None
        for masque in range(1, 1 << len(pool)):
            combo = [pool[i] for i in range(len(pool)) if masque >> i & 1]
            if len(combo) < 2:
                continue
            net = sum(-r["montant"] if r["signe"] == "+" else r["montant"] for r in combo)
            if abs(net - e["montant"]) < 0.015 and (meilleur is None or len(combo) < len(meilleur)):
                meilleur = combo
        if meilleur:
            for r in meilleur:
                r["cocher"] = False
                r["doublon"] = True
                r["raison"] = "Déjà dans ton fichier, en montant net : « %s » %s $" % (e["libelle"], ("%.2f" % e["montant"]).replace(".", ","))
            e["_regroupe"] = True
    # entrées d'argent (et avances de fonds poker) déjà inscrites dans les revenus du mois
    revs = [dict(r) for r in (revenus or []) if r.get("montants")]
    for passe in ("+", "avance"):
        for r in res:
            if r["cocher"] or r["doublon"]:
                continue
            if passe == "+" and r["signe"] != "+":
                continue
            if passe == "avance" and not re.search(r"avance de fonds", sans_accents(r["description_brute"]).lower()):
                continue
            cands = [v for v in revs if any(abs(m - r["montant"]) < 0.005 for m in v["montants"]) and
                     (not v.get("date") or not r["date"] or abs((_jour(v["date"]) - _jour(r["date"])).days) <= 3)]
            if cands:
                v = min(cands, key=lambda v: abs((_jour(v["date"]) - _jour(r["date"])).days) if v.get("date") and r["date"] else 9)
                revs.remove(v)
                r["raison"] = "Déjà dans tes revenus" + (" : « %s »" % v["libelle"] if v.get("libelle") else "")
    global DERNIERS_RESTANTS
    DERNIERS_RESTANTS = [e for e in restants if not e.get("_regroupe")]   # pour le diagnostic
    return res
