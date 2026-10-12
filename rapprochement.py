"""Rapprochement bancaire mensuel (lecture seule) : relevés Amex / Desjardins crédit / Desjardins compte
comparés aux dépenses et revenus de l'app.

Tout se fait sur l'appareil (navigateur via Pyodide, ou Mac) : les relevés ne sont ni envoyés, ni
enregistrés dans l'Excel. Entrées acceptées :
  - CSV / Excel (export de la banque) ;
  - PDF de relevé officiel (texte extrait ici même, sans bibliothèque externe) ;
  - captures d'écran (mots + positions lus par Tesseract dans le navigateur).
"""
import base64
import datetime
import io
import itertools
import re
import unicodedata
import zlib

COMPTES = {"amex": "Amex Cobalt", "desj_credit": "Desjardins crédit", "desj_compte": "Desjardins compte"}
MOIS = {"jan": 1, "janv": 1, "fev": 2, "fevr": 2, "feb": 2, "mar": 3, "mars": 3, "avr": 4, "apr": 4, "mai": 5,
        "may": 5, "jun": 6, "juin": 6, "jui": 6, "jul": 7, "juil": 7, "aou": 8, "aout": 8, "aug": 8, "sep": 9,
        "sept": 9, "oct": 10, "nov": 11, "dec": 12}
MOIS_NOMS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
VILLES = {"montreal", "westmount", "laval", "longueuil", "qc", "quebec", "ca", "can", "outremont", "verdun", "on",
          "brossard", "toronto", "mtl", "beaconsfield", "delson", "pointe", "claire", "gb", "london", "san", "francisco"}


def sans_accents(t):
    return "".join(c for c in unicodedata.normalize("NFKD", t or "") if not unicodedata.combining(c))


def norm(t):
    return sans_accents(t).lower()


def jour(iso):
    return datetime.date.fromisoformat(iso)


def fr_date(iso):
    d = jour(iso)
    return "%d %s" % (d.day, MOIS_NOMS[d.month - 1])


# ---------------------------------------------------------------------------- montants et dates
RX_MT = re.compile(r"(?<![\w.,])([-−–]?)\s?\$?\s?(\d{1,3}(?:[   ,.]\d{3})*|\d+)([.,])(\d{2})(?![\d])\s?\$?\s?(CR)?", re.I)


def montant(txt):
    """« 1 700.00 », « 1,337.13 », « 20,60 », « -751.28 », « 44,62CR » → (valeur, est_credit) ou None."""
    m = None
    for m in RX_MT.finditer(txt or ""):
        pass
    if not m:
        return None
    ent = re.sub(r"[   ,.]", "", m.group(2))
    v = float(ent + "." + m.group(4))
    return v, bool(m.group(1)) or bool(m.group(5))


def montants_fin(txt):
    """Tous les montants d'une ligne, dans l'ordre (pour séparer montant et solde)."""
    res = []
    for m in RX_MT.finditer(txt or ""):
        ent = re.sub(r"[   ,.]", "", m.group(2))
        res.append((float(ent + "." + m.group(4)), bool(m.group(1)) or bool(m.group(5)), m.start(), m.end()))
    return res


def annee_proche(jr, mo, an_ref, mo_ref):
    """Année la plus proche du mois de référence (relevé de décembre lu en janvier, etc.)."""
    best = None
    for an in (an_ref - 1, an_ref, an_ref + 1):
        try:
            d = datetime.date(an, mo, jr)
        except ValueError:
            continue
        dist = abs((an - an_ref) * 12 + mo - mo_ref)
        if best is None or dist < best[0]:
            best = (dist, d)
    return best[1].isoformat() if best else None


def mois_num(t):
    t = norm(t).strip(". ")
    return MOIS.get(t) or MOIS.get(t[:4]) or MOIS.get(t[:3])


def date_libre(txt, an_ref, mo_ref, ordre=None):
    """« 2026/09/01 », « 2026-09-01 », « 01/09/2026 », « 09/01/2026 », « 1 Sep 2026 », « Sep 1, 2026 »."""
    t = (txt or "").strip()
    m = re.fullmatch(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})(?:[ T].*)?", t)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", t)
    if m:
        a, b, an = int(m.group(1)), int(m.group(2)), int(m.group(3))
        an = an + 2000 if an < 100 else an
        if ordre == "md" or (ordre is None and b > 12):
            a, b = b, a
        try:
            return datetime.date(an, b, a).isoformat()
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2})\s+([A-Za-zÀ-ÿ]{3,9})\.?,?\s+(\d{4})", t)
    if m and mois_num(m.group(2)):
        try:
            return datetime.date(int(m.group(3)), mois_num(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            return None
    m = re.fullmatch(r"([A-Za-zÀ-ÿ]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})", t)
    if m and mois_num(m.group(1)):
        try:
            return datetime.date(int(m.group(3)), mois_num(m.group(1)), int(m.group(2))).isoformat()
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------- PDF : texte avec positions
GLYPHES = {"space": " ", "exclam": "!", "quotedbl": '"', "numbersign": "#", "dollar": "$", "percent": "%",
           "ampersand": "&", "quoteright": "'", "quotesingle": "'", "parenleft": "(", "parenright": ")",
           "asterisk": "*", "plus": "+", "comma": ",", "hyphen": "-", "minus": "-", "period": ".", "slash": "/",
           "colon": ":", "semicolon": ";", "less": "<", "equal": "=", "greater": ">", "question": "?", "at": "@",
           "bracketleft": "[", "backslash": "\\", "bracketright": "]", "underscore": "_", "quoteleft": "'",
           "braceleft": "{", "bar": "|", "braceright": "}", "endash": "–", "emdash": "—", "guillemotleft": "«",
           "guillemotright": "»", "nbspace": " ", "germandbls": "ß", "oe": "œ", "OE": "Œ", "ae": "æ", "AE": "Æ",
           "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7",
           "eight": "8", "nine": "9", "degree": "°", "section": "§", "bullet": "•", "quotedblleft": '"',
           "quotedblright": '"', "ellipsis": "…", "Euro": "€", "sterling": "£"}
ACCENTS = {"acute": "ACUTE", "grave": "GRAVE", "circumflex": "CIRCUMFLEX", "dieresis": "DIAERESIS",
           "cedilla": "CEDILLA", "tilde": "TILDE", "ring": "RING ABOVE"}


def glyphe(nom):
    if nom in GLYPHES:
        return GLYPHES[nom]
    if len(nom) == 1:
        return nom
    m = re.fullmatch(r"uni([0-9A-Fa-f]{4})", nom)
    if m:
        return chr(int(m.group(1), 16))
    m = re.fullmatch(r"([A-Za-z])(acute|grave|circumflex|dieresis|cedilla|tilde|ring)", nom)
    if m:
        try:
            return unicodedata.lookup("LATIN %s LETTER %s WITH %s" % ("CAPITAL" if m.group(1).isupper() else "SMALL",
                                                                     m.group(1).upper(), ACCENTS[m.group(2)]))
        except KeyError:
            return ""
    return ""


class PDF:
    """Lecteur PDF minimal : objets (y compris flux d'objets), pages, polices (ToUnicode, encodages simples),
    opérateurs de texte. Suffit pour les relevés générés par les banques ; un PDF numérisé (image) n'a pas de texte."""

    def __init__(self, data):
        self.data = data
        self.objs = {}
        for m in re.finditer(rb"(\d+)\s+(\d+)\s+obj\b(.*?)\bendobj", data, re.S):
            self.objs[int(m.group(1))] = m.group(3)
        for num, raw in list(self.objs.items()):
            if re.search(rb"/Type\s*/ObjStm", raw):
                try:
                    self._objstm(raw)
                except Exception:
                    pass

    # --- objets
    def _objstm(self, raw):
        d, s = self.flux(raw)
        n = int(re.search(rb"/N\s+(\d+)", d).group(1))
        first = int(re.search(rb"/First\s+(\d+)", d).group(1))
        nums = [int(x) for x in s[:first].split()]
        for i in range(n):
            num, off = nums[2 * i], nums[2 * i + 1]
            fin = nums[2 * i + 3] if i + 1 < n else len(s) - first
            if num not in self.objs:
                self.objs[num] = s[first + off:first + fin]

    def flux(self, raw):
        m = re.search(rb"stream\r?\n", raw)
        if not m:
            return raw, b""
        d = raw[:m.start()]
        fin = raw.rfind(b"endstream")
        s = raw[m.end():fin if fin > 0 else len(raw)]
        fm = re.search(rb"/Filter\s*(\[[^\]]*\]|/[A-Za-z0-9]+)", d)
        filtres = re.findall(rb"/([A-Za-z0-9]+)", fm.group(1)) if fm else []
        for flt in filtres:
            if flt in (b"ASCII85Decode", b"A85"):
                t = re.sub(rb"\s", b"", s)
                if t.startswith(b"<~"):
                    t = t[2:]
                fin = t.find(b"~>")
                t = t[:fin] if fin >= 0 else t
                try:
                    s = base64.a85decode(t)
                except ValueError:
                    return d, b""
            elif flt in (b"ASCIIHexDecode", b"AHx"):
                t = re.sub(rb"[^0-9A-Fa-f]", b"", s.split(b">")[0])
                s = bytes.fromhex((t + b"0" if len(t) % 2 else t).decode())
            elif flt in (b"FlateDecode", b"Fl"):
                for w in (zlib.MAX_WBITS, -zlib.MAX_WBITS):
                    try:
                        s = zlib.decompressobj(w).decompress(s)
                        break
                    except zlib.error:
                        continue
        return d, s

    def obj(self, x):
        """Référence « 12 0 R » → contenu de l'objet ; sinon renvoie tel quel."""
        if isinstance(x, (bytes, bytearray)):
            m = re.fullmatch(rb"\s*(\d+)\s+\d+\s+R\s*", x)
            if m:
                return self.objs.get(int(m.group(1)), b"")
        return x

    @staticmethod
    def valeur(d, cle):
        """Valeur brute de /Cle dans un dictionnaire (référence, nom, tableau ou sous-dictionnaire)."""
        m = re.search(rb"/" + cle + rb"(?![A-Za-z0-9])\s*", d)
        if not m:
            return None
        i = m.end()
        if d[i:i + 2] == b"<<":
            prof, j = 0, i
            while j < len(d):
                if d[j:j + 2] == b"<<":
                    prof += 1; j += 2; continue
                if d[j:j + 2] == b">>":
                    prof -= 1; j += 2
                    if prof == 0:
                        return d[i:j]
                    continue
                j += 1
            return d[i:]
        if d[i:i + 1] == b"[":
            prof, j = 0, i
            while j < len(d):
                if d[j:j + 1] == b"[":
                    prof += 1
                elif d[j:j + 1] == b"]":
                    prof -= 1
                    if prof == 0:
                        return d[i:j + 1]
                j += 1
            return d[i:]
        m2 = re.match(rb"(\d+\s+\d+\s+R|/[^\s/<>\[\]()]+|-?[\d.]+|\([^)]*\)|true|false)", d[i:])
        return m2.group(1) if m2 else None

    def dico(self, x):
        x = self.obj(x)
        if x is None:
            return b""
        i = x.find(b"stream")
        return x[:i] if i >= 0 else x

    # --- pages
    def pages(self):
        cat = next((r for r in self.objs.values() if re.search(rb"/Type\s*/Catalog", r)), None)
        res = []
        if cat is not None:
            racine = self.valeur(self.dico(cat), b"Pages")
            self._kids(racine, res, {}, 0)
        if not res:
            for num in sorted(self.objs):
                r = self.objs[num]
                if re.search(rb"/Type\s*/Page(?!s)", r):
                    res.append((self.dico(r), {}))
        return res

    def _kids(self, ref, res, herite, prof):
        if prof > 30:
            return
        d = self.dico(ref)
        h = dict(herite)
        for cle in (b"Resources", b"MediaBox"):
            v = self.valeur(d, cle)
            if v is not None:
                h[cle] = v
        if re.search(rb"/Type\s*/Pages", d):
            kids = self.valeur(d, b"Kids") or b"[]"
            for k in re.findall(rb"\d+\s+\d+\s+R", kids):
                self._kids(k, res, h, prof + 1)
        elif re.search(rb"/Type\s*/Page", d):
            res.append((d, h))

    # --- polices
    def police(self, ref):
        d = self.dico(ref)
        f = {"cid": bool(re.search(rb"/Subtype\s*/Type0", d)), "map": {}, "enc": "cp1252", "diff": {}}
        tu = self.valeur(d, b"ToUnicode")
        if tu is not None and re.match(rb"\d", tu):
            _, s = self.flux(self.obj(tu))
            f["map"] = self.cmap(s)
        enc = self.valeur(d, b"Encoding")
        if enc is not None:
            e = self.obj(enc) if re.match(rb"\d", enc) else enc
            if b"Differences" in e:
                diff = self.valeur(e, b"Differences") or b""
                code = 0
                for tok in re.findall(rb"\d+|/[^\s/\[\]]+", diff):
                    if tok[:1] == b"/":
                        f["diff"][code] = glyphe(tok[1:].decode("latin-1"))
                        code += 1
                    else:
                        code = int(tok)
            if b"MacRoman" in e:
                f["enc"] = "mac_roman"
        if f["cid"] and not f["map"]:
            f["enc"] = "utf-16-be"
        return f

    @staticmethod
    def cmap(s):
        mp = {}
        for bloc in re.findall(rb"beginbfchar(.*?)endbfchar", s, re.S):
            for a, b in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", bloc):
                try:
                    mp[(int(a, 16), len(a) // 2)] = bytes.fromhex(b.decode()).decode("utf-16-be", "ignore")
                except ValueError:
                    pass
        for bloc in re.findall(rb"beginbfrange(.*?)endbfrange", s, re.S):
            for a, b, c in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])", bloc):
                lo, hi, n = int(a, 16), int(b, 16), len(a) // 2
                if c.startswith(b"["):
                    for k, h in enumerate(re.findall(rb"<([0-9A-Fa-f]*)>", c)):
                        try:
                            mp[(lo + k, n)] = bytes.fromhex(h.decode()).decode("utf-16-be", "ignore")
                        except ValueError:
                            pass
                else:
                    base = bytes.fromhex(c[1:-1].decode())
                    v0 = int.from_bytes(base, "big") if base else 0
                    for k in range(min(hi - lo + 1, 65536)):
                        try:
                            mp[(lo + k, n)] = (v0 + k).to_bytes(max(2, len(base)), "big").decode("utf-16-be", "ignore")
                        except (OverflowError, ValueError):
                            pass
        return mp

    @staticmethod
    def decode(f, b):
        if f["map"]:
            n = 2 if f["cid"] else 1
            out = []
            i = 0
            while i < len(b):
                code = int.from_bytes(b[i:i + n], "big")
                ch = f["map"].get((code, n))
                if ch is None and n == 1:
                    ch = f["diff"].get(code) or bytes([code]).decode("cp1252", "ignore")
                out.append(ch or "")
                i += n
            return "".join(out)
        if f["diff"]:
            return "".join(f["diff"].get(c) or bytes([c]).decode("cp1252", "ignore") for c in b)
        if f["enc"] == "utf-16-be":
            return b.decode("utf-16-be", "ignore")
        return b.decode(f["enc"], "ignore")

    # --- contenu
    @staticmethod
    def jetons(s):
        i, n = 0, len(s)
        while i < n:
            c = s[i:i + 1]
            if c in b" \t\r\n\f\x00":
                i += 1
            elif c == b"%":
                j = s.find(b"\n", i)
                i = n if j < 0 else j + 1
            elif c == b"(":
                prof, j, out = 1, i + 1, bytearray()
                while j < n and prof:
                    ch = s[j:j + 1]
                    if ch == b"\\":
                        nx = s[j + 1:j + 2]
                        esc = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b", b"f": b"\f", b"(": b"(", b")": b")", b"\\": b"\\"}
                        if nx in esc:
                            out += esc[nx]; j += 2
                        elif nx.isdigit():
                            m = re.match(rb"[0-7]{1,3}", s[j + 1:j + 4])
                            out.append(int(m.group(0), 8) & 255); j += 1 + len(m.group(0))
                        elif nx in (b"\r", b"\n"):
                            j += 2
                        else:
                            j += 1
                        continue
                    if ch == b"(":
                        prof += 1
                    elif ch == b")":
                        prof -= 1
                        if prof == 0:
                            break
                    out += ch
                    j += 1
                yield ("s", bytes(out))
                i = j + 1
            elif c == b"<" and s[i + 1:i + 2] != b"<":
                j = s.find(b">", i)
                h = re.sub(rb"\s", b"", s[i + 1:j])
                if len(h) % 2:
                    h += b"0"
                try:
                    yield ("s", bytes.fromhex(h.decode()))
                except ValueError:
                    yield ("s", b"")
                i = j + 1
            elif s[i:i + 2] in (b"<<", b">>"):
                yield ("d", s[i:i + 2]); i += 2
            elif c in b"[]":
                yield ("a", c); i += 1
            elif c == b"/":
                m = re.match(rb"/[^\s/\[\]()<>{}%]*", s[i:])
                yield ("n", m.group(0)); i += len(m.group(0))
            else:
                m = re.match(rb"[^\s/\[\]()<>{}%]+", s[i:])
                tok = m.group(0) if m else s[i:i + 1]
                i += len(tok)
                if re.fullmatch(rb"[+\-]?(\d+\.?\d*|\.\d+)", tok):
                    yield ("v", float(tok))
                elif tok == b"BI":
                    j = s.find(b"EI", i)                       # image incrustée : sauter les données
                    i = n if j < 0 else j + 2
                else:
                    yield ("o", tok)

    def texte_page(self, d, herite):
        mb = self.valeur(d, b"MediaBox") or herite.get(b"MediaBox") or b"[0 0 612 792]"
        b = [float(x) for x in re.findall(rb"-?[\d.]+", self.obj(mb))[:4]] or [0, 0, 612, 792]
        W, H = b[2] - b[0], b[3] - b[1]
        res = self.valeur(d, b"Resources") or herite.get(b"Resources") or b""
        cont = self.valeur(d, b"Contents")
        flux = b""
        if cont is not None:
            refs = re.findall(rb"\d+\s+\d+\s+R", cont)
            for r in refs:
                o = self.obj(r)
                if re.match(rb"\s*\[", self.dico(o) or b"") and b"stream" not in o:
                    for r2 in re.findall(rb"\d+\s+\d+\s+R", o):
                        flux += self.flux(self.obj(r2))[1] + b"\n"
                else:
                    flux += self.flux(o)[1] + b"\n"
        mots = []
        self._execute(flux, self.dico(res) if re.match(rb"\d", res or b"") else res, [1, 0, 0, 1, 0, 0], mots, 0)
        out = []
        for t, x, y, taille, larg in mots:
            t = t.replace(" ", " ")
            if not t.strip():
                continue
            out.append({"t": t.strip(), "x0": x - b[0], "x1": x - b[0] + larg, "y0": H - (y - b[1]) - taille, "y1": H - (y - b[1]), "c": 99})
        return {"largeur": W, "hauteur": H, "mots": out}

    def _execute(self, flux, res, ctm0, mots, prof):
        if prof > 3:
            return
        polices, xobj = {}, {}
        fdict = self.valeur(res, b"Font")
        if fdict is not None:
            fd = self.dico(fdict) if re.match(rb"\d", fdict) else fdict
            for nom, ref in re.findall(rb"/([^\s/<>\[\]]+)\s*(\d+\s+\d+\s+R)", fd):
                polices[nom] = ref
        xd = self.valeur(res, b"XObject")
        if xd is not None:
            xdd = self.dico(xd) if re.match(rb"\d", xd) else xd
            for nom, ref in re.findall(rb"/([^\s/<>\[\]]+)\s*(\d+\s+\d+\s+R)", xdd):
                xobj[nom] = ref
        cache = {}

        def mul(a, b):
            return [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2],
                    a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]]
        ctm, pile = list(ctm0), []
        tm = lm = [1, 0, 0, 1, 0, 0]
        fs, tl, police = 10.0, 0.0, None
        ops = []

        def montre(chaines):
            nonlocal tm
            f = police or {"cid": False, "map": {}, "enc": "cp1252", "diff": {}}
            txt = ""
            for c in chaines:
                if isinstance(c, float):
                    if c < -180:
                        txt += " "
                    continue
                txt += self.decode(f, c)
            m = mul(tm, ctm)
            taille = abs(fs * (m[3] if m[3] else m[0])) or fs
            larg = len(txt) * taille * 0.5
            mots.append((txt, m[4], m[5], taille, larg))
            tm = [tm[0], tm[1], tm[2], tm[3], tm[4] + larg / (abs(ctm[0]) or 1), tm[5]]

        for typ, v in self.jetons(flux):
            if typ != "o":
                ops.append((typ, v))
                continue
            o = v
            nums = [x[1] for x in ops if x[0] == "v"]
            if o == b"q":
                pile.append(list(ctm))
            elif o == b"Q":
                ctm = pile.pop() if pile else ctm
            elif o == b"cm" and len(nums) >= 6:
                ctm = mul(nums[-6:], ctm)
            elif o == b"BT":
                tm = lm = [1, 0, 0, 1, 0, 0]
            elif o == b"Tf" and ops:
                nm = next((x[1][1:] for x in reversed(ops) if x[0] == "n"), None)
                fs = nums[-1] if nums else fs
                if nm is not None and nm in polices:
                    if nm not in cache:
                        try:
                            cache[nm] = self.police(polices[nm])
                        except Exception:
                            cache[nm] = None
                    police = cache[nm]
            elif o == b"TL" and nums:
                tl = nums[-1]
            elif o in (b"Td", b"TD") and len(nums) >= 2:
                if o == b"TD":
                    tl = -nums[-1]
                lm = mul([1, 0, 0, 1, nums[-2], nums[-1]], lm)
                tm = list(lm)
            elif o == b"Tm" and len(nums) >= 6:
                lm = tm = nums[-6:]
            elif o == b"T*":
                lm = mul([1, 0, 0, 1, 0, -tl], lm); tm = list(lm)
            elif o == b"Tj":
                s = [x[1] for x in ops if x[0] == "s"]
                if s:
                    montre([s[-1]])
            elif o in (b"'", b'"'):
                lm = mul([1, 0, 0, 1, 0, -tl], lm); tm = list(lm)
                s = [x[1] for x in ops if x[0] == "s"]
                if s:
                    montre([s[-1]])
            elif o == b"TJ":
                arr, dedans = [], False
                for x in ops:
                    if x == ("a", b"["):
                        arr, dedans = [], True
                    elif x == ("a", b"]"):
                        dedans = False
                    elif dedans:
                        arr.append(x[1] if x[0] == "s" else (float(x[1]) if x[0] == "v" else 0.0))
                montre(arr)
            elif o == b"Do":
                nm = next((x[1][1:] for x in reversed(ops) if x[0] == "n"), None)
                if nm in xobj:
                    raw = self.obj(xobj[nm])
                    dx, sx = self.flux(raw)
                    if re.search(rb"/Subtype\s*/Form", dx):
                        mat = self.valeur(dx, b"Matrix")
                        mm = [float(x) for x in re.findall(rb"-?[\d.]+", mat or b"")[:6]] if mat else [1, 0, 0, 1, 0, 0]
                        rx = self.valeur(dx, b"Resources")
                        rr = (self.dico(rx) if rx is not None and re.match(rb"\d", rx) else rx) or res
                        self._execute(sx, rr, mul(mm if len(mm) == 6 else [1, 0, 0, 1, 0, 0], ctm), mots, prof + 1)
            ops = []


def pdf_pages(b64):
    data = base64.b64decode(b64)
    if not data.startswith(b"%PDF"):
        raise ValueError("Ce fichier n'est pas un PDF.")
    p = PDF(data)
    return [p.texte_page(d, h) for d, h in p.pages()]


# ---------------------------------------------------------------------------- lignes de texte
def lignes_de(pages):
    """[{largeur,hauteur,mots}] → lignes de texte de haut en bas (toutes les pages à la suite)."""
    import releve
    out = []
    for pg in pages:
        mots = [dict(m, yc=(m["y0"] + m["y1"]) / 2, h=max(1, m["y1"] - m["y0"])) for m in pg.get("mots") or [] if (m.get("t") or "").strip()]
        for L in releve._grouper(mots):
            out.append(re.sub(r"\s+", " ", L["texte"]).strip())
    return out


SIGNES = {
    "amex": [r"american express(?! *$)", r"\bcobalt\b", r"new transactions for", r"payment received", r"your transactions",
             r"posting\s*date", r"transaction\s*date", r"new payments", r"statement period"],
    "desj_credit": [r"date d.?inscription", r"\bremises\b", r"visa desjardins odyssee", r"carte de credit desjardins",
                    r"mastercard desjardins", r"transactions facturees", r"non facturees", r"date de transaction", r"\binscription\b"],
    "desj_compte": [r"solde reporte", r"\bretrait\b", r"\bdep[oé]t\b", r"\bsolde\b", r"operations courantes", r"\b(eop|eor)\b",
                    r"epargne avec operations", r"\bfolio\b", r"\bcode\b"],
}


def emetteur_texte(txt):
    """Émetteur d'après les mots propres à chaque relevé (en-têtes de colonnes, titres)."""
    t = norm(txt)
    # les descriptions (« Paiement facture /American Express ») ne doivent pas l'emporter sur les en-têtes
    scores = {em: sum(1 for rx in l if re.search(rx, t)) for em, l in SIGNES.items()}
    if re.search(r"paiement facture.*american express", t):
        scores["amex"] -= 1
    best = max(scores, key=lambda k: scores[k])
    return best if scores[best] > 0 else None


def periode_texte(txt, an_ref, mo_ref):
    """Période du relevé si elle est écrite (« du 2026-08-12 au 2026-09-11 », « Aug 12, 2026 to Sep 11, 2026 »)."""
    t = sans_accents(txt)
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*(?:au|a|-|to|–)\s*(\d{4}-\d{2}-\d{2})", t)
    if m:
        return [m.group(1), m.group(2)]
    m = re.search(r"(\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{4})\s*(?:au|a|-|to|–)\s*(\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{4})", t, re.I)
    if m:
        a, b = date_libre(m.group(1), an_ref, mo_ref), date_libre(m.group(2), an_ref, mo_ref)
        if a and b:
            return [a, b]
    m = re.search(r"([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4})\s*(?:-|to|–|au)\s*([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4})", t, re.I)
    if m:
        a, b = date_libre(m.group(1), an_ref, mo_ref), date_libre(m.group(2), an_ref, mo_ref)
        if a and b:
            return [a, b]
    return None


def masque(t):
    """Numéros de carte ou de compte : 4 derniers chiffres au maximum."""
    return re.sub(r"[\d*•Xx][\d*•Xx \-]{6,}(\d{4})\b", lambda m: "••••" + m.group(1) if sum(c.isdigit() or c in "*•" for c in m.group(0)) >= 8 else m.group(0), t or "")


def nettoie(desc):
    d = re.sub(r"\s+", " ", (desc or "").replace(" / ", " ").replace("/", " ")).strip(" -*")
    mots = [w for w in d.split() if not re.fullmatch(r"#?\d{2,}|\d{3}-\d{3,}-?\d*|\(?\d{3}\)?-?\d{3}-\d{4}", w)]
    while len(mots) > 1 and norm(mots[-1]).strip(",.") in VILLES:
        mots.pop()
    d = " ".join(mots)
    if d.isupper():
        petits = {"de", "du", "des", "la", "le", "les", "et", "au", "aux", "a", "en", "of", "the", "and"}
        sigles = {"iga", "hec", "saq", "saaq", "stm", "bmo", "td", "pc", "rbc", "kfc", "bp", "uqam", "ups", "cn", "via", "amc", "lcbo", "h&m", "a&w"}
        d = " ".join(w if w in petits and k else (w.upper() if w in sigles else w.capitalize())
                     for k, w in enumerate(d.lower().split()))
    return d.strip()


MO = r"(?:jan|janv|f[ée]vr?|feb|mars?|avr|apr|mai|may|juin|jun|juil|jul|ao[uû]t?|aug|sept?|oct|nov|d[ée]c)\.?"


def _date_amex(txt, an, mo):
    m = re.fullmatch(r"(" + MO + r")\s+(\d{1,2})|(\d{1,2})\s+(" + MO + r")", txt.strip(), re.I)
    if not m:
        return None
    mois = mois_num(m.group(1) or m.group(4))
    j = int(m.group(2) or m.group(3))
    return annee_proche(j, mois, an, mo) if mois else None


def lignes_amex(lignes, an, mo):
    D = r"(?:" + MO + r"\s+\d{1,2}|\d{1,2}\s+" + MO + r")"
    rx = re.compile(r"^(?P<d1>" + D + r")\s+(?:(?P<d2>" + D + r")\s+)?(?P<reste>.+)$", re.I)
    out = []
    for L in lignes:
        m = rx.match(L)
        if not m:
            continue
        mt = montants_fin(m.group("reste"))
        if not mt:
            continue
        v, credit, deb, _ = mt[-1]
        desc = m.group("reste")[:deb].strip()
        if not desc or re.match(r"total", norm(desc)):
            continue
        d1 = _date_amex(m.group("d1"), an, mo)
        if not d1:
            continue
        tx = {"date": d1, "desc_brute": desc, "montant": v, "sens": "C" if credit else "D"}
        if m.group("d2"):
            tx["date_insc"] = _date_amex(m.group("d2"), an, mo)
        else:
            tx["note_date"] = "Une seule date lue : prise comme date de transaction"
        out.append(tx)
    return out


def lignes_desj_credit(lignes, an, mo):
    rx = re.compile(r"^(\d{1,2})\s*(\d{2})\s+(\d{1,2})\s*(\d{2})\s+(.+)$")
    out = []
    for L in lignes:
        m = rx.match(L)
        if not m or not (1 <= int(m.group(2)) <= 12 and 1 <= int(m.group(4)) <= 12):
            continue
        reste = re.sub(r"\s\d{1,2}[,.]\d{2}\s?[%Y°]", " ", m.group(5))            # colonne « Remises » (0,50 %)
        mt = montants_fin(reste)
        if not mt:
            continue
        v, credit, deb, _ = mt[-1]
        desc = reste[:deb].strip()
        if not desc:
            continue
        out.append({"date": annee_proche(int(m.group(1)), int(m.group(2)), an, mo),
                    "date_insc": annee_proche(int(m.group(3)), int(m.group(4)), an, mo),
                    "desc_brute": desc, "montant": v, "sens": "C" if credit else "D"})
    return out


RX_CREDIT_COMPTE = re.compile(r"virement interac de|depot|paie\b|credit d.impot|interet|avance fonds|annulation|remboursement", re.I)


def lignes_desj_compte(lignes, an, mo):
    rx = re.compile(r"^(\d{1,2})\s*(" + MO + r")\s+(.*)$", re.I)
    out, solde, attente, derniere = [], None, None, None
    for L in lignes:
        n = norm(L)
        if out and re.search(r"epargne et de placement|epargne stable|compte d.epargne a terme", n):
            break                                      # autre compte du même relevé
        if "solde reporte" in n or "solde d.ouverture" in n:
            mt = montants_fin(L)
            if mt:
                solde = mt[-1][0]
            continue
        m = rx.match(L)
        if m:
            date = annee_proche(int(m.group(1)), mois_num(m.group(2)), an, mo)
            reste = m.group(3)
            reste = re.sub(r"^[A-Z]{2,4}\s+(?=\S)", "", reste)           # code d'opération (VMW, PWW, RA…)
            derniere = date
        elif attente is not None:
            date, reste = attente["date"], attente["desc"] + " " + L
        else:
            continue
        mt = montants_fin(reste)
        if not mt:
            attente = {"date": date, "desc": reste}
            continue
        attente = None
        desc = reste[:mt[0][2]].strip()
        if re.match(r"total|solde", norm(desc)):
            continue
        if len(mt) >= 2:
            v, nouveau = mt[-2][0], mt[-1][0]
            if solde is not None and abs(solde - v - nouveau) < 0.011:
                sens = "D"
            elif solde is not None and abs(solde + v - nouveau) < 0.011:
                sens = "C"
            else:
                sens = "C" if RX_CREDIT_COMPTE.search(norm(desc)) else "D"
            solde = nouveau
        else:
            v = mt[0][0]
            sens = "C" if RX_CREDIT_COMPTE.search(norm(desc)) else "D"
            if solde is not None:
                solde = solde - v if sens == "D" else solde + v
        out.append({"date": date or derniere, "desc_brute": desc, "montant": v, "sens": sens})
    return out


RX_RELEVE = re.compile(r"posting\s*date|transaction\s*date|date d.?inscription|solde reporte|new transactions for|transactions facturees|"
                       r"statement period|periode du|date de transaction|retrait.{0,40}dep[oe]t.{0,40}solde")


def est_releve(txt):
    """Relevé officiel (en-têtes de colonnes) plutôt qu'une capture de l'app bancaire."""
    return bool(RX_RELEVE.search(norm(txt)))


def depuis_mots(pages, an, mo, emetteur=None, suite=None, releve_pdf=False):
    """Relevé (PDF ou capture du relevé) ou capture de l'app bancaire → (émetteur, transactions, période, notes).
    `suite` : état partagé entre captures successives de l'app (en-têtes de date et sorte de compte)."""
    lignes = lignes_de(pages)
    txt = "\n".join(lignes)
    notes, txs, em = [], [], emetteur
    if releve_pdf or est_releve(txt):
        em = emetteur or emetteur_texte(txt)
        if em == "amex":
            txs = lignes_amex(lignes, an, mo)
        elif em == "desj_credit":
            txs = lignes_desj_credit(lignes, an, mo)
        elif em == "desj_compte":
            txs = lignes_desj_compte(lignes, an, mo)
    if not txs:
        # capture de l'app bancaire (formats déjà reconnus par l'import de captures)
        if releve_pdf:
            return em, [], periode_texte(txt, an, mo), notes
        import releve
        suite = suite if suite is not None else {}
        src_vu, n_sans = None, 0
        for img in pages:
            r, src, suite["date"] = releve.analyse_image(img, an, suite.get("date"), suite.get("src"))
            suite["src"] = src_vu = src
            for t in r:
                d = t.get("date")
                if not d:
                    n_sans += 1
                    continue
                try:
                    date = datetime.date(d["annee"] if d["annee"] else an, d["mois"], d["jour"]).isoformat()
                except ValueError:
                    continue
                txs.append({"date": date, "desc_brute": t["description_brute"], "montant": t["montant"], "sens": "C" if t["signe"] == "+" else "D",
                            "note_date": "Une seule date à l'écran : date de transaction" if src == "amex" else "Date de l'opération (affichée par l'app)"})
        em = emetteur or {"amex": "amex", "desjardins_credit": "desj_credit", "desjardins_compte": "desj_compte"}.get(src_vu)
        if txs:
            notes.append("Capture de l'app bancaire.")
        if n_sans:
            notes.append("%d transaction%s sans date (en attente) non prise%s en compte." % (n_sans, "s" if n_sans > 1 else "", "s" if n_sans > 1 else ""))
    return em, txs, periode_texte(txt, an, mo), notes


def lire_images(images, noms, cle_mois):
    """Captures choisies ensemble : lues dans l'ordre, l'en-tête de date d'une capture de l'app se reporte
    à la suivante (comme l'import de captures)."""
    suite, res = {}, []
    for img, nom in zip(images, noms):
        res.append(lire({"nom": nom, "type": "mots", "pages": [img]}, cle_mois, suite))
    return res


# ---------------------------------------------------------------------------- CSV / Excel
def lignes_csv(texte):
    import csv
    t = texte.lstrip("﻿")
    echant = t[:4000]
    delim = max([",", ";", "\t"], key=lambda c: echant.count(c))
    return [[c.strip() for c in r] for r in csv.reader(io.StringIO(t), delimiter=delim) if any(c.strip() for c in r)]


def lignes_xlsx(b64):
    import os
    import tempfile
    import xlsx as X
    data = base64.b64decode(b64)
    if not data.startswith(b"PK"):
        raise ValueError("Ce fichier Excel est dans l'ancien format (.xls) : exporte-le plutôt en CSV.")
    d = tempfile.mkdtemp()
    p = os.path.join(d, "r.xlsx")
    with open(p, "wb") as fh:
        fh.write(data)
    try:
        f = X.Classeur(p).feuilles[0]
    finally:
        try:
            os.remove(p)
            os.rmdir(d)
        except OSError:
            pass
    def num_col(c):
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        return n
    rows = {}
    for c in f.cellules.values():
        v = c.valeur
        if isinstance(v, tuple):
            v = None
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if float(v).is_integer() and 36000 < v < 55000:
                v = X.serie_en_date(v).isoformat()
            else:
                v = "%.2f" % v
        rows.setdefault(c.lig, {})[num_col(c.col)] = "" if v is None else str(v)
    out = []
    for lig in sorted(rows):
        r = rows[lig]
        out.append([r.get(i, "") for i in range(1, max(r) + 1)])
    return [r for r in out if any(x.strip() for x in r)]


def nombre(c):
    c = (c or "").replace(" ", " ").strip()
    if not c:
        return None
    m = re.fullmatch(r"\(?([-−]?)\s?\$?\s?([\d ,.]+?)\s?\$?\)?\s?(CR)?", c, re.I)
    if not m or not re.search(r"\d", m.group(2)):
        return None
    s = m.group(2).replace(" ", "")
    if re.search(r",\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return None
    neg = bool(m.group(1)) or c.startswith("(") or bool(m.group(3))
    return -v if neg else v


def depuis_tableau(rows, an, mo, emetteur=None):
    """Lignes d'un CSV / Excel → (émetteur, transactions, notes). Avec en-tête (Amex, la plupart des exports)
    ou sans (AccèsD : institution, folio, compte, date, n°, description, …, retrait, dépôt, …, solde)."""
    notes = []
    MOTS = r"date|description|montant|amount|retrait|depot|debit|credit|withdrawal|deposit|details|solde|balance|marchand|merchant"
    entete, i0 = None, 0
    for k, r in enumerate(rows[:15]):
        n = [norm(c) for c in r]
        if sum(1 for c in n if re.search(MOTS, c) and not re.search(r"\d{3}", c)) >= 2:
            entete, i0 = n, k + 1
            break
    corps = rows[i0:]
    # ordre jour/mois des dates « 01/09/2026 » : celui qui place le plus de dates près du mois choisi
    ordre = None
    tests = [c for r in corps[:200] for c in r if re.fullmatch(r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}", c)]
    if tests:
        def score(o):
            s = 0
            for c in tests:
                d = date_libre(c, an, mo, o)
                if d and abs((jour(d) - datetime.date(an, mo, 15)).days) <= 50:
                    s += 1
            return s
        ordre = "dm" if score("dm") >= score("md") else "md"
    txs = []
    if entete:
        idx = lambda rx, sauf=None: next((i for i, c in enumerate(entete) if re.search(rx, c) and not (sauf and re.search(sauf, c))), None)
        i_date = idx(r"date", r"inscri|post|process|trait|compta")
        i_insc = idx(r"date.*(inscri|post|process|trait|compta)|(inscri|post|process)")
        i_desc = idx(r"description|details|libel|marchand|merchant|commer")
        i_mt = idx(r"^montant|amount|^montant ")
        i_deb = idx(r"retrait|debit|withdrawal|achat|purchase", r"date")
        i_cre = idx(r"depot|credit|deposit|paiement|payment", r"date|carte")
        i_sol = idx(r"solde|balance")
        for r in corps:
            g = lambda i: r[i] if i is not None and i < len(r) else ""
            d = date_libre(g(i_date), an, mo, ordre)
            if not d:
                continue
            if i_deb is not None or i_cre is not None:
                vd, vc = nombre(g(i_deb)), nombre(g(i_cre))
                if vd:
                    v, sens = abs(vd), ("C" if vd < 0 else "D")
                elif vc:
                    v, sens = abs(vc), ("D" if vc < 0 else "C")
                else:
                    continue
            else:
                vm = nombre(g(i_mt))
                if not vm:
                    continue
                # Amex et cartes : positif = achat ; un compte (avec solde) : négatif = retrait
                if i_sol is not None:
                    v, sens = abs(vm), ("D" if vm < 0 else "C")
                else:
                    v, sens = abs(vm), ("C" if vm < 0 else "D")
            tx = {"date": d, "desc_brute": g(i_desc) or " ".join(x for x in r if re.search(r"[A-Za-z]{3}", x))[:80], "montant": v, "sens": sens}
            di = date_libre(g(i_insc), an, mo, ordre)
            if di:
                tx["date_insc"] = di
            txs.append(tx)
        em = emetteur or ("desj_compte" if i_sol is not None else None)
    else:
        # AccèsD (sans en-tête) : 1re date = date de l'opération, description = texte le plus long,
        # dernier nombre = solde ; le sens se déduit de la variation du solde
        em = emetteur
        solde, col_sens = None, {}
        brutes = []
        for r in corps:
            di = next((i for i, c in enumerate(r) if date_libre(c, an, mo, ordre)), None)
            if di is None:
                continue
            d = date_libre(r[di], an, mo, ordre)
            textes = [(i, c) for i, c in enumerate(r) if i > di and re.search(r"[A-Za-zÀ-ÿ]{3}", c) and nombre(c) is None]
            if not textes:
                continue
            idesc = max(textes, key=lambda x: len(x[1]))[0]
            nums = [(i, nombre(c)) for i, c in enumerate(r) if i > idesc and nombre(c) is not None]
            if not nums:
                continue
            if any(re.fullmatch(r"(EOP|EOR|ES|EP|ET|CC|PR|MC|VD|VISA)", c) for c in r[:di]):
                em = em or ("desj_credit" if any(c in ("VISA", "MC") for c in r[:di]) else "desj_compte")
            brutes.append((d, r[idesc], nums))
        for d, desc, nums in brutes:
            if len(nums) >= 2:
                (ci, v), nouveau = nums[0], nums[-1][1]
                sens = None
                if solde is not None:
                    if abs(solde - abs(v) - nouveau) < 0.011:
                        sens = "D"
                    elif abs(solde + abs(v) - nouveau) < 0.011:
                        sens = "C"
                if sens:
                    col_sens[ci] = sens
                solde = nouveau
                brutes_sens = sens or col_sens.get(ci)
            else:
                ci, v = nums[0]
                brutes_sens = col_sens.get(ci)
            if brutes_sens is None:
                brutes_sens = "C" if v < 0 or RX_CREDIT_COMPTE.search(norm(desc)) else "D"
            txs.append({"date": d, "desc_brute": desc, "montant": abs(v), "sens": brutes_sens, "_col": ci})
        # 1re ligne (solde précédent inconnu) : sens d'après sa colonne, apprise sur les lignes suivantes
        for t in txs:
            c = t.pop("_col", None)
            if c in col_sens:
                t["sens"] = col_sens[c]
        if txs:
            notes.append("Fichier sans en-tête : lu comme un export AccèsD (sens d'après le solde).")
    return em, txs, notes


# ---------------------------------------------------------------------------- lecture d'un fichier
def lire(fichier, cle_mois, suite=None):
    """fichier : {nom, type: csv|xlsx|pdf|mots, texte|b64|pages, emetteur?} → transactions normalisées."""
    an, mo = (int(x) for x in cle_mois.split("-"))
    em_force = fichier.get("emetteur") or None
    notes, periode = [], None
    t = fichier.get("type")
    try:
        if t == "csv":
            em, txs, notes = depuis_tableau(lignes_csv(fichier.get("texte") or ""), an, mo, em_force)
        elif t == "xlsx":
            em, txs, notes = depuis_tableau(lignes_xlsx(fichier.get("b64") or ""), an, mo, em_force)
        elif t == "pdf":
            pages = pdf_pages(fichier.get("b64") or "")
            if not any(p["mots"] for p in pages):
                return {"nom": masque(fichier.get("nom")), "emetteur": em_force, "txs": [], "periode": None,
                        "erreur": "Aucun texte dans ce PDF (relevé numérisé ?) : utilise plutôt l'export CSV ou des captures."}
            em, txs, periode, notes = depuis_mots(pages, an, mo, em_force or None, None, True)
        elif t == "mots":
            em, txs, periode, notes = depuis_mots(fichier.get("pages") or [{}], an, mo, em_force, suite)
        else:
            return {"nom": masque(fichier.get("nom")), "emetteur": em_force, "txs": [], "periode": None, "erreur": "Format non reconnu."}
    except ValueError as e:
        return {"nom": masque(fichier.get("nom")), "emetteur": em_force, "txs": [], "periode": None, "erreur": str(e)}
    nom = norm(fichier.get("nom") or "")
    if not em:
        em = "amex" if re.search(r"amex|american|cobalt", nom) else ("desj_credit" if re.search(r"visa|credit|mastercard", nom) else
                                                                    ("desj_compte" if re.search(r"desjardins|compte|releve|accesd|eop", nom) else None))
    for k, x in enumerate(txs):
        x["montant"] = round(x["montant"], 2)
        x["desc_brute"] = masque(x["desc_brute"])
        x["desc"] = nettoie(x["desc_brute"])
    dates = sorted(x["date"] for x in txs if x.get("date"))
    lu = [dates[0], dates[-1]] if dates else None
    return {"nom": masque(fichier.get("nom")), "emetteur": em, "txs": txs, "periode": periode, "periode_tx": lu, "notes": notes,
            "deux_dates": any(x.get("date_insc") for x in txs)}


# ---------------------------------------------------------------------------- rapprochement
def _cle_desc(t):
    return re.sub(r"[^a-z]", "", norm(t))[:8]


def _mots(t):
    return {w for w in re.findall(r"[a-z]{4,}", norm(t)) if w not in VILLES and w not in ("montreal", "payment", "paiement")}


GENERIQUES = {"restaurant", "resto", "depanneur", "marche", "store", "magasin", "boutique", "inc", "ltee", "canada", "montreal",
              "payment", "paiement", "achat", "purchase", "online", "www", "com", "fictif", "fictive", "exemple", "test"}


def score_desc(a, b, cat=None):
    """0 = rien en commun ; 1 = même catégorie ou mot générique ; 2 = même nom de commerce ; 3 = même début."""
    if not a or not b:
        return 0
    ma, mb = _mots(a) - GENERIQUES, _mots(b) - GENERIQUES
    ka, kb = _cle_desc(a)[:5], _cle_desc(b)[:5]
    s = 0
    if ma & mb:
        s = 2
    if len(ka) >= 4 and ka == kb:
        s = 3
    if not s and (_mots(a) & _mots(b)):
        s = 1
    if not s and cat:
        ca, cb = cat(a), cat(b)
        if ca and cb and ca == cb and ca != "Autres":
            s = 1
    return s


def semblable(a, b, cat=None):
    if not a or not b:
        return False
    if _mots(a) & _mots(b):
        return True
    ka, kb = _cle_desc(a)[:5], _cle_desc(b)[:5]
    if len(ka) >= 4 and ka == kb:
        return True
    if cat:
        ca, cb = cat(a), cat(b)
        return bool(ca and cb and ca == cb and ca != "Autres")
    return False


RX_PAIE_CARTE = r"^paiement\b|payment|merci pour votre paiement|paiement recu"
RX_PAIE_COMPTE = r"paiement facture|paiement carte|visa desjardins|mastercard desjardins|american express|\bamex\b"
NATURES = [
    ("exclu_auth", r"temp auth|auth hold|autorisation temp"),
    ("placement", r"fidelity|wealthsimple|questrade|disnat|courtage|placements? (en ligne|direct)|\bceli\b|\breer\b|\bceliapp\b"),
    ("avance_compte", r"avance fonds carte|avance de fonds carte"),
    ("avance_carte", r"avance d.argent|avance de fonds|cash advance"),
    ("annulation", r"^annulation"),
]


def nature(tx, partenaire, nom_famille, exclusions):
    d = norm(tx["desc_brute"])
    if tx["compte"] == "desj_compte":
        if tx["sens"] == "D" and re.search(RX_PAIE_COMPTE, d):
            return "paiement"
    elif tx["sens"] == "C" and re.search(RX_PAIE_CARTE, d):
        return "paiement"
    for nat, rx in NATURES:
        if re.search(rx, d):
            return nat
    if partenaire and norm(partenaire) in d and re.search(r"interac|virement|transfer", d):
        return "partenaire"
    if nom_famille and len(nom_famille) >= 3 and re.search(r"virement interac de|e-transfer from|virement de", d) and re.search(r"\b" + re.escape(norm(nom_famille)) + r"\b", d):
        return "famille"
    for e in exclusions or []:
        if e and len(e.strip()) >= 3 and norm(e.strip()) in d:
            return "exclu"
    return None


def fusionne(fichiers):
    """Relevés qui se chevauchent : une transaction ne compte qu'une fois. Pour chaque compte, chaque
    transaction (date, montant, sens, début de description) est comptée au plus le nombre de fois où elle
    apparaît dans UN même relevé."""
    par = {}
    for i, f in enumerate(fichiers):
        em = f.get("emetteur")
        if not em:
            continue
        compte = par.setdefault(em, {})
        vus = {}
        for t in f.get("txs") or []:
            k = (t["date"], round(t["montant"], 2), t["sens"])          # la description lue peut varier d'une capture à l'autre
            vus.setdefault(k, []).append(t)
        for k, l in vus.items():
            deja = compte.get(k, [])
            if len(l) > len(deja):
                compte[k] = [dict(x, compte=em, fichier=f.get("nom")) for x in l]
    res = []
    for em, compte in par.items():
        for l in compte.values():
            res.extend(l)
    res.sort(key=lambda t: (t["date"], t["compte"]))
    for k, t in enumerate(res):
        t["id"] = "b%d" % k
    return res


def couverture(fichiers, cle_mois):
    an, mo = (int(x) for x in cle_mois.split("-"))
    debut = datetime.date(an, mo, 1)
    fin = (datetime.date(an + (mo == 12), mo % 12 + 1, 1) - datetime.timedelta(days=1))
    res = {}
    for em in COMPTES:
        per = []
        for f in fichiers:
            if f.get("emetteur") != em:
                continue
            p = f.get("periode") or f.get("periode_tx")
            if p:
                per.append((jour(p[0]), jour(p[1]), bool(f.get("periode"))))
        if not per:
            continue
        per.sort()
        manque, cur = [], debut
        for a, b, _ in per:
            if a > cur and a <= fin:
                manque.append((cur, min(fin, a - datetime.timedelta(days=1))))
            if b + datetime.timedelta(days=1) > cur:
                cur = b + datetime.timedelta(days=1)
        if cur <= fin:
            manque.append((cur, fin))
        ecrite = all(x[2] for x in per)
        res[em] = {"periodes": [[a.isoformat(), b.isoformat()] for a, b, _ in per],
                   "manque": [[a.isoformat(), b.isoformat()] for a, b in manque], "periode_ecrite": ecrite,
                   "fin": max(b for _, b, _ in per).isoformat()}
    return res


def analyse(fichiers, cle_mois, app, prec=None, options=None):
    """fichiers : résultats de lire() (émetteur éventuellement corrigé par l'utilisateur) ;
    app : modele() du mois ; prec : modele() du mois précédent (revenus) ; options : partenaire, nom_famille,
    exclusions, categorie (fonction), modifiable."""
    o = options or {}
    cat = o.get("categorie")
    part = o.get("partenaire") or ""
    an, mo = (int(x) for x in cle_mois.split("-"))
    banque = fusionne(fichiers)
    dans = lambda d: bool(d) and d[:7] == cle_mois
    cov = couverture(fichiers, cle_mois)
    groupes = {k: [] for k in ("concordantes", "manquantes", "absentes", "ecarts", "exclues", "normales", "revenus_exercice")}

    def ligne_b(t, **kw):
        x = {"id": t["id"], "source": "releve", "compte": t["compte"], "date": t["date"], "desc": t["desc"], "desc_brute": t["desc_brute"],
             "montant": t["montant"], "sens": t["sens"], "fichier": t.get("fichier"), "date_insc": t.get("date_insc"), "note_date": t.get("note_date")}
        x.update(kw)
        return x

    # ---------- côté app
    rx_part = re.compile(r"\(carte " + re.escape(part) + r"\b", re.I) if part else None
    A, normales_app = [], []
    for t in app.get("transactions") or []:
        if not t.get("utilise") or not (t.get("deboursement") or 0) > 0:
            continue
        x = {"id": "t%d" % t["ligne"], "date": t.get("date"), "montant": round(t["deboursement"], 2), "desc": t.get("description") or "",
             "type": t.get("type") or "", "ligne": t["ligne"]}
        if t.get("a_rembourser") or (rx_part and rx_part.search(x["desc"])):
            normales_app.append(x)
        else:
            A.append(x)
    for r in app.get("reports") or []:
        if r.get("montant") and not r.get("auto") and r["montant"] > 0:
            A.append({"id": "r%d" % r["ligne"], "date": None, "montant": round(abs(r["montant"]), 2), "desc": r.get("libelle") or "Autre montant", "type": "report"})
    V = []
    for r in app.get("revenus") or []:
        if r.get("utilise") and r.get("montant"):
            ms = [round(r["montant"], 2)] + [float(x) for x in re.findall(r"(?<![\d.])\d+\.\d{2}(?![\d])", str(r.get("montant_txt") or ""))[:1]]
            V.append({"id": "v%d" % r["ligne"], "date": r.get("date"), "montants": ms, "montant": round(r["montant"], 2), "desc": r.get("description") or "Revenu"})
    Vprec = []
    for r in (prec or {}).get("revenus") or []:
        if r.get("utilise") and r.get("montant"):
            Vprec.append({"date": r.get("date"), "montant": round(r["montant"], 2), "desc": r.get("description") or "Revenu"})
    retraits = [x for x in app.get("retraits") or [] if (x.get("montant") or 0) > 0]

    # ---------- côté relevé : natures (exclusions, cas normaux)
    B, Cr, R, avances, places = [], [], [], [], []
    annul = [t for t in banque if dans(t["date"]) and nature(t, part, o.get("nom_famille"), o.get("exclusions")) == "annulation"]
    for t in banque:
        if not dans(t["date"]):
            continue
        n = nature(t, part, o.get("nom_famille"), o.get("exclusions"))
        if n == "paiement":
            groupes["exclues"].append(ligne_b(t, raison="Paiement de carte de crédit (jamais inscrit dans l'Excel)"))
        elif n == "avance_compte":
            groupes["exclues"].append(ligne_b(t, raison="Avance de fonds reçue de ta carte (transfert interne)"))
        elif n == "famille":
            groupes["exclues"].append(ligne_b(t, raison="Virement familial (non inscrit)"))
        elif n == "exclu":
            groupes["exclues"].append(ligne_b(t, raison="Exclu (mot-clé de tes exclusions)"))
        elif n == "partenaire":
            groupes["normales"].append(ligne_b(t, raison="Virement %s %s : règlement du split" % ("de" if t["sens"] == "C" else "à", part)))
        elif n == "annulation":
            groupes["exclues"].append(ligne_b(t, raison="Opération annulée"))
        elif n == "exclu_auth":
            groupes["exclues"].append(ligne_b(t, raison="Autorisation temporaire (sera annulée)"))
        elif n == "avance_carte":
            avances.append(t)
        elif n == "placement" and t["sens"] == "D":
            places.append(t)
        elif t["sens"] == "D":
            B.append(t)
        elif t["compte"] == "desj_compte":
            R.append(t)
        else:
            Cr.append(t)
    # retrait annulé : la paire (retrait + annulation du même montant) est exclue
    for a in annul:
        x = next((t for t in B if abs(t["montant"] - a["montant"]) < 0.005 and abs((jour(t["date"]) - jour(a["date"])).days) <= 5 and
                  _mots(t["desc_brute"]) & _mots(a["desc_brute"])), None)
        if x:
            B.remove(x)
            groupes["exclues"].append(ligne_b(x, raison="Opération annulée ensuite (%s)" % fr_date(a["date"])))

    # ---------- dépenses : 1 pour 1, regroupées, nettes, écarts
    Adans = [a for a in A if a["date"] is None or dans(a["date"])]
    ecart_j = lambda a, t: abs((jour(a["date"]) - jour(t["date"])).days) if a.get("date") and t.get("date") else 0
    paires = []
    for t in sorted(B, key=lambda t: t["date"]):
        cands = [a for a in Adans if abs(a["montant"] - t["montant"]) < 0.005 and ecart_j(a, t) <= 3]
        if cands:
            a = min(cands, key=lambda a: (not semblable(a["desc"], t["desc_brute"], cat), a["date"] is None, ecart_j(a, t)))
            Adans.remove(a)
            paires.append((t, a))
    for t, a in paires:
        B.remove(t)
        groupes["concordantes"].append(ligne_b(t, app={"id": a["id"], "desc": a["desc"], "montant": a["montant"], "date": a["date"], "type": a["type"]}))
    for a in sorted([a for a in Adans if a["date"]], key=lambda a: -a["montant"]):
        pool = [t for t in B if ecart_j(a, t) <= 3][:14]
        trouve = None
        for n in (2, 3, 4):
            for combo in itertools.combinations(pool, n):
                if abs(sum(t["montant"] for t in combo) - a["montant"]) < 0.015:
                    trouve = combo
                    break
            if trouve:
                break
        if not trouve:
            # montant net (achats − crédits du même commerçant), ex. GGPoker
            cle = _cle_desc(a["desc"])[:5]
            pool2 = [t for t in B + Cr if ecart_j(a, t) <= 3 and len(cle) >= 4 and _cle_desc(t["desc_brute"])[:5] == cle][:12]
            if any(t["sens"] == "C" for t in pool2):
                for k in range(2, len(pool2) + 1):
                    for combo in itertools.combinations(pool2, k):
                        if abs(sum(t["montant"] * (1 if t["sens"] == "D" else -1) for t in combo) - a["montant"]) < 0.015:
                            trouve = combo
                            break
                    if trouve:
                        break
        if trouve:
            Adans.remove(a)
            for t in trouve:
                (B if t in B else Cr).remove(t)
                groupes["concordantes"].append(ligne_b(t, app={"id": a["id"], "desc": a["desc"], "montant": a["montant"], "date": a["date"], "type": a["type"]},
                                                       note="Regroupé dans l'app : « %s » %s" % (a["desc"], ("%.2f $" % a["montant"]).replace(".", ","))))
    # écarts : même commerce, montant ou date différents (paires les plus sûres d'abord)
    def candidat(a, t, hors_mois=False):
        sc = score_desc(a["desc"], t["desc_brute"], cat)
        d, j = abs(a["montant"] - t["montant"]), ecart_j(a, t)
        if d < 0.005:
            return sc >= (1 if hors_mois or j <= 10 else 9) and (j <= 3 if hors_mois else 3 < j <= 10)
        if j > 3 or hors_mois:
            return False
        if sc >= 2:
            return d <= max(10.0, 0.5 * t["montant"])
        if sc == 1:
            return d <= max(2.0, 0.1 * t["montant"])
        return False
    paires = sorted(((score_desc(a["desc"], t["desc_brute"], cat), -abs(a["montant"] - t["montant"]), -ecart_j(a, t), k1, k2)
                     for k1, t in enumerate(B) for k2, a in enumerate(Adans) if a.get("date") and candidat(a, t)), reverse=True)
    pris_b, pris_a = set(), set()
    for _, _, _, k1, k2 in paires:
        if k1 in pris_b or k2 in pris_a:
            continue
        pris_b.add(k1); pris_a.add(k2)
        t, a = B[k1], Adans[k2]
        if abs(a["montant"] - t["montant"]) < 0.005:
            quoi = "Date différente : %s au relevé, %s dans l'app" % (fr_date(t["date"]), fr_date(a["date"]))
        else:
            quoi = "Montant différent : %s au relevé, %s dans l'app" % (("%.2f $" % t["montant"]).replace(".", ","), ("%.2f $" % a["montant"]).replace(".", ","))
        groupes["ecarts"].append(ligne_b(t, app={"id": a["id"], "desc": a["desc"], "montant": a["montant"], "date": a["date"], "type": a["type"]}, raison=quoi))
    B = [t for k, t in enumerate(B) if k not in pris_b]
    Adans = [a for k, a in enumerate(Adans) if k not in pris_a]
    # dépense de l'app dont la transaction est datée du mois voisin (ex. achat du 30 septembre inscrit le 1er octobre)
    voisins = [t for t in banque if not dans(t["date"]) and t["sens"] == "D" and not nature(t, part, o.get("nom_famille"), o.get("exclusions"))]
    for a in list(Adans):
        if not a.get("date"):
            continue
        t = next((t for t in voisins if candidat(a, t, True)), None)
        if t:
            voisins.remove(t)
            Adans.remove(a)
            groupes["ecarts"].append(ligne_b(t, app={"id": a["id"], "desc": a["desc"], "montant": a["montant"], "date": a["date"], "type": a["type"]},
                                             raison="Transaction datée du %s au relevé (autre mois), %s dans l'app" % (fr_date(t["date"]), fr_date(a["date"])), hors_mois=True))
    for t in B:
        groupes["manquantes"].append(ligne_b(t, genre="depense"))
    for t in Cr:
        groupes["manquantes"].append(ligne_b(t, genre="credit", raison="Crédit sur la carte (remboursement ?) : à expliquer"))
    for a in Adans:
        if a["date"] is None:
            groupes["absentes"].append({"id": a["id"], "source": "app", "date": None, "desc": a["desc"], "montant": a["montant"], "type": a["type"],
                                        "raison": "Autre montant (tableau Remboursement), sans date"})
        else:
            groupes["absentes"].append({"id": a["id"], "source": "app", "date": a["date"], "desc": a["desc"], "montant": a["montant"], "type": a["type"]})
    for x in normales_app:
        if dans(x["date"]):
            groupes["normales"].append({"id": x["id"], "source": "app", "date": x["date"], "desc": x["desc"], "montant": x["montant"],
                                        "raison": "Payé par %s avec sa carte : pas dans tes relevés" % part})

    # ---------- placements (épargne du mois)
    tot_ret = round(sum(x["montant"] for x in retraits), 2)
    restant = list(retraits)
    for t in places:
        x = next((r for r in restant if abs(r["montant"] - t["montant"]) < 0.005), None)
        if x:
            restant.remove(x)
            groupes["concordantes"].append(ligne_b(t, app={"id": "epargne", "desc": x.get("libelle", "Épargne").replace("(-) ", ""), "montant": x["montant"], "date": None, "type": "épargne"}))
        else:
            groupes["ecarts"].append(ligne_b(t, raison="Transfert vers tes placements : %s d'épargne inscrite dans l'app ce mois-ci" % ("%.2f $" % tot_ret).replace(".", ",")))

    # ---------- revenus (comptabilité d'exercice : le tableau Revenus du mois contient ce qui est GAGNÉ
    # dans le mois, daté du jour où l'argent arrive — parfois au début du mois suivant)
    Vdispo = list(V)
    fin_mois = (datetime.date(an + (mo == 12), mo % 12 + 1, 1) - datetime.timedelta(days=1)).isoformat()
    poker = [v for v in Vdispo if "poker" in norm(v["desc"])]
    for t in avances:
        v = next((v for v in poker if any(abs(m - t["montant"]) < 0.005 for m in v["montants"]) and (not v["date"] or abs((jour(v["date"]) - jour(t["date"])).days) <= 3)), None)
        if v:
            poker.remove(v)
            Vdispo.remove(v)
            groupes["concordantes"].append(ligne_b(t, app={"id": v["id"], "desc": v["desc"], "montant": v["montant"], "date": v["date"], "type": "revenu"}, note="Avance poker = « Gain Poker »"))
        else:
            groupes["manquantes"].append(ligne_b(t, genre="avance", raison="Avance d'argent : « Gain Poker » non trouvé dans tes revenus"))
    jv = lambda v, t: abs((jour(v["date"]) - jour(t["date"])).days) if v.get("date") else 0
    Rdispo = list(R)
    for t in sorted(R, key=lambda t: t["date"]):
        cands = [v for v in Vdispo if any(abs(m - t["montant"]) < 0.005 for m in v["montants"]) and jv(v, t) <= 6]
        if cands:
            v = min(cands, key=lambda v: (jv(v, t), not semblable(v["desc"], t["desc_brute"])))
            Vdispo.remove(v)
            Rdispo.remove(t)
            groupes["concordantes"].append(ligne_b(t, app={"id": v["id"], "desc": v["desc"], "montant": v["montant"], "date": v["date"], "type": "revenu"}))
    for v in sorted([v for v in Vdispo if v["date"]], key=lambda v: -v["montant"]):
        pool = [t for t in Rdispo if jv(v, t) <= 6][:12]
        trouve = None
        for n in (2, 3):
            for combo in itertools.combinations(pool, n):
                if abs(sum(t["montant"] for t in combo) - v["montant"]) < 0.015:
                    trouve = combo
                    break
            if trouve:
                break
        if trouve:
            Vdispo.remove(v)
            for t in trouve:
                Rdispo.remove(t)
                groupes["concordantes"].append(ligne_b(t, app={"id": v["id"], "desc": v["desc"], "montant": v["montant"], "date": v["date"], "type": "revenu"},
                                                       note="Regroupé dans l'app : « %s »" % v["desc"]))
    # revenus de l'app sans dépôt dans le mois : reçus au début du mois suivant, à recevoir, ou absents
    apres = [t for t in banque if t["compte"] == "desj_compte" and t["sens"] == "C" and t["date"] > fin_mois]
    fin_compte = (cov.get("desj_compte") or {}).get("fin")
    for v in Vdispo:
        base = {"id": v["id"], "source": "app", "date": v["date"], "desc": v["desc"], "montant": v["montant"], "genre": "revenu"}
        t = next((t for t in apres if any(abs(m - t["montant"]) < 0.005 for m in v["montants"]) and (not v["date"] or jv(v, t) <= 6)), None)
        if t:
            apres.remove(t)
            groupes["revenus_exercice"].append(dict(base, statut="recu", raison="Reçu le %s (gagné ce mois-ci)" % fr_date(t["date"]), releve=ligne_b(t)))
        elif v["date"] and (not fin_compte or v["date"] > fin_compte):
            groupes["revenus_exercice"].append(dict(base, statut="a_recevoir", raison="À recevoir — attendu le %s" % fr_date(v["date"])))
        elif not fin_compte:
            groupes["revenus_exercice"].append(dict(base, statut="a_recevoir", raison="À recevoir"))
        else:
            groupes["absentes"].append(dict(base, raison="Revenu de l'app sans dépôt au relevé"))
    # dépôts du début du mois qui appartiennent au mois précédent (gagnés avant, inscrits dans le mois précédent)
    Vprec_ici = [v for v in Vprec if not v.get("date") or v["date"][:7] == cle_mois]
    for t in Rdispo:
        vp = next((v for v in Vprec_ici if abs(v["montant"] - t["montant"]) < 0.005 and (not v.get("date") or abs((jour(v["date"]) - jour(t["date"])).days) <= 6)), None)
        if vp:
            Vprec_ici.remove(vp)
            groupes["revenus_exercice"].append(ligne_b(t, statut="mois_precedent", genre="revenu",
                                                       raison="Revenu du mois précédent (« %s », inscrit dans le mois précédent)" % vp["desc"]))
        else:
            groupes["manquantes"].append(ligne_b(t, genre="revenu", raison="Entrée d'argent absente de tes revenus"))

    # ---------- sommaire par compte
    som = {}
    for em, nom in COMPTES.items():
        if em not in cov and not any(t["compte"] == em for t in banque):
            continue
        lst = lambda g: [x for x in groupes[g] if x.get("compte") == em and x.get("source") == "releve"]
        dep = [t for t in banque if t["compte"] == em and dans(t["date"]) and t["sens"] == "D"]
        exclus_ids = {x["id"] for x in groupes["exclues"] + groupes["normales"]}
        dep_compte = round(sum(t["montant"] for t in dep if t["id"] not in exclus_ids), 2)
        app_ok = round(sum(x["montant"] for x in lst("concordantes") if x["sens"] == "D") + sum(x["app"]["montant"] for x in lst("ecarts") if x.get("app") and x["sens"] == "D" and not x.get("hors_mois")), 2)
        som[em] = {"nom": nom, "total_releve": dep_compte, "total_app": app_ok, "ecart": round(dep_compte - app_ok, 2),
                   "n": {g: len(lst(g)) for g in ("concordantes", "manquantes", "ecarts", "exclues", "normales")},
                   "couverture": cov.get(em)}
    som_app = {"absentes": len(groupes["absentes"])}
    return {"mois": cle_mois, "comptes": som, "groupes": groupes, "app": som_app,
            "modifiable": bool(o.get("modifiable")), "n_releve": sum(1 for t in banque if dans(t["date"]))}
