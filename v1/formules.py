"""Outils sur les formules Excel : décalage de références et petit évaluateur.

Aucune dépendance externe (bibliothèque standard Python uniquement).
"""
import re

# ---------------------------------------------------------------------------
# Adresses de cellules
# ---------------------------------------------------------------------------

def col_en_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_en_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def decoupe_adresse(adr):
    m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", adr)
    return m.group(1), int(m.group(2))


# Référence : préfixe de feuille optionnel, puis A1 ou $A$1, puis :B2 optionnel
_FEUILLE = r"(?:'((?:[^']|'')+)'|([A-Za-z_À-ſ][\wÀ-ſ.]*))!"
_CELL = r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
_REF_RE = re.compile(r"(?:" + _FEUILLE + r")?" + _CELL + r"(?::" + _CELL + r")?(?![\w(!])")


def _iter_hors_chaines(formule):
    """Découpe la formule en segments (texte, est_chaine)."""
    i, n, debut = 0, len(formule), 0
    while i < n:
        if formule[i] == '"':
            if i > debut:
                yield formule[debut:i], False
            j = i + 1
            while j < n:
                if formule[j] == '"':
                    if j + 1 < n and formule[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            yield formule[i:j + 1], True
            i = debut = j + 1
        else:
            i += 1
    if debut < n:
        yield formule[debut:], False


def transforme_refs(formule, fn):
    """Applique fn(feuille, (c1, r1, abs_c1, abs_r1), fin_ou_None) -> texte à chaque référence.

    feuille = nom de feuille (None si aucune). fn renvoie None pour laisser intact.
    """
    out = []
    for seg, est_chaine in _iter_hors_chaines(formule):
        if est_chaine:
            out.append(seg)
            continue
        pos = 0
        res = []
        for m in _REF_RE.finditer(seg):
            s = m.start()
            if s > 0 and (seg[s - 1].isalnum() or seg[s - 1] in "_.$"):
                continue
            feuille = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
            debut = (m.group(4), int(m.group(6)), m.group(3) == "$", m.group(5) == "$")
            fin = None
            if m.group(8):
                fin = (m.group(8), int(m.group(10)), m.group(7) == "$", m.group(9) == "$")
            nouveau = fn(feuille, debut, fin)
            if nouveau is None:
                continue
            res.append(seg[pos:s])
            res.append(nouveau)
            pos = m.end()
        res.append(seg[pos:])
        out.append("".join(res))
    return "".join(out)


def _fmt_cell(c, r, ac, ar):
    return ("$" if ac else "") + c + ("$" if ar else "") + str(r)


def prefixe_feuille(nom):
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", nom):
        return nom + "!"
    return "'" + nom.replace("'", "''") + "'!"


def _rebatit(feuille_txt, debut, fin):
    s = _fmt_cell(*debut)
    if fin:
        s += ":" + _fmt_cell(*fin)
    return (prefixe_feuille(feuille_txt) if feuille_txt else "") + s


def copie_formule(formule, d_lig, d_col=0):
    """Décalage « recopie » : seules les parties relatives bougent."""
    def fn(feuille, debut, fin):
        def dec(p):
            c, r, ac, ar = p
            if not ac:
                c = num_en_col(col_en_num(c) + d_col)
            if not ar:
                r = r + d_lig
            return (c, r, ac, ar)
        return _rebatit(feuille, dec(debut), dec(fin) if fin else None)
    return transforme_refs(formule, fn)


def insere_lignes_formule(formule, feuille_cible, feuille_courante, a_partir, n):
    """Décalage « insertion de n lignes à la ligne a_partir » dans feuille_cible.

    feuille_courante = feuille où se trouve la formule (pour les références sans préfixe).
    """
    def fn(feuille, debut, fin):
        f = feuille if feuille is not None else feuille_courante
        if f != feuille_cible:
            return None
        def dec(p):
            c, r, ac, ar = p
            return (c, r + n if r >= a_partir else r, ac, ar)
        return _rebatit(feuille, dec(debut), dec(fin) if fin else None)
    return transforme_refs(formule, fn)


def renomme_feuille_formule(formule, ancien, nouveau):
    def fn(feuille, debut, fin):
        if feuille != ancien:
            return None
        return _rebatit(nouveau, debut, fin)
    return transforme_refs(formule, fn)


def decale_plage_txt(sqref, a_partir, n):
    """Décale une liste de plages « A1:B2 C3 » (sqref) pour une insertion de lignes."""
    morceaux = []
    for p in sqref.split():
        morceaux.append(insere_lignes_formule(p, "__x__", "__x__", a_partir, n))
    return " ".join(morceaux)


# ---------------------------------------------------------------------------
# Évaluateur de formules (sous-ensemble)
# ---------------------------------------------------------------------------

class NonSupporte(Exception):
    pass


class ErreurExcel(Exception):
    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code


_TOKEN_RE = re.compile(r"""
    (?P<espace>\s+)
  | (?P<chaine>"(?:[^"]|"")*")
  | (?P<ref>(?:'(?:[^']|'')+'!|[A-Za-z_À-ſ][\wÀ-ſ.]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![\w(]))
  | (?P<nombre>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|\.\d+)
  | (?P<fonction>[A-Za-z_][\w.]*\()
  | (?P<bool>TRUE|FALSE)
  | (?P<op><=|>=|<>|[-+*/^&=<>%(),:;{}])
""", re.X)


def _tokens(f):
    pos, out = 0, []
    while pos < len(f):
        m = _TOKEN_RE.match(f, pos)
        if not m:
            raise NonSupporte("jeton inconnu: " + f[pos:pos + 10])
        pos = m.end()
        k = m.lastgroup
        if k == "espace":
            continue
        out.append((k, m.group(k)))
    return out


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise ErreurExcel("#VALUE!")


def _aplatit(v):
    if isinstance(v, list):
        for x in v:
            yield from _aplatit(x)
    else:
        yield v


def _wild(motif):
    rx = ""
    i = 0
    while i < len(motif):
        ch = motif[i]
        if ch == "~" and i + 1 < len(motif):
            rx += re.escape(motif[i + 1]); i += 2; continue
        rx += ".*" if ch == "*" else "." if ch == "?" else re.escape(ch)
        i += 1
    return re.compile("^" + rx + "$", re.I | re.S)


class Evaluateur:
    """Évalue une formule. `lire(feuille, adresse)` fournit la valeur d'une cellule."""

    def __init__(self, lire, feuille):
        self.lire = lire
        self.feuille = feuille

    def evalue(self, formule):
        if formule.startswith("="):
            formule = formule[1:]
        self.t = _tokens(formule)
        self.i = 0
        v = self._comparaison()
        if self.i != len(self.t):
            raise NonSupporte("reste: %r" % (self.t[self.i:],))
        if isinstance(v, list):
            v = next(_aplatit(v), None)
        return v

    # --- analyse ---
    def _voir(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def _prendre(self, val=None):
        k, v = self._voir()
        if val is not None and v != val:
            raise NonSupporte("attendu %s" % val)
        self.i += 1
        return k, v

    def _comparaison(self):
        g = self._concat()
        while self._voir()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self._prendre()[1]
            d = self._concat()
            g = self._compare(g, op, d)
        return g

    def _compare(self, a, op, b):
        if isinstance(a, str) or isinstance(b, str):
            a2 = "" if a is None else str(a).lower() if not isinstance(a, (int, float)) else a
            b2 = "" if b is None else str(b).lower() if not isinstance(b, (int, float)) else b
            if type(a2) != type(b2):
                # en Excel, texte > nombre
                a2, b2 = (1 if isinstance(a2, str) else 0), (1 if isinstance(b2, str) else 0)
        else:
            a2, b2 = _num(a), _num(b)
        return {"=": a2 == b2, "<>": a2 != b2, "<": a2 < b2, ">": a2 > b2,
                "<=": a2 <= b2, ">=": a2 >= b2}[op]

    def _concat(self):
        g = self._addition()
        while self._voir()[1] == "&":
            self._prendre()
            d = self._addition()
            g = self._txt(g) + self._txt(d)
        return g

    def _txt(self, v):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "VRAI" if v else "FAUX"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    def _addition(self):
        g = self._mult()
        while self._voir()[1] in ("+", "-"):
            op = self._prendre()[1]
            d = self._mult()
            g = _num(g) + _num(d) if op == "+" else _num(g) - _num(d)
        return g

    def _mult(self):
        g = self._puissance()
        while self._voir()[1] in ("*", "/"):
            op = self._prendre()[1]
            d = self._puissance()
            if op == "*":
                g = _num(g) * _num(d)
            else:
                if _num(d) == 0:
                    raise ErreurExcel("#DIV/0!")
                g = _num(g) / _num(d)
        return g

    def _puissance(self):
        g = self._unaire()
        while self._voir()[1] == "^":
            self._prendre()
            g = _num(g) ** _num(self._unaire())
        return g

    def _unaire(self):
        if self._voir()[1] == "-":
            self._prendre()
            return -_num(self._unaire())
        if self._voir()[1] == "+":
            self._prendre()
            return self._unaire()
        v = self._primaire()
        while self._voir()[1] == "%":
            self._prendre()
            v = _num(v) / 100
        return v

    def _primaire(self):
        k, v = self._voir()
        if k == "nombre":
            self._prendre()
            return float(v)
        if k == "chaine":
            self._prendre()
            return v[1:-1].replace('""', '"')
        if k == "bool":
            self._prendre()
            return v == "TRUE"
        if k == "ref":
            self._prendre()
            return self._ref(v)
        if k == "fonction":
            self._prendre()
            return self._fonction(v[:-1].upper())
        if v == "(":
            self._prendre()
            r = self._comparaison()
            self._prendre(")")
            return r
        raise NonSupporte("primaire %r" % v)

    def _ref(self, txt):
        feuille = self.feuille
        if "!" in txt:
            f, txt = txt.rsplit("!", 1)
            feuille = f[1:-1].replace("''", "'") if f.startswith("'") else f
        txt = txt.replace("$", "")
        if ":" in txt:
            a, b = txt.split(":")
            ca, ra = decoupe_adresse(a)
            cb, rb = decoupe_adresse(b)
            return [[self.lire(feuille, num_en_col(c) + str(r))
                     for c in range(col_en_num(ca), col_en_num(cb) + 1)]
                    for r in range(ra, rb + 1)]
        return self.lire(feuille, txt)

    def _args_bruts(self):
        args = []
        if self._voir()[1] == ")":
            self._prendre()
            return args
        while True:
            debut = self.i
            prof = 0
            while True:
                k, v = self._voir()
                if k is None:
                    raise NonSupporte("parenthèse")
                if v == "(" or k == "fonction":
                    prof += 1
                elif v == ")":
                    if prof == 0:
                        break
                    prof -= 1
                elif v in (",", ";") and prof == 0:
                    break
                self.i += 1
            args.append((debut, self.i))
            sep = self._prendre()[1]
            if sep == ")":
                return args

    def _eval_arg(self, borne):
        sauve_t, sauve_i = self.t, self.i
        sous = Evaluateur(self.lire, self.feuille)
        sous.t = self.t[borne[0]:borne[1]]
        sous.i = 0
        if not sous.t:
            return None
        v = sous._comparaison()
        if sous.i != len(sous.t):
            raise NonSupporte("arg")
        self.t, self.i = sauve_t, sauve_i
        return v

    def _fonction(self, nom):
        args = self._args_bruts()
        ev = lambda k: self._eval_arg(args[k])
        if nom == "IF":
            c = ev(0)
            if isinstance(c, list):
                raise NonSupporte("IF matriciel")
            vrai = bool(_num(c)) if not isinstance(c, str) else bool(c)
            if vrai:
                return ev(1) if len(args) > 1 else True
            return ev(2) if len(args) > 2 else False
        if nom == "IFERROR":
            try:
                return ev(0)
            except ErreurExcel:
                return ev(1)
        if nom in ("SUM", "MIN", "MAX"):
            vals = []
            for k in range(len(args)):
                v = ev(k)
                if isinstance(v, list):
                    vals += [x for x in _aplatit(v) if isinstance(x, (int, float)) and not isinstance(x, bool)]
                else:
                    vals.append(_num(v))
            if nom == "SUM":
                return sum(vals)
            return (min if nom == "MIN" else max)(vals) if vals else 0.0
        if nom == "ROUND":
            x, d = _num(ev(0)), int(_num(ev(1)))
            import decimal
            q = decimal.Decimal(1).scaleb(-d)
            return float(decimal.Decimal(repr(x)).quantize(q, rounding=decimal.ROUND_HALF_UP))
        if nom == "ABS":
            return abs(_num(ev(0)))
        if nom == "N":
            v = ev(0)
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
        if nom == "ISNUMBER":
            v = ev(0)
            return isinstance(v, (int, float)) and not isinstance(v, bool)
        if nom == "TRIM":
            return re.sub(r" +", " ", self._txt(ev(0)).strip(" "))
        if nom == "COUNTIF":
            plage, crit = ev(0), ev(1)
            rx = _wild(str(crit))
            return float(sum(1 for x in _aplatit(plage) if x is not None and rx.match(self._txt(x))))
        if nom in ("AND", "OR"):
            vals = [bool(_num(ev(k))) for k in range(len(args))]
            return all(vals) if nom == "AND" else any(vals)
        raise NonSupporte("fonction " + nom)


def deplace_refs_formule(formule, feuille_cible, feuille_courante, rect, d_lig, d_col=0):
    """Comme « Couper / Coller » dans Excel : les références (de feuille_cible) qui pointent
    entièrement à l'intérieur du rectangle (c1, r1, c2, r2) sont déplacées de (d_lig, d_col)."""
    c1, r1, c2, r2 = rect

    def dedans(p):
        c, r = col_en_num(p[0]), p[1]
        return c1 <= c <= c2 and r1 <= r <= r2

    def fn(feuille, debut, fin):
        f = feuille if feuille is not None else feuille_courante
        if f != feuille_cible:
            return None
        if not dedans(debut) or (fin is not None and not dedans(fin)):
            return None

        def dec(p):
            c, r, ac, ar = p
            return (num_en_col(col_en_num(c) + d_col), r + d_lig, ac, ar)
        return _rebatit(feuille, dec(debut), dec(fin) if fin else None)
    return transforme_refs(formule, fn)


def deplace_plage_txt(sqref, rect, d_lig, d_col=0):
    return " ".join(deplace_refs_formule(p, "__x__", "__x__", rect, d_lig, d_col) for p in sqref.split())
