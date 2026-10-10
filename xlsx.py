"""Lecture et écriture « chirurgicale » d'un fichier .xlsx.

Le fichier n'est jamais réécrit par une bibliothèque Excel : on modifie
directement le XML des seules cellules concernées. Tout le reste (graphiques,
listes déroulantes, mises en forme conditionnelles, formules dynamiques…)
est conservé à l'octet près.
"""
import datetime
import html
import os
import re
import shutil
import tempfile
import zipfile

from formules import (col_en_num, num_en_col, decoupe_adresse, copie_formule,
                      insere_lignes_formule, renomme_feuille_formule, decale_plage_txt,
                      prefixe_feuille, deplace_refs_formule, deplace_plage_txt)

EPOQUE = datetime.datetime(1899, 12, 30)

_CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)


def esc(t):
    return html.escape(t, quote=False).replace('"', "&quot;")


def unesc(t):
    return html.unescape(t)


def serie_en_date(x):
    return (EPOQUE + datetime.timedelta(days=float(x))).date()


def date_en_serie(d):
    return (datetime.datetime(d.year, d.month, d.day) - EPOQUE).days


class Cellule:
    __slots__ = ("adr", "col", "lig", "valeur", "formule", "style", "type", "xml", "partagee")

    def __init__(self):
        self.formule = None
        self.partagee = None


class Feuille:
    def __init__(self, classeur, nom, chemin):
        self.classeur = classeur
        self.nom = nom
        self.chemin = chemin
        self.analyse()

    @property
    def xml(self):
        return self.classeur.parties[self.chemin].decode("utf-8")

    @xml.setter
    def xml(self, v):
        self.classeur.parties[self.chemin] = v.encode("utf-8")

    def analyse(self):
        x = self.xml
        self.cellules = {}
        maitres = {}
        sst = self.classeur.chaines
        for m in _CELL_RE.finditer(x):
            c = Cellule()
            c.col, c.lig = m.group(1), int(m.group(2))
            c.adr = c.col + str(c.lig)
            attrs = m.group(3)
            inner = m.group(4) or ""
            c.xml = m.group(0)
            sm = re.search(r'\bs="(\d+)"', attrs)
            c.style = sm.group(1) if sm else None
            tm = re.search(r'\bt="(\w+)"', attrs)
            c.type = tm.group(1) if tm else "n"
            fm = re.search(r"<f([^>]*?)(?:/>|>(.*?)</f>)", inner, re.S)
            if fm:
                fa, ft = fm.group(1), fm.group(2)
                si = re.search(r'si="(\d+)"', fa)
                if 't="shared"' in fa and si:
                    if ft:
                        maitres[si.group(1)] = (c.col, c.lig, unesc(ft))
                        c.formule = unesc(ft)
                    else:
                        c.partagee = si.group(1)
                        c.formule = ""
                else:
                    c.formule = unesc(ft or "")
            vm = re.search(r"<v>(.*?)</v>", inner, re.S)
            v = unesc(vm.group(1)) if vm else None
            if c.type == "s" and v is not None:
                v = sst[int(v)]
            elif c.type == "inlineStr":
                v = "".join(unesc(t) for t in re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S))
            elif c.type == "b":
                v = v == "1"
            elif c.type == "e":
                v = ("#ERR", v)
            elif c.type in ("str",):
                v = v if v is not None else ""
            elif v is not None:
                try:
                    v = float(v)
                except ValueError:
                    pass
            c.valeur = v
            self.cellules[c.adr] = c
        for c in self.cellules.values():
            if c.partagee is not None and c.partagee in maitres:
                mc, ml, mf = maitres[c.partagee]
                c.formule = copie_formule(mf, c.lig - ml, col_en_num(c.col) - col_en_num(mc))

    # --- lecture ---
    def cell(self, adr):
        return self.cellules.get(adr)

    def val(self, adr):
        c = self.cellules.get(adr)
        return c.valeur if c else None

    def texte(self, adr):
        v = self.val(adr)
        if v is None or isinstance(v, tuple):
            return ""
        if isinstance(v, float):
            return str(v)
        return str(v)

    def max_lig(self):
        return max((c.lig for c in self.cellules.values()), default=0)

    def cherche(self, texte, cols=None, debut=1, fin=None, exact=True):
        """Première cellule dont le texte normalisé correspond."""
        cible = normalise(texte)
        res = []
        for c in self.cellules.values():
            if not isinstance(c.valeur, str) or c.formule:
                continue
            if cols and c.col not in cols:
                continue
            if c.lig < debut or (fin and c.lig > fin):
                continue
            t = normalise(c.valeur)
            if (t == cible) if exact else t.startswith(cible):
                res.append(c)
        res.sort(key=lambda c: (c.lig, col_en_num(c.col)))
        return res[0] if res else None

    # --- écriture ---
    def ecrit(self, adr, valeur=None, formule=None, vider_formule=True, style=None):
        """Écrit une valeur (texte, nombre, date) ou une formule dans une cellule,
        en conservant son style (ou en appliquant `style` s'il est fourni)."""
        col, lig = decoupe_adresse(adr)
        anc = self.cellules.get(adr)
        if style is None:
            style = anc.style if anc else None
        if anc is not None and anc.partagee is None and anc.formule and 't="shared"' in anc.xml and not vider_formule:
            raise ValueError("cellule maîtresse partagée")
        s_attr = ' s="%s"' % style if style else ""
        if formule is not None:
            f = formule[1:] if formule.startswith("=") else formule
            if isinstance(valeur, (int, float)) and not isinstance(valeur, bool):
                nouveau = '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (adr, s_attr, esc(f), repr(float(valeur)))
            else:
                nouveau = '<c r="%s"%s><f>%s</f></c>' % (adr, s_attr, esc(f))
        elif valeur is None or valeur == "":
            nouveau = '<c r="%s"%s/>' % (adr, s_attr)
        elif isinstance(valeur, bool):
            nouveau = '<c r="%s"%s t="b"><v>%d</v></c>' % (adr, s_attr, 1 if valeur else 0)
        elif isinstance(valeur, (datetime.date, datetime.datetime)):
            nouveau = '<c r="%s"%s><v>%d</v></c>' % (adr, s_attr, date_en_serie(valeur))
        elif isinstance(valeur, (int, float)):
            v = repr(float(valeur))
            if v.endswith(".0"):
                v = v[:-2]
            nouveau = '<c r="%s"%s><v>%s</v></c>' % (adr, s_attr, v)
        else:
            nouveau = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
                adr, s_attr, esc(str(valeur)))
        self.remplace_xml_cellule(adr, nouveau)

    def remplace_formules(self, nouvelles):
        """Plusieurs formules en une seule passe ({adr: formule sans « = »}, cellules existantes) : bien plus rapide
        que ecrit() cellule par cellule sur une grande feuille. Les suiveurs non visés d'une formule partagée sont matérialisés."""
        x = self.xml
        cibles = set(nouvelles)
        for adr in nouvelles:
            c = self.cellules[adr]
            if c.partagee is None and c.formule and 't="shared"' in c.xml:
                si = re.search(r'si="(\d+)"', c.xml).group(1)
                for d in self.cellules.values():
                    if d.partagee == si and d.adr not in cibles:
                        x = x.replace(d.xml, re.sub(r"<f[^>]*/>", "<f>%s</f>" % esc(d.formule), d.xml, count=1), 1)
        for adr, fo in nouvelles.items():
            c = self.cellules[adr]
            s_attr = ' s="%s"' % c.style if c.style else ""
            i = x.find(c.xml)
            if i < 0:
                raise RuntimeError("cellule introuvable " + adr)
            x = x[:i] + '<c r="%s"%s><f>%s</f></c>' % (adr, s_attr, esc(fo[1:] if fo.startswith("=") else fo)) + x[i + len(c.xml):]
        self.xml = x
        self.analyse()

    def ecrit_xml_brut(self, adr, xml_cellule):
        self.remplace_xml_cellule(adr, xml_cellule)

    def remplace_xml_cellule(self, adr, nouveau):
        col, lig = decoupe_adresse(adr)
        x = self.xml
        anc = self.cellules.get(adr)
        if anc is not None:
            # si c'est le maître d'une formule partagée, on matérialise ses suiveurs
            if anc.partagee is None and anc.formule and 't="shared"' in anc.xml:
                x = self._materialise_partage(x, anc)
            i = x.find(anc.xml)
            if i < 0:
                raise RuntimeError("cellule introuvable " + adr)
            x = x[:i] + nouveau + x[i + len(anc.xml):]
        else:
            m = re.search(r'<row r="%d"([^>]*?)(/>|>)' % lig, x)
            if m is None:
                # créer la ligne au bon endroit
                pos = None
                for rm in _ROW_RE.finditer(x):
                    if int(rm.group(1)) > lig:
                        pos = rm.start()
                        break
                ligne = '<row r="%d">%s</row>' % (lig, nouveau)
                if pos is None:
                    pos = x.find("</sheetData>")
                    if pos < 0:
                        x = x.replace("<sheetData/>", "<sheetData>" + ligne + "</sheetData>")
                    else:
                        x = x[:pos] + ligne + x[pos:]
                else:
                    x = x[:pos] + ligne + x[pos:]
            else:
                if m.group(2) == "/>":
                    x = x[:m.start()] + '<row r="%d"%s>%s</row>' % (lig, m.group(1), nouveau) + x[m.end():]
                else:
                    fin_ligne = x.find("</row>", m.end())
                    contenu = x[m.end():fin_ligne]
                    pos = len(contenu)
                    for cm in _CELL_RE.finditer(contenu):
                        if col_en_num(cm.group(1)) > col_en_num(col):
                            pos = cm.start()
                            break
                    abs_pos = m.end() + pos
                    x = x[:abs_pos] + nouveau + x[abs_pos:]
        self.xml = x
        self.analyse()

    def _materialise_partage(self, x, maitre):
        si = re.search(r'si="(\d+)"', maitre.xml).group(1)
        for c in list(self.cellules.values()):
            if c.partagee == si:
                nv = re.sub(r"<f[^>]*/>", "<f>%s</f>" % esc(c.formule), c.xml, count=1)
                x = x.replace(c.xml, nv, 1)
        return x

    def formule_modele(self, col, lig, lig_min, lig_max):
        """XML d'une cellule recopiée depuis la cellule à formule la plus proche
        de la même colonne (entre lig_min et lig_max), ajustée à la ligne lig."""
        meilleur = None
        for r in sorted(range(lig_min, lig_max + 1), key=lambda r: abs(r - lig)):
            c = self.cellules.get(col + str(r))
            if c is not None and c.formule and r != lig:
                meilleur = c
                break
        if meilleur is None:
            return None
        return recopie_xml_cellule(meilleur, lig)


def recopie_xml_cellule(c, lig, vider_valeur=True):
    """Construit le XML d'une copie de la cellule c placée à la ligne lig."""
    d = lig - c.lig
    adr = c.col + str(lig)
    attrs = re.match(r'<c r="[A-Z]+\d+"([^>]*?)(/?>)', c.xml).group(1)
    if c.formule:
        f = copie_formule(c.formule, d)
        fm = re.search(r"<f([^>]*?)(?:/>|>)", c.xml)
        fattrs = fm.group(1) if fm else ""
        if 't="array"' in fattrs:
            fx = '<f t="array" ref="%s">%s</f>' % (adr, esc(f))
        else:
            fx = "<f>%s</f>" % esc(f)
        return '<c r="%s"%s>%s<v></v></c>' % (adr, attrs, fx) if 't="str"' in attrs else \
            '<c r="%s"%s>%s</c>' % (adr, attrs, fx)
    attrs = re.sub(r'\s+t="\w+"', "", attrs)
    if vider_valeur:
        return '<c r="%s"%s/>' % (adr, attrs)
    return c.xml.replace('r="%s"' % c.adr, 'r="%s"' % adr, 1)


def normalise(t):
    t = str(t).replace(" ", " ").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", t).strip().lower()


class Classeur:
    def __init__(self, chemin):
        self.chemin = chemin
        with zipfile.ZipFile(chemin) as z:
            self.ordre = z.namelist()
            self.infos = {i.filename: i for i in z.infolist()}
            self.parties = {n: z.read(n) for n in self.ordre}
        self._lit_chaines()
        self._lit_feuilles()

    def _txt(self, nom):
        return self.parties[nom].decode("utf-8")

    def _ecrit_txt(self, nom, txt):
        if nom not in self.parties:
            self.ordre.append(nom)
        self.parties[nom] = txt.encode("utf-8")

    def _lit_chaines(self):
        self.chaines = []
        if "xl/sharedStrings.xml" in self.parties:
            x = self._txt("xl/sharedStrings.xml")
            for si in re.findall(r"<si>(.*?)</si>", x, re.S):
                # on ignore les annotations phonétiques
                si = re.sub(r"<rPh.*?</rPh>", "", si, flags=re.S)
                self.chaines.append("".join(unesc(t) for t in re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))

    def _lit_feuilles(self):
        wb = self._txt("xl/workbook.xml")
        rels = self._txt("xl/_rels/workbook.xml.rels")
        cibles = dict(re.findall(r'<Relationship [^>]*?Id="(\w+)"[^>]*?Target="([^"]+)"', rels))
        cibles.update({k: v for v, k in re.findall(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="(\w+)"', rels)})
        self.feuilles = []
        for nom, rid in re.findall(r'<sheet name="([^"]*)" sheetId="\d+"[^>]*?r:id="(\w+)"', wb):
            cible = cibles[rid]
            chemin = cible.lstrip("/") if cible.startswith("/") else "xl/" + cible
            self.feuilles.append(Feuille(self, unesc(nom), chemin))
        self.par_nom = {f.nom: f for f in self.feuilles}

    def feuille(self, nom):
        return self.par_nom[nom]

    # ------------------------------------------------------------------
    def enregistre(self, chemin=None):
        chemin = chemin or self.chemin
        # forcer Excel à tout recalculer à l'ouverture et retirer la chaîne de calcul
        wb = self._txt("xl/workbook.xml")
        if "<calcPr" in wb:
            if "fullCalcOnLoad" not in wb:
                wb = re.sub(r"<calcPr", '<calcPr fullCalcOnLoad="1"', wb, count=1)
        else:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self._ecrit_txt("xl/workbook.xml", wb)
        if "xl/calcChain.xml" in self.parties:
            del self.parties["xl/calcChain.xml"]
            self.ordre.remove("xl/calcChain.xml")
            ct = self._txt("[Content_Types].xml")
            ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
            self._ecrit_txt("[Content_Types].xml", ct)
            r = self._txt("xl/_rels/workbook.xml.rels")
            r = re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", r)
            self._ecrit_txt("xl/_rels/workbook.xml.rels", r)
        dossier = os.path.dirname(os.path.abspath(chemin))
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dossier)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
                for n in self.ordre:
                    if n in self.parties:
                        z.writestr(n, self.parties[n])
            os.replace(tmp, chemin)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    # ------------------------------------------------------------------
    # Insertion de lignes (comme « Insérer des lignes » dans Excel)
    # ------------------------------------------------------------------
    def insere_lignes(self, nom_feuille, a_partir, n, lig_modele=None, cols_modele=None):
        """Insère n lignes vides avant la ligne `a_partir` et y recopie les
        formules de la ligne `lig_modele` (par défaut la ligne au-dessus)."""
        f = self.feuille(nom_feuille)
        lig_modele = lig_modele or a_partir - 1
        # matérialiser les formules partagées (plus simple et sûr à décaler)
        x = f.xml
        for c in f.cellules.values():
            if c.partagee is not None:
                nv = re.sub(r"<f[^>]*/>", "<f>%s</f>" % esc(c.formule), c.xml, count=1)
                x = x.replace(c.xml, nv, 1)
        x = re.sub(r'<f t="shared" ref="[^"]*" si="\d+">', "<f>", x)
        x = re.sub(r'<f t="shared" si="\d+" ref="[^"]*">', "<f>", x)
        f.xml = x
        f.analyse()
        mrow = re.search(r'<row r="%d"([^>]*?)(?:/>|>)' % lig_modele, f.xml)
        row_attrs = mrow.group(1) if mrow else ""

        def decale(txt_xml, feuille_courante):
            """Formule décalée ; le texte XML d'origine est conservé tel quel si rien ne change."""
            brut = unesc(txt_xml)
            nouv = insere_lignes_formule(brut, nom_feuille, feuille_courante, a_partir, n)
            return txt_xml if nouv == brut else esc(nouv)

        def decale_formules_txt(txt, feuille_courante):
            def rep(m):
                return m.group(1) + decale(m.group(2), feuille_courante) + m.group(3)
            return re.sub(r"(<f(?: [^>]*)?(?<!/)>)(.*?)(</f>)", rep, txt, flags=re.S)

        # 1) cette feuille
        x = f.xml
        # lignes et cellules (du bas vers le haut, via une fonction)
        def rep_row(m):
            r = int(m.group(1))
            return '<row r="%d"' % (r + n if r >= a_partir else r)
        x = re.sub(r'<row r="(\d+)"', rep_row, x)

        def rep_c(m):
            r = int(m.group(2))
            return '<c r="%s%d"' % (m.group(1), r + n if r >= a_partir else r)
        x = re.sub(r'<c r="([A-Z]+)(\d+)"', rep_c, x)
        x = decale_formules_txt(x, nom_feuille)
        # attributs ref / sqref
        x = re.sub(r'(<f [^>]*?ref=")([^"]+)(")',
                   lambda m: m.group(1) + decale_plage_txt(m.group(2), a_partir, n) + m.group(3), x)
        for balise in ("dimension", "mergeCell", "autoFilter", "hyperlink"):
            x = re.sub(r'(<%s [^>]*?ref=")([^"]+)(")' % balise,
                       lambda m: m.group(1) + decale_plage_txt(m.group(2), a_partir, n) + m.group(3), x)
        x = re.sub(r'(\ssqref=")([^"]+)(")',
                   lambda m: m.group(1) + decale_plage_txt(m.group(2), a_partir, n) + m.group(3), x)
        x = re.sub(r"(<xm:sqref>)(.*?)(</xm:sqref>)",
                   lambda m: m.group(1) + decale_plage_txt(m.group(2), a_partir, n) + m.group(3), x)
        for balise in ("formula", "formula1", "formula2", "xm:f"):
            x = re.sub(r"(<%s>)(.*?)(</%s>)" % (balise, balise),
                       lambda m: m.group(1) + decale(m.group(2), nom_feuille) + m.group(3),
                       x, flags=re.S)
        f.xml = x
        f.analyse()
        # reconstruire avec les formules décalées de la ligne modèle (qui n'a pas bougé si < a_partir)
        lig_mod_apres = lig_modele + n if lig_modele >= a_partir else lig_modele
        modele2 = {c.col: c for c in f.cellules.values() if c.lig == lig_mod_apres}
        blocs = []
        for k in range(n):
            lig = a_partir + k
            cells = []
            for col in sorted(modele2, key=col_en_num):
                c = modele2[col]
                if (cols_modele and col not in cols_modele) or not c.formule:
                    cells.append(recopie_xml_cellule_sans_formule(c, lig))
                else:
                    cells.append(recopie_xml_cellule(c, lig))
            blocs.append('<row r="%d"%s>%s</row>' % (lig, row_attrs, "".join(cells)))
        x = f.xml
        pos = None
        for rm in _ROW_RE.finditer(x):
            if int(rm.group(1)) >= a_partir + n:
                pos = rm.start()
                break
        if pos is None:
            pos = x.find("</sheetData>")
        x = x[:pos] + "".join(blocs) + x[pos:]
        f.xml = x
        f.analyse()

        # 2) les autres feuilles qui pointent vers celle-ci
        for g in self.feuilles:
            if g is f:
                continue
            gx = g.xml
            if "!" not in gx:
                continue
            ng = decale_formules_txt(gx, g.nom)
            if ng != gx:
                g.xml = ng
                g.analyse()
        # 3) noms définis
        wb = self._txt("xl/workbook.xml")
        wb = re.sub(r"(<definedName [^>]*>)(.*?)(</definedName>)",
                    lambda m: m.group(1) + decale(m.group(2), None) + m.group(3), wb)
        self._ecrit_txt("xl/workbook.xml", wb)
        # 4) graphiques et ancrages du dessin
        for nom, data in list(self.parties.items()):
            if nom.startswith("xl/charts/chart") and nom.endswith(".xml"):
                t = data.decode("utf-8")
                t2 = re.sub(r"(<c:f>)(.*?)(</c:f>)",
                            lambda m: m.group(1) + decale(m.group(2), None) + m.group(3), t)
                if t2 != t:
                    self.parties[nom] = t2.encode("utf-8")
        dessin = self.dessin_de(f)
        if dessin:
            t = self._txt(dessin)
            def rep_anc(m):
                r = int(m.group(2))
                return m.group(1) + str(r + n if r >= a_partir - 1 else r) + m.group(3)
            t = re.sub(r"(<xdr:row>)(\d+)(</xdr:row>)", rep_anc, t)
            self._ecrit_txt(dessin, t)

    # ------------------------------------------------------------------
    # Déplacement d'une plage (comme « Couper » puis « Coller » dans Excel)
    # ------------------------------------------------------------------
    def deplace_plage(self, nom_feuille, rect, d_lig, d_col=0):
        """Déplace les cellules du rectangle (c1, r1, c2, r2) de (d_lig, d_col).
        Formules, fusions, mises en forme conditionnelles et validations suivent."""
        f = self.feuille(nom_feuille)
        c1, r1, c2, r2 = rect
        # formules partagées → explicites (plus sûr à déplacer)
        x = f.xml
        for c in f.cellules.values():
            if c.partagee is not None:
                nv = re.sub(r"<f[^>]*/>", "<f>%s</f>" % esc(c.formule), c.xml, count=1)
                x = x.replace(c.xml, nv, 1)
        x = re.sub(r'<f t="shared" ref="[^"]*" si="\d+">', "<f>", x)
        x = re.sub(r'<f t="shared" si="\d+" ref="[^"]*">', "<f>", x)
        f.xml = x
        f.analyse()

        def tf_formules(txt, courante):
            def rep(m):
                return m.group(1) + esc(deplace_refs_formule(unesc(m.group(2)), nom_feuille, courante,
                                                             rect, d_lig, d_col)) + m.group(3)
            return re.sub(r"(<f(?: [^>]*)?(?<!/)>)(.*?)(</f>)", rep, txt, flags=re.S)

        # 1) formules de toutes les feuilles
        for g in self.feuilles:
            gx = g.xml
            ng = tf_formules(gx, g.nom)
            if g is f:
                for balise in ("formula", "formula1", "formula2", "xm:f"):
                    ng = re.sub(r"(<%s>)(.*?)(</%s>)" % (balise, balise),
                                lambda m: m.group(1) + esc(deplace_refs_formule(unesc(m.group(2)), nom_feuille,
                                                                                nom_feuille, rect, d_lig, d_col)) + m.group(3),
                                ng, flags=re.S)
                ng = re.sub(r'(<f [^>]*?ref=")([^"]+)(")',
                            lambda m: m.group(1) + deplace_plage_txt(m.group(2), rect, d_lig, d_col) + m.group(3), ng)
                ng = re.sub(r'(<mergeCell ref=")([^"]+)(")',
                            lambda m: m.group(1) + deplace_plage_txt(m.group(2), rect, d_lig, d_col) + m.group(3), ng)
                ng = re.sub(r'(\ssqref=")([^"]+)(")',
                            lambda m: m.group(1) + deplace_plage_txt(m.group(2), rect, d_lig, d_col) + m.group(3), ng)
                ng = re.sub(r"(<xm:sqref>)(.*?)(</xm:sqref>)",
                            lambda m: m.group(1) + deplace_plage_txt(m.group(2), rect, d_lig, d_col) + m.group(3), ng)
            if ng != gx:
                g.xml = ng
                g.analyse()
        wb = self._txt("xl/workbook.xml")
        wb = re.sub(r"(<definedName [^>]*>)(.*?)(</definedName>)",
                    lambda m: m.group(1) + esc(deplace_refs_formule(unesc(m.group(2)), nom_feuille, None,
                                                                    rect, d_lig, d_col)) + m.group(3), wb)
        self._ecrit_txt("xl/workbook.xml", wb)
        for nom, data in list(self.parties.items()):
            if nom.startswith("xl/charts/chart") and nom.endswith(".xml"):
                t = data.decode("utf-8")
                t2 = re.sub(r"(<c:f>)(.*?)(</c:f>)",
                            lambda m: m.group(1) + esc(deplace_refs_formule(unesc(m.group(2)), nom_feuille, None,
                                                                            rect, d_lig, d_col)) + m.group(3), t)
                if t2 != t:
                    self.parties[nom] = t2.encode("utf-8")

        # 2) déplacer les cellules elles-mêmes
        f = self.feuille(nom_feuille)
        a_deplacer = sorted((c for c in f.cellules.values()
                             if c1 <= col_en_num(c.col) <= c2 and r1 <= c.lig <= r2),
                            key=lambda c: (c.lig, col_en_num(c.col)))
        x = f.xml
        nouveaux = []
        for c in a_deplacer:
            x = x.replace(c.xml, "", 1)
            na = num_en_col(col_en_num(c.col) + d_col) + str(c.lig + d_lig)
            nouveaux.append((na, c.xml.replace('r="%s"' % c.adr, 'r="%s"' % na, 1)))
        f.xml = x
        f.analyse()
        for na, cx in nouveaux:
            if na in f.cellules:
                f.remplace_xml_cellule(na, cx)
            else:
                f.remplace_xml_cellule(na, cx)
        # 3) dimension
        x = f.xml
        der = f.max_lig()
        x = re.sub(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>',
                   lambda m: '<dimension ref="%s%s:%s%d"/>' % (m.group(1), m.group(2), m.group(3),
                                                              max(int(m.group(4)), der)), x)
        f.xml = x
        f.analyse()

    def ajoute_style(self, base, num_fmt_id):
        """Crée (ou retrouve) un style identique à `base` mais avec un autre format de nombre."""
        st = self._txt("xl/styles.xml")
        m = re.search(r"(<cellXfs[^>]*>)(.*?)(</cellXfs>)", st, re.S)
        xfs = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", m.group(2), re.S)
        nv = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % num_fmt_id, xfs[int(base)], count=1)
        if "applyNumberFormat" not in nv:
            nv = nv.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        if nv in xfs:
            return str(xfs.index(nv))
        xfs.append(nv)
        corps = "".join(xfs)
        debut = re.sub(r'count="\d+"', 'count="%d"' % len(xfs), m.group(1))
        st = st[:m.start()] + debut + corps + m.group(3) + st[m.end():]
        self._ecrit_txt("xl/styles.xml", st)
        return str(len(xfs) - 1)

    def rels_de(self, chemin):
        d, b = os.path.split(chemin)
        return d + "/_rels/" + b + ".rels"

    def dessin_de(self, f):
        r = self.rels_de(f.chemin)
        if r not in self.parties:
            return None
        m = re.search(r'Target="([^"]*drawings/[^"]+)"', self._txt(r))
        if not m:
            return None
        return os.path.normpath(os.path.join(os.path.dirname(f.chemin), m.group(1))).replace("\\", "/")

    # ------------------------------------------------------------------
    # Copie d'une feuille (comme « Déplacer ou copier… > Créer une copie »)
    # ------------------------------------------------------------------
    def copie_feuille(self, source, nouveau_nom):
        src = self.feuille(source)
        ct = self._txt("[Content_Types].xml")

        def libre(prefixe, ext=".xml"):
            k = 1
            while "%s%d%s" % (prefixe, k, ext) in self.parties:
                k += 1
            return "%s%d%s" % (prefixe, k, ext)

        def ajoute_ct(partie, type_contenu):
            nonlocal ct
            ct = ct.replace("</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>' % (partie, type_contenu))

        def type_de(partie):
            m = re.search(r'<Override PartName="/%s" ContentType="([^"]+)"' % re.escape(partie), ct)
            return m.group(1) if m else None

        def resout(base, cible):
            return os.path.normpath(os.path.join(os.path.dirname(base), cible)).replace("\\", "/")

        def relatif(base, cible):
            return os.path.relpath(cible, os.path.dirname(base)).replace("\\", "/")

        def clone_partie(chemin, nouveau_chemin, transforme=None):
            """Clone une partie et (récursivement) ses relations."""
            data = self._txt(chemin)
            if transforme:
                data = transforme(data)
            self._ecrit_txt(nouveau_chemin, data)
            t = type_de(chemin)
            if t:
                ajoute_ct(nouveau_chemin, t)
            r = self.rels_de(chemin)
            if r in self.parties:
                rels = self._txt(r)
                def rep(m):
                    cible = m.group(2)
                    if 'TargetMode="External"' in m.group(0):
                        return m.group(0)
                    abs_c = resout(chemin, cible)
                    if abs_c.startswith("xl/printerSettings"):
                        return m.group(0)
                    rep_dir, base = os.path.split(abs_c)
                    prefixe = re.sub(r"\d+\.\w+$", "", base)
                    ext = os.path.splitext(base)[1]
                    nouvelle = libre(rep_dir + "/" + prefixe, ext)
                    tf = None
                    if "/charts/chart" in abs_c:
                        tf = lambda d: re.sub(r"(<c:f>)(.*?)(</c:f>)", lambda mm: mm.group(1) + esc(
                            renomme_feuille_formule(unesc(mm.group(2)), source, nouveau_nom)) + mm.group(3), d)
                    clone_partie(abs_c, nouvelle, tf)
                    return m.group(1) + relatif(nouveau_chemin, nouvelle) + m.group(3)
                rels = re.sub(r'(Target=")([^"]+)(")', rep, rels)
                self._ecrit_txt(self.rels_de(nouveau_chemin), rels)

        nouveau_chemin = libre("xl/worksheets/sheet")

        def tf_feuille(d):
            d = d.replace(' tabSelected="1"', "")
            d = re.sub(r'xr:uid="\{[0-9A-F-]+\}"', lambda m: m.group(0), d)
            return d
        clone_partie(src.chemin, nouveau_chemin, tf_feuille)
        self._ecrit_txt("[Content_Types].xml", ct)

        # classeur : relation + feuille
        rels = self._txt("xl/_rels/workbook.xml.rels")
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', rels)]
        rid = "rId%d" % (max(ids) + 1)
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="%s"/></Relationships>'
                            % (rid, nouveau_chemin[3:]))
        self._ecrit_txt("xl/_rels/workbook.xml.rels", rels)
        wb = self._txt("xl/workbook.xml")
        sid = max(int(i) for i in re.findall(r'sheetId="(\d+)"', wb)) + 1
        nb_avant = len(re.findall(r"<sheet ", wb))
        wb = wb.replace("</sheets>", '<sheet name="%s" sheetId="%d" r:id="%s"/></sheets>' % (esc(nouveau_nom), sid, rid))
        # nom défini du filtre automatique
        idx_src = [f.nom for f in self.feuilles].index(source)
        m = re.search(r'<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">(.*?)</definedName>' % idx_src, wb)
        if m:
            txt = renomme_feuille_formule(unesc(m.group(1)), source, nouveau_nom)
            wb = wb.replace("</definedNames>",
                            '<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s</definedName></definedNames>'
                            % (nb_avant, esc(txt)))
        self._ecrit_txt("xl/workbook.xml", wb)
        # docProps/app.xml
        if "docProps/app.xml" in self.parties:
            a = self._txt("docProps/app.xml")
            m = re.search(r"(<TitlesOfParts><vt:vector size=\")(\d+)(\"[^>]*>)(.*?)(</vt:vector>)", a, re.S)
            if m:
                k = int(m.group(2)) + 1
                a = a[:m.start()] + m.group(1) + str(k) + m.group(3) + m.group(4) + \
                    "<vt:lpstr>%s</vt:lpstr>" % esc(nouveau_nom) + m.group(5) + a[m.end():]
                a = re.sub(r"(<vt:lpstr>Feuilles de calcul</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)",
                           lambda mm: mm.group(1) + str(int(mm.group(2)) + 1), a)
                a = re.sub(r"(<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)",
                           lambda mm: mm.group(1) + str(int(mm.group(2)) + 1), a)
                self._ecrit_txt("docProps/app.xml", a)
        f = Feuille(self, nouveau_nom, nouveau_chemin)
        self.feuilles.append(f)
        self.par_nom[nouveau_nom] = f
        return f

    # ------------------------------------------------------------------
    # Suppression d'une feuille (comme « Supprimer » sur l'onglet)
    # ------------------------------------------------------------------
    def supprime_feuille(self, nom):
        f = self.feuille(nom)
        idx = self.feuilles.index(f)
        wb = self._txt("xl/workbook.xml")
        rels = self._txt("xl/_rels/workbook.xml.rels")
        cible = f.chemin[3:] if f.chemin.startswith("xl/") else "/" + f.chemin
        m = re.search(r'<Relationship [^>]*?Id="(\w+)"[^>]*?Target="/?(?:xl/)?%s"[^>]*/>' % re.escape(f.chemin[3:]), rels) or \
            re.search(r'<Relationship [^>]*?Target="/?(?:xl/)?%s"[^>]*?Id="(\w+)"[^>]*/>' % re.escape(f.chemin[3:]), rels)
        rid = m.group(1)
        rels = rels.replace(m.group(0), "")
        wb = re.sub(r'<sheet [^>]*?r:id="%s"[^>]*/>' % rid, "", wb)
        nom_q = "'" + nom.replace("'", "''") + "'"

        def dn(mm):
            tout, loc, txt = mm.group(0), mm.group(1), mm.group(2)
            if loc is not None:
                k = int(loc)
                if k == idx:
                    return ""
                if k > idx:
                    return tout.replace('localSheetId="%d"' % k, 'localSheetId="%d"' % (k - 1))
                return tout
            t = unesc(txt)
            if (nom_q + "!") in t or re.search(r"(^|[^\w'])%s!" % re.escape(nom), t):
                return ""
            return tout
        wb = re.sub(r'<definedName [^>]*?(?:localSheetId="(\d+)")?[^>]*>(.*?)</definedName>', dn, wb)
        wb = re.sub(r"<definedNames>\s*</definedNames>", "", wb)
        nb = len(re.findall(r"<sheet ", wb))
        wb = re.sub(r'activeTab="(\d+)"', lambda mm: 'activeTab="%d"' % min(int(mm.group(1)), nb - 1), wb)
        wb = re.sub(r'firstSheet="(\d+)"', lambda mm: 'firstSheet="%d"' % min(int(mm.group(1)), nb - 1), wb)
        self._ecrit_txt("xl/workbook.xml", wb)
        self._ecrit_txt("xl/_rels/workbook.xml.rels", rels)

        # parties liées (dessins, graphiques, commentaires…) qui ne servent qu'à cette feuille
        def liens(partie):
            r = self.rels_de(partie)
            if r not in self.parties:
                return []
            out = []
            for mm in re.finditer(r'<Relationship [^>]*/>', self._txt(r)):
                if 'TargetMode="External"' in mm.group(0):
                    continue
                t = re.search(r'Target="([^"]+)"', mm.group(0)).group(1)
                out.append(t.lstrip("/") if t.startswith("/") else
                           os.path.normpath(os.path.join(os.path.dirname(partie), t)).replace("\\", "/"))
            return out
        a_retirer = set()
        pile = [f.chemin]
        while pile:
            p = pile.pop()
            if p in a_retirer:
                continue
            a_retirer.add(p)
            pile.extend(x for x in liens(p) if x in self.parties)
        gardees = set()
        for p in self.parties:
            if p in a_retirer or p.endswith(".rels"):
                continue
            for x in liens(p):
                gardees.add(x)
        gardees.update(x for x in liens("xl/workbook.xml"))
        a_retirer -= gardees
        ct = self._txt("[Content_Types].xml")
        for p in a_retirer:
            for q in (p, self.rels_de(p)):
                if q in self.parties:
                    del self.parties[q]
                    self.ordre.remove(q)
            ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(p), "", ct)
        self._ecrit_txt("[Content_Types].xml", ct)
        # docProps/app.xml
        if "docProps/app.xml" in self.parties:
            a = self._txt("docProps/app.xml")
            mm = re.search(r"(<TitlesOfParts><vt:vector size=\")(\d+)(\"[^>]*>)(.*?)(</vt:vector>)", a, re.S)
            if mm:
                items = re.findall(r"<vt:lpstr>.*?</vt:lpstr>", mm.group(4), re.S)
                lp = "<vt:lpstr>%s</vt:lpstr>" % esc(nom)
                if lp in items:
                    items.remove(lp)
                    a = a[:mm.start()] + mm.group(1) + str(len(items)) + mm.group(3) + "".join(items) + mm.group(5) + a[mm.end():]
                    a = re.sub(r"(<vt:lpstr>(?:Feuilles de calcul|Worksheets)</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)",
                               lambda x: x.group(1) + str(int(x.group(2)) - 1), a)
                    self._ecrit_txt("docProps/app.xml", a)
        self.feuilles.remove(f)
        del self.par_nom[nom]

    # ------------------------------------------------------------------
    def remplace_dans_parties(self, prefixe, ancien, nouveau):
        for nom in list(self.parties):
            if nom.startswith(prefixe) and nom.endswith(".xml"):
                t = self._txt(nom)
                if ancien in t:
                    self._ecrit_txt(nom, t.replace(ancien, nouveau))
        for f in self.feuilles:
            f.analyse()


def recopie_xml_cellule_sans_formule(c, lig):
    adr = c.col + str(lig)
    attrs = re.match(r'<c r="[A-Z]+\d+"([^>]*?)(/?>)', c.xml).group(1)
    attrs = re.sub(r'\s+t="\w+"', "", attrs)
    attrs = re.sub(r'\s+cm="\d+"', "", attrs)
    return '<c r="%s"%s/>' % (adr, attrs)


def copie_sauvegarde(chemin, dossier, garder=200):
    os.makedirs(dossier, exist_ok=True)
    base = os.path.splitext(os.path.basename(chemin))[0]
    horo = datetime.datetime.now().strftime("%Y-%m-%d %Hh%Mm%Ss")
    dest = os.path.join(dossier, "%s – %s.xlsx" % (base, horo))
    k = 2
    while os.path.exists(dest):
        dest = os.path.join(dossier, "%s – %s (%d).xlsx" % (base, horo, k))
        k += 1
    shutil.copy2(chemin, dest)
    fichiers = sorted(p for p in os.listdir(dossier) if p.endswith(".xlsx"))
    for p in fichiers[:-garder]:
        try:
            os.remove(os.path.join(dossier, p))
        except OSError:
            pass
    return dest
