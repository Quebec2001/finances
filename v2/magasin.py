"""Magasin de données de la v2 : les tables de schema.sql gardées en mémoire et enregistrées en JSON (« Finances.json »).
(Pas de SQLite dans le navigateur : aucune bibliothèque externe à charger, fichier lisible, même structure de tables.)
Comprend le petit sous-ensemble de SQL utilisé par l'import : INSERT [OR REPLACE] INTO t[(cols)] VALUES (?…) et
UPDATE t SET col=? WHERE id=?."""
import json, os, re
ICI = os.path.dirname(os.path.abspath(__file__))


def lit_schema():
    s = re.sub(r"--[^\n]*", "", open(os.path.join(ICI, "schema.sql"), encoding="utf-8").read())
    tables = {}
    for nom, corps in re.findall(r"CREATE TABLE (\w+)\s*\((.*?)\);", s, re.S):
        cols, cle = [], None
        for part in re.split(r",(?![^()]*\))", corps):
            part = part.strip()
            m = re.match(r"PRIMARY KEY\s*\(([^)]*)\)", part)
            if m:
                cle = [x.strip() for x in m.group(1).split(",")]
                continue
            m = re.match(r"(\w+)\s+\w+", part)
            if m:
                cols.append(m.group(1))
                if "PRIMARY KEY" in part:
                    cle = [m.group(1)]
        tables[nom] = {"cols": cols, "cle": cle}
    return tables


class Curseur:
    def __init__(self, i): self.lastrowid = i


class Magasin:
    def __init__(self, donnees=None):
        self.schema = lit_schema()
        self.t = {n: [] for n in self.schema}
        if donnees:
            for n, l in donnees.items():
                if n in self.t: self.t[n] = l

    def execute(self, sql, params=()):
        sql = " ".join(sql.split())
        m = re.match(r"INSERT (OR REPLACE )?INTO (\w+)\s*(?:\(([^)]*)\))? VALUES \(([^)]*)\)", sql)
        if m:
            rempl, nom, cols, vals = m.groups()
            sch = self.schema[nom]
            cols = [x.strip() for x in cols.split(",")] if cols else sch["cols"]
            it = iter(params)
            vals = [next(it) if v.strip() == "?" else (int(v) if re.fullmatch(r"-?\d+", v.strip()) else v.strip().strip("'"))
                    for v in vals.split(",")]
            row = {c: None for c in sch["cols"]}
            row.update(dict(zip(cols, vals)))
            l = self.t[nom]
            if sch["cle"] == ["id"] or ("id" in sch["cols"] and sch["cle"] is None):
                if row.get("id") is None:
                    row["id"] = max((r["id"] for r in l), default=0) + 1
            elif sch["cle"]:
                k = tuple(row[c] for c in sch["cle"])
                dbl = [r for r in l if tuple(r[c] for c in sch["cle"]) == k]
                if dbl:
                    if not rempl: raise ValueError(f"doublon {nom} {k}")
                    l.remove(dbl[0])
            l.append(row)
            return Curseur(row.get("id"))
        m = re.match(r"UPDATE (\w+) SET (\w+)=\? WHERE id=\?", sql)
        if m:
            nom, col = m.groups()
            for r in self.t[m.group(1)]:
                if r["id"] == params[1]: r[col] = params[0]
            return Curseur(None)
        raise ValueError("SQL non pris en charge : " + sql)

    def json(self):
        return self.t

    def enregistre(self, chemin):
        with open(chemin, "w", encoding="utf-8") as f:
            json.dump(self.t, f, ensure_ascii=False, separators=(",", ":"))
