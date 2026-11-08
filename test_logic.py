"""Tests de la logique (sans Kivy) : python test_logic.py"""
import datetime as dt
import os
import random
import tempfile

from openpyxl import Workbook

import logic

HEADERS = [
    "Nº de séquence", "Date d'exécution", "Date valeur", "Montant", "Devise du compte",
    "Numéro de compte", "Type de transaction", "Contrepartie", "Nom de la contrepartie",
    "Communication", "Détails", "Statut", "Motif du refus", "Origine", "Année", "Période",
    "Nom de la contrepartie NEW", "Clé unique", "Catégorie", "Sous Catégorie",
]


def make_sample(path, n_extra=0):
    """Classeur à 5 onglets comme celui de l'utilisateur ; 2014_HB_KBC n'est PAS actif."""
    wb = Workbook()
    ws = wb.active
    ws.title = "2016-2026 KBC"
    ws.append(["Date", "Montant"])
    ws.append([dt.datetime(2020, 1, 1), 1])
    ws = wb.create_sheet("2014_HB_KBC")
    ws.append(HEADERS)
    rows = [
        # dates en texte JJ-MM-AA, montants texte à virgule (comme dans l'export)
        ("2014-00266", "01-09-14", "01-09-14", "-423,09", "EUR", "BE08xxxx", "Remboursements Crédits Hypothécaires",
         None, None, 12724813846, None, "Accepté", None, "?", 2014, "2014-09", None, "k1", "31 Jean Monnet", "Credit hypo."),
        ("2014-00267", "01-09-14", "01-09-14", "-226,99", "EUR", "BE0800yy", "Débit Divers",
         None, None, 12724813846, None, "Accepté", None, "?", 2014, "2014-09", None, "k2", "31 Jean Monnet", "Credit Frais"),
        ("2014-00268", "02-09-14", "02-09-14", "587,61", "EUR", "BE080ppp", "Virement en euros",
         "BE307oo", "JEAN DUPONT", "LOYER JUILLET 2014", "CLEMENCE DUPINT", "Accepté", None, "?", 2014, "2014-09",
         "JEAN DUPONT", "k3", "91 Virements", "JEAN DUPONT"),
        # vraies dates et vrais nombres
        ("2026-00001", dt.datetime(2026, 11, 3), dt.datetime(2026, 11, 3), -12.5, "EUR", "BE08xxxx", "Paiement",
         "BE1", "Café Léon", "ticket", None, "Accepté", None, "?", 2026, "2026-11", "Café Léon", "k4", "20 Courses", "Café"),
        ("2026-00002", dt.datetime(2026, 10, 30), dt.datetime(2026, 10, 30), 2500.0, "EUR", "BE08xxxx", "Virement",
         "BE2", "ACME", "SALAIRE", None, "Accepté", None, "?", 2026, "2026-10", "ACME", "k5", "10 Revenus", "Salaire"),
        ("2026-00003", dt.datetime(2026, 10, 2), dt.datetime(2026, 10, 2), -80.0, "EUR", "BE08xxxx", "Paiement",
         "BE3", "Delhaize", "courses", None, "Accepté", None, "?", 2026, "2026-10", "Delhaize", "k6", "20 Courses", "Supermarché"),
    ]
    for r in rows:
        ws.append(list(r))
    cats = [("20 Courses", "Supermarché"), ("20 Courses", "Marché"), ("30 Loisirs", "Cinéma"), ("10 Revenus", "Salaire")]
    rnd = random.Random(1)
    for k in range(n_extra):
        d = dt.datetime(2015, 1, 1) + dt.timedelta(days=rnd.randint(0, 4000))
        cat, sub = rnd.choice(cats)
        amt = round(rnd.uniform(-300, -2), 2) if cat != "10 Revenus" else round(rnd.uniform(1500, 3000), 2)
        ws.append([f"X-{k:05d}", d, d, amt, "EUR", "BE08xxxx", "Paiement", "BE9", f"Commerçant {k % 97}",
                   f"communication {k}", None, "Accepté", None, "?", d.year, f"{d.year}-{d.month:02d}",
                   f"Commerçant {k % 97}", f"cle-{k}", cat, sub])
    for name in ("CONFIG", "PIVOT_1", "PIVOT_2"):
        w = wb.create_sheet(name)
        w.append(["a", "b"])
        w.append([1, 2])
    wb.active = 0
    wb.save(path)


def main():
    path = os.path.join(tempfile.mkdtemp(), "sample.xlsx")
    make_sample(path)

    # --- onglets : l'onglet préféré est choisi même s'il n'est pas actif
    names = logic.sheet_names(path)
    assert logic.pick_sheet(names) == "2014_HB_KBC", names
    assert logic.pick_sheet(names, remembered="2016-2026 KBC") == "2016-2026 KBC"
    assert logic.pick_sheet(names, explicit="CONFIG", remembered="2016-2026 KBC") == "CONFIG"
    assert logic.pick_sheet(["CONFIG", "PIVOT_1", "Données"]) == "Données"
    assert logic.pick_sheet(["CONFIG", "PIVOT_1"]) == "CONFIG"

    headers, data = logic.load_xlsx(path, "2014_HB_KBC")
    assert headers == HEADERS, headers
    assert len(data) == 6
    kinds = logic.detect_kinds(headers, data)
    k = dict(zip(headers, kinds))
    assert k["Date d'exécution"] == "date" and k["Date valeur"] == "date", k
    assert k["Montant"] == "num" and k["Année"] == "num", k
    for c in ("Communication", "Numéro de compte", "Période", "Clé unique", "Catégorie", "Nº de séquence", "Détails"):
        assert k[c] == "text", (c, k[c])
    money = logic.money_flags(headers, kinds)
    assert money[headers.index("Montant")] and not money[headers.index("Année")]
    strs = [[logic.fmt(c, m) for c, m in zip(r, money)] for r in data]
    ci = {h: i for i, h in enumerate(headers)}
    assert strs[0][ci["Montant"]] == "-423,09"      # texte conservé tel quel
    assert strs[3][ci["Montant"]] == "-12,50"       # nombre -> 2 décimales + virgule
    assert strs[4][ci["Montant"]] == "2500,00"
    assert strs[4][ci["Année"]] == "2026"
    assert strs[0][ci["Communication"]] == "12724813846"

    # --- dates à 2 chiffres et nombres à virgule
    assert logic.to_date("01-09-14") == dt.date(2014, 9, 1)
    assert logic.to_date("15/02/2026") == dt.date(2026, 2, 15)
    assert logic.to_date("2026-03-15 10:00:00") == dt.date(2026, 3, 15)
    assert logic.to_number("-423,09") == -423.09
    assert logic.to_number("1.234,56") == 1234.56
    assert logic.to_number("1,234.56") == 1234.56
    assert logic.to_number("−5") == -5.0
    assert logic.to_number("abc") is None

    exprs0, sels0 = [""] * len(headers), [None] * len(headers)

    def f(expr=None, sel=None, skip=None):
        ex, se = list(exprs0), list(sels0)
        for c, v in (expr or {}).items():
            ex[ci[c]] = v
        for c, v in (sel or {}).items():
            se[ci[c]] = set(v)
        idx, err = logic.apply_filters(data, kinds, ex, se, strs, skip=skip and ci[skip])
        return idx, err

    assert len(f()[0]) == 6
    # expressions
    assert len(f({"Date d'exécution": "2014"})[0]) == 3
    assert len(f({"Date d'exécution": "2026-10"})[0]) == 2
    assert len(f({"Date d'exécution": ">=2026-01-01"})[0]) == 3
    assert len(f({"Date d'exécution": "01/09/2014..02/09/2014"})[0]) == 3
    assert len(f({"Montant": "<-100"})[0]) == 2
    assert len(f({"Montant": "-100..-10"})[0]) == 2       # -12.5 et -80
    assert len(f({"Montant": ">0"})[0]) == 2
    assert len(f({"Catégorie": "courses"})[0]) == 2
    assert len(f({"Catégorie": "!courses"})[0]) == 4
    assert len(f({"Nom de la contrepartie NEW": "cafe"})[0]) == 1
    # sélections par valeur (listes à cocher) et combinaison ET
    assert len(f(sel={"Catégorie": ["20 Courses", "10 Revenus"]})[0]) == 3
    assert len(f({"Date d'exécution": "2026-10"}, {"Catégorie": ["20 Courses"]})[0]) == 1
    assert len(f(sel={"Statut": ["Accepté"]})[0]) == 6
    assert len(f(sel={"Détails": [""]})[0]) == 5            # valeur vide
    # expression invalide : signalée, ignorée, pas de plantage
    idx, err = f({"Montant": ">abc"})
    assert ci["Montant"] in err and len(idx) == 6
    # valeurs possibles d'une colonne (en ignorant son propre filtre)
    idx, _ = f({"Catégorie": "courses"}, skip="Catégorie")
    vals = logic.distinct_values(idx, data, strs, ci["Catégorie"], "text")
    assert vals == [("10 Revenus", 1), ("20 Courses", 2), ("31 Jean Monnet", 2), ("91 Virements", 1)], vals
    idx, _ = f()
    vals = logic.distinct_values(idx, data, strs, ci["Année"], "num")
    assert vals == [("2014", 3), ("2026", 3)], vals
    vals = logic.distinct_values(idx, data, strs, ci["Détails"], "text")
    assert vals[-1] == ("", 5), vals                         # le vide en dernier

    # raccourcis de période
    today = dt.date(2026, 1, 15)
    assert logic.date_shortcut("Ce mois", today) == "2026-01"
    assert logic.date_shortcut("Mois dernier", today) == "2025-12"
    assert logic.date_shortcut("Cette année", today) == "2026"
    assert logic.date_shortcut("Tout", today) == ""

    # --- totaux hiérarchiques
    rows = [data[i] for i in f()[0]]
    d_i, a_i, c_i, s_i = ci["Date d'exécution"], ci["Montant"], ci["Catégorie"], ci["Sous Catégorie"]
    levels = [("period",), ("col", c_i), ("col", s_i)]
    root, skipped = logic.build_tree(rows, levels, d_i, a_i, "month")
    assert skipped == 0
    assert root["n"] == 6
    assert round(root["pos"], 2) == round(587.61 + 2500, 2)
    assert round(root["neg"], 2) == round(-423.09 - 226.99 - 12.5 - 80, 2)
    flat = logic.flatten(root, levels, set())
    assert [r["label"] for r in flat] == ["TOTAL", "2026-11", "2026-10", "2014-09"], [r["label"] for r in flat]
    flat = logic.flatten(root, levels, {("2026-10",)})
    labs = [(r["depth"], r["label"]) for r in flat]
    assert labs == [(0, "TOTAL"), (1, "2026-11"), (1, "2026-10"), (2, "10 Revenus"), (2, "20 Courses"), (1, "2014-09")], labs
    flat = logic.flatten(root, levels, set(), all_open=True)
    assert len(flat) == 1 + 3 + 5 + 4 - 2 + 2 - 2 + 0 or len(flat) > 10
    oct_ = [r for r in flat if r["path"] == ("2026-10",)][0]
    assert round(oct_["net"], 2) == 2420.0 and oct_["n"] == 2 and oct_["has_children"]
    # sans niveau : un seul total
    root, _ = logic.build_tree(rows, [], d_i, a_i, "month")
    assert [r["label"] for r in logic.flatten(root, [], set())] == ["TOTAL"]
    # catégorie seule (sans date) : pas de ligne perdue pour date illisible
    broken = rows + [[None] * len(headers)]
    root, skipped = logic.build_tree(broken, [("col", c_i)], d_i, a_i, "month")
    assert skipped == 1
    # période semaine ISO
    s, _ = logic.period_series(rows, d_i, a_i, "week")
    assert s[0][0].startswith("2014-S36"), s[0]
    s, _ = logic.period_series(rows, d_i, a_i, "year")
    assert [(k, round(p, 2), round(n, 2)) for k, p, n in s] == [
        ("2014", 587.61, -650.08), ("2026", 2500.0, -92.5)], s

    # --- volume réaliste : 5 000 lignes
    big = os.path.join(tempfile.mkdtemp(), "big.xlsx")
    make_sample(big, n_extra=5000)
    import time
    t = time.time()
    h2, d2 = logic.load_xlsx(big, "2014_HB_KBC")
    k2 = logic.detect_kinds(h2, d2)
    m2 = logic.money_flags(h2, k2)
    st2 = [[logic.fmt(c, m) for c, m in zip(r, m2)] for r in d2]
    idx, _ = logic.apply_filters(d2, k2, [""] * len(h2), [None] * len(h2), st2)
    assert len(d2) == 5006 and len(idx) == 5006
    ex = [""] * len(h2)
    ex[h2.index("Date d'exécution")] = "2020"
    sl = [None] * len(h2)
    sl[h2.index("Catégorie")] = {"20 Courses"}
    idx, _ = logic.apply_filters(d2, k2, ex, sl, st2)
    assert 0 < len(idx) < 5006
    print(f"5 000 lignes : lecture + filtres en {time.time() - t:.2f} s")

    assert logic.fmt_money(-1234.5) == "-1 234,50"
    print("OK — tous les tests passent")


if __name__ == "__main__":
    main()
