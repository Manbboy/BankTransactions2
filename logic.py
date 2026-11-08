"""Logique pure (sans Kivy) : lecture Excel, filtres, sélections, regroupements."""
import re
import unicodedata
import datetime as dt

from openpyxl import load_workbook

DATE_FORMATS = (
    "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d",
    "%d-%m-%y", "%d/%m/%y", "%d.%m.%y",
)
_OPS = (">=", "<=", "!=", ">", "<", "=")

PREFERRED_SHEETS = ("2014_HB_KBC",)
TECH_PREFIXES = ("config", "pivot")

NUM_NAMES = ("montant", "amount", "solde", "credit", "debit", "total")
TEXT_WORDS = (
    "numero", "communication", "cle", "unique", "detail", "compte", "iban", "bic",
    "nom", "contrepartie", "description", "type", "categorie", "origine", "periode",
    "statut", "motif", "devise", "sequence",
)


# ---------------------------------------------------------------- conversions
def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        s = re.split(r"[ T]", v.strip())[0]
        for f in DATE_FORMATS:
            try:
                return dt.datetime.strptime(s, f).date()
            except ValueError:
                pass
    return None


def to_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = re.sub(r"[\s €$£]", "", v).replace("−", "-")
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):      # 1.234,56
                s = s.replace(".", "").replace(",", ".")
            else:                                 # 1,234.56
                s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def norm(v):
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFD", s.casefold())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def fmt(v, money=False):
    """Texte affiché pour une cellule. money=True : toujours 2 décimales."""
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        if v.hour == v.minute == v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, bool):
        return str(v)
    if money and isinstance(v, (int, float)):   # un montant entier s'affiche aussi 2500,00
        return f"{float(v):.2f}".replace(".", ",")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.2f}".replace(".", ",")
    return str(v)


def fmt_money(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


# -------------------------------------------------------------------- lecture
def sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def pick_sheet(names, explicit=None, remembered=None):
    for c in (explicit, remembered):
        if c in names:
            return c
    for p in PREFERRED_SHEETS:
        if p in names:
            return p
    for n in names:
        if not n.lower().startswith(TECH_PREFIXES):
            return n
    return names[0]


def load_xlsx(path, sheet=None):
    """(en-têtes, lignes) de la feuille demandée (sinon la première).
    La première ligne de la feuille contient les en-têtes."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        try:
            ws.reset_dimensions()  # certains fichiers déclarent une taille erronée
        except Exception:  # noqa: BLE001
            pass
        it = ws.iter_rows(values_only=True)
        header = next(it, None)
        if not header:
            return [], []
        headers, seen = [], {}
        for i, h in enumerate(header):
            name = str(h).strip() if h is not None and str(h).strip() else f"Col{i + 1}"
            seen[name] = seen.get(name, 0) + 1
            if seen[name] > 1:
                name = f"{name} ({seen[name]})"
            headers.append(name)
        n = len(headers)
        data = []
        for r in it:
            if r is None or all(c is None for c in r):
                continue
            r = list(r)[:n]
            r += [None] * (n - len(r))
            data.append(r)
        return headers, data
    finally:
        wb.close()


def detect_kinds(headers, data, sample=300, threshold=0.75):
    kinds = []
    for i, h in enumerate(headers):
        vals = [r[i] for r in data[:sample * 3] if r[i] is not None][:sample]
        nd = nn = 0
        for v in vals:
            if to_date(v) is not None:
                nd += 1
            # un long numéro (communication structurée, IBAN…) n'est pas un montant
            if to_number(v) is not None and not (isinstance(v, str) and re.fullmatch(r"\d{9,}", v.strip())):
                if not (isinstance(v, (int, float)) and abs(v) >= 1e9):
                    nn += 1
        total = len(vals) or 1
        if nd / total >= threshold:
            kind = "date"
        elif nn / total >= threshold:
            kind = "num"
        else:
            kind = "text"
        hn = norm(h)
        if hn.startswith("date"):
            if nd / total >= 0.5:
                kind = "date"
        elif hn.startswith(NUM_NAMES):
            if nn / total >= 0.5:
                kind = "num"
        elif any(w in hn for w in TEXT_WORDS):
            kind = "text"
        kinds.append(kind)
    return kinds


def money_flags(headers, kinds):
    return [k == "num" and norm(h).startswith(NUM_NAMES) for h, k in zip(headers, kinds)]


# --------------------------------------------------------------------- filtres
def _val(kind, v):
    return to_number(v) if kind == "num" else to_date(v)


def _operand(kind, s):
    x = _val(kind, s.strip())
    if x is None:
        raise ValueError(f"valeur invalide : {s!r}")
    return x


def make_predicate(kind, expr):
    """Transforme l'expression saisie en fonction. None = pas de filtre.
    Lève ValueError si l'expression est invalide.

    texte  : mot (contient), a|b (ou), !mot (exclut)
    nombre : 12 | >50 | <=10 | !=0 | 10..20 | ..20 | 10..
    date   : 2026 | 2026-03 | >=2026-01-15 | 2026-01-01..2026-03-31 (ou JJ/MM/AAAA)
    """
    expr = (expr or "").strip()
    if not expr:
        return None

    if kind == "text":
        neg = expr.startswith("!")
        if neg:
            expr = expr[1:]
        terms = [norm(t.strip()) for t in expr.split("|") if t.strip()]
        if not terms:
            return None

        def p_text(v):
            s = norm(v)
            hit = any(t in s for t in terms)
            return (not hit) if neg else hit

        return p_text

    if kind == "date" and re.fullmatch(r"\d{4}(-\d{2})?", expr):
        def p_prefix(v):
            d = to_date(v)
            return d is not None and d.isoformat().startswith(expr)

        return p_prefix

    if ".." in expr:
        a, b = expr.split("..", 1)
        lo = _operand(kind, a) if a.strip() else None
        hi = _operand(kind, b) if b.strip() else None

        def p_range(v):
            x = _val(kind, v)
            return x is not None and (lo is None or x >= lo) and (hi is None or x <= hi)

        return p_range

    op, rest = "=", expr
    for o in _OPS:
        if expr.startswith(o):
            op, rest = o, expr[len(o):]
            break
    ref = _operand(kind, rest)

    def p_cmp(v):
        x = _val(kind, v)
        if x is None:
            return op == "!="
        return {
            ">=": x >= ref, "<=": x <= ref, ">": x > ref,
            "<": x < ref, "=": x == ref, "!=": x != ref,
        }[op]

    return p_cmp


def apply_filters(data, kinds, exprs, sels, strs, skip=None):
    """Filtres indépendants par colonne (expression + valeurs cochées), combinés en ET.
    Retourne (indices_des_lignes_retenues, {colonne: message_d_erreur}).
    `skip` : colonne ignorée (pour proposer ses valeurs possibles)."""
    preds, errors = [], {}
    for i, e in enumerate(exprs):
        if i == skip:
            continue
        try:
            p = make_predicate(kinds[i], e)
        except ValueError as ex:
            errors[i] = str(ex)
            p = None
        if p is not None:
            preds.append((i, p))
    checks = [(i, s) for i, s in enumerate(sels) if s and i != skip]
    if not preds and not checks:
        return list(range(len(data))), errors
    out = []
    for k, r in enumerate(data):
        ok = True
        for i, p in preds:
            if not p(r[i]):
                ok = False
                break
        if ok:
            sr = strs[k]
            for i, s in checks:
                if sr[i] not in s:
                    ok = False
                    break
        if ok:
            out.append(k)
    return out, errors


def distinct_values(idx, data, strs, col, kind):
    """[(texte, nombre_de_lignes)] trié naturellement (dates, nombres, texte)."""
    cnt = {}
    for i in idx:
        s = strs[i][col]
        e = cnt.get(s)
        if e is None:
            cnt[s] = [1, data[i][col]]
        else:
            e[0] += 1

    def key(item):
        s, (n, raw) = item
        if s == "":
            return (1, 0, "")
        if kind == "num":
            x = to_number(raw)
            return (0, x if x is not None else float("inf"), s)
        if kind == "date":
            d = to_date(raw)
            return (0, d.toordinal() if d else 10 ** 9, s)
        return (0, 0, norm(s))

    return [(s, n) for s, (n, _raw) in sorted(cnt.items(), key=key)]


def date_shortcut(name, today=None):
    """Expression de filtre de date pour un raccourci ('' = tout)."""
    today = today or dt.date.today()
    if name == "Ce mois":
        return f"{today.year}-{today.month:02d}"
    if name == "Mois dernier":
        first = today.replace(day=1)
        last = first - dt.timedelta(days=1)
        return f"{last.year}-{last.month:02d}"
    if name == "Cette année":
        return str(today.year)
    return ""


# ------------------------------------------------------------------ agrégation
PERIODS = {"Jour": "day", "Semaine": "week", "Mois": "month", "Année": "year"}


def period_key(d, period):
    if period == "day":
        return d.isoformat()
    if period == "week":
        y, w, _ = d.isocalendar()
        monday = d - dt.timedelta(days=d.weekday())
        return f"{y}-S{w:02d} (du {monday:%d/%m})"
    if period == "month":
        return f"{d.year}-{d.month:02d}"
    return str(d.year)


def period_series(rows, date_idx, amount_idx, period):
    """Série chronologique [(période, entrées, sorties)], lignes_ignorées."""
    acc, skipped = {}, 0
    for r in rows:
        d, a = to_date(r[date_idx]), to_number(r[amount_idx])
        if d is None or a is None:
            skipped += 1
            continue
        k = period_key(d, period)
        p, n = acc.get(k, (0.0, 0.0))
        acc[k] = (p + a, n) if a > 0 else (p, n + a)
    return [(k, p, n) for k, (p, n) in sorted(acc.items())], skipped


def _node(label, path):
    return {"label": label, "path": path, "n": 0, "pos": 0.0, "neg": 0.0, "ch": {}}


def _add(node, a):
    node["n"] += 1
    if a > 0:
        node["pos"] += a
    else:
        node["neg"] += a


def build_tree(rows, levels, date_idx, amount_idx, period):
    """Totaux hiérarchiques. levels : liste de ("period",) ou ("col", index).
    Retourne (racine, lignes_ignorées)."""
    root = _node("TOTAL", ())
    skipped = 0
    for r in rows:
        a = to_number(r[amount_idx])
        if a is None:
            skipped += 1
            continue
        keys, bad = [], False
        for lv in levels:
            if lv[0] == "period":
                d = to_date(r[date_idx])
                if d is None:
                    bad = True
                    break
                keys.append(period_key(d, period))
            else:
                keys.append(fmt(r[lv[1]]) or "(vide)")
        if bad:
            skipped += 1
            continue
        node = root
        _add(node, a)
        for k in keys:
            child = node["ch"].get(k)
            if child is None:
                child = _node(k, node["path"] + (k,))
                node["ch"][k] = child
            _add(child, a)
            node = child
    return root, skipped


def flatten(root, levels, expanded, all_open=False):
    """Liste à plat des lignes visibles de l'arbre (selon dépliés)."""
    out = []

    def children(node, depth):
        items = list(node["ch"].values())
        if depth < len(levels) and levels[depth][0] == "period":
            items.sort(key=lambda c: c["label"], reverse=True)   # récent en premier
        else:
            items.sort(key=lambda c: norm(c["label"]))
        return items

    def walk(node, depth):
        has = bool(node["ch"])
        is_open = has and (depth == 0 or all_open or node["path"] in expanded)
        out.append({
            "depth": depth, "path": node["path"], "label": node["label"],
            "n": node["n"], "pos": node["pos"], "neg": node["neg"],
            "net": node["pos"] + node["neg"], "has_children": has, "open": is_open,
        })
        if is_open:
            for c in children(node, depth):
                walk(c, depth + 1)

    walk(root, 0)
    return out
