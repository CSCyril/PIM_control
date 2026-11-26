import os
import re
import json
import shutil
import threading
import datetime
import urllib.request as urlreq
from collections import Counter, defaultdict

from flask import Flask, jsonify, request, render_template, send_file
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
SRC = os.environ.get("PIM_SOURCE", "")
XLSX = os.environ.get("PIM_FILE") or os.path.join(DATA_DIR, "pim_source_pour_controle.xlsx")
HOST = os.environ.get("PIM_HOST", "127.0.0.1")
PORT = int(os.environ.get("PIM_PORT", "8765"))

os.makedirs(os.path.dirname(os.path.abspath(XLSX)), exist_ok=True)
if not os.path.exists(XLSX):
    if SRC and os.path.exists(SRC):
        shutil.copy(SRC, XLSX)
    else:
        raise SystemExit(
            "Fichier Excel introuvable : %s\n"
            "Copiez votre fichier source dans data/ (ex. data/pim_source_pour_controle.xlsx)\n"
            "ou lancez avec : PIM_FILE=/chemin/vers/votre_fichier.xlsx python3 app.py" % XLSX)

backup = os.path.join(DATA_DIR, "sauvegarde_%s_%s" % (
    datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), os.path.basename(XLSX)))
if not os.path.exists(backup):
    shutil.copy(XLSX, backup)

REFS_DIR = os.environ.get("REFS_DIR") or os.path.join(BASE, "refs")

attributes = {}
attr_options = {}
categories = {}
families = {}
variants = {}
attr_groups = {}
attr_types = {}
sku_prefixes = []
refs_loaded = False
col_meta = {}
unknown_columns = []


def _find_ref(name_part):
    if not os.path.isdir(REFS_DIR):
        return None
    for f in sorted(os.listdir(REFS_DIR)):
        if f.lower().endswith(".xlsx") and name_part.lower() in f.lower():
            return os.path.join(REFS_DIR, f)
    return None


def _sheet_rows(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def _header_idx(rows, key="code", max_scan=15):
    for i, r in enumerate(rows[:max_scan]):
        for c in r:
            if c and str(c).strip().lower() == key.lower():
                return i
    return None


def _col(hdr, *names):
    for n in names:
        for i, h in enumerate(hdr):
            if h and str(h).strip().lower() == n.lower():
                return i
    for n in names:
        for i, h in enumerate(hdr):
            if h and str(h).strip().lower().startswith(n.lower()):
                return i
    return None


def _lab(hdr, row, *names):
    i = _col(hdr, *names)
    if i is None or i >= len(row) or row[i] in (None, ""):
        return None
    return str(row[i]).strip()


def attr_code_of(col):
    if not col or not isinstance(col, str):
        return None
    return col.split("-")[0]


def load_refs():
    global attributes, attr_options, categories, families, variants
    global attr_groups, attr_types, sku_prefixes, refs_loaded, col_meta, unknown_columns
    attributes, attr_options, categories, families, variants = {}, {}, {}, {}, {}
    attr_groups, attr_types, sku_prefixes = {}, {}, []

    p = _find_ref("DICTIONARY ATTRIBUTES")
    if p:
        rows = _sheet_rows(p, "edited")
        hi = _header_idx(rows)
        if hi is not None:
            hdr = rows[hi]
            for r in rows[hi + 1:]:
                code = _lab(hdr, r, "code")
                if not code:
                    continue
                attributes[code] = {
                    "label": _lab(hdr, r, "label") or code,
                    "type": _lab(hdr, r, "type"),
                    "group": _lab(hdr, r, "PARENT CODE"),
                    "level": _lab(hdr, r, "managed"),
                    "definition": _lab(hdr, r, "definition"),
                }

    p = _find_ref("ATTRIBUTE OPTIONS")
    if p:
        rows = _sheet_rows(p, "EXPORT")
        hi = _header_idx(rows)
        if hi is not None:
            hdr = rows[hi]
            for r in rows[hi + 1:]:
                a = _lab(hdr, r, "attribute")
                c = _lab(hdr, r, "code")
                if not a or not c:
                    continue
                attr_options.setdefault(a, {})[c] = _lab(hdr, r, "label-fr_FR", "label-en_SOP", "label-en_US") or c

    p = _find_ref("DICTIONARY CATEGORIES")
    if p:
        rows = _sheet_rows(p)
        hi = _header_idx(rows)
        if hi is not None:
            hdr = rows[hi]
            for r in rows[hi + 1:]:
                c = _lab(hdr, r, "code")
                if c:
                    categories[c] = _lab(hdr, r, "label-fr_FR", "label-en_SOP", "label-en_US") or c

    p = _find_ref("DICTIONARY FAMILIES")
    if p:
        rows = _sheet_rows(p)
        hi = _header_idx(rows)
        if hi is not None:
            hdr = rows[hi]
            ai = _col(hdr, "attributes")
            mi = _col(hdr, "requirements")
            for r in rows[hi + 1:]:
                c = _lab(hdr, r, "code")
                if not c:
                    continue
                attrs, mand = set(), set()
                if ai is not None and ai < len(r) and r[ai]:
                    attrs = {x.strip() for x in str(r[ai]).replace(";", ",").split(",") if x.strip()}
                if mi is not None and mi < len(r) and r[mi]:
                    mand = {x.strip() for x in str(r[mi]).replace(";", ",").split(",") if x.strip()}
                families[c] = {"label": _lab(hdr, r, "label-fr_FR", "label-en_SOP") or c,
                               "attributes": sorted(attrs), "mandatory": sorted(mand)}

    p = _find_ref("LINKINGTABLE") or _find_ref("FAMILIE VARIANTS")
    if p:
        rows = _sheet_rows(p)
        hi = _header_idx(rows, "variant original", max_scan=5)
        if hi is None:
            hi = 0
        hdr = rows[hi]
        fi = _col(hdr, "Variant 1 FAM")
        axes_cols = [i for i, h in enumerate(hdr)
                     if h and re.fullmatch(r"variant \d+", str(h).strip().lower())]
        for r in rows[hi + 1:]:
            v = _lab(hdr, r, "variant original")
            if not v:
                continue
            fam = str(r[fi]).strip() if fi is not None and fi < len(r) and r[fi] else None
            axes = [str(r[i]).strip() for i in axes_cols if i < len(r) and r[i]]
            variants[v] = {"family": fam, "axes": axes}

    p = _find_ref("ATTRIBUTE GROUPS")
    if p:
        rows = _sheet_rows(p, "EXPORT")
        hi = _header_idx(rows)
        if hi is not None:
            hdr = rows[hi]
            for r in rows[hi + 1:]:
                c = _lab(hdr, r, "code")
                if c:
                    attr_groups[c] = _lab(hdr, r, "label-fr_FR", "label-en_US") or c

    p = _find_ref("ATTRIBUTE TYPES")
    if p:
        rows = _sheet_rows(p)
        hi = _header_idx(rows, "type")
        if hi is not None:
            hdr = rows[hi]
            for r in rows[hi + 1:]:
                t = _lab(hdr, r, "type")
                if t:
                    attr_types[t] = _lab(hdr, r, "description") or ""

    p = _find_ref("PREFIXES")
    if p:
        rows = _sheet_rows(p)
        hi = _header_idx(rows, "subject")
        if hi is not None:
            hdr = rows[hi]
            pi = _col(hdr, "prefix")
            for r in rows[hi + 1:]:
                if pi is not None and pi < len(r) and r[pi]:
                    sku_prefixes.append(str(r[pi]).strip())

    if os.path.isdir(REFS_DIR):
        for f in sorted(os.listdir(REFS_DIR)):
            if "addendum" in f.lower() and f.lower().endswith(".xlsx"):
                w = openpyxl.load_workbook(os.path.join(REFS_DIR, f), read_only=True, data_only=True)
                if "VARIANTS" in w.sheetnames:
                    rows = [list(r) for r in w["VARIANTS"].iter_rows(values_only=True)]
                    hi = _header_idx(rows, "variant original", max_scan=5) or 0
                    hdr = rows[hi]
                    fi = _col(hdr, "Variant 1 FAM")
                    axes_cols = [i for i, h in enumerate(hdr)
                                 if h and re.fullmatch(r"variant \d+", str(h).strip().lower())]
                    for r in rows[hi + 1:]:
                        v = _lab(hdr, r, "variant original")
                        if not v:
                            continue
                        fam = str(r[fi]).strip() if fi is not None and fi < len(r) and r[fi] else None
                        axes = [str(r[i]).strip() for i in axes_cols if i < len(r) and r[i]]
                        variants[v] = {"family": fam, "axes": axes}
                if "FAMILIES" in w.sheetnames:
                    rows = [list(r) for r in w["FAMILIES"].iter_rows(values_only=True)]
                    hi = _header_idx(rows, "code")
                    if hi is not None:
                        hdr = rows[hi]
                        ai = _col(hdr, "attributes")
                        mi = _col(hdr, "requirements")
                        for r in rows[hi + 1:]:
                            c = _lab(hdr, r, "code")
                            if not c:
                                continue
                            attrs, mand = set(), set()
                            if ai is not None and ai < len(r) and r[ai]:
                                attrs = {x.strip() for x in str(r[ai]).replace(";", ",").split(",") if x.strip()}
                            if mi is not None and mi < len(r) and r[mi]:
                                mand = {x.strip() for x in str(r[mi]).replace(";", ",").split(",") if x.strip()}
                            families[c] = {"label": _lab(hdr, r, "label-fr_FR", "label-en_US") or c,
                                           "attributes": sorted(attrs), "mandatory": sorted(mand)}
                w.close()

    refs_loaded = bool(attributes or attr_options or categories or families or variants)

    col_meta = {}
    unknown_columns = []
    fixed = {"sku", "categories", "enabled", "family", "parent", "code", "family_variant"}
    for col in list(pr_hdr) + list(pm_hdr):
        base = attr_code_of(col)
        meta = None
        if base and base in attributes:
            a = attributes[base]
            meta = {"code": base, "label": a["label"], "type": a["type"], "group": a["group"],
                    "group_label": attr_groups.get(a["group"]) if a["group"] else None,
                    "is_unit": isinstance(col, str) and col.endswith("-unit"),
                    "is_locale": isinstance(col, str) and "-" in col and not col.endswith("-unit")}
            if base in attr_options:
                meta["options"] = attr_options[base]
        col_meta[col] = meta
        if base and col not in fixed and base not in attributes:
            unknown_columns.append(col)


app = Flask(__name__)
lock = threading.RLock()

wb = openpyxl.load_workbook(XLSX)
pm_ws = wb["PM"]
pr_ws = wb["produits"]

pm_hdr = []
pr_hdr = []
pm_col = {}
pr_col = {}
unit_pairs = []
mandatory_pr = []

pm_rows = []
pr_rows = []
pm_by_code = defaultdict(list)
pr_by_parent = defaultdict(list)
pr_by_sku = defaultdict(list)
product_checks = {}
pm_checks = {}
pm_list = []
distinct_pm = {}
distinct_pr = {}
last_saved = None

NUM_RE = re.compile(r"^-?\d+([.,]\d+)?$")

LABELS = {
    "pm_doublon_conflictuel": "PM présent en plusieurs lignes avec des valeurs différentes",
    "pm_doublon_identique": "PM présent en plusieurs lignes identiques (doublon à nettoyer)",
    "pm_sans_produit": "Aucun produit rattaché à ce PM",
    "pm_categorie_mismatch": "Produits dont les catégories diffèrent de celles du PM",
    "pm_famille_mismatch": "Produits dont la famille diffère de la famille du PM",
    "axe_variant_vide": "Produit sans valeur pour un axe de variante du PM",
    "axe_variant_doublon": "Combinaison d'axes de variante en double parmi les produits du PM",
    "parent_sans_pm": "Produit dont le parent n'existe pas dans la feuille PM",
    "sku_doublon": "SKU présent en double dans la feuille produits",
    "desactive": "Produit désactivé (enabled = 0)",
    "valeur_sans_unite": "Valeur renseignée sans unité associée",
    "unite_sans_valeur": "Unité renseignée sans valeur associée",
    "champ_obligatoire_vide": "Champ obligatoire vide",
    "categorie_diff_pm": "Catégories du produit différentes de celles du PM",
    "famille_diff_pm": "Famille du produit différente de la famille du PM",
    "option_invalide": "Valeur d'option inconnue du dictionnaire Akeneo",
    "famille_inconnue": "Famille non déclarée dans le dictionnaire des familles",
    "family_variant_non_reference": "family_variant absent de la LINKINGTABLE Akeneo",
    "attribut_inconnu": "Colonne du fichier PIM sans attribut correspondant dans le dictionnaire",
}


def s(v):
    return "" if v is None else str(v)


def parse_value(raw):
    if raw is None:
        return None
    t = str(raw).strip()
    if t == "":
        return None
    if NUM_RE.match(t):
        t2 = t.replace(",", ".")
        f = float(t2)
        if "." not in t2 and "," not in t2:
            return int(f)
        return f
    return t


def scan():
    global pm_hdr, pr_hdr, pm_col, pr_col, unit_pairs, mandatory_pr
    global pm_rows, pr_rows, pm_by_code, pr_by_parent, pr_by_sku
    global product_checks, pm_checks, pm_list, distinct_pm, distinct_pr

    pm_hdr = [c.value for c in pm_ws[1]]
    pr_hdr = [c.value for c in pr_ws[1]]
    pm_col = {h: i + 1 for i, h in enumerate(pm_hdr)}
    pr_col = {h: i + 1 for i, h in enumerate(pr_hdr)}

    unit_pairs = [(h, h + "-unit") for h in pr_hdr
                  if h and isinstance(h, str) and (h + "-unit") in pr_col]

    pm_rows = []
    for r in range(2, pm_ws.max_row + 1):
        vals = {h: pm_ws.cell(row=r, column=pm_col[h]).value for h in pm_hdr}
        pm_rows.append({"row": r, "values": vals})

    pr_rows = []
    for r in range(2, pr_ws.max_row + 1):
        vals = {h: pr_ws.cell(row=r, column=pr_col[h]).value for h in pr_hdr}
        pr_rows.append({"row": r, "values": vals})

    pm_by_code = defaultdict(list)
    for e in pm_rows:
        code = s(e["values"].get("code"))
        if code:
            pm_by_code[code].append(e["row"])
    pr_by_parent = defaultdict(list)
    pr_by_sku = defaultdict(list)
    for e in pr_rows:
        pr_by_parent[s(e["values"].get("parent"))].append(e["row"])
        pr_by_sku[s(e["values"].get("sku"))].append(e["row"])

    n = len(pr_rows)
    fixed = {"sku", "categories", "enabled", "family", "parent"}
    mandatory_pr = [h for h in pr_hdr
                    if h and isinstance(h, str) and not h.endswith("-unit")
                    and h not in fixed
                    and sum(1 for e in pr_rows if s(e["values"].get(h))) / max(n, 1) >= 0.9]

    distinct_pm = {h: sorted({s(e["values"].get(h)) for e in pm_rows if s(e["values"].get(h))})[:500]
                   for h in pm_hdr}
    distinct_pr = {h: sorted({s(e["values"].get(h)) for e in pr_rows if s(e["values"].get(h))})[:500]
                   for h in pr_hdr}

    compute_checks()


def product_checks_for(vals):
    checks = []
    parent = s(vals.get("parent"))
    sku = s(vals.get("sku"))
    if not parent or parent not in pm_by_code:
        checks.append({"type": "parent_sans_pm", "severity": "warning",
                       "msg": "Parent « %s » introuvable dans la feuille PM" % (parent or "(vide)")})
    if sku and len(pr_by_sku.get(sku, [])) > 1:
        checks.append({"type": "sku_doublon", "severity": "warning",
                       "msg": "SKU présent %d fois dans la feuille produits" % len(pr_by_sku[sku])})
    if s(vals.get("enabled")) in ("0", "False", "false"):
        checks.append({"type": "desactive", "severity": "info", "msg": "Produit désactivé (enabled = 0)"})
    for base, unit in unit_pairs:
        v, u = s(vals.get(base)), s(vals.get(unit))
        if v and not u:
            checks.append({"type": "valeur_sans_unite", "severity": "warning",
                           "msg": "%s = %s sans unité (%s)" % (base, v, unit)})
        if u and not v:
            checks.append({"type": "unite_sans_valeur", "severity": "warning",
                           "msg": "%s = %s sans valeur (%s)" % (unit, u, base)})
    for m in mandatory_pr:
        if not s(vals.get(m)):
            checks.append({"type": "champ_obligatoire_vide", "severity": "warning",
                           "msg": "Champ obligatoire vide : %s" % m})
    famv = s(vals.get("family"))
    if famv and families and famv not in families:
        checks.append({"type": "famille_inconnue", "severity": "info",
                       "msg": "Famille « %s » non déclarée dans le dictionnaire des familles" % famv})
    for col in pr_hdr:
        if not isinstance(col, str) or col in ("sku", "categories", "enabled", "family", "parent") or col.endswith("-unit"):
            continue
        base = attr_code_of(col)
        opts = attr_options.get(base)
        v = s(vals.get(col))
        if opts and v:
            bad = [t.strip() for t in v.split(",") if t.strip() and t.strip() not in opts]
            if bad:
                checks.append({"type": "option_invalide", "severity": "warning",
                               "msg": "%s : valeur(s) hors dictionnaire : %s (attendu %s_OPTIONxxxx)"
                                      % (col, ", ".join(bad[:3]), base)})
    if parent in pm_by_code:
        pmrow = next(e for e in pm_rows if e["row"] == pm_by_code[parent][0])
        pmc = s(pmrow["values"].get("categories"))
        if pmc and s(vals.get("categories")) and s(vals.get("categories")) != pmc:
            checks.append({"type": "categorie_diff_pm", "severity": "warning",
                           "msg": "Catégories produit « %s » ≠ PM « %s »" % (s(vals.get("categories")), pmc)})
        fam = s(pmrow["values"].get("family_variant")).split("_")[0]
        if fam and s(vals.get("family")) and s(vals.get("family")) != fam:
            checks.append({"type": "famille_diff_pm", "severity": "warning",
                           "msg": "Famille produit « %s » ≠ famille PM « %s »" % (s(vals.get("family")), fam)})
    return checks


def compute_checks():
    global product_checks, pm_checks, pm_list
    product_checks = {}
    for e in pr_rows:
        product_checks[e["row"]] = product_checks_for(e["values"])

    pm_checks = {}
    all_codes = set(pm_by_code) | {p for p in pr_by_parent if p}
    for code in sorted(all_codes):
        checks = []
        rows = pm_by_code.get(code, [])
        prods = [e for e in pr_rows if s(e["values"].get("parent")) == code]
        if len(rows) > 1:
            vals_list = [next(e for e in pm_rows if e["row"] == r)["values"] for r in rows]
            fields = [h for h in pm_hdr if len({s(v.get(h)) for v in vals_list}) > 1]
            if fields:
                checks.append({"type": "pm_doublon_conflictuel", "severity": "warning",
                               "msg": "%d lignes PM pour ce code, valeurs différentes sur : %s"
                                      % (len(rows), ", ".join(fields))})
            else:
                checks.append({"type": "pm_doublon_identique", "severity": "info",
                               "msg": "%d lignes PM identiques pour ce code" % len(rows)})
        if not prods:
            checks.append({"type": "pm_sans_produit", "severity": "warning",
                           "msg": "Aucun produit rattaché à ce PM"})
        if rows and prods:
            first = next(e for e in pm_rows if e["row"] == rows[0])["values"]
            pmc = s(first.get("categories"))
            cat_bad = [s(p["values"].get("sku")) for p in prods
                       if pmc and s(p["values"].get("categories")) != pmc]
            if cat_bad:
                checks.append({"type": "pm_categorie_mismatch", "severity": "warning",
                               "msg": "%d produit(s) avec catégories ≠ PM : %s"
                                      % (len(cat_bad), ", ".join(cat_bad[:8]))})
            fv = s(first.get("family_variant"))
            entry = variants.get(fv) if fv else None
            if fv and not entry:
                checks.append({"type": "family_variant_non_reference", "severity": "warning",
                               "msg": "family_variant « %s » absent de la LINKINGTABLE (axes déduits du nom)" % fv})
            if entry:
                fam = entry.get("family") or fv.split("_")[0]
                axes = [a for a in entry.get("axes", []) if a in pr_col]
            else:
                fam = fv.split("_")[0] if fv else ""
                axes = [a for a in fv.split("_")[1:] if a in pr_col] if fv else []
            fam_bad = [s(p["values"].get("sku")) for p in prods
                       if fam and s(p["values"].get("family")) and s(p["values"].get("family")) != fam]
            if fam_bad:
                checks.append({"type": "pm_famille_mismatch", "severity": "warning",
                               "msg": "%d produit(s) avec famille ≠ PM (%s) : %s"
                                      % (len(fam_bad), fam, ", ".join(fam_bad[:8]))})
            for col in pm_hdr:
                if not isinstance(col, str) or col in ("code", "family_variant", "categories"):
                    continue
                base = attr_code_of(col)
                opts = attr_options.get(base)
                v = s(first.get(col))
                if opts and v:
                    bad = [t.strip() for t in v.split(",") if t.strip() and t.strip() not in opts]
                    if bad:
                        checks.append({"type": "option_invalide", "severity": "warning",
                                       "msg": "%s : valeur(s) hors dictionnaire : %s (attendu %s_OPTIONxxxx)"
                                              % (col, ", ".join(bad[:3]), base)})
            if axes:
                missing = [s(p["values"].get("sku")) for p in prods
                           if any(not s(p["values"].get(a)) for a in axes)]
                if missing:
                    checks.append({"type": "axe_variant_vide", "severity": "warning",
                                   "msg": "%d produit(s) sans valeur d'axe de variante (%s) : %s"
                                          % (len(missing), ", ".join(axes), ", ".join(missing[:8]))})
                combos = Counter(tuple(s(p["values"].get(a)) for a in axes) for p in prods)
                dups = [c for c, n in combos.items() if n > 1]
                if dups:
                    checks.append({"type": "axe_variant_doublon", "severity": "warning",
                                   "msg": "%d combinaison(s) d'axes en double : %s"
                                          % (len(dups), "; ".join("/".join(c) for c in dups[:5]))})
        for p in prods:
            for c in product_checks[p["row"]]:
                if c["type"] in ("categorie_diff_pm", "famille_diff_pm"):
                    continue
                if c["type"] == "parent_sans_pm" and rows:
                    continue
                cc = dict(c)
                cc["sku"] = s(p["values"].get("sku"))
                checks.append(cc)
        pm_checks[code] = checks

    pm_list = []
    for code in sorted(all_codes):
        pm_list.append({
            "code": code,
            "has_pm": code in pm_by_code,
            "nb_pm_rows": len(pm_by_code.get(code, [])),
            "nb_products": len(pr_by_parent.get(code, [])),
            "nb_issues": len(pm_checks.get(code, [])),
        })


def save():
    global last_saved
    wb.save(XLSX)
    last_saved = datetime.datetime.now().strftime("%H:%M:%S")


def pm_detail(code):
    rows = pm_by_code.get(code, [])
    pm_rows_out = [e for e in pm_rows if e["row"] in rows]
    prods = [{"row": e["row"], "values": e["values"],
              "checks": product_checks.get(e["row"], [])}
             for e in pr_rows if s(e["values"].get("parent")) == code]
    return {
        "code": code,
        "pm_rows": pm_rows_out,
        "products": prods,
        "checks": pm_checks.get(code, []),
        "has_pm": bool(rows),
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/state")
def api_state():
    with lock:
        return jsonify({
            "pm_hdr": pm_hdr,
            "pr_hdr": pr_hdr,
            "unit_pairs": unit_pairs,
            "mandatory_pr": mandatory_pr,
            "pm_list": pm_list,
            "distinct_pm": distinct_pm,
            "distinct_pr": distinct_pr,
            "labels": LABELS,
            "last_saved": last_saved,
            "file": os.path.basename(XLSX),
            "backup": os.path.basename(backup),
            "refs": {
                "loaded": refs_loaded,
                "dir": REFS_DIR,
                "summary": {
                    "attributes": len(attributes),
                    "options": sum(len(v) for v in attr_options.values()),
                    "categories": len(categories),
                    "families": len(families),
                    "variants": len(variants),
                    "groups": len(attr_groups),
                    "types": len(attr_types),
                    "prefixes": len(sku_prefixes),
                },
                "col_meta": col_meta,
                "unknown_columns": unknown_columns,
                "families": {k: v["label"] for k, v in families.items()},
                "categories": categories,
                "variants": sorted(variants.keys()),
            },
        })


@app.route("/api/pm/<path:code>")
def api_pm(code):
    with lock:
        return jsonify(pm_detail(code))


@app.route("/api/pm/row/<int:r>", methods=["PUT"])
def api_pm_update(r):
    data = request.get_json(force=True)
    field = data.get("field")
    with lock:
        if field not in pm_col or r < 2 or r > pm_ws.max_row:
            return jsonify({"error": "champ ou ligne invalide"}), 400
        pm_ws.cell(row=r, column=pm_col[field]).value = parse_value(data.get("value"))
        scan()
        save()
        return jsonify({"ok": True, "saved_at": last_saved})


@app.route("/api/pm/row/<int:r>", methods=["DELETE"])
def api_pm_delete(r):
    with lock:
        if r < 2 or r > pm_ws.max_row:
            return jsonify({"error": "ligne invalide"}), 400
        pm_ws.delete_rows(r)
        scan()
        save()
        return jsonify({"ok": True, "saved_at": last_saved})


@app.route("/api/pm", methods=["POST"])
def api_pm_create():
    data = request.get_json(force=True)
    code = s(data.get("code"))
    values = data.get("values") or {}
    with lock:
        if not code or code in pm_by_code:
            return jsonify({"error": "code vide ou déjà existant"}), 400
        r = pm_ws.max_row + 1
        pm_ws.cell(row=r, column=pm_col["code"]).value = code
        for f, v in values.items():
            if f in pm_col:
                pm_ws.cell(row=r, column=pm_col[f]).value = parse_value(v)
        scan()
        save()
        return jsonify({"ok": True, "row": r, "saved_at": last_saved})


@app.route("/api/codes/next")
def api_codes_next():
    with lock:
        nums = [int(m.group(1)) for c in pm_by_code
                for m in [re.fullmatch(r"NSPLD(\d+)", c)] if m]
        mx = max(nums) if nums else 0
        nxt = mx + 1
        return jsonify({"next": "NSPLD%03d" % nxt, "max": mx, "count": len(pm_by_code)})


def _next_free_codes(n, start):
    out = []
    cur = start
    with lock:
        while len(out) < n:
            cur += 1
            c = "NSPLD%03d" % cur
            if c not in pm_by_code:
                out.append(c)
    return out


@app.route("/api/pm/<path:code>/split_suggestions")
def api_split_suggestions(code):
    with lock:
        d = pm_detail(code)
        prods = d["products"]
        fv = s(d["pm_rows"][0]["values"].get("family_variant")) if d["pm_rows"] else ""
        axes = [a for a in fv.split("_")[1:] if a in pr_col] if fv else []
        suggestions = []
        criteria = [("Axe variante : %s" % a, a) for a in axes]
        criteria += [("Coloris EUINF0233", "EUINF0233"), ("Sélection EUINF0400", "EUINF0400"),
                     ("Catégorie", "categories"), ("Marque EUINF0010", "EUINF0010")]
        for label, field in criteria:
            if field not in pr_col:
                continue
            groups = defaultdict(list)
            for p in prods:
                groups[s(p["values"].get(field)) or "(vide)"].append(p["row"])
            if len(groups) > 1:
                suggestions.append({"label": label, "field": field,
                                    "groups": {k: v for k, v in groups.items()}})
        nums = [int(m.group(1)) for c in pm_by_code for m in [re.fullmatch(r"NSPLD(\d+)", c)] if m]
        start = max(nums) if nums else 0
        return jsonify({
            "code": code, "family_variant": fv, "axes": axes,
            "products": [{"row": p["row"], "sku": s(p["values"].get("sku")),
                          "name": s(p["values"].get("EUMKG0160-fr_FR")),
                          "values": {k: p["values"].get(k) for k in
                                     (["EUINF0233", "EUINF0400", "categories", "EUINF0010"] + axes)}} for p in prods],
            "suggestions": suggestions,
            "next_codes": _next_free_codes(max(2, len(suggestions) + 1), start),
        })


@app.route("/api/pm/<path:code>/split", methods=["POST"])
def api_split(code):
    data = request.get_json(force=True)
    groups = data.get("groups") or []
    with lock:
        src = pm_by_code.get(code) or []
        if not src:
            return jsonify({"error": "PM inconnu"}), 404
        src_vals = next(e for e in pm_rows if e["row"] == src[0])["values"]
        valid_rows = {p["row"] for p in pr_rows if s(p["values"].get("parent")) == code}
        created, errors = [], []
        used_codes = set()
        for g in groups:
            new_code = s(g.get("new_code"))
            rows = g.get("rows") or []
            values = g.get("values") or {}
            if not new_code or new_code in pm_by_code or new_code in used_codes:
                errors.append({"group": new_code, "error": "code vide, existant ou dupliqué"})
                continue
            bad_rows = [r for r in rows if r not in valid_rows]
            if bad_rows:
                errors.append({"group": new_code, "error": "lignes produit invalides: %s" % bad_rows})
                continue
            merged = dict(src_vals)
            merged.update({k: v for k, v in values.items() if k in pm_col})
            merged["code"] = new_code
            for mand in ("EUMKG0030-fr_FR", "EUMKG0120-fr_FR"):
                if not s(merged.get(mand)):
                    errors.append({"group": new_code, "error": "champ obligatoire vide : %s" % mand})
                    break
            else:
                r = pm_ws.max_row + 1
                for f, v in merged.items():
                    if f in pm_col:
                        pm_ws.cell(row=r, column=pm_col[f]).value = parse_value(v)
                for row in rows:
                    pr_ws.cell(row=row, column=pr_col["parent"]).value = new_code
                used_codes.add(new_code)
                created.append({"new_code": new_code, "row": r, "nb_produits": len(rows)})
        if created:
            scan()
            save()
        return jsonify({"created": created, "errors": errors, "saved_at": last_saved})


@app.route("/api/produit/row/<int:r>", methods=["PUT"])
def api_pr_update(r):
    data = request.get_json(force=True)
    field = data.get("field")
    with lock:
        if field not in pr_col or r < 2 or r > pr_ws.max_row:
            return jsonify({"error": "champ ou ligne invalide"}), 400
        pr_ws.cell(row=r, column=pr_col[field]).value = parse_value(data.get("value"))
        scan()
        save()
        return jsonify({"ok": True, "saved_at": last_saved})


@app.route("/api/produit/row/<int:r>", methods=["DELETE"])
def api_pr_delete(r):
    with lock:
        if r < 2 or r > pr_ws.max_row:
            return jsonify({"error": "ligne invalide"}), 400
        pr_ws.delete_rows(r)
        scan()
        save()
        return jsonify({"ok": True, "saved_at": last_saved})


@app.route("/api/produit", methods=["POST"])
def api_pr_create():
    data = request.get_json(force=True)
    parent = s(data.get("parent"))
    with lock:
        r = pr_ws.max_row + 1
        if "parent" in pr_col:
            pr_ws.cell(row=r, column=pr_col["parent"]).value = parent or None
        if "enabled" in pr_col:
            pr_ws.cell(row=r, column=pr_col["enabled"]).value = 1
        scan()
        save()
        return jsonify({"ok": True, "row": r, "saved_at": last_saved})


@app.route("/api/checks")
def api_checks():
    with lock:
        groups = defaultdict(list)
        for code, checks in pm_checks.items():
            for c in checks:
                item = {"pm": code, "msg": c["msg"], "severity": c["severity"]}
                if c.get("sku"):
                    item["sku"] = c["sku"]
                groups[c["type"]].append(item)
        for e in pr_rows:
            parent = s(e["values"].get("parent"))
            if not parent:
                for c in product_checks[e["row"]]:
                    groups[c["type"]].append({"pm": "(sans parent)",
                                              "sku": s(e["values"].get("sku")),
                                              "msg": c["msg"], "severity": c["severity"]})
        for col in unknown_columns:
            groups["attribut_inconnu"].append({"pm": "-",
                                               "msg": "Colonne « %s » : aucun attribut « %s » dans le dictionnaire"
                                                      % (col, attr_code_of(col)),
                                               "severity": "info"})
        out = [{"type": t, "label": LABELS.get(t, t), "count": len(items),
                "items": sorted(items, key=lambda x: (x["pm"], x.get("sku", "")))[:200]}
               for t, items in sorted(groups.items(), key=lambda kv: -len(kv[1]))]
        return jsonify(out)


@app.route("/api/save", methods=["POST"])
def api_save():
    with lock:
        save()
        return jsonify({"ok": True, "saved_at": last_saved})


@app.route("/api/export")
def api_export():
    with lock:
        save()
    return send_file(XLSX, as_attachment=True,
                     download_name=os.path.basename(XLSX).replace(".xlsx", "_controle.xlsx"))


# ---------- Assistant IA ----------
LLM_API_URL = os.environ.get("LLM_API_URL", "https://api.mistral.ai/v1/chat/completions")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
LLM_MODEL = os.environ.get("LLM_MODEL", "mistral-large-latest")

SYSTEM_PROMPT = (
    "Tu es un assistant expert du PIM Akeneo. On te donne le contexte complet d'un PM "
    "(fiche PM, produits liés avec leurs attributs et contrôles de cohérence, métadonnées des "
    "colonnes issues du dictionnaire Akeneo : libellé, type, options valides).\n"
    "L'utilisateur te demande de modifier ce PM et/ou ses produits.\n"
    "Réponds UNIQUEMENT avec un objet JSON (aucun texte avant/après) :\n"
    "{\n"
    "  \"explication\": \"résumé en français des modifications proposées\",\n"
    "  \"patch\": [ {\"sheet\": \"PM\" ou \"produits\", \"row\": <numéro de ligne Excel>, "
    "\"field\": \"<nom de colonne exact>\", \"value\": <nouvelle valeur>} ]\n"
    "}\n"
    "Règles : n'utilise que des feuilles, numéros de ligne et colonnes présents dans le contexte ; "
    "pour un attribut à options, value doit être un code d'option fourni dans le contexte "
    "(ou null pour vider) ; value null = vider la cellule ; ne modifie que ce qui est demandé.\n"
    "Tu peux aussi créer de nouveaux PM (opération de scindage) avec des items de la forme "
    "{\"op\": \"create_pm\", \"code\": \"NSPLDxxx\", \"values\": {\"family_variant\": ..., \"categories\": ..., "
    "\"EUMKG0030-fr_FR\": ..., \"EUMKG0040-fr_FR\": ..., \"EUMKG0120-fr_FR\": ...}} "
    "et déplacer des produits vers ce nouveau PM avec "
    "{\"sheet\": \"produits\", \"row\": <ligne>, \"field\": \"parent\", \"value\": \"NSPLDxxx\"}. "
    "Propose des codes NSPLD libres (supérieurs au plus grand code existant) et remplis les champs obligatoires."
)


def assistant_context(code):
    with lock:
        d = pm_detail(code)
        meta = {}
        for col, m in col_meta.items():
            if not m:
                continue
            mm = {k: m[k] for k in ("code", "label", "type", "group_label") if k in m}
            if "options" in m:
                items = list(m["options"].items())
                mm["options"] = dict(items[:80])
                if len(items) > 80:
                    mm["options_note"] = "liste tronquée, %d options au total" % len(items)
            meta[col] = mm
        return {
            "pm": d,
            "colonnes": meta,
            "labels_controles": LABELS,
            "familles": {k: v["label"] for k, v in families.items()},
        }


def enrich_patch(patch):
    out = []
    with lock:
        for it in patch:
            sheet, row, field = it.get("sheet"), it.get("row"), it.get("field")
            ws = pm_ws if sheet == "PM" else pr_ws if sheet == "produits" else None
            colmap = pm_col if sheet == "PM" else pr_col
            old = None
            if ws is not None and isinstance(row, int) and field in colmap and 2 <= row <= ws.max_row:
                old = ws.cell(row=row, column=colmap[field]).value
            out.append({"sheet": sheet, "row": row, "field": field,
                        "value": it.get("value"), "old": old})
    return out


@app.route("/api/assistant", methods=["POST"])
def api_assistant():
    data = request.get_json(force=True)
    code = s(data.get("pm"))
    message = s(data.get("message"))
    ctx = assistant_context(code)
    user_prompt = ("CONTEXTE (PM %s) :\n```json\n%s\n```\n\nDEMANDE :\n%s\n\n"
                   "Réponds UNIQUEMENT avec le JSON demandé." % (
                       code, json.dumps(ctx, ensure_ascii=False)[:150000], message))
    if LLM_API_KEY:
        try:
            req = urlreq.Request(LLM_API_URL, data=json.dumps({
                "model": LLM_MODEL,
                "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                             {"role": "user", "content": user_prompt}],
                "temperature": 0.2,
            }).encode(), headers={"Content-Type": "application/json",
                                 "Authorization": "Bearer " + LLM_API_KEY})
            with urlreq.urlopen(req, timeout=120) as resp:
                out = json.loads(resp.read())
            content = out["choices"][0]["message"]["content"]
            m = re.search(r"\{.*\}", content, re.S)
            ai = json.loads(m.group(0)) if m else {}
            return jsonify({"mode": "api", "model": LLM_MODEL,
                            "explication": ai.get("explication", ""),
                            "patch": enrich_patch(ai.get("patch", []))})
        except Exception as e:
            return jsonify({"mode": "manuel",
                            "note": "Appel API indisponible (%s) — utilisation du mode manuel." % type(e).__name__,
                            "prompt": SYSTEM_PROMPT + "\n\n" + user_prompt})
    return jsonify({"mode": "manuel",
                    "note": "Aucune clé API configurée (LLM_API_KEY) — mode manuel : copiez le prompt, "
                            "envoyez-le à votre IA, puis collez sa réponse JSON.",
                    "prompt": SYSTEM_PROMPT + "\n\n" + user_prompt})


@app.route("/api/assistant/preview", methods=["POST"])
def api_assistant_preview():
    data = request.get_json(force=True)
    patch = data.get("patch", [])
    if not isinstance(patch, list):
        return jsonify({"error": "patch doit être une liste"}), 400
    return jsonify({"patch": enrich_patch(patch)})


@app.route("/api/assistant/apply", methods=["POST"])
def api_assistant_apply():
    data = request.get_json(force=True)
    patch = data.get("patch", [])
    if not isinstance(patch, list):
        return jsonify({"error": "patch doit être une liste"}), 400
    applied, errors = [], []
    with lock:
        for it in patch:
            if it.get("op") == "create_pm":
                code = s(it.get("code"))
                values = it.get("values") or {}
                if not code or code in pm_by_code:
                    errors.append({"item": it, "error": "code PM vide ou déjà existant"})
                    continue
                r = pm_ws.max_row + 1
                pm_ws.cell(row=r, column=pm_col["code"]).value = code
                for f, v in values.items():
                    if f in pm_col:
                        pm_ws.cell(row=r, column=pm_col[f]).value = parse_value(v)
                applied.append({"op": "create_pm", "code": code, "row": r})
                continue
            sheet, row, field, value = it.get("sheet"), it.get("row"), it.get("field"), it.get("value")
            ws = pm_ws if sheet == "PM" else pr_ws if sheet == "produits" else None
            colmap = pm_col if sheet == "PM" else pr_col
            if (ws is None or not isinstance(row, int) or field not in colmap
                    or row < 2 or row > ws.max_row):
                errors.append({"item": it, "error": "feuille/ligne/champ invalide"})
                continue
            old = ws.cell(row=row, column=colmap[field]).value
            new = parse_value(value)
            ws.cell(row=row, column=colmap[field]).value = new
            applied.append({"sheet": sheet, "row": row, "field": field,
                            "old": old, "new": new})
        if applied:
            scan()
            save()
    return jsonify({"applied": applied, "errors": errors, "saved_at": last_saved})


scan()
load_refs()
scan()

if __name__ == "__main__":
    print("Fichier de travail :", XLSX)
    print("Sauvegarde de départ :", backup)
    app.run(host=HOST, port=PORT, debug=False)
