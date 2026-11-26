import os
import re
import shutil
import threading
import datetime
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

app = Flask(__name__)
lock = threading.Lock()

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
            fam = s(first.get("family_variant")).split("_")[0]
            fam_bad = [s(p["values"].get("sku")) for p in prods
                       if fam and s(p["values"].get("family")) and s(p["values"].get("family")) != fam]
            if fam_bad:
                checks.append({"type": "pm_famille_mismatch", "severity": "warning",
                               "msg": "%d produit(s) avec famille ≠ PM (%s) : %s"
                                      % (len(fam_bad), fam, ", ".join(fam_bad[:8]))})
            axes = [a for a in s(first.get("family_variant")).split("_")[1:] if a in pr_col]
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
    with lock:
        if not code or code in pm_by_code:
            return jsonify({"error": "code vide ou déjà existant"}), 400
        r = pm_ws.max_row + 1
        pm_ws.cell(row=r, column=pm_col["code"]).value = code
        scan()
        save()
        return jsonify({"ok": True, "row": r, "saved_at": last_saved})


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


scan()

if __name__ == "__main__":
    print("Fichier de travail :", XLSX)
    print("Sauvegarde de départ :", backup)
    app.run(host=HOST, port=PORT, debug=False)
