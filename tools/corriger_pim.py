#!/usr/bin/env python3
"""Correction automatique du fichier de contrôle PIM + analyse des variantes.

Usage : python3 tools/corriger_pim.py <fichier_source.xlsx> <dossier_refs> <dossier_sortie>
Produit dans <dossier_sortie> :
  - pim_source_pour_controle_CORRIGE.xlsx
  - variants_propositions.csv
  - fixlog.json (détail des corrections pour le rapport)
"""
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

import openpyxl

SRC = sys.argv[1]
REFS_DIR = sys.argv[2]
OUT_DIR = sys.argv[3]
os.makedirs(OUT_DIR, exist_ok=True)
OUT_XLSX = os.path.join(OUT_DIR, "pim_source_pour_controle_CORRIGE.xlsx")
OUT_CSV = os.path.join(OUT_DIR, "variants_propositions.csv")
OUT_LOG = os.path.join(OUT_DIR, "fixlog.json")

ACCENTS = set("éèêëàâäùûüçÉÈÊËÀÂÄÙÛÜÇ²")


def s(v):
    return "" if v is None else str(v)


def norm(t):
    t = unicodedata.normalize("NFD", s(t))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", t.lower()).strip()


def rows_of(path, sheet=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
    out = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return out


def find_ref(part):
    for f in sorted(os.listdir(REFS_DIR)):
        if f.lower().endswith(".xlsx") and part.lower() in f.lower():
            return os.path.join(REFS_DIR, f)
    return None


def header_idx(rows, key, max_scan=15):
    for i, r in enumerate(rows[:max_scan]):
        for c in r:
            if c and str(c).strip().lower() == key.lower():
                return i
    return None


def col(hdr, *names):
    for n in names:
        for i, h in enumerate(hdr):
            if h and str(h).strip().lower() == n.lower():
                return i
    for n in names:
        for i, h in enumerate(hdr):
            if h and str(h).strip().lower().startswith(n.lower()):
                return i
    return None


def lab(hdr, row, *names):
    i = col(hdr, *names)
    if i is None or i >= len(row) or row[i] in (None, ""):
        return None
    return str(row[i]).strip()


# ---------- chargement des référentiels ----------
options = defaultdict(dict)          # attr -> code -> label
label2code = defaultdict(dict)       # attr -> norm(label) -> code
p = find_ref("ATTRIBUTE OPTIONS")
rows = rows_of(p, "EXPORT")
hi = header_idx(rows, "attribute")
hdr = rows[hi]
for r in rows[hi + 1:]:
    a, c = lab(hdr, r, "attribute"), lab(hdr, r, "code")
    if not a or not c:
        continue
    l = lab(hdr, r, "label-fr_FR", "label-en_SOP", "label-en_US") or c
    options[a][c] = l
    label2code[a][norm(l)] = c

p = find_ref("DICTIONARY FAMILIES")
fam_rows = rows_of(p)
hi = header_idx(fam_rows, "code")
families = {}
hdr = fam_rows[hi]
for r in fam_rows[hi + 1:]:
    c = lab(hdr, r, "code")
    if c:
        families[c] = lab(hdr, r, "label-fr_FR", "label-en_SOP") or c

p = find_ref("LINKINGTABLE")
lk_rows = rows_of(p)
hi = header_idx(lk_rows, "variant original", max_scan=5) or 0
hdr = lk_rows[hi]
fi = col(hdr, "Variant 1 FAM")
axes_cols = [i for i, h in enumerate(hdr) if h and re.fullmatch(r"variant \d+", str(h).strip().lower())]
variants = {}
for r in lk_rows[hi + 1:]:
    v = lab(hdr, r, "variant original")
    if v:
        variants[v] = (str(r[fi]).strip() if fi is not None and fi < len(r) and r[fi] else v.split("_")[0],
                       [str(r[i]).strip() for i in axes_cols if i < len(r) and r[i]])


def axes_of(fv):
    if fv in variants:
        fam, axes = variants[fv]
        return fam or fv.split("_")[0], [a for a in axes]
    parts = fv.split("_")
    return (parts[0] if parts else ""), parts[1:]


# ---------- chargement du fichier à corriger ----------
wb = openpyxl.load_workbook(SRC)
pm_ws, pr_ws = wb["PM"], wb["produits"]
pm_hdr = [c.value for c in pm_ws[1]]
pr_hdr = [c.value for c in pr_ws[1]]
pm_col = {h: i + 1 for i, h in enumerate(pm_hdr)}
pr_col = {h: i + 1 for i, h in enumerate(pr_hdr)}

log = defaultdict(list)
def L(cat, msg):
    log[cat].append(msg)


def pm_row(r):
    return {h: pm_ws.cell(row=r, column=pm_col[h]).value for h in pm_hdr}


def pr_row(r):
    return {h: pr_ws.cell(row=r, column=pr_col[h]).value for h in pr_hdr}


def completeness(vals, hdr):
    return sum(1 for h in hdr if h != "code" and s(vals.get(h)))


# ---------- 1. dédoublonnage des lignes PM ----------
pm_by_code = defaultdict(list)
for r in range(2, pm_ws.max_row + 1):
    code = s(pm_ws.cell(row=r, column=pm_col["code"]).value)
    if code:
        pm_by_code[code].append(r)

pm_rows_to_delete = []
for code, rows in pm_by_code.items():
    if len(rows) < 2:
        continue
    vals = {r: pm_row(r) for r in rows}
    keep = max(rows, key=lambda r: (completeness(vals[r], pm_hdr), -r))
    for r in rows:
        if r == keep:
            continue
        for h in pm_hdr:
            if h == "code":
                continue
            if not s(vals[keep].get(h)) and s(vals[r].get(h)):
                pm_ws.cell(row=keep, column=pm_col[h]).value = vals[r][h]
                L("pm_dedoublonnage", f"{code} ligne {keep} : champ {h} rempli depuis la ligne {r} ({s(vals[r][h])[:60]})")
            elif s(vals[keep].get(h)) and s(vals[r].get(h)) and s(vals[keep][h]) != s(vals[r][h]):
                L("pm_conflit_conserve", f"{code} ligne {keep} conserve {h}={s(vals[keep][h])[:40]} (ligne {r}: {s(vals[r][h])[:40]})")
        pm_rows_to_delete.append(r)
    L("pm_dedoublonnage", f"{code} : {len(rows)} lignes -> 1 conservée (ligne {keep}), {len(rows)-1} supprimée(s)")

for r in sorted(pm_rows_to_delete, reverse=True):
    pm_ws.delete_rows(r)
L("pm_dedoublonnage_resume", f"{len(pm_rows_to_delete)} lignes PM dupliquées supprimées")

# ---------- relecture des index après suppression ----------
pm_by_code = defaultdict(list)
for r in range(2, pm_ws.max_row + 1):
    code = s(pm_ws.cell(row=r, column=pm_col["code"]).value)
    if code:
        pm_by_code[code].append(r)

# ---------- 2. dédoublonnage des SKU ----------
pr_by_sku = defaultdict(list)
for r in range(2, pr_ws.max_row + 1):
    sku = s(pr_ws.cell(row=r, column=pr_col["sku"]).value)
    if sku:
        pr_by_sku[sku].append(r)

pr_rows_to_delete = []
for sku, rows in pr_by_sku.items():
    if len(rows) < 2:
        continue
    vals = {r: pr_row(r) for r in rows}
    if len({tuple(s(v) for v in vals[r].values()) for r in rows}) == 1:
        keep = rows[0]
        L("sku_doublon_identique", f"{sku} : {len(rows)} lignes identiques -> {len(rows)-1} supprimée(s)")
    else:
        def score(r):
            v = vals[r]
            return (sum(1 for ch in s(v.get("EUMKG0160-fr_FR")) + s(v.get("EUINF0400")) if ch in ACCENTS),
                    completeness(v, pr_hdr))
        keep = max(rows, key=lambda r: (score(r), -r))
        diffs = [h for h in pr_hdr if len({s(vals[r].get(h)) for r in rows}) > 1]
        L("sku_doublon_conflictuel", f"{sku} : lignes différentes sur {diffs} -> ligne {keep} conservée, {len(rows)-1} supprimée(s)")
    for r in rows:
        if r != keep:
            pr_rows_to_delete.append(r)

for r in sorted(pr_rows_to_delete, reverse=True):
    pr_ws.delete_rows(r)
L("sku_doublonnage_resume", f"{len(pr_rows_to_delete)} lignes produit dupliquées supprimées ({len(pr_by_sku)} SKU concernés)")

# ---------- index produits ----------
pr_by_parent = defaultdict(list)
pr_rows = list(range(2, pr_ws.max_row + 1))
for r in pr_rows:
    pr_by_parent[s(pr_ws.cell(row=r, column=pr_col["parent"]).value)].append(r)

# ---------- 3. création des fiches PM manquantes (parents orphelins nommés) ----------
created_pm = []
for parent, rows in sorted(pr_by_parent.items()):
    if not parent or parent in pm_by_code:
        continue
    fams = Counter(s(pr_ws.cell(row=r, column=pr_col["family"]).value) for r in rows)
    fam = fams.most_common(1)[0][0]
    cats = Counter(s(pr_ws.cell(row=r, column=pr_col["categories"]).value) for r in rows)
    cat = cats.most_common(1)[0][0]
    best_axis, best_n = "", 0
    for cand in ("EUINF0233", "EUINF0400", "EUINF0240", "EUINF0250", "EUINF0260", "EUINF0140", "EULOG0040", "EULOG0050"):
        if cand not in pr_col:
            continue
        vals = {s(pr_ws.cell(row=r, column=pr_col[cand]).value) for r in rows
                if s(pr_ws.cell(row=r, column=pr_col[cand]))}
        if len(vals) > best_n:
            best_axis, best_n = cand, len(vals)
    axis = best_axis if best_n > 1 else ("EUINF0233" if any(
        s(pr_ws.cell(row=r, column=pr_col["EUINF0233"]).value) for r in rows) else "")
    fv = fam + ("_" + axis if axis else "")
    r0 = rows[0]
    name = s(pr_ws.cell(row=r0, column=pr_col["EUMKG0160-fr_FR"]).value) or s(pr_ws.cell(row=r0, column=pr_col["EUMKG0130-fr_FR"]).value) or parent
    new_r = pm_ws.max_row + 1
    pm_ws.cell(row=new_r, column=pm_col["code"]).value = parent
    pm_ws.cell(row=new_r, column=pm_col["family_variant"]).value = fv
    pm_ws.cell(row=new_r, column=pm_col["categories"]).value = cat
    pm_ws.cell(row=new_r, column=pm_col["EUMKG0030-fr_FR"]).value = name[:100]
    pm_ws.cell(row=new_r, column=pm_col["EUMKG0040-fr_FR"]).value = (name[:90] + "s") if not name.lower().endswith("s") else name[:100]
    pm_ws.cell(row=new_r, column=pm_col["EUMKG0120-fr_FR"]).value = name[:40].upper()
    pm_by_code[parent] = [new_r]
    created_pm.append(parent)
    L("pm_cree", f"{parent} : fiche PM créée (family_variant={fv}, catégories={cat}, nom depuis {s(pr_ws.cell(row=r0, column=pr_col['sku']).value)})")

# ---------- helpers ----------
def pm_of(code):
    rows = pm_by_code.get(code) or []
    return pm_row(rows[0]) if rows else None


def guess_brand(name):
    up = " " + norm(name).upper().replace(" ", " ") + " "
    cands = [c for c, l in options.get("EUINF0010", {}).items()
             if len(l) >= 3 and re.search(r"\b" + re.escape(l.upper()) + r"\b", up)]
    return cands[0] if len(cands) == 1 else None


def guess_option(attr, name):
    n = " " + norm(name) + " "
    for c, l in options.get(attr, {}).items():
        nl = norm(l)
        if not nl:
            continue
        if len(nl) >= 4:
            if nl in n:
                return c
        elif re.search(r"\b" + re.escape(nl) + r"\b", n):
            return c
    return None


# ---------- 4. remplissage des champs PM ----------
for r in range(2, pm_ws.max_row + 1):
    v = pm_row(r)
    code = s(v["code"])
    name = s(v.get("EUMKG0030-fr_FR")) or s(v.get("EUMKG0120-fr_FR")) or code
    if not s(v.get("EUMKG0120-fr_FR")) and s(v.get("EUMKG0030-fr_FR")):
        base = s(v["EUMKG0030-fr_FR"])
        short = re.split(r" - | pour |,", base)[0]
        pm_ws.cell(row=r, column=pm_col["EUMKG0120-fr_FR"]).value = short[:40].upper()
        L("pm_eumkg0120_rempli", f"{code} : EUMKG0120-fr_FR = {short[:40].upper()!r} (depuis EUMKG0030)")
    if not s(v.get("EUINF0010")):
        b = guess_brand(name)
        if b:
            pm_ws.cell(row=r, column=pm_col["EUINF0010"]).value = b
            L("pm_marque", f"{code} : EUINF0010 = {b} ({options['EUINF0010'][b]})")
    for attr in ("EUINF0020", "EUINF0030"):
        if not s(v.get(attr)):
            g = guess_option(attr, name)
            if g:
                pm_ws.cell(row=r, column=pm_col[attr]).value = g
                L("pm_option", f"{code} : {attr} = {g} ({options[attr][g][:50]})")
    for attr in ("EUINF0010", "EUINF0020", "EUINF0030"):
        val = s(v.get(attr))
        if val and val not in options.get(attr, {}):
            m = label2code.get(attr, {}).get(norm(val))
            if m:
                pm_ws.cell(row=r, column=pm_col[attr]).value = m
                L("pm_option_normalisee", f"{code} : {attr} {val[:40]!r} -> {m}")
            else:
                L("pm_option_invalide_restee", f"{code} : {attr} = {val[:60]!r} (hors dictionnaire, non reconnu)")

# ---------- 5. corrections produits ----------
OPT_PR = [h for h in pr_hdr if h in ("EUINF0110", "EUINF0140", "EUINF0400", "EUFIL130")]
color_codes = set()
for r in pr_rows:
    c = s(pr_ws.cell(row=r, column=pr_col["EUINF0233"]).value)
    if c:
        color_codes.add(c)


def euFil130_range(label):
    t = s(label)
    m = re.match(r"<\s*(\d+)", t)
    if m:
        return 0, int(m.group(1))
    m = re.search(r"(\d+)\s*-\s*(\d+)", t)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d+)\s*mm", t)
    if m:
        v = int(m.group(1))
        return v, v
    return None


eufil_ranges = [(c, euFil130_range(l)) for c, l in options.get("EUFIL130", {}).items()]
eufil_ranges = [(c, rg) for c, rg in eufil_ranges if rg]

for r in pr_rows:
    v = pr_row(r)
    sku = s(v["sku"])
    parent = s(v["parent"])
    pm = pm_of(parent) if parent else None
    # catégories / famille alignées sur le PM
    if pm:
        if s(pm.get("categories")) and s(v.get("categories")) != s(pm["categories"]):
            old = s(v["categories"])
            pr_ws.cell(row=r, column=pr_col["categories"]).value = pm["categories"]
            L("produit_categories_alignees", f"{sku} : {old} -> {pm['categories']} (PM {parent})")
        fam, _ = axes_of(s(pm.get("family_variant")))
        if fam and s(v.get("family")) and s(v["family"]) != fam:
            old = s(v["family"])
            pr_ws.cell(row=r, column=pr_col["family"]).value = fam
            L("produit_famille_alignee", f"{sku} : {old} -> {fam} (PM {parent})")
    # normalisation des valeurs d'options (correspondance libellé -> code)
    for h in OPT_PR:
        val = s(v.get(h))
        if val and val not in options.get(h, {}):
            m = label2code.get(h, {}).get(norm(val))
            if m:
                pr_ws.cell(row=r, column=pr_col[h]).value = m
                L("option_normalisee", f"{sku} : {h} {val[:40]!r} -> {m}")
            else:
                L("option_invalide_restee", f"{sku} : {h} = {val[:70]!r} (hors dictionnaire)")
    # EUFIL130 : extraire le diamètre du texte libre -> option de plage + unité
    if "EUFIL130" in pr_col and s(v.get("EUFIL130")) and not s(v.get("EUFIL130-unit")):
        txt = s(v["EUFIL130"])
        m = re.search(r"[Øø]\s*(\d+)", txt)
        if m:
            d = int(m.group(1))
            hit = next((c for c, (lo, hi) in eufil_ranges if lo <= d <= hi), None)
            if hit:
                pr_ws.cell(row=r, column=pr_col["EUFIL130"]).value = hit
                pr_ws.cell(row=r, column=pr_col["EUFIL130-unit"]).value = "MILLIMETER"
                L("eufil130_corrige", f"{sku} : Ø{d} mm -> {hit} ({options['EUFIL130'][hit]}) + unité MILLIMETER")
    # coloris EUINF0233 manquante : retrouver le code couleur dans le nom
    if not s(v.get("EUINF0233")) and s(v.get("EUMKG0160-fr_FR")):
        name_up = s(v["EUMKG0160-fr_FR"]).upper()
        for c in sorted(color_codes, key=len, reverse=True):
            if not c or not any(ch.isalpha() for ch in c):
                continue
            if re.search(r"(?<![A-Z0-9])" + re.escape(c.upper()) + r"(?![A-Z0-9])", name_up):
                pr_ws.cell(row=r, column=pr_col["EUINF0233"]).value = c
                L("coloris_retrouve", f"{sku} : EUINF0233 = {c} (extrait du nom)")
                break
        else:
            if not s(v.get("EUINF0233")):
                L("coloris_non_retrouve", f"{sku} : coloris à renseigner manuellement (nom: {s(v.get('EUMKG0160-fr_FR'))[:60]!r})")
    # champs obligatoires vides
    if not s(v.get("EUMKG0160-fr_FR")):
        src_txt = s(v.get("EUFIL130")) or s(v.get("EUMKG0130-fr_FR"))
        if pm and not src_txt:
            src_txt = s(pm.get("EUMKG0030-fr_FR"))
        if src_txt:
            title = src_txt[0].upper() + src_txt[1:]
            pr_ws.cell(row=r, column=pr_col["EUMKG0160-fr_FR"]).value = title[:255]
            L("eumkg0160_rempli", f"{sku} : titre = {title[:60]!r}")
    if not s(v.get("EUMKG0130-fr_FR")):
        src_txt = s(v.get("EUMKG0160-fr_FR"))
        if src_txt:
            pr_ws.cell(row=r, column=pr_col["EUMKG0130-fr_FR"]).value = src_txt[:120]
            L("eumkg0130_rempli", f"{sku} : description = {src_txt[:60]!r}")

# ---------- 6. analyse des variantes ----------
CANDIDATES = ["EUINF0233", "EUINF0140", "EUINF0240", "EUINF0250", "EUINF0260", "EULOG0040", "EULOG0050", "EUFIL130", "EUINF0400"]


def analyze():
    pm_by_code2 = defaultdict(list)
    for r in range(2, pm_ws.max_row + 1):
        c = s(pm_ws.cell(row=r, column=pm_col["code"]).value)
        if c:
            pm_by_code2[c].append(r)
    pr_by_parent2 = defaultdict(list)
    for r in range(2, pr_ws.max_row + 1):
        pr_by_parent2[s(pr_ws.cell(row=r, column=pr_col["parent"]).value)].append(r)
    out = []
    for code, rows in sorted(pm_by_code2.items()):
        v = pm_row(rows[0])
        fv = s(v["family_variant"])
        fam, axes = axes_of(fv)
        prods = pr_by_parent2.get(code, [])
        stats = {}
        for c in CANDIDATES:
            if c not in pr_col:
                continue
            vals = [s(pr_ws.cell(row=r, column=pr_col[c]).value) for r in prods]
            filled = [x for x in vals if x]
            stats[c] = (len(filled), len(set(filled)))
        aidx = [pr_col[a] for a in axes if a in pr_col]
        combos = Counter(tuple(s(pr_ws.cell(row=r, column=k).value) for k in aidx) for r in prods) if aidx else Counter()
        ndup = sum(n - 1 for n in combos.values() if n > 1)
        missing = sum(1 for r in prods if aidx and any(not s(pr_ws.cell(row=r, column=k).value) for k in aidx))
        disc_declared = [a for a in axes if a in stats and stats[a][1] > 1]
        non_disc = [a for a in axes if a not in disc_declared]
        free_disc = [c for c in CANDIDATES if c not in axes and c in stats and stats[c][1] > 1 and stats[c][0] >= max(1, int(0.8 * len(prods)))]
        proposal = "OK"
        if non_disc or ndup or missing:
            keep = disc_declared or []
            extra = [c for c in free_disc if c not in keep]
            sugg = keep + extra[:1]
            if not sugg and free_disc:
                sugg = free_disc[:1]
            if ndup and not extra and not keep:
                proposal = "doublons de combinaisons sans axe discriminant : fusionner les produits ou ajouter un axe"
            elif non_disc:
                proposal = ("axe %s non discriminant -> remplacer par %s" % (", ".join(non_disc), sugg[0] if sugg else "aucun axe disponible"))
                if sugg and len(sugg) > 1:
                    proposal += f" ; axes proposés : {'_'.join([fam] + sugg)}"
            elif ndup:
                proposal = f"{ndup} combinaison(s) en double -> ajouter l'axe {extra[0] if extra else '?'}"
            elif missing:
                proposal = f"{missing} produit(s) sans valeur d'axe -> à renseigner ({', '.join(axes)})"
        out.append({
            "pm": code, "family_variant": fv, "family": fam, "axes": ",".join(axes),
            "nb_produits": len(prods), "nb_combinaisons": len(combos), "doublons": ndup, "manquants": missing,
            "axes_discriminants": ",".join(disc_declared), "axes_non_discriminants": ",".join(non_disc),
            "autres_axes_utiles": ",".join(free_disc), "proposition": proposal,
        })
    return out


variant_rows = analyze()
with open(OUT_CSV, "w", encoding="utf-8") as f:
    f.write("pm;family_variant;family;axes;nb_produits;nb_combinaisons;doublons;manquants;axes_discriminants;axes_non_discriminants;autres_axes_utiles;proposition\n")
    for v in variant_rows:
        f.write(";".join(str(v[k]).replace(";", ",") for k in ("pm", "family_variant", "family", "axes", "nb_produits", "nb_combinaisons", "doublons", "manquants", "axes_discriminants", "axes_non_discriminants", "autres_axes_utiles", "proposition")) + "\n")

wb.save(OUT_XLSX)

summary = {k: len(v) for k, v in log.items()}
with open(OUT_LOG, "w", encoding="utf-8") as f:
    json.dump({"summary": summary, "details": {k: v[:400] for k, v in log.items()},
               "variants": variant_rows}, f, ensure_ascii=False, indent=1)

print("== RESUME DES CORRECTIONS ==")
for k, n in sorted(summary.items(), key=lambda x: -x[1]):
    print(f"  {k}: {n}")
print("fichier corrigé :", OUT_XLSX)
print("variantes :", OUT_CSV)
n_pb = sum(1 for v in variant_rows if v["proposition"] != "OK")
print(f"PM avec proposition de variante: {n_pb} / {len(variant_rows)}")
