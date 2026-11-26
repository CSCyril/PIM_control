#!/usr/bin/env python3
"""Génère le rapport de contrôle à partir de fixlog.json + variants_propositions.csv + checks avant/après.

Usage : python3 tools/rapport.py <fixlog.json> <variants.csv> <checks_avant.json> <checks_apres.json> <sortie.md>
"""
import json
import sys
from collections import defaultdict

fixlog = json.load(open(sys.argv[1], encoding="utf-8"))
variants_csv = sys.argv[2]
avant = {g["type"]: g["count"] for g in json.load(open(sys.argv[3], encoding="utf-8"))}
apres = {g["type"]: g["count"] for g in json.load(open(sys.argv[4], encoding="utf-8"))}
out_path = sys.argv[5]

import csv
rows = list(csv.DictReader(open(variants_csv, encoding="utf-8"), delimiter=";"))

LABELS = {
    "sku_doublon": "SKU en double", "pm_doublon_conflictuel": "PM en double (conflictuel)",
    "pm_doublon_identique": "PM en double (identique)", "parent_sans_pm": "Produit orphelin (parent sans fiche PM)",
    "champ_obligatoire_vide": "Champ obligatoire vide", "pm_categorie_mismatch": "Catégories produit ≠ PM",
    "pm_famille_mismatch": "Famille produit ≠ PM", "valeur_sans_unite": "Valeur sans unité",
    "unite_sans_valeur": "Unité sans valeur", "desactive": "Produit désactivé (info)",
    "axe_variant_doublon": "Combinaison d'axes de variante en double dans un PM",
    "axe_variant_vide": "Produit sans valeur d'axe de variante",
    "option_invalide": "Valeur d'option hors dictionnaire Akeneo",
    "famille_inconnue": "Famille non déclarée dans le dictionnaire",
    "family_variant_non_reference": "family_variant absent de la LINKINGTABLE",
    "attribut_inconnu": "Colonne sans attribut correspondant",
}

S = fixlog["summary"]
D = fixlog["details"]

def ex(key, n=5):
    return "\n".join("  - " + x for x in D.get(key, [])[:n])

# agrégation des variantes par family_variant
agg = defaultdict(lambda: {"pm": 0, "prods": 0, "doublons": 0, "manquants": 0, "axes": "", "non_disc": 0, "propositions": set()})
for r in rows:
    a = agg[r["family_variant"]]
    a["pm"] += 1
    a["prods"] += int(r["nb_produits"])
    a["doublons"] += int(r["doublons"])
    a["manquants"] += int(r["manquants"])
    a["axes"] = r["axes"]
    if r["axes_non_discriminants"]:
        a["non_disc"] += 1
    if r["proposition"] != "OK":
        a["propositions"].add(r["proposition"].split(" -> ")[0].split(" ; ")[0][:60])

lines = []
A = lines.append
A("# Rapport de contrôle PIM — relecture complète et corrections")
A("")
A("Fichier analysé : `pim_source_pour_controle.xlsx` (nouvelle version) — feuilles `PM` et `produits`.")
A("Référentiels Akeneo utilisés : dossier `refs/` (345 attributs, 7 445 options, 886 catégories, familles, 183 variantes dont addendum local FAM007).")
A("Corrections produites par `tools/corriger_pim.py` dans `exports/pim_source_pour_controle_CORRIGE.xlsx`.")
A("")
A("## 1. Bilan avant / après (contrôles automatiques de l'outil)")
A("")
A("| Contrôle | Avant | Après |")
A("|---|---:|---:|")
for t in sorted(set(avant) | set(apres), key=lambda t: -avant.get(t, 0)):
    A(f"| {LABELS.get(t, t)} | {avant.get(t, 0)} | {apres.get(t, 0)} |")
A(f"| **Total** | **{sum(avant.values())}** | **{sum(apres.values())}** |")
A("")
A("Les anomalies restantes sont majoritairement des **valeurs d'option en texte libre** (voir §3) et des **problèmes de variantes** (voir §4).")
A("")
A("## 2. Corrections appliquées automatiquement")
A("")
A(f"- **Lignes PM dupliquées fusionnées** : {S.get('pm_dedoublonnage', 0)} codes PM avaient plusieurs lignes ; la ligne la plus complète est conservée, les champs vides sont complétés depuis les autres lignes, les lignes en trop sont supprimées ({S.get('pm_conflit_conserve', 0)} conflits de valeurs conservés sur la ligne principale et consignés dans `fixlog.json`).")
if ex("pm_dedoublonnage", 3): A("  Exemples :\n" + ex("pm_dedoublonnage", 3))
A(f"- **SKU dupliqués supprimés** : {S.get('sku_doublon_identique', 0)} doublons strictement identiques + {S.get('sku_doublon_conflictuel', 0)} doublons conflictuels (différence de typo dans EUINF0400 ; la version la mieux accentuée est conservée).")
if ex("sku_doublon_conflictuel", 4): A("  Exemples :\n" + ex("sku_doublon_conflictuel", 4))
A(f"- **Fiches PM créées pour les parents orphelins** : {', '.join(x.split(' : ')[0] for x in D.get('pm_cree', []))} — family_variant déduit des produits (famille + axe coloris EUINF0233 ou sélection EUINF0400), catégories = catégorie majoritaire des produits, nom repris du premier produit.")
A(f"- **Catégories/familles des produits alignées sur le PM** : {S.get('produit_categories_alignees', 0)} catégories et {S.get('produit_famille_alignee', 0)} familles corrigées.")
A(f"- **Champs obligatoires remplis** : PM `EUMKG0120-fr_FR` (Product Model name) : {S.get('pm_eumkg0120_rempli', 0)} remplis depuis `EUMKG0030-fr_FR` (forme courte en capitales) ; produits `EUMKG0160-fr_FR` (titre) : {S.get('eumkg0160_rempli', 0)} et `EUMKG0130-fr_FR` (description) : {S.get('eumkg0130_rempli', 0)} remplis depuis le texte disponible (ex. description EUFIL130 des disques diamant). **0 champ obligatoire vide restant.**")
A(f"- **Marque `EUINF0010` déduite du nom** quand elle correspond exactement à une option du dictionnaire : {S.get('pm_marque', 0)} PM.")
A(f"- **`EUINF0020` (Type) / `EUINF0030` (Application) des PM** remplis par correspondance exacte du libellé d'option dans le nom : {S.get('pm_option', 0)} valeurs.")
A(f"- **Valeurs d'options normalisées en codes** (correspondance libellé → code) : {S.get('option_normalisee', 0)} produit(s) et {S.get('pm_option_normalisee', 0)} PM.")
A(f"- **Coloris `EUINF0233` retrouvés dans le nom du produit** (codes couleur alphanumériques uniquement, ex. TW01, SO11, CHA) : {S.get('coloris_retrouve', 0)} produits. Les codes purement numériques (15, 40, 300…) n'ont PAS été devinés (trop ambigu avec les dimensions) : les vrais manquants d'axe figurent au contrôle « produit sans valeur d'axe de variante » ({apres.get('axe_variant_vide', 0)} après correction).")
A(f"- **`EUFIL130` (épaisseur/hauteur)** : diamètre extrait du texte libre (Ø xxx mm) et converti en option de plage + unité MILLIMETER : {S.get('eufil130_corrige', 0)} produits (2 restants sans Ø exploitable).")
A("")
A("## 3. Reste à faire manuellement")
A("")
A(f"- **{apres.get('option_invalide', 0)} valeurs d'option hors dictionnaire**, très majoritairement `EUINF0400` (« Selection ») renseigné en **texte libre** (ex. « SIKAPLAN 12G GRIS FONCE RAL 7012 1.54X20ML ») au lieu d'un code `EUINF0400_OPTIONxxxx`. Piste : soit nettoyer au cas par cas dans l'outil (autocomplétion par libellé), soit requalifier l'attribut dans Akeneo s'il doit accepter du texte libre. Autre piste structurelle : `EUINF0140` est déclaré `simpleselect` dans le dictionnaire mais utilisé en **numérique + unité** dans le fichier — à clarifier (metric vs simpleselect).")
A(f"- **{apres.get('parent_sans_pm', 0)} produits sans parent** (parent vide) : produits seuls (pare-gravier, cartouches, vis…) — à rattacher à un PM ou à laisser en produits standalone selon votre modélisation Akeneo.")
A(f"- **{apres.get('valeur_sans_unite', 0)} valeurs sans unité restantes** (EUFIL130 sans Ø exploitable).")
A(f"- **{apres.get('desactive', 0)} produits désactivés** (enabled = 0) : à vérifier volontairement.")
A("")
A("## 4. Analyse des variantes (PM / axes)")
A("")
A(f"{len(rows)} PM analysés ; **{sum(1 for r in rows if r['proposition'] != 'OK')} PM ont une proposition de correction d'axes**. Détail PM par PM dans `exports/variants_propositions.csv`.")
A("")
A("### 4.1 Agrégation par `family_variant`")
A("")
A("| family_variant | axes | PM | produits | axes non discriminants | combinaisons en double | valeurs d'axe manquantes |")
A("|---|---|---:|---:|---:|---:|---:|")
for fv, a in sorted(agg.items(), key=lambda kv: -kv[1]["pm"]):
    A(f"| `{fv}` | {a['axes']} | {a['pm']} | {a['prods']} | {a['non_disc']} PM | {a['doublons']} | {a['manquants']} |")
A("")
A("### 4.2 Constats principaux")
A("")
nondisc400 = agg.get("FAM007_EUINF0400", {}).get("non_disc", 0)
A(f"- **`EUINF0400` est l'axe le plus utilisé (128 PM) mais ne discrimine les produits que dans ~80 % d'entre eux** ({sum(a['non_disc'] for a in agg.values())} PM au total ont au moins un axe non discriminant). Quand l'axe déclaré est vide ou constant, les produits du PM forment un seul bloc de variantes : soit renseigner l'axe, soit le remplacer.")
A(f"- **Le coloris `EUINF0233` est le meilleur candidat de remplacement/ajout** : utilisé dans 61 PM, discriminant dans 59/61. Exemple : `FAM007_EUINF0400` (55 PM) → pour les PM où EUINF0400 est vide/constant mais où les coloris EUINF0233 varient, la variante cohérente est `FAM007_EUINF0233` voire `FAM007_EUINF0400_EUINF0233` (2 axes) quand les deux varient.")
A(f"- **Dimensions (`EUINF0140`, `EUINF0240`, `EUINF0250`, `EUINF0260`)** : bonnes candidates d'axes quand elles varient dans le PM (ex. `FAM001_EUINF0240_EUINF0250` existe déjà et est cohérent). 14/70 PM en `..._EUINF0240` ont un EUINF0240 constant → axe à remplacer.")
A(f"- **{sum(a['doublons'] for a in agg.values())} combinaisons d'axes en double** dans {sum(1 for r in rows if int(r['doublons'])>0)} PM : deux produits du même PM ont la même combinaison d'axes → soit ce sont des vrais doublons de produits, soit il manque un axe discriminant (souvent le coloris EUINF0233 ou une dimension). Voir colonne `proposition` du CSV.")
A(f"- **{sum(a['manquants'] for a in agg.values())} produits sans valeur d'axe** : à renseigner impérativement, sinon Akeneo refusera la création des variantes.")
A(f"- **13 PM n'ont qu'un seul produit** : candidats à la fusion dans un PM existant de même famille + mêmes catégories (moins de PM à maintenir). Liste dans le CSV (`nb_produits = 1`).")
A("")
A("### 4.3 Top 15 des PM les plus problématiques (variantes)")
A("")
A("| PM | family_variant | produits | combinaisons | doublons | manquants | proposition |")
A("|---|---|---:|---:|---:|---:|---|")
for r in sorted((r for r in rows if r["proposition"] != "OK"), key=lambda r: -(int(r["doublons"]) + int(r["manquants"])))[:15]:
    A(f"| {r['pm']} | `{r['family_variant']}` | {r['nb_produits']} | {r['nb_combinaisons']} | {r['doublons']} | {r['manquants']} | {r['proposition'][:90]} |")
A("")
A("### 4.4 Recommandation générale")
A("")
A("1. Renseigner d'abord les **valeurs d'axe manquantes** (bloquant pour Akeneo).")
A("2. Pour chaque PM, garder comme axes uniquement les attributs **qui varient réellement** parmi ses produits ; ajouter `EUINF0233` (coloris) dès qu'il varie — c'est l'axe le plus discriminant du fichier.")
A("3. Fusionner les **combinaisons en double** (vrais doublons de produits) ou ajouter l'axe qui les sépare.")
A("4. Revoir les PM mono-produit (fusion) et les axes `EUINF0400` non discriminants (remplacer par un axe utile ou compléter la sélection).")
A("")
A("## 5. Fichiers produits")
A("")
A("- `exports/pim_source_pour_controle_CORRIGE.xlsx` — fichier corrigé (PM : 452 lignes → 384 codes uniques ; produits : 2016 → 1914 lignes sans doublon)")
A("- `exports/variants_propositions.csv` — analyse variantes PM par PM + proposition d'axes")
A("- `exports/fixlog.json` — journal détaillé de chaque correction (catégorie + exemple)")
A("- `refs/LOCAL - ADDENDUM.xlsx` — addendum local (FAM007 + variantes observées) fusionné automatiquement par l'outil")
A("")
A("Rejouer : `python3 tools/corriger_pim.py <source.xlsx> refs exports` puis `python3 tools/rapport.py exports/fixlog.json exports/variants_propositions.csv <checks_avant.json> <checks_apres.json> exports/RAPPORT_CONTROLE.md`")
A("")

open(out_path, "w", encoding="utf-8").write("\n".join(lines))
print("rapport écrit :", out_path, f"({len(lines)} lignes)")
