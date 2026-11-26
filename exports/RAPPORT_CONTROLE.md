# Rapport de contrôle PIM — relecture complète et corrections

Fichier analysé : `pim_source_pour_controle.xlsx` (nouvelle version) — feuilles `PM` et `produits`.
Référentiels Akeneo utilisés : dossier `refs/` (345 attributs, 7 445 options, 886 catégories, familles, 183 variantes dont addendum local FAM007).
Corrections produites par `tools/corriger_pim.py` dans `exports/pim_source_pour_controle_CORRIGE.xlsx`.

## 1. Bilan avant / après (contrôles automatiques de l'outil)

| Contrôle | Avant | Après |
|---|---:|---:|
| Valeur d'option hors dictionnaire Akeneo | 1349 | 1217 |
| SKU en double | 204 | 0 |
| Combinaison d'axes de variante en double dans un PM | 67 | 46 |
| Produit orphelin (parent sans fiche PM) | 60 | 28 |
| Produit sans valeur d'axe de variante | 44 | 45 |
| PM en double (conflictuel) | 37 | 0 |
| Produit désactivé (info) | 30 | 28 |
| PM en double (identique) | 22 | 0 |
| Champ obligatoire vide | 14 | 0 |
| Catégories produit ≠ PM | 12 | 0 |
| Valeur sans unité | 10 | 2 |
| Famille produit ≠ PM | 8 | 0 |
| **Total** | **1857** | **1366** |

Les anomalies restantes sont majoritairement des **valeurs d'option en texte libre** (voir §3) et des **problèmes de variantes** (voir §4).

## 2. Corrections appliquées automatiquement

- **Lignes PM dupliquées fusionnées** : 59 codes PM avaient plusieurs lignes ; la ligne la plus complète est conservée, les champs vides sont complétés depuis les autres lignes, les lignes en trop sont supprimées (130 conflits de valeurs conservés sur la ligne principale et consignés dans `fixlog.json`).
  Exemples :
  - NSPLD052 : 2 lignes -> 1 conservée (ligne 2), 1 supprimée(s)
  - NSPLD021 : 2 lignes -> 1 conservée (ligne 6), 1 supprimée(s)
  - NSPLD027 : 3 lignes -> 1 conservée (ligne 9), 2 supprimée(s)
- **SKU dupliqués supprimés** : 98 doublons strictement identiques + 4 doublons conflictuels (différence de typo dans EUINF0400 ; la version la mieux accentuée est conservée).
  Exemples :
  - LD_00002623 : lignes différentes sur ['EUINF0400'] -> ligne 327 conservée, 1 supprimée(s)
  - LD_00002688 : lignes différentes sur ['EUINF0400'] -> ligne 246 conservée, 1 supprimée(s)
  - LD_00002759 : lignes différentes sur ['EUINF0400'] -> ligne 375 conservée, 1 supprimée(s)
  - LD_00002882 : lignes différentes sur ['EUINF0400'] -> ligne 413 conservée, 1 supprimée(s)
- **Fiches PM créées pour les parents orphelins** : NSPLD046, NSPLD047, NSPLD048, NSPLD461 — family_variant déduit des produits (famille + axe coloris EUINF0233 ou sélection EUINF0400), catégories = catégorie majoritaire des produits, nom repris du premier produit.
- **Catégories/familles des produits alignées sur le PM** : 54 catégories et 53 familles corrigées.
- **Champs obligatoires remplis** : PM `EUMKG0120-fr_FR` (Product Model name) : 47 remplis depuis `EUMKG0030-fr_FR` (forme courte en capitales) ; produits `EUMKG0160-fr_FR` (titre) : 10 et `EUMKG0130-fr_FR` (description) : 4 remplis depuis le texte disponible (ex. description EUFIL130 des disques diamant). **0 champ obligatoire vide restant.**
- **Marque `EUINF0010` déduite du nom** quand elle correspond exactement à une option du dictionnaire : 8 PM.
- **`EUINF0020` (Type) / `EUINF0030` (Application) des PM** remplis par correspondance exacte du libellé d'option dans le nom : 92 valeurs.
- **Valeurs d'options normalisées en codes** (correspondance libellé → code) : 0 produit(s) et 0 PM.
- **Coloris `EUINF0233` retrouvés dans le nom du produit** (codes couleur alphanumériques uniquement, ex. TW01, SO11, CHA) : 17 produits. Les codes purement numériques (15, 40, 300…) n'ont PAS été devinés (trop ambigu avec les dimensions) : les vrais manquants d'axe figurent au contrôle « produit sans valeur d'axe de variante » (45 après correction).
- **`EUFIL130` (épaisseur/hauteur)** : diamètre extrait du texte libre (Ø xxx mm) et converti en option de plage + unité MILLIMETER : 8 produits (2 restants sans Ø exploitable).

## 3. Reste à faire manuellement

- **1217 valeurs d'option hors dictionnaire**, très majoritairement `EUINF0400` (« Selection ») renseigné en **texte libre** (ex. « SIKAPLAN 12G GRIS FONCE RAL 7012 1.54X20ML ») au lieu d'un code `EUINF0400_OPTIONxxxx`. Piste : soit nettoyer au cas par cas dans l'outil (autocomplétion par libellé), soit requalifier l'attribut dans Akeneo s'il doit accepter du texte libre. Autre piste structurelle : `EUINF0140` est déclaré `simpleselect` dans le dictionnaire mais utilisé en **numérique + unité** dans le fichier — à clarifier (metric vs simpleselect).
- **28 produits sans parent** (parent vide) : produits seuls (pare-gravier, cartouches, vis…) — à rattacher à un PM ou à laisser en produits standalone selon votre modélisation Akeneo.
- **2 valeurs sans unité restantes** (EUFIL130 sans Ø exploitable).
- **28 produits désactivés** (enabled = 0) : à vérifier volontairement.

## 4. Analyse des variantes (PM / axes)

384 PM analysés ; **104 PM ont une proposition de correction d'axes**. Détail PM par PM dans `exports/variants_propositions.csv`.

### 4.1 Agrégation par `family_variant`

| family_variant | axes | PM | produits | axes non discriminants | combinaisons en double | valeurs d'axe manquantes |
|---|---|---:|---:|---:|---:|---:|
| `FAM007_EUINF0400` | EUINF0400 | 46 | 229 | 12 PM | 42 | 55 |
| `FAM007_EUINF0233` | EUINF0233 | 33 | 127 | 1 PM | 0 | 0 |
| `FAM004_EUINF0400` | EUINF0400 | 25 | 173 | 1 PM | 36 | 45 |
| `FAM003_EUINF0260` | EUINF0260 | 24 | 116 | 3 PM | 1 | 0 |
| `FAM001_EUINF0400` | EUINF0400 | 22 | 165 | 2 PM | 33 | 38 |
| `FAM006_EUINF0400` | EUINF0400 | 18 | 55 | 6 PM | 15 | 20 |
| `FAM007_EUINF0290` | EUINF0290 | 16 | 63 | 16 PM | 0 | 0 |
| `FAM005_EUINF0400` | EUINF0400 | 15 | 70 | 3 PM | 6 | 10 |
| `FAM007_EUINF0260` | EUINF0260 | 14 | 106 | 0 PM | 0 | 0 |
| `FAM001_EUINF0233` | EUINF0233 | 13 | 42 | 1 PM | 2 | 0 |
| `FAM007_EUINF0240` | EUINF0240 | 13 | 32 | 0 PM | 0 | 0 |
| `FAM007_EUINF0140_EUINF0240` | EUINF0140,EUINF0240 | 13 | 60 | 9 PM | 0 | 0 |
| `FAM005_EUINF0140` | EUINF0140 | 10 | 43 | 0 PM | 2 | 0 |
| `FAM001_EUINF0250` | EUINF0250 | 9 | 25 | 0 PM | 2 | 0 |
| `FAM007_EUINF0140` | EUINF0140 | 9 | 46 | 0 PM | 0 | 0 |
| `FAM005_EUINF0140_EUINF0240` | EUINF0140,EUINF0240 | 8 | 47 | 1 PM | 3 | 3 |
| `FAM001_EUINF0240_EUINF0250` | EUINF0240,EUINF0250 | 8 | 36 | 5 PM | 0 | 0 |
| `FAM004_EUINF0240` | EUINF0240 | 7 | 23 | 2 PM | 6 | 7 |
| `FAM001_EULOG0050` | EULOG0050 | 7 | 16 | 0 PM | 0 | 0 |
| `FAM007_EUINF0250` | EUINF0250 | 7 | 26 | 0 PM | 0 | 0 |
| `FAM004_EUINF0140` | EUINF0140 | 6 | 13 | 1 PM | 0 | 0 |
| `FAM007_EUINF0240_EUINF0250` | EUINF0240,EUINF0250 | 6 | 38 | 3 PM | 0 | 0 |
| `FAM005_EUINF0233_EUINF0140_EUINF0240` | EUINF0233,EUINF0140,EUINF0240 | 5 | 89 | 1 PM | 0 | 0 |
| `FAM007_EULOG0050` | EULOG0050 | 4 | 11 | 0 PM | 0 | 0 |
| `FAM001_EULOG0040` | EULOG0040 | 3 | 7 | 0 PM | 0 | 0 |
| `FAM001_EUINF0260` | EUINF0260 | 3 | 7 | 0 PM | 0 | 0 |
| `FAM001_EUINF0233_EULOG0040` | EUINF0233,EULOG0040 | 3 | 17 | 0 PM | 0 | 0 |
| `FAM007_EULOG0040` | EULOG0040 | 3 | 6 | 0 PM | 0 | 0 |
| `FAM005_EULOG0040` | EULOG0040 | 2 | 4 | 0 PM | 0 | 0 |
| `FAM002_EUINF0400` | EUINF0400 | 2 | 18 | 0 PM | 9 | 10 |
| `FAM003_EUINF0240_EUINF0250_EUINF0260` | EUINF0240,EUINF0250,EUINF0260 | 2 | 6 | 2 PM | 0 | 0 |
| `FAM004_EUINF0140_EUINF0260` | EUINF0140,EUINF0260 | 2 | 8 | 0 PM | 0 | 3 |
| `FAM001_EUINF0140` | EUINF0140 | 2 | 10 | 0 PM | 0 | 0 |
| `FAM006_EUINF0240` | EUINF0240 | 2 | 3 | 1 PM | 0 | 0 |
| `FAM004_EUINF0140_EUINF0240` | EUINF0140,EUINF0240 | 2 | 32 | 0 PM | 0 | 0 |
| `FAM007_EUINF0240_EUINF0250_EUINF0260` | EUINF0240,EUINF0250,EUINF0260 | 2 | 21 | 1 PM | 0 | 0 |
| `FAM007_EUINF0233_EUINF0400` | EUINF0233,EUINF0400 | 2 | 9 | 2 PM | 4 | 9 |
| `FAM004_EUINF0250` | EUINF0250 | 2 | 4 | 0 PM | 0 | 0 |
| `FAM007_EUINF0233_EULOG0040` | EUINF0233,EULOG0040 | 2 | 9 | 1 PM | 0 | 0 |
| `FAM005_EUINF0233` | EUINF0233 | 1 | 15 | 0 PM | 0 | 0 |
| `FAM006_EUINF0140` | EUINF0140 | 1 | 2 | 0 PM | 0 | 0 |
| `FAM005_EUINF0240` | EUINF0240 | 1 | 9 | 0 PM | 4 | 0 |
| `FAM005_EUINF0250` | EUINF0250 | 1 | 12 | 0 PM | 2 | 0 |
| `FAM005_EUINF0260` | EUINF0260 | 1 | 6 | 0 PM | 0 | 0 |
| `FAM001_EUINF0240` | EUINF0240 | 1 | 2 | 0 PM | 0 | 0 |
| `FAM004_EULOG0040` | EULOG0040 | 1 | 2 | 0 PM | 0 | 0 |
| `FAM006_EULOG0040` | EULOG0040 | 1 | 3 | 0 PM | 0 | 0 |
| `FAM004_EUINF0233` | EUINF0233 | 1 | 3 | 0 PM | 0 | 0 |
| `FAM007_EUFIL130_EUINF0250` | EUFIL130,EUINF0250 | 1 | 9 | 0 PM | 5 | 0 |
| `FAM007_EUINF0233_EUINF0250` | EUINF0233,EUINF0250 | 1 | 7 | 1 PM | 0 | 5 |
| `FAM007_EUINF0233_EUINF0240_EUINF0250_EUINF0260` | EUINF0233,EUINF0240,EUINF0250,EUINF0260 | 1 | 4 | 1 PM | 0 | 0 |

### 4.2 Constats principaux

- **`EUINF0400` est l'axe le plus utilisé (128 PM) mais ne discrimine les produits que dans ~80 % d'entre eux** (76 PM au total ont au moins un axe non discriminant). Quand l'axe déclaré est vide ou constant, les produits du PM forment un seul bloc de variantes : soit renseigner l'axe, soit le remplacer.
- **Le coloris `EUINF0233` est le meilleur candidat de remplacement/ajout** : utilisé dans 61 PM, discriminant dans 59/61. Exemple : `FAM007_EUINF0400` (55 PM) → pour les PM où EUINF0400 est vide/constant mais où les coloris EUINF0233 varient, la variante cohérente est `FAM007_EUINF0233` voire `FAM007_EUINF0400_EUINF0233` (2 axes) quand les deux varient.
- **Dimensions (`EUINF0140`, `EUINF0240`, `EUINF0250`, `EUINF0260`)** : bonnes candidates d'axes quand elles varient dans le PM (ex. `FAM001_EUINF0240_EUINF0250` existe déjà et est cohérent). 14/70 PM en `..._EUINF0240` ont un EUINF0240 constant → axe à remplacer.
- **172 combinaisons d'axes en double** dans 46 PM : deux produits du même PM ont la même combinaison d'axes → soit ce sont des vrais doublons de produits, soit il manque un axe discriminant (souvent le coloris EUINF0233 ou une dimension). Voir colonne `proposition` du CSV.
- **205 produits sans valeur d'axe** : à renseigner impérativement, sinon Akeneo refusera la création des variantes.
- **13 PM n'ont qu'un seul produit** : candidats à la fusion dans un PM existant de même famille + mêmes catégories (moins de PM à maintenir). Liste dans le CSV (`nb_produits = 1`).

### 4.3 Top 15 des PM les plus problématiques (variantes)

| PM | family_variant | produits | combinaisons | doublons | manquants | proposition |
|---|---|---:|---:|---:|---:|---|
| NSPLD698 | `FAM007_EUINF0400` | 18 | 1 | 17 | 18 | doublons de combinaisons sans axe discriminant : fusionner les produits ou ajouter un axe |
| NSPLD795 | `FAM004_EUINF0400` | 22 | 8 | 14 | 15 | 14 combinaison(s) en double -> ajouter l'axe ? |
| NSPLD475 | `FAM001_EUINF0400` | 19 | 8 | 11 | 12 | 11 combinaison(s) en double -> ajouter l'axe EUINF0240 |
| NSPLD555 | `FAM001_EUINF0400` | 17 | 7 | 10 | 11 | 10 combinaison(s) en double -> ajouter l'axe EULOG0040 |
| NSPLD498 | `FAM002_EUINF0400` | 14 | 5 | 9 | 10 | 9 combinaison(s) en double -> ajouter l'axe ? |
| NSPLD692 | `FAM001_EUINF0400` | 10 | 1 | 9 | 10 | doublons de combinaisons sans axe discriminant : fusionner les produits ou ajouter un axe |
| NSPLD821 | `FAM006_EUINF0400` | 9 | 1 | 8 | 9 | doublons de combinaisons sans axe discriminant : fusionner les produits ou ajouter un axe |
| NSPLD993 | `FAM007_EUINF0400` | 9 | 1 | 8 | 9 | axe EUINF0400 non discriminant -> remplacer par EUINF0240 |
| NSPLD501 | `FAM004_EUINF0240` | 8 | 2 | 6 | 7 | axe EUINF0240 non discriminant -> remplacer par EUINF0400 |
| NSPLD986 | `FAM007_EUINF0400` | 6 | 1 | 5 | 6 | axe EUINF0400 non discriminant -> remplacer par EUINF0250 |
| NSPLD056 | `FAM004_EUINF0400` | 9 | 5 | 4 | 5 | 4 combinaison(s) en double -> ajouter l'axe ? |
| NSPLD549 | `FAM004_EUINF0400` | 9 | 5 | 4 | 5 | 4 combinaison(s) en double -> ajouter l'axe EUINF0140 |
| NSPLD730 | `FAM004_EUINF0400` | 11 | 7 | 4 | 5 | 4 combinaison(s) en double -> ajouter l'axe EUINF0240 |
| NSPLD988 | `FAM007_EUINF0233_EUINF0400` | 6 | 3 | 3 | 6 | axe EUINF0400 non discriminant -> remplacer par EUINF0233 |
| NSPLD502 | `FAM004_EUINF0400` | 12 | 9 | 3 | 4 | 3 combinaison(s) en double -> ajouter l'axe EUINF0250 |

### 4.4 Recommandation générale

1. Renseigner d'abord les **valeurs d'axe manquantes** (bloquant pour Akeneo).
2. Pour chaque PM, garder comme axes uniquement les attributs **qui varient réellement** parmi ses produits ; ajouter `EUINF0233` (coloris) dès qu'il varie — c'est l'axe le plus discriminant du fichier.
3. Fusionner les **combinaisons en double** (vrais doublons de produits) ou ajouter l'axe qui les sépare.
4. Revoir les PM mono-produit (fusion) et les axes `EUINF0400` non discriminants (remplacer par un axe utile ou compléter la sélection).

## 5. Fichiers produits

- `exports/pim_source_pour_controle_CORRIGE.xlsx` — fichier corrigé (PM : 452 lignes → 384 codes uniques ; produits : 2016 → 1914 lignes sans doublon)
- `exports/variants_propositions.csv` — analyse variantes PM par PM + proposition d'axes
- `exports/fixlog.json` — journal détaillé de chaque correction (catégorie + exemple)
- `refs/LOCAL - ADDENDUM.xlsx` — addendum local (FAM007 + variantes observées) fusionné automatiquement par l'outil

Rejouer : `python3 tools/corriger_pim.py <source.xlsx> refs exports` puis `python3 tools/rapport.py exports/fixlog.json exports/variants_propositions.csv <checks_avant.json> <checks_apres.json> exports/RAPPORT_CONTROLE.md`
