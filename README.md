# Contrôle PIM — PM par PM

Interface web locale pour parcourir un fichier PIM (`pim_source_pour_controle.xlsx`)
PM par PM, voir les attributs de chaque PM et de ses produits liés, vérifier la
cohérence automatiquement et corriger les données en direct (les modifications sont
enregistrées dans le fichier Excel à chaque changement).

## Lancement

```bash
cd pim-controle
pip install flask openpyxl   # si pas déjà installé
cp /chemin/de/votre/pim_source_pour_controle.xlsx data/   # votre fichier source
python3 app.py
```

Puis ouvrir **http://127.0.0.1:8765** dans un navigateur.

Le serveur n'écoute que sur `127.0.0.1` (boucle locale uniquement).

## Fichier de travail

- Par défaut, l'application travaille sur **votre copie** placée dans `data/pim_source_pour_controle.xlsx`
  (le fichier source d'origine n'est jamais modifié).
- Une **sauvegarde de départ** est créée automatiquement dans `data/sauvegarde_*.xlsx`.
- Pour éditer un autre fichier : `PIM_FILE=/chemin/vers/fichier.xlsx python3 app.py`
- Le bouton **⬇ Exporter Excel** télécharge le fichier corrigé, prêt à remplacer la source.

## Fonctionnalités

- **Navigation PM par PM** : liste filtrable/recherchable, boutons Précédent / Suivant,
  et « ⚠ Suivant avec anomalie » pour ne parcourir que les PM problématiques.
- **Fiche PM** : tous les attributs de la feuille `PM`, éditables (clic dans un champ).
  Si un code PM existe en plusieurs lignes (doublons), chaque ligne est accessible par onglet
  et peut être supprimée individuellement. Pour un parent sans fiche PM, un bouton permet
  de créer la ligne PM.
- **Produits liés** : tableau de tous les produits dont `parent` = code du PM, avec tous
  leurs attributs éditables **cellule par cellule** (clic → modification → Entrée/blur pour
  enregistrer). Listes déroulantes d'aide pour les champs existants (parent = codes PM,
  unités, familles, etc.). Ajout / suppression de produit possible.
- **Contrôles de cohérence automatiques** (par PM et par produit) :
  PM en double (identique ou conflictuel), PM sans produit, catégories/famille produit ≠ PM,
  axes de variante (`family_variant`) vides ou en double, produit orphelin (parent sans PM),
  SKU en double, produit désactivé, valeur sans unité / unité sans valeur, champs
  obligatoires vides (remplis à ≥ 90 % dans la source).
- **Référentiels Akeneo** (dossier `refs/`, voir ci-dessous) : autocomplétion par options
  officielles (code + libellé), libellés/types/groupes d'attributs en en-tête de colonnes,
  axes de variante résolus via la LINKINGTABLE, et contrôles supplémentaires :
  valeur d'option hors dictionnaire, famille non déclarée, `family_variant` absent de la
  LINKINGTABLE, colonne PIM sans attribut correspondant.
- **Tableau des anomalies** global : toutes les anomalies regroupées par type, cliquables
  pour ouvrir directement le PM concerné.
- **Sauvegarde automatique** dans le fichier Excel à chaque modification, plus bouton
  💾 Enregistrer et ⬇ Exporter Excel.

## Référentiels Akeneo (`refs/`)

Le dossier `refs/` contient les exports du dictionnaire Akeneo (10 fichiers `.xlsx`) :

| Fichier | Usage dans l'outil |
|---|---|
| DICTIONARY ATTRIBUTES | libellé, type, groupe, niveau (PM/produit) de chaque attribut → en-têtes de colonnes et fiches |
| DICTIONARY ATTRIBUTE OPTIONS | options officielles par attribut → **autocomplétion code + libellé** et contrôle « valeur hors dictionnaire » |
| DICTIONARY CATEGORIES | codes catégories + libellés → autocomplétion du champ `categories` |
| DICTIONARY FAMILIES | familles + attributs + obligatoires → autocomplétion `family`, contrôle « famille non déclarée » |
| LINKINGTABLE (ou FAMILIE VARIANTS) | `family_variant` → famille + axes de variante (référence officielle pour les contrôles d'axes) |
| DICTIONARY ATTRIBUTE GROUPS | libellés des groupes d'attributs (INF, MKG…) |
| DICTIONARY ATTRIBUTE TYPES | description des types d'attributs |
| DICTIONARY PREFIXES | préfixes de SKU (informatif) |
| DICTIONARY AKENEO TREE | arbre (vide dans la source fournie) |

Sans `refs/` (ou avec `REFS_DIR=/autre/dossier`), l'outil fonctionne comme avant :
l'autocomplétion retombe sur les valeurs distinctes présentes dans le fichier PIM.

Un fichier `refs/*ADDENDUM*.xlsx` (feuilles `VARIANTS` et/ou `FAMILIES`) est fusionné en
priorité : il sert à déclarer localement FAM007 et les `family_variant` observés dans le PIM.

## Champs description élargis

Dans la fiche PM, les attributs de type `textarea` du dictionnaire (ex. `EUMKG0030-fr_FR`
description courte, `EUMKG0040-fr_FR` description longue) s'affichent en **grandes zones de
texte** (5 lignes, pleine largeur sur 2 colonnes). Dans le tableau produits, un clic sur une
cellule texte/textarea longue ouvre aussi une zone de texte multi-lignes (Ctrl+Entrée pour
enregistrer, Échap pour annuler).

## Assistant IA par PM

Bouton **🤖 Assistant IA** sur chaque fiche PM : décrivez la modification souhaitée en français
(ex. « raccourcis le nom du PM et harmonise les descriptions »). L'assistant reçoit le contexte
complet (fiche PM, produits liés + contrôles, métadonnées des colonnes et options du dictionnaire)
et renvoie un **patch JSON** prévisualisable (tableau ancienne → nouvelle valeur, cases à cocher)
puis appliquable en un clic — chaque modification est validée (feuille/ligne/colonne existante)
et **enregistrée dans le xlsx**.

Deux modes :
- **API directe** : définissez `LLM_API_KEY` (et éventuellement `LLM_API_URL`, `LLM_MODEL`,
  défaut Mistral `https://api.mistral.ai/v1/chat/completions` / `mistral-large-latest`).
  La clé reste côté serveur, jamais affichée ni journalisée.
- **Mode manuel (hors-ligne)** : sans clé, l'interface génère le prompt complet copiable ;
  collez-le à n'importe quelle IA, puis collez sa réponse JSON dans le panneau pour
  prévisualiser et appliquer le patch.

## Relecture / correction automatique (`tools/`, `exports/`)

- `python3 tools/corriger_pim.py <source.xlsx> refs exports` : applique les corrections sûres
  (dédoublonnage des lignes PM et des SKU, création des fiches PM des parents orphelins,
  alignement catégories/familles produit ↔ PM, remplissage des champs obligatoires
  EUMKG0120/EUMKG0160/EUMKG0130, marque EUINF0010 et Type/Application déduits du nom,
  normalisation libellé→code des options, coloris EUINF0233 retrouvés dans les noms,
  EUFIL130 : Ø du texte → option de plage + unité) et produit :
  `exports/pim_source_pour_controle_CORRIGE.xlsx`, `exports/variants_propositions.csv`
  (analyse PM par PM des axes de variante + proposition), `exports/fixlog.json` (journal détaillé).
- `python3 tools/rapport.py exports/fixlog.json exports/variants_propositions.csv <checks_avant.json> <checks_apres.json> exports/RAPPORT_CONTROLE.md`
  : génère le rapport de contrôle (bilan avant/après + analyse approfondie des variantes).

## Structure du fichier attendu

- Feuille `PM` : `code`, `family_variant`, `categories`, + attributs (ex. `EUINF0010`…).
- Feuille `produits` : `sku`, `categories`, `enabled`, `family`, `parent` (lien vers `PM.code`),
  + attributs, avec les colonnes `<attribut>-unit` pour les unités.
