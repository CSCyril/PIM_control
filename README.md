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

## Structure du fichier attendu

- Feuille `PM` : `code`, `family_variant`, `categories`, + attributs (ex. `EUINF0010`…).
- Feuille `produits` : `sku`, `categories`, `enabled`, `family`, `parent` (lien vers `PM.code`),
  + attributs, avec les colonnes `<attribut>-unit` pour les unités.
