# Shine — mise en forme des exports de droits pour revue

Transforme des exports Active Directory (CSV de membres de groupes) en **un classeur Excel soigné par
dossier d'export**, que l'ayant droit relit et valide (liste déroulante ou cases à cocher + commentaire)
avant de le renvoyer.

**L'outil met en forme, il ne traite pas.** Une ligne de CSV = une ligne dans le classeur, rien de plus :

* pas de fusion Full / Read : une personne présente dans `_Full` **et** dans `_Read` donne deux lignes,
  chacune validée séparément (les doublons entre fichiers sont voulus) ;
* pas de dédoublonnage, pas de rapprochement de noms de ressources, pas d'interprétation du DN
  (type de compte, service…), pas de détection d'anomalies métier ;
* le seul « rapport » conservé est `rapport_lecture.csv` : fichiers ou lignes que l'outil n'a pas pu lire,
  pour qu'aucune ligne ne disparaisse en silence (il n'est créé que s'il y a un problème).

Phase 1 : lecture et génération. La réinjection des classeurs revus (`ingest`, phase 2) est **préparée**
(clés stables, onglet `_meta`) mais pas implémentée.

Python 3.10+. Dépendances : `XlsxWriter`, `PyYAML` (tests : `pytest`, `openpyxl`). 100 % local.

```
pip install -r requirements.txt            # (+ requirements-dev.txt pour les tests)
python revue_droits.py --input ..\Nov                      # sortie dans Documents\shine
python revue_droits.py --input ..\Nov --output D:\autre\dossier
python revue_droits.py --input ..\Nov --config config.yaml \
       --checkbox-mode native --only Export_E.Test7 --period 2026-11 --export-date 2026-11-19
python -m pytest
```

## Utilisation simple : `Shine.exe`

Double-clic → une fenêtre demande **le dossier global** (ex. `Liaison`) → mise en forme → message de fin avec les
compteurs → le dossier de sortie (`Documents\shine`) s'ouvre. Aucune console, aucune option à connaître.
Avec des arguments (`Shine.exe --input … --output …`), l'exécutable se comporte comme la ligne de commande
décrite ci-dessous. Un fichier `shine.yaml` posé **à côté de `Shine.exe`** (ou de `shine.py`) est chargé
automatiquement (date limite, contact, `show_dn`, `owners`…) : il suffit d'y mettre ce qui change.

### Construire `Shine.exe` (`build_exe.bat`)

À faire une seule fois, sur Windows, avec un accès internet : double-cliquer `build_exe.bat`. Le résultat est
`dist\Shine.exe`, à copier où l'on veut (le dossier `build\`, `dist\` et `Shine.spec` sont des fichiers
temporaires, ignorés par Git).

Ce que fait le script, dans l'ordre :

1. Il cherche un Python 3.10+ déjà installé (`py -3`, puis `python`, puis une installation utilisateur précédente).
2. S'il n'en trouve pas, il télécharge Python 3.12.8 depuis python.org (`amd64` ou `arm64` selon le PC).
3. Il vérifie la signature de l'installateur (éditeur « Python Software Foundation ») et s'arrête si elle n'est
   pas valide.
4. Il installe Python **pour l'utilisateur courant, sans droits administrateur**, et l'ajoute au PATH.
5. Il installe `XlsxWriter`, `PyYAML` et `PyInstaller`, puis construit `Shine.exe` (sans console, avec
   `config.yaml` intégré).

En cas d'échec, un message indique quoi faire et la fenêtre reste ouverte.

| Problème | Cause probable | Solution |
|---|---|---|
| « téléchargement de Python impossible » | proxy ou pare-feu d'entreprise qui bloque python.org | installer Python 3.10+ à la main (https://www.python.org/downloads/, cocher *Add python.exe to PATH*), puis relancer le `.bat`, qui le détecte |
| « signature de l'installateur non valide » | fichier corrompu ou altéré | relancer ; si cela persiste, installer Python à la main |
| échec de `pip install` | proxy qui bloque PyPI (pypi.org) | configurer le proxy pour `pip`, ou faire autoriser PyPI |
| Windows SmartScreen / antivirus sur `Shine.exe` | exécutable non signé | le faire autoriser par la sécurité, ou le signer avec un certificat de signature de code |

Le script n'a pas été testé sur tous les postes : au premier essai, relever le message affiché en cas de souci.

| Option | Rôle |
|---|---|
| `--input` | dossier parcouru **récursivement** (nombre de dossiers et de fichiers quelconque) |
| `--output` | dossier racine de sortie (défaut : `output.root`, soit `~/Documents/shine`, créé si besoin). Peut se trouver dans `--input` : il est ignoré à la lecture |
| `--config` | YAML fusionné avec `config.yaml` (les listes remplacent, les dictionnaires fusionnent) |
| `--checkbox-mode` | `dropdown` (défaut) ou `native` |
| `--only` | ne traiter que ce dossier (nom ou libellé), option répétable |
| `--period` | période du nom de fichier (défaut : mois courant) |
| `--export-date` | date de l'export affichée (défaut : date de modification des CSV) |
| `-v` | messages détaillés (jamais de données personnelles) |

Code retour : `0` ok, `2` entrée ou configuration invalide. Un fichier illisible n'arrête rien.

## Sortie

Dans le dossier racine (`Documents\shine` par défaut) :

```
shine\
  Revue_NTFS_E.Test7_191126\Revue_NTFS_E.Test7_191126.xlsx     un dossier par Export_<X>.<Nom>
  Revue_NTFS_F.Test10_191126\...                                (JJMMAA = date du jour de génération)
  rapport_lecture_191126.csv                                      seulement s'il y a un problème de lecture
```

Le dossier est celui que l'on transmet à l'ayant droit ; le rapport de lecture, interne, reste à côté.
Nom, date et racine sont réglables : `output.root`, `output.folder_pattern` (`Revue_NTFS_{label}_{date}`),
`output.date_format` (`%d%m%y`). `{label}` vaut `E.Test7` (voir `input.workbook_label`) ou le nom d'ayant droit
si `owners:` regroupe plusieurs dossiers. Relancer le même jour écrase les fichiers du jour.

## Entrées

* Dossiers `Export_<X>.<Nom>` (motif `input.folder_pattern`) ; un classeur par dossier. Pour regrouper
  plusieurs dossiers d'un même ayant droit : `owners:` (dossier → libellé). Les onglets de même nom sont alors
  suffixés par le code dossier ; rien n'est jamais fusionné.
* Fichiers `<Ressource>_Full.csv` (droit « Modification ») et `<Ressource>_Read.csv` (« Lecture »).
  Autre suffixe : ignoré et signalé dans `rapport_lecture.csv`.
* Format : les exports sont des CSV « doublement quotés » (chaque ligne est un champ unique qui contient du
  CSV) ; le parseur fait un double parsing et détecte aussi un CSV normal (`,` `;` tabulation). UTF-8 avec BOM,
  repli cp1252. Un fichier de 5 octets est un groupe vide : la ressource affiche « Aucun membre ».
* Les `?` déjà présents dans les noms (`L?LLL`) sont recopiés tels quels, jamais corrigés.

## Le classeur

`Instructions` · `Synthèse` · une feuille par ressource · `_meta` (masqué).

* **Instructions** : objectif, mode d'emploi, date limite et contact (`review.*`), légende des couleurs.
* **Feuille de ressource** : en-tête (ressource, dossier d'origine, groupes sources et effectifs, date de
  l'export), commentaire général, puis un tableau filtrable à volets figés :
  `Nom | Identifiant | Droit actuel | Accès légitime ? | Commentaire | Clé (masquée)`.
  Les lignes sont triées : Modification puis Lecture, puis nom, puis identifiant. Colonne `DN (annuaire)`
  optionnelle (`excel.show_dn: true`). Sous le tableau : bloc **Personnes à ajouter** (`excel.add_rows`
  lignes : Nom / Identifiant, Droit souhaité, Justification).
* **Synthèse** : par ressource, membres Full / Read, lignes à traiter, traitées, à préciser, reste,
  avancement — par formules, avec lien vers l'onglet et totaux (`SUBTOTAL`, sensibles aux filtres).
* **Non traité = cellule vide.** Couleurs par mise en forme conditionnelle : vert = validé, rouge = à retirer,
  jaune = à vérifier, gris zébré = non traité.
* Mise en page : police unique, couleurs en paramètres (`excel.colors`), impression A4 paysage ajustée en
  largeur, titres répétés.

### Liste déroulante ou cases à cocher natives

| | `dropdown` (défaut) | `native` |
|---|---|---|
| Compatibilité | toutes versions d'Excel, LibreOffice | Excel 365 (LibreOffice affiche `FALSE`/`TRUE`) |
| Colonnes | `Accès légitime ?` (Oui / Non / À vérifier) | `Conserver` et `Retirer` |
| Non traité | cellule vide | deux cases décochées |

Une case native n'a que deux états ; avec une seule case, « non traité » et « à retirer » seraient
indiscernables. Le mode natif utilise donc deux cases ; les deux cochées = incohérent, compté « à préciser ».

### Protection

Feuilles protégées **sans mot de passe**. Seules les cellules de décision, de commentaire et du bloc d'ajout
sont déverrouillées ; filtre autorisé. **Limite d'Excel** : le tri est refusé dès qu'une cellule de la plage est
verrouillée, d'où la consigne « filtrez plutôt que triez » (la clé masquée par ligne protège de toute façon la
relecture). Un onglet masqué se réaffiche en deux clics : `_meta` ne contient que des clés techniques.

## Sécurité et confidentialité

* Toute valeur commençant par `=`, `+`, `-` ou `@` est préfixée de `'` (injection de formules), y compris dans
  `rapport_lecture.csv` ; classeur créé avec `strings_to_formulas=False`.
* Aucune donnée personnelle dans la console, les logs ni les messages d'erreur (compteurs uniquement).
* Sortie déterministe : mêmes données et même horodatage → mêmes cellules, mêmes clés, même empreinte.

## Phase 2 — contrat de relecture (`ingest`, non implémenté)

* Aucune cellule fusionnée dans les zones de données (seul le commentaire général, hors tableau, est fusionné).
* Chaque ligne porte une **clé stable** (colonne masquée) :
  `sha1(dossier | sous-dossier | ressource | niveau | identifiant | n° d'occurrence)[:16]`, casse ignorée,
  indépendante du chemin d'entrée. Le niveau (`full` / `read`) et le numéro d'occurrence font que deux lignes
  identiques d'un même fichier gardent des clés distinctes.
* `_meta` (schéma v2) contient trois tableaux Excel : `MetaKV_*` (`schema_version`, `tool_version`,
  `generated_at`, `period`, `workbook_label`, `folders`, `export_dates`, `checkbox_mode`, `decision_values`,
  `n_resources`, `n_rows`, `data_hash`, `row_key`) ; `MetaSheets_*` (`sheet, folder, subpath, resource,
  group_full, group_read, header_row, first_data_row, last_data_row, n_rows, add_header_row, add_first_row,
  add_last_row, comment_cell, decision_cols, key_col, data_hash`) ; `MetaRows_*` (`sheet, key, sam, level,
  initial_row`).
* Un classeur dont `schema_version` ou `data_hash` ne correspond pas à l'export de référence doit être refusé.

## Organisation

`shine.py` (lanceur : fenêtre / `.exe`) · `build_exe.bat` · `revue_droits.py` (CLI) · `config.py` / `config.yaml` · `model.py` (dataclasses) · `parser.py` (lecture CSV
tolérante) · `compiler.py` (parcours, lecture, lignes triées) · `excel.py` (classeurs) · `tests/` (le jeu
NTFTEST est reconstruit dans `conftest.py` + cas limites : fichier vide, CSV normal, ligne cassée, même
personne en Full et Read, noms d'onglet > 31 caractères, injection, volume ~30 000 lignes).

## Licence

Licence non commerciale (voir `LICENSE`, FR + EN). Copyright (c) 2026 ReV. Conçu par ReV, développé avec
l'assistance de Claude (Anthropic).

* Utilisation, copie, modification et redistribution libres et gratuites, y compris en interne dans une entreprise.
* Interdit sans accord écrit préalable et **redevances** : vendre l'outil, l'intégrer dans un logiciel ou un service
  payant, ou l'utiliser dans une offre facturée à des tiers.
* Ce n'est pas une licence « open source » au sens de l'OSI (« source disponible »).

Rappel Git : ne jamais versionner d'exports CSV ni de classeurs générés (ils contiennent des données personnelles) ;
`.gitignore` les exclut. Les tests reconstruisent leurs propres données de test.
