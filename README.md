<p align="center"><img src="assets/shine_logo_1024.png" alt="Logo Shine : un S rose style graffiti" width="160"></p>

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

Python 3.10+. Dépendances : `XlsxWriter`, `PyYAML`, `openpyxl` (lecture des classeurs renvoyés ; tests : `pytest`). 100 % local.

```
pip install -r requirements.txt            # (+ requirements-dev.txt pour les tests)
python revue_droits.py --input ..\Nov                      # sortie dans Documents\shine
python revue_droits.py --input ..\Nov --output D:\autre\dossier
python revue_droits.py --input ..\Nov --config config.yaml \
       --checkbox-mode native --only Export_E.Test7 --period 2026-11 --export-date 2026-11-19
python -m pytest
```

## Utilisation simple : `Shine.exe`

Double-clic → une petite fenêtre propose deux actions : **mettre en forme des exports** ou **extraire le récapitulatif**
(voir plus bas). Pour la mise en forme, une fenêtre demande **le dossier global** (ex. `Liaison`), puis **la date limite de retour** des classeurs
(JJ/MM/AAAA, non passée ; une date est proposée par défaut) → mise en forme → message de fin avec les
compteurs → le dossier de sortie (`Documents\shine`) s'ouvre. Aucune console, aucune option à connaître.
Avec des arguments (`Shine.exe --input … --output …`), l'exécutable se comporte comme la ligne de commande
décrite ci-dessous. Le contact (et l'organisation) se renseignent dans `config.yaml` (`review.contact`, `review.organisation`) ; la date limite,
elle, est saisie au lancement. Il n'y a qu'**un seul fichier de configuration : `config.yaml`**. Il est embarqué dans `Shine.exe` ; un `config.yaml`
posé **à côté de `Shine.exe`** le remplace (fusion clé par clé) sans reconstruire l'exécutable : copiez-y le fichier,
modifiez le contact, la version, `owners`, `show_dn`… Les clés absentes gardent leur valeur par défaut.

### Extraire le récapitulatif d'un classeur renvoyé

Second choix de la fenêtre de lancement : on désigne **un dossier** (comme pour la compilation) contenant les classeurs
renvoyés par les ayants droit. Shine le parcourt, sous-dossiers compris, et crée **à côté de chaque classeur Shine** un
fichier `Recap_<nom du classeur>.xlsx`, puis ouvre le dossier. Les autres fichiers Excel, les temporaires `~$…` et les
`Recap_…` déjà extraits sont ignorés.

Le fichier extrait contient le **bloc de validation** (statut Signé / Non signé, nom, fonction, date, attestation,
signature), les compteurs d'ajouts et de retraits, et le **tableau de l'onglet Récapitulatif**, en **valeurs figées** (aucune
formule, aucun lien externe : il reste valable après un envoi par mail). C'est une copie fidèle, sans analyse.
Les valeurs sont celles **enregistrées par Excel** dans le classeur renvoyé : si l'ayant droit ne l'a pas enregistré, le récapitulatif
est vide et Shine le signale. Le résumé affiché ne contient que des compteurs et des noms de fichiers, jamais de noms de personnes.

En ligne de commande : `Shine.exe --extract-recap <classeur.xlsx ou dossier> […] [--output <dossier>]`
(par défaut, à côté de chaque classeur).

### Construire `Shine.exe` (`build_exe.bat`)

À faire une seule fois, sur Windows, avec un accès internet : double-cliquer `build_exe.bat`. Le résultat est
`dist\Shine.exe`, à copier où l'on veut (les dossiers `build\` et `dist\` sont temporaires, ignorés par Git ;
la recette de construction est `config.spec`, versionnée, que `build_exe.bat` passe à PyInstaller).

Le script **vérifie d'abord ce qui est déjà installé** et ne télécharge que ce qui manque. Si tout est présent, il
n'a même pas besoin d'internet. Dans l'ordre :

1. **Python 3.10+** : il cherche une installation existante (`py -3`, puis `python`, puis une installation utilisateur
   précédente). Seulement s'il n'en trouve aucune, il télécharge Python 3.12.8 depuis python.org (`amd64` ou `arm64`),
   vérifie la signature de l'installateur (éditeur « Python Software Foundation », arrêt si invalide) et l'installe
   **pour l'utilisateur courant, sans droits administrateur**, avec ajout au PATH.
2. **Composants** : il contrôle `XlsxWriter` (≥ 3.2), `PyYAML` (≥ 6) et `PyInstaller` ; il n'installe par `pip` que
   ceux qui manquent ou sont trop anciens.
3. **`Shine.exe`** : s'il existe déjà dans `dist\` et qu'aucun fichier `.py` ou `config.yaml` n'est plus récent, il
   n'est pas reconstruit. `build_exe.bat /force` force la reconstruction.

En cas d'échec, un message indique quoi faire et la fenêtre reste ouverte.

| Problème | Cause probable | Solution |
|---|---|---|
| « téléchargement de Python impossible » | proxy ou pare-feu d'entreprise qui bloque python.org | installer Python 3.10+ à la main (https://www.python.org/downloads/, cocher *Add python.exe to PATH*), puis relancer le `.bat`, qui le détecte |
| « signature de l'installateur non valide » | fichier corrompu ou altéré | relancer ; si cela persiste, installer Python à la main |
| « Impossible d'écrire les classeurs… » / `FileCreateError` | 1) le classeur du jour est ouvert dans Excel (Shine s'arrête avec un message et ne crée aucun fichier de secours) ; 2) Windows bloque l'écriture dans le dossier de sortie : « Accès contrôlé aux dossiers », antivirus, OneDrive ; 3) chemin trop long | fermer le classeur dans Excel ; autoriser `Shine.exe` ou choisir un autre dossier (`--output`, ou `output.root` dans `config.yaml`) ; raccourcir le chemin |
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
| `--deadline` | date limite de retour JJ/MM/AAAA, affichée dans les onglets En-tête et Instructions (par défaut : `review.deadline`) |
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
  Un code de plusieurs lettres (`Export_ABC.Nom`) est accepté. Les fichiers placés dans des **sous-dossiers** d'un export
  (`Export_E.Nom/Alpha/`, `Beta/`…) sont traités : un onglet par ressource et par sous-dossier, nommé
  `<Ressource>_<sous-dossier>` (ex. `Res1_Alpha`) ; le sous-dossier figure dans « Dossier d'origine » et dans la Synthèse.
* Fichiers `<Ressource>_Full.csv` (droit « Modification ») et `<Ressource>_Read.csv` (« Lecture »).
  Un fichier **sans suffixe** (ex. `Alpha.csv`, `Beta.csv`) est traité comme une ressource au nom du fichier, avec le droit
  « **Non précisé** » (Shine ne suppose rien) ; la Synthèse ajoute alors la colonne « Membres (droit non précisé) ».
  `resources.unsuffixed: ignore` revient à l'ancien comportement (fichier ignoré et signalé dans `rapport_lecture.csv`).
* Format : les exports sont des CSV « doublement quotés » (chaque ligne est un champ unique qui contient du
  CSV) ; le parseur fait un double parsing et détecte aussi un CSV normal (`,` `;` tabulation). UTF-8 avec BOM,
  repli cp1252. Un fichier de 5 octets est un groupe vide : la ressource affiche « Aucun membre ».
* Les `?` déjà présents dans les noms (`L?LLL`) sont recopiés tels quels, jamais corrigés.

## Le classeur

`En-tête` · `Instructions` · `Synthèse` · `Récapitulatif` · une feuille par ressource · `_calc` et `_meta` (masqués).

* **En-tête** (page de garde, d'après le modèle fourni) : en haut à droite **Version / Indice / Date** ; au centre le
  titre « Revue des droits NTFS » et le périmètre ; **Classification du document** (`1-Public`, `2-Restreint`,
  `3-Confidentiel` : le niveau de `review.classification` est mis en évidence) ; tableau **Mise à jour**
  `Auteur | Date | Version | Nature de la modification | Approbation | Validation | Status`. La première ligne est
  écrite par Shine (création du classeur ; « Validation » reprend le statut de la signature) ; `excel.history_rows`
  lignes libres suivent (6 par défaut) pour tracer les modifications ; les dates se choisissent dans une liste
  déroulante (du jour de génération à `excel.date_choices_days` jours après, 90 par défaut ; saisie libre possible). Réglages : `review.version`, `review.index`,
  `review.classification`, `review.author`, `review.approval` (case « Approbation » de la 1re ligne), `review.footer_notice` (mention rouge en pied de page, vide par défaut),
  couleurs `excel.colors.cover_*`.
* **Instructions** : objectif, mode d'emploi (`texts.steps_*`), consigne de date limite (`texts.return_line`,
  `review.deadline`), contact, puis le bloc **Validation et signature** (nom, fonction, date, attestation « Oui »,
  signature saisie ; le statut passe à « Signé » quand tout est renseigné et s'affiche aussi dans l'En-tête), et
  la légende des couleurs.
* **Signature électronique** : la saisie du nom + la date valent signature électronique *simple* (traçabilité
  interne). Une signature numérique certifiée demande un certificat et se fait à la main, en dernier : Fichier ›
  Informations › Protéger le classeur › Ajouter une signature numérique (toute modification ultérieure
  l'invalide). Shine ne signe rien lui-même.
* **Feuille de ressource** : en-tête (ressource, dossier d'origine, groupes sources et effectifs, date de
  l'export), puis un tableau filtrable à volets figés :
  `Nom | Identifiant | Droit actuel | Accès autorisé ? | Commentaire | Clé (masquée)`, et **sous le tableau** la
  zone de commentaire général (un lien en haut de page y renvoie).
  Les lignes sont triées : Modification puis Lecture, puis nom, puis identifiant. Colonne `DN (annuaire)`
  optionnelle (`excel.show_dn: true`).
* **Synthèse** : par ressource, membres Full / Read, lignes à traiter, traitées, reste, avancement — par
  formules, avec lien vers l'onglet et totaux (`SUBTOTAL`, sensibles aux filtres).
  En dessous du total : le tableau **Personnes à ajouter** (`excel.add_rows` lignes, 20 par défaut) :
  `Nom | Identifiant (optionnel) | Groupe de ressource | Justification`. « Groupe de ressource » est une liste déroulante des
  **fichiers CSV du classeur** (nom sans extension, ex. `Activites_etudes_Full`, `Activites_etudes_Read`) : la personne
  est à ajouter à ce groupe ; le droit (Lecture / Modification) s'en déduit. Quand un classeur regroupe plusieurs
  dossiers (`owners:`), les libellés sont préfixés du dossier (`Export_E.Test7 / Documentation_INTERNE_Full`).
  Chaque feuille de ressource renvoie à ce tableau par un lien.
* **Récapitulatif** : tableau **calculé automatiquement** (aucune saisie) qui liste les demandes à transmettre :
  `Action | Propriétaire du groupe | Nom | Identifiant | Droit souhaité | Justification | Ressource`.
  « Ajout » = lignes remplies du tableau « Personnes à ajouter » ; « Retrait » = lignes « Non » (ou « Retirer »
  seule cochée), avec leur commentaire, le droit « Aucun (retrait) » et le groupe CSV concerné. Il fonctionne par
  formules ordinaires (`INDEX`/`MATCH` sur la feuille masquée `_calc`) : ni fonction volatile, ni tableau dynamique,
  donc toutes versions d'Excel. Capacité : `excel.add_rows` + au plus `excel.recap_max_removals` retraits
  (500 par défaut) ; au-delà, un avertissement s'affiche en haut de l'onglet.
* **Non traité = cellule vide.** Couleurs par mise en forme conditionnelle : vert = accès autorisé, rouge = à
  retirer, gris zébré = non traité. Mise en page : police unique, couleurs en paramètres (`excel.colors`, seules
  les couleurs utilisées y figurent), impression A4 paysage ajustée en largeur, titres répétés.
* **Incohérence Lecture / Modification** : si une même personne figure dans les deux fichiers, Shine ne la détecte
  pas (pas d'analyse) ; l'instruction n° 4 demande au propriétaire de la corriger lui-même, en la justifiant en
  commentaire. Chaque ligne se valide séparément.

### Liste déroulante ou cases à cocher natives

| | `dropdown` (défaut) | `native` |
|---|---|---|
| Compatibilité | toutes versions d'Excel, LibreOffice | Excel 365 (LibreOffice affiche `FALSE`/`TRUE`) |
| Colonnes | `Accès autorisé ?` (Oui / Non) | `Conserver` et `Retirer` |
| Non traité | cellule vide | deux cases décochées |

Une case native n'a que deux états ; avec une seule case, « non traité » et « à retirer » seraient
indiscernables. Le mode natif utilise donc deux cases ; les deux cochées = incohérent, compté **non traité**
(ni vert ni rouge, absent du récapitulatif).

### Protection

Feuilles protégées **sans mot de passe**. Seules les cellules de décision, de commentaire et du bloc d'ajout
sont déverrouillées ; filtre autorisé. **Limite d'Excel** : le tri est refusé dès qu'une cellule de la plage est
verrouillée, d'où la consigne « filtrez plutôt que triez » (la clé masquée par ligne protège de toute façon la
relecture). Un onglet masqué se réaffiche en deux clics : `_meta` et `_calc` ne contiennent que des clés techniques et des
formules de calcul. Le récapitulatif est dérivé : l'ingestion de la phase 2 ne le relit pas.

## Sécurité et confidentialité

* Toute valeur commençant par `=`, `+`, `-` ou `@` est préfixée de `'` (injection de formules), y compris dans
  `rapport_lecture.csv` ; classeur créé avec `strings_to_formulas=False`.
* Aucune donnée personnelle dans la console, les logs ni les messages d'erreur (compteurs uniquement).
* Sortie déterministe : mêmes données et même horodatage → mêmes cellules, mêmes clés, même empreinte.

## Phase 2 — contrat de relecture (`ingest`, non implémenté)

* Aucune cellule fusionnée dans les zones de données des feuilles de ressource (seul le commentaire général, sous le
  tableau, est fusionné). Exception : dans le tableau « Personnes à ajouter » de la Synthèse, le groupe de ressource
  occupe C:E et la justification F:H par ligne (valeurs lues en colonnes C et F ; nom en A, identifiant facultatif en B).
* Chaque ligne porte une **clé stable** (colonne masquée) :
  `sha1(dossier | sous-dossier | ressource | niveau | identifiant | n° d'occurrence)[:16]`, casse ignorée,
  indépendante du chemin d'entrée. Le niveau (`full` / `read`) et le numéro d'occurrence font que deux lignes
  identiques d'un même fichier gardent des clés distinctes.
* `_meta` (schéma v5) contient quatre tableaux Excel : `MetaKV_*` (`schema_version`, `tool_version`,
  `generated_at`, `period`, `workbook_label`, `folders`, `export_dates`, `checkbox_mode`, `decision_values`,
  `n_resources`, `n_rows`, `data_hash`, `row_key`, et pour le tableau d'ajout : `add_sheet`, `add_header_row`,
  `add_first_row`, `add_last_row`, `add_col_name`, `add_col_sam`, `add_col_group`, `add_col_justification` ; pour l'en-tête et le récapitulatif :
  `cover_sheet`, `cover_history_rows`, `signature_sheet`, `signature_cells`, `recap_sheet`, `recap_header_row`, `recap_first_row`,
  `recap_last_row`, `recap_note`) ; `MetaSheets_*` (`sheet, folder, subpath, resource,
  group_full, group_read, header_row, first_data_row, last_data_row, n_rows, comment_cell, decision_cols, key_col,
  data_hash, group_other`) ; `MetaRows_*` (`sheet, key, sam, level, initial_row`) ; `MetaGroups_*` (`group, sheet, folder,
  subpath, resource, level, file, owner` : source de la liste « Groupe de ressource »).
* Un classeur dont `schema_version` ou `data_hash` ne correspond pas à l'export de référence doit être refusé.

## Organisation

`shine.py` (lanceur : fenêtre / `.exe`) · `extract.py` (extraction du récapitulatif) · `build_exe.bat` · `revue_droits.py` (CLI) · `config.py` / `config.yaml` · `model.py` (dataclasses) · `parser.py` (lecture CSV
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
