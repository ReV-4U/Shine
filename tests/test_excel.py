import datetime as dt
import re
import zipfile

import openpyxl
import pytest

from compiler import compile_all, group_books
from conftest import dn_user, write_export
from excel import NO_MEMBER, build_book, make_sheet_name, safe

NOW = dt.datetime(2026, 11, 19, 10, 0, 0)
U = lambda n, sam, ou="TEST_DIRECTION": (n, dn_user(n, ou), sam)  # noqa: E731


def build_all(root, cfg, out, mode="dropdown", only=None):
    res = compile_all(root, cfg, only=only, export_date=dt.date(2026, 11, 19))
    out.mkdir(parents=True, exist_ok=True)
    paths = {}
    for b in group_books(res.scopes, cfg):
        p = out / f"{b.label}.xlsx"
        build_book(p, b, cfg, mode, "2026-11", NOW)
        paths[b.label] = p
    return res, paths


@pytest.fixture
def books(ntftest, cfg, tmp_path):
    return build_all(ntftest, cfg, tmp_path)[1]


def get_table(ws, prefix):
    return ws.tables[next(k for k in ws.tables if k.startswith(prefix))]


def table_rows(ws, prefix):
    rows = list(ws[get_table(ws, prefix).ref])
    return [c.value for c in rows[0]], [[c.value for c in r] for r in rows[1:]]


def find_row(ws, text, col=1):
    return next(c.row for row in ws.iter_rows(min_col=col, max_col=col) for c in row if c.value == text)


# --------------------------------------------------------------------------- noms d'onglet / injection
def test_sheet_names():
    used = set()
    names = [make_sheet_name("A" * 40, used), make_sheet_name("A" * 41, used), make_sheet_name("a" * 40, used),
             make_sheet_name("Bad[name]:*?/\\x", used), make_sheet_name("'quoted'", used),
             make_sheet_name("Synthèse", used), make_sheet_name("", used)]
    assert all(1 <= len(n) <= 31 for n in names)
    assert len({n.casefold() for n in names}) == len(names)
    assert not any(re.search(r"[\[\]:*?/\\]", n) for n in names)
    assert not any(n.startswith("'") or n.endswith("'") for n in names)
    assert names[5] != "Synthèse"


@pytest.mark.parametrize("v,expected", [("=1+1", "'=1+1"), ("+cmd", "'+cmd"), ("-x", "'-x"), ("@SUM(A1)", "'@SUM(A1)"),
                                        ("Normal", "Normal"), ("", ""), (None, "")])
def test_safe(v, expected):
    assert safe(v) == expected


def test_formula_injection_neutralised_everywhere(tmp_path, cfg):
    root = tmp_path / "in"
    evil = [('=HYPERLINK("http://x")', "=cmd"), ("+SUM(1)", "+id"), ("-2+3", "-id"), ("@x", "@id")]
    write_export(root / "Export_A.Un" / "=EVIL_Full.csv", [(n, "=dn" + str(i), s) for i, (n, s) in enumerate(evil)])
    write_export(root / "Export_A.Un" / "=EVIL_Read.csv", [])
    cfg["excel"]["show_dn"] = True
    _, paths = build_all(root, cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["A.Un"])
    ok_formulas = ('="Progression', "=COUNTIF", "=SUBTOTAL", "=E", "=IF", "=IFERROR", "=SUMPRODUCT", "=SUM(", "='", "=B", "=I", "=$C", "=TODAY()")
    seen = 0
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    assert str(c.value).startswith(ok_formulas), (ws.title, c.coordinate, c.value)   # uniquement nos formules
                if isinstance(c.value, str) and any(x in c.value for x in ("cmd", "HYPERLINK", "SUM(1)", "-2+3", "=dn")):
                    assert c.value[0] == "'", c.value
                    seen += 1
    assert seen >= 8                                                  # noms, identifiants, DN, nom de ressource
    assert any("EVIL" in n for n in wb.sheetnames)


# --------------------------------------------------------------------------- structure
def test_workbook_structure(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    assert wb.sheetnames == ["En-tête", "Instructions", "Synthèse", "Récapitulatif", "Activites_etudes", "Documentation_INTERNE",
                             "TEST_CHINE", "_calc", "_meta"]
    state = {ws.title: ws.sheet_state for ws in wb}
    assert state["_meta"] == "hidden" and state["_calc"] == "hidden" and state["En-tête"] == "visible"
    assert wb.active.title == "En-tête"
    assert all(len(n) <= 31 for n in wb.sheetnames)


def test_resource_sheet_layout(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    assert ws["B1"].value == "Activites_etudes" and ws["B2"].value == "Export_E.Test7"
    assert "Modification : Activites_etudes_Full (3)" in ws["B3"].value and "Lecture : Activites_etudes_Read (4)" in ws["B3"].value
    assert ws["B4"].value == dt.datetime(2026, 11, 19)
    header, rows = table_rows(ws, "Membres")
    assert header == ["Nom", "Identifiant", "Droit actuel", "Accès autorisé ?", "Commentaire", "Clé"]   # ni service, ni type, ni DN
    # une ligne par ligne d'export : les 3 comptes apparaissent en Modification ET en Lecture, sans fusion
    assert [(r[1], r[2]) for r in rows] == [("KKK", "Modification"), ("MMM", "Modification"), ("GGG", "Modification"),
                                            ("CASA", "Lecture"), ("MMM", "Lecture"), ("GGG", "Lecture"), ("KKK", "Lecture")]
    assert [r[0] for r in rows][:3] == ["AAAA BBBB", "CCCC XXXX", "DDDD L?LLL"] and rows[6][0] == "ZZZZ BBBB"
    assert ws.freeze_panes == "A7"
    assert ws.page_setup.orientation == "landscape" and ws.page_setup.fitToHeight == 0
    assert ws.sheet_properties.pageSetUpPr.fitToPage and ws.print_title_rows == "$6:$6"
    assert sorted(str(m) for m in ws.merged_cells.ranges) == ["B15:E15", "B3:D3"]   # groupes sources + commentaire général (sous le tableau) ; aucune fusion dans les données
    assert ws["A15"].value == "Commentaire général" and ws["B5"].value is None and ws["A5"].value is None
    assert get_table(ws, "Membres").autoFilter is not None
    assert ws.max_row == 15 and ws.column_dimensions["F"].hidden              # plus de bloc d'ajout sur les feuilles de ressource
    assert "Personnes à ajouter" not in " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None
                                                 and not str(c.value).startswith("↓"))
    assert ws["E3"].hyperlink.location == "'Synthèse'!A11" and ws["E3"].value.startswith("↓ Personnes à ajouter")
    assert ws["E4"].hyperlink.location == "'Activites_etudes'!A15" and ws["E4"].value.startswith("↓ Commentaire général")
    assert ws.print_area == "'Activites_etudes'!$A$1:$E$15"


def test_default_state_is_untreated(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    assert all(ws.cell(r, 4).value is None for r in range(7, 14))      # liste vide = non traité


def test_validation_and_conditional_formats(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    dv = {str(d.sqref): d for d in ws.data_validations.dataValidation}
    assert list(dv) == ["D7:D13"] and dv["D7:D13"].formula1 == '"Oui,Non"' and dv["D7:D13"].errorStyle in (None, "stop")
    cf = [(str(rng.sqref), r.formula[0], r.stopIfTrue) for rng, rules in ws.conditional_formatting._cf_rules.items() for r in rules]
    assert [f for _, f, _ in cf] == ['$D7="Oui"', '$D7="Non"', "MOD(ROW(),2)=0", "MOD(ROW(),2)=1"]
    assert all(s == "A7:D13" and stop for s, _, stop in cf)            # vert / rouge / gris zébré (plus de jaune « à vérifier »)


def test_no_interpretation_of_dn_in_workbook(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    text = " ".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value is not None)
    for word in ("TEST_DIRECTION", "METHODES", "Utilisateur", "Compte générique", "Groupe imbriqué", "anomalie"):
        assert word not in text


def test_show_dn_option(ntftest, cfg, tmp_path):
    cfg["excel"]["show_dn"] = True
    _, paths = build_all(ntftest, cfg, tmp_path)
    ws = openpyxl.load_workbook(paths["E.Test7"])["Activites_etudes"]
    header, rows = table_rows(ws, "Membres")
    assert header == ["Nom", "Identifiant", "DN (annuaire)", "Droit actuel", "Accès autorisé ?", "Commentaire", "Clé"]
    assert rows[0][2] == "CN=BBBB AAAA,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"    # DN brut
    assert str(ws.data_validations.dataValidation[0].sqref) == "E7:E13"


def test_protection_without_password_and_unlocked_cells(books):
    z = zipfile.ZipFile(books["E.Test7"])
    xml = z.read("xl/worksheets/sheet5.xml").decode()
    prot = re.search(r"<sheetProtection[^>]*>", xml).group(0)
    assert "password" not in prot.lower()
    assert 'sort="0"' in prot and 'autoFilter="0"' in prot          # filtre et tri autorisés
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["Activites_etudes"]
    assert ws.protection.sheet
    unlocked = {c.coordinate for row in ws.iter_rows() for c in row if not c.protection.locked}
    assert {f"{col}{r}" for col in "DE" for r in range(7, 14)} | {"B15"} <= unlocked
    for r in range(7, 14):                                           # nom, identifiant, droit, clé : toujours verrouillés
        assert all(ws.cell(r, c).protection.locked for c in (1, 2, 3, 6))
    assert wb["Synthèse"].protection.sheet and wb["Instructions"].protection.sheet


def test_synthese(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Synthèse"]
    assert ws.freeze_panes == "A5"                                    # seules les lignes d'en-tête sont figées, pas la colonne A
    header, rows = table_rows(ws, "Synthese")
    assert header == ["Ressource", "Dossier", "Membres Full", "Membres Read", "Lignes à traiter", "Traitées", "Reste", "Avancement"]
    assert {r[0]: r[2:5] for r in rows} == {"Activites_etudes": [3, 4, 7], "Documentation_INTERNE": [3, 1, 4], "TEST_CHINE": [3, 0, 3]}
    assert ws["A5"].hyperlink.location == "'Activites_etudes'!A1"
    assert ws["F5"].value.startswith("=COUNTIF('Activites_etudes'!$D$7:$D$13,\"Oui\")")
    assert ws["A8"].value == "Total" and ws["E8"].value == "=SUBTOTAL(109,E5:E7)"     # total hors tableau, respecte les filtres


def test_native_mode(ntftest, cfg, tmp_path):
    _, paths = build_all(ntftest, cfg, tmp_path, mode="native")
    z = zipfile.ZipFile(paths["E.Test7"])
    assert any("featurePropertyBag" in n for n in z.namelist())      # vraies cases à cocher
    wb = openpyxl.load_workbook(paths["E.Test7"])
    ws = wb["Activites_etudes"]
    header, rows = table_rows(ws, "Membres")
    assert header == ["Nom", "Identifiant", "Droit actuel", "Conserver", "Retirer", "Commentaire", "Clé"]
    assert all(ws.cell(r, c).value is False for r in range(7, 14) for c in (4, 5))    # décoché = non traité
    cf = [r.formula[0] for rules in ws.conditional_formatting._cf_rules.values() for r in rules]
    assert cf[:2] == ["AND($D7=TRUE,$E7<>TRUE)", "AND($E7=TRUE,$D7<>TRUE)"] and len(cf) == 4
    assert not any(d.formula1 and "Oui" in d.formula1 for d in ws.data_validations.dataValidation)
    assert ws.column_dimensions["G"].hidden
    assert "SUMPRODUCT" in wb["Synthèse"]["F5"].value


def test_fully_empty_resource(tmp_path, cfg):
    root = tmp_path / "in"
    write_export(root / "Export_A.Un" / "Vide_Full.csv", [])
    write_export(root / "Export_A.Un" / "Vide_Read.csv", [])
    _, paths = build_all(root, cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["A.Un"])
    ws = wb["Vide"]
    assert ws["A7"].value == NO_MEMBER and ws.protection.sheet and ws["D7"].protection.locked      # rien à décider
    assert [c.value for c in ws[7]][1:4] == [None] * 3
    ws_syn = wb["Synthèse"]                                           # le tableau d'ajout (Synthèse) reste disponible
    t = find_row(ws_syn, "Personnes à ajouter")
    assert [ws_syn.cell(t + 2, c).value for c in (1, 2, 3, 6)] == ["Nom", "Identifiant (optionnel)", "Groupe de ressource", "Justification"]
    _, rows = table_rows(wb["Synthèse"], "Synthese")
    assert rows[0][2:6] == [0, 0, 0, 0]
    assert rows[0][6] == "=E5-F5"                                     # reste : formule, valeur 0


def test_empty_read_group_is_shown_with_zero(books):
    ws = openpyxl.load_workbook(books["G.Test1"])["MEXIQUE"]            # MEXIQUE_Read : 5 octets, groupe vide
    assert "Lecture : MEXIQUE_Read (0)" in ws["B3"].value and "Modification : MEXIQUE_Full (3)" in ws["B3"].value
    assert ws.max_row >= 9 and sum(1 for r in range(7, 10) if ws.cell(r, 3).value == "Modification") == 3


# --------------------------------------------------------------------------- _meta (contrat phase 2)
def test_meta_contract(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    meta = wb["_meta"]
    kv = {r[0].value: r[1].value for r in meta[get_table(meta, "MetaKV").ref] if r[0].value}
    assert kv["schema_version"] == "5" and kv["folders"] == "Export_E.Test7" and kv["checkbox_mode"] == "dropdown"
    assert kv["generated_at"] == "2026-11-19T10:00:00" and kv["n_rows"] == "14" and len(kv["data_hash"]) == 64
    syn = wb["Synthèse"]
    assert kv["add_sheet"] == "Synthèse" and syn.cell(int(kv["add_header_row"]), 1).value == "Nom"
    assert kv["add_col_name"] == "A" and kv["add_col_sam"].startswith("B") and kv["add_col_group"].startswith("C") \
        and kv["add_col_justification"].startswith("F")
    assert int(kv["add_last_row"]) - int(kv["add_first_row"]) + 1 == 20
    sh = [[c.value for c in r] for r in meta[get_table(meta, "MetaSheets").ref]]
    hdr, sh = sh[0], [dict(zip(sh[0], r)) for r in sh[1:]]
    assert {"sheet", "first_data_row", "last_data_row", "comment_cell", "key_col", "data_hash"} <= set(hdr)
    assert not {"add_header_row", "add_first_row", "add_last_row"} & set(hdr)       # le tableau d'ajout est global (MetaKV)
    meta_rows = [[c.value for c in r] for r in meta[get_table(meta, "MetaRows").ref]][1:]
    for info in sh:
        ws = wb[info["sheet"]]
        keys_in_sheet = [ws[f"{info['key_col']}{r}"].value for r in range(info["first_data_row"], info["last_data_row"] + 1)]
        keys_in_meta = [m[1] for m in meta_rows if m[0] == info["sheet"]]
        assert keys_in_sheet == keys_in_meta and len(set(keys_in_sheet)) == len(keys_in_sheet) == info["n_rows"]
        assert ws[info["comment_cell"]].protection.locked is False
    assert len({m[1] for m in meta_rows}) == len(meta_rows)            # clés uniques dans le classeur
    assert {m[3] for m in meta_rows} == {"full", "read"}


# --------------------------------------------------------------------------- tableau « Personnes à ajouter » (Synthèse)
def _stems(ntftest, folder):
    return sorted(p.stem for p in (ntftest / "Nov").rglob("*.csv") if p.parent.name == folder)


def test_add_people_table_in_synthese(ntftest, books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["Synthèse"]
    t = find_row(ws, "Personnes à ajouter")
    assert t == 11 and "groupe de ressource" in ws.cell(t + 1, 1).value and "identifiant" in ws.cell(t + 1, 1).value
    assert [ws.cell(t + 2, c).value for c in (1, 2, 3, 6)] == ["Nom", "Identifiant (optionnel)", "Groupe de ressource", "Justification"]
    first, last = t + 3, t + 22                                         # 20 lignes vides
    assert all(ws.cell(r, c).value is None for r in range(first, last + 1) for c in (1, 2, 3, 6))
    assert ws.cell(last + 1, 1).value is None
    # saisie : A, B, groupe (C:E) et justification (F:I) déverrouillés, le reste de la feuille verrouillé
    for r in (first, last):
        assert all(not ws.cell(r, c).protection.locked for c in (1, 2, 3, 6))
    assert ws["A5"].protection.locked and ws[f"A{t}"].protection.locked
    merged = {str(m) for m in ws.merged_cells.ranges}
    for r in (t + 2, first, last):
        assert f"C{r}:E{r}" in merged and f"F{r}:H{r}" in merged
    assert len(merged) == 2 * 21                                   # en-tête + 20 lignes, deux fusions par ligne
    # liste déroulante « Groupe de ressource » : plage de l'onglet masqué _meta = noms des fichiers CSV (sans extension)
    dv = {str(d.sqref): d for d in ws.data_validations.dataValidation}[f"C{first}:C{last}"]
    meta = wb["_meta"]
    ref = get_table(meta, "MetaGroups").ref
    group_rows = [[c.value for c in r] for r in meta[ref]]
    assert group_rows[0] == ["group", "sheet", "folder", "subpath", "resource", "level", "file", "owner"]
    assert {r[7] for r in group_rows[1:]} == {"E.Test7"}              # propriétaire du groupe = <X>.<Nom>
    groups = [r[0] for r in group_rows[1:]]
    col = ref.split(":")[0][:-1]
    assert dv.formula1 == f"'_meta'!${col}$2:${col}${len(groups) + 1}"
    assert sorted(groups) == _stems(ntftest, "Export_E.Test7")
    assert "Activites_etudes_Full" in groups and "Activites_etudes_Read" in groups
    assert {r[5] for r in group_rows[1:]} == {"full", "read"}
    assert wb["_meta"].sheet_state == "hidden" and ws.protection.sheet


def test_add_people_groups_prefixed_when_several_folders(ntftest, cfg, tmp_path):
    cfg["owners"] = {"Export_E.Test7": "Dupont", "Export_F.Test10": "Dupont"}
    _, paths = build_all(ntftest, cfg, tmp_path)
    meta = openpyxl.load_workbook(paths["Dupont"])["_meta"]
    groups = [r[0].value for r in meta[get_table(meta, "MetaGroups").ref]][1:]
    assert "Export_E.Test7 / Documentation_INTERNE_Full" in groups and "Export_F.Test10 / Documentation_INTERNE_Full" in groups
    assert len(set(groups)) == len(groups) and all(" / " in g for g in groups)


def test_add_people_no_resource_means_no_list(tmp_path, cfg):
    root = tmp_path / "in"
    (root / "Export_A.Un").mkdir(parents=True)
    res = compile_all(root, cfg)
    assert res.scopes == [] or all(not s.resources for s in res.scopes)


def test_hash_and_determinism(ntftest, cfg, tmp_path):
    _, a = build_all(ntftest, cfg, tmp_path / "a")
    _, b = build_all(ntftest, cfg, tmp_path / "b")

    def dump(p):
        wb = openpyxl.load_workbook(p)
        return [(ws.title, [[c.value for c in r] for r in ws.iter_rows()]) for ws in wb]
    assert dump(a["E.Test7"]) == dump(b["E.Test7"]) and dump(a["G.Test1"]) == dump(b["G.Test1"])
    (ntftest / "Nov/Liaison/Export_E.Test7/Activites_etudes_Full.csv").write_bytes(
        (ntftest / "Nov/Liaison/Export_E.Test7/Activites_etudes_Read.csv").read_bytes())
    _, c = build_all(ntftest, cfg, tmp_path / "c")
    h = lambda p: {r[0].value: r[1].value for r in openpyxl.load_workbook(p)["_meta"]["A2:B14"]}["data_hash"]  # noqa: E731
    assert h(a["E.Test7"]) != h(c["E.Test7"]) and h(a["G.Test1"]) == h(c["G.Test1"])


# --------------------------------------------------------------------------- noms longs, regroupement
def test_long_names_collide_safely(tmp_path, cfg):
    root = tmp_path / "in"
    for n in ("Ressource_tres_longue_nom_numero_un", "Ressource_tres_longue_nom_numero_deux", "Ressource[a:b]*?"):
        write_export(root / "Export_A.Un" / f"{n}_Full.csv", [U("A A", "aa")])
        write_export(root / "Export_A.Un" / f"{n}_Read.csv", [])
    _, paths = build_all(root, cfg, tmp_path / "out")
    res_names = openpyxl.load_workbook(paths["A.Un"]).sheetnames[4:-2]
    assert len(res_names) == 3 and len({n.casefold() for n in res_names}) == 3
    assert all(len(n) <= 31 and not re.search(r"[\[\]:*?/\\]", n) for n in res_names)


def test_owner_grouping_suffixes_duplicate_sheet_names(ntftest, cfg, tmp_path):
    cfg["owners"] = {"Export_E.Test7": "Dupont", "Export_F.Test10": "Dupont"}
    _, paths = build_all(ntftest, cfg, tmp_path)
    assert sorted(paths) == ["Dupont", "G.Test1"]
    wb = openpyxl.load_workbook(paths["Dupont"])
    assert "Documentation_INTERNE_E" in wb.sheetnames and "Documentation_INTERNE_F" in wb.sheetnames
    assert "Documentation_INTERNE" not in wb.sheetnames
    _, rows = table_rows(wb["Synthèse"], "Synthese")
    assert len(rows) == 6 and {r[1] for r in rows} == {"Export_E.Test7", "Export_F.Test10"}


def test_unreadable_group_shown_in_synthese(tmp_path, cfg, monkeypatch):
    import compiler
    root = tmp_path / "in"
    write_export(root / "Export_A.Un" / "R_Full.csv", [U("A A", "aa")])
    write_export(root / "Export_A.Un" / "R_Read.csv", [U("B B", "bb")])
    real = compiler.parse_file
    monkeypatch.setattr(compiler, "parse_file",
                        lambda p: (_ for _ in ()).throw(OSError("x")) if p.name == "R_Read.csv" else real(p))
    _, paths = build_all(root, cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["A.Un"])
    _, rows = table_rows(wb["Synthèse"], "Synthese")
    assert rows[0][2:5] == [1, "illisible", 1]
    assert "fichier illisible" in wb["R"]["B3"].value


# --------------------------------------------------------------------------- onglet « En-tête »
def frow(ws, text):                                                      # libellés de l'onglet En-tête : colonne B
    return find_row(ws, text, col=2)


def _cell_of(ws, label, col=2):
    return ws.cell(frow(ws, label), col)


def test_cover_sheet(books):
    """L'En-tête reprend le modèle fourni : version / indice / date, titre, classification, tableau « Mise à jour »."""
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["En-tête"]
    assert [ws.cell(r, 7).value for r in (1, 2, 3)] == ["Version", "Indice", "Date"]
    assert ws["H1"].value == "1.0" and ws["H2"].value == "A" and ws["H3"].value == dt.datetime(2026, 11, 19)
    assert ws["C2"].value == "Revue des droits NTFS" and ws["C2"].font.b and ws["C3"].value == "E.Test7 – 2026-11"
    assert {"C2:F2", "C3:F3", "B5:D5", "B9:H9"} <= {str(m) for m in ws.merged_cells.ranges}
    # classification : les trois niveaux, celui de la configuration est mis en évidence
    assert ws["B5"].value == "Classification du document"
    assert [ws.cell(6, c).value for c in (2, 3, 4)] == ["1-Public", "2-Restreint", "3-Confidentiel"]
    assert ws["D6"].font.b and ws["D6"].fill.fgColor.rgb.endswith("0B5394") and not ws["B6"].font.b
    assert "noms et identifiants" in ws["B7"].value
    # tableau « Mise à jour »
    assert ws["B9"].value == "Mise à jour" and ws["B9"].fill.fgColor.rgb.endswith("0B5394")
    assert [ws.cell(10, c).value for c in range(2, 9)] == ["Auteur", "Date", "Version", "Nature de la modification",
                                                           "Approbation", "Validation", "Status"]
    assert ws["B11"].value == "Shine (génération automatique)" and ws["C11"].value == dt.datetime(2026, 11, 19)
    assert ws["D11"].value == "1.0" and "Création du classeur" in ws["E11"].value and "Shine v2.3.0" in ws["E11"].value
    assert ws["G11"].value == "='Instructions'!$C$" + str(frow(openpyxl.load_workbook(books["E.Test7"])["Instructions"],
                                                            "Statut de la validation")) and ws["H11"].value == "Émis"
    assert ws["B11"].fill.fgColor.rgb.endswith("CFE2F3")
    assert ws["B11"].protection.locked and ws["G11"].protection.locked and not ws["F11"].protection.locked   # approbation : à saisir
    for r in range(12, 18):                                                 # 6 lignes libres, déverrouillées
        assert all(not ws.cell(r, c).protection.locked and ws.cell(r, c).fill.fgColor.rgb.endswith("CFE2F3") for c in range(2, 9))
    assert ws.cell(18, 2).value is None and ws.protection.sheet and ws.page_setup.orientation == "portrait"
    # plus de bloc de signature dans l'En-tête (il est dans Instructions)
    assert not any(c.value == "Signature (saisissez votre nom)" for row in ws.iter_rows() for c in row)


def test_cover_config_values(ntftest, cfg, tmp_path):
    cfg["review"].update(version="2.3", index="C", classification="2-Restreint", author="Équipe Sécurité",
                         footer_notice="Document propriété de Exemple & Cie")
    cfg["excel"]["history_rows"] = 3
    _, paths = build_all(ntftest, cfg, tmp_path)
    ws = openpyxl.load_workbook(paths["E.Test7"])["En-tête"]
    assert ws["H1"].value == "2.3" and ws["H2"].value == "C" and ws["D11"].value == "2.3" and ws["B11"].value == "Équipe Sécurité"
    assert ws["C6"].font.b and not ws["D6"].font.b                              # 2-Restreint mis en évidence
    assert not ws["B14"].protection.locked and ws["B15"].fill.fill_type is None            # 3 lignes libres seulement (12 à 14)
    footer = zipfile.ZipFile(paths["E.Test7"]).read("xl/worksheets/sheet1.xml").decode()
    assert "Document propriété de Exemple &amp;&amp; Cie" in footer and "KFF0000" in footer


@pytest.mark.parametrize("bad", ["4-Secret", "", None])
def test_config_rejects_bad_classification(bad, cfg):
    import config as _config
    cfg["review"]["classification"] = bad
    with pytest.raises(_config.ConfigError):
        _config.validate(cfg)


# --------------------------------------------------------------------------- Instructions / couleurs / libellés
def test_instructions_texts(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Instructions"]
    texts = [str(c.value) for row in ws.iter_rows() for c in row if c.value]
    joined = " ".join(texts)
    assert any(t.startswith("Chaque ligne se valide séparément.") and "Lecture et en Modification" in t and "vous-même" in t for t in texts)
    assert any(t.startswith("Veuillez traiter la revue des droits NTFS avant le [date à renseigner].") for t in texts)
    assert "Enregistrez le fichier" not in joined and "renvoyez-le à" not in joined       # ancienne phrase retirée
    assert "À vérifier" not in joined and "À préciser" not in joined and "légitime" not in joined.lower()
    assert "Accès autorisé ?" in joined and "Récapitulatif" in joined and "Validation et signature" in joined
    assert [ws.cell(r, 1).value for r in range(1, 40) if isinstance(ws.cell(r, 1).value, int)] == list(range(1, 10))


def test_instructions_signature_block(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["Instructions"]
    top = frow(ws, "Nom et prénom")
    assert [ws.cell(top + i, 2).value for i in range(6)] == [
        "Nom et prénom", "Fonction", "Date de validation (jour / mois / année)",
        "Je certifie avoir examiné toutes les lignes et le récapitulatif", "Signature (saisissez votre nom)",
        "Statut de la validation"]
    assert all(not ws.cell(top + i, 3).protection.locked for i in (0, 1, 3, 4))            # champs à remplir
    assert all(not ws.cell(top + 2, col).protection.locked for col in (3, 4, 5))            # jour | mois | année sur une ligne
    status = ws.cell(top + 5, 3)
    assert status.protection.locked and 'IF(AND(LEN(C' in status.value and 'COUNTA(C' in status.value
    assert '"Signé","Non signé"' in status.value
    assert ws.cell(frow(ws, "Validation et signature"), 2).value == "Validation et signature"
    dv = {str(d.sqref): d for d in ws.data_validations.dataValidation}
    # date : jour | mois | année dans trois listes sur la même ligne (pas de calendrier possible dans un .xlsx)
    assert dv[f"C{top + 2}"].formula1.startswith('"1,2,3') and dv[f"C{top + 2}"].formula1.rstrip('"').endswith("31")
    assert "01 - Janvier" in dv[f"D{top + 2}"].formula1 and "12 - Décembre" in dv[f"D{top + 2}"].formula1
    assert dv[f"E{top + 2}"].formula1.count(",") == 2
    assert f"C{top + 3}" not in dv                                                           # attestation = case à cocher
    assert not any(d.type == "date" for d in ws.data_validations.dataValidation)
    cal = wb["_calc"]
    # dates proposées (historique de l'En-tête) à partir du jour d'ouverture du classeur (TODAY)
    assert cal["Q2"].value == "=TODAY()" and cal["Q3"].value == "=TODAY()+1" and cal["Q92"].value == "=TODAY()+90" and cal["Q93"].value is None
    text = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "signature numérique" in text and "En-tête" not in text.replace("Terminez", "")  # plus de renvoi vers l'En-tête
    meta = {r[0].value: r[1].value for r in openpyxl.load_workbook(books["E.Test7"])["_meta"][
        get_table(wb["_meta"], "MetaKV").ref] if r[0].value}
    assert meta["signature_sheet"] == "Instructions" and f"statut C{top + 5}" in meta["signature_cells"]


def test_no_removed_wording_anywhere(books, ntftest, cfg, tmp_path):
    _, native = build_all(ntftest, cfg, tmp_path / "n", mode="native")
    for p in (books["E.Test7"], native["E.Test7"]):
        wb = openpyxl.load_workbook(p)
        text = " ".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value is not None)
        assert "À vérifier" not in text and "À préciser" not in text and "Accès légitime" not in text
    assert "pending" not in cfg["excel"]["colors"] and "pending" not in cfg["texts"]["legend"]


# --------------------------------------------------------------------------- onglet « Récapitulatif »
def test_recap_structure(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["Récapitulatif"]
    header, rows = table_rows(ws, "Recap")
    assert header == ["Action", "Propriétaire du groupe", "Nom", "Identifiant", "Droit souhaité", "Justification", "Ressource"]
    assert len(rows) == 20 + 14                                        # 20 lignes d'ajout + autant que de lignes de ressources
    assert ws["A3"].value == "Ajouts demandés" and ws["B3"].value == "='_calc'!$I$21"
    assert ws["A4"].value == "Retraits demandés" and ws["B4"].value == "='_calc'!$B$15"
    assert ws.column_dimensions["H"].hidden and ws.freeze_panes == "A8" and ws.protection.sheet
    assert all(not c.protection.locked is False for row in ws.iter_rows(min_row=8) for c in row)    # aucune zone de saisie
    cached = openpyxl.load_workbook(books["E.Test7"], data_only=True)["Récapitulatif"]        # à vide : valeurs en cache = vides
    assert all(c.value in (None, "") for row in cached.iter_rows(min_row=8, max_col=7) for c in row)


def test_recap_calc_sheet_is_wired_to_decisions(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    calc = wb["_calc"]
    assert calc["B2"].value == "=IF('Activites_etudes'!$D$7=\"Non\",1,0)" and calc["B3"].value.startswith("=B2+IF('Activites_etudes'!$D$8")
    assert calc["F2"].value == "Activites_etudes_Full" and calc["F5"].value == "Activites_etudes_Read" and calc["G2"].value == "E.Test7"
    assert calc["I2"].value.startswith("=IF(LEN(TRIM('Synthèse'!$A$14&") and calc["I21"].value.startswith("=I20+IF(")
    assert "'_meta'!" in calc["N2"].value and 'IFERROR' in calc["O2"].value
    assert max(c.row for row in calc.iter_rows(max_col=15) for c in row if c.value is not None) == 21  # hors colonne Q (dates)


def test_recap_cap_and_warning(ntftest, cfg, tmp_path):
    cfg["excel"]["recap_max_removals"] = 5
    _, paths = build_all(ntftest, cfg, tmp_path)
    ws = openpyxl.load_workbook(paths["E.Test7"])["Récapitulatif"]
    assert len(table_rows(ws, "Recap")[1]) == 20 + 5
    assert "recap_max_removals" in ws["A5"].value and "25" in ws["A5"].value


@pytest.mark.parametrize("bad", [{"add_rows": 0}, {"recap_max_removals": -1}, {"history_rows": "x"}])
def test_config_rejects_bad_counts(bad, cfg):
    import config as _config
    cfg["excel"].update(bad)
    with pytest.raises(_config.ConfigError):
        _config.validate(cfg)


# --------------------------------------------------------------------------- récapitulatif : recalcul réel (LibreOffice)
import shutil
import subprocess


def _recalc(path, tmp):
    out = tmp / "lo"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(path)],
                   capture_output=True, timeout=240, check=True)
    return openpyxl.load_workbook(out / path.name, data_only=True)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice absent")
@pytest.mark.parametrize("mode", ["dropdown", "native"])
def test_recap_live_with_libreoffice(ntftest, cfg, tmp_path, mode):
    _, paths = build_all(ntftest, cfg, tmp_path / "b", mode=mode)
    wb = openpyxl.load_workbook(paths["E.Test7"])
    a, d = wb["Activites_etudes"], wb["Documentation_INTERNE"]
    if mode == "dropdown":
        a["D7"], a["D8"], a["D10"], d["D7"] = "Non", "Oui", "Non", "Non"
    else:                                                      # Conserver = D, Retirer = E ; les deux cochées = non traité
        a["E7"], a["D8"], a["E10"], d["E7"], a["D9"], a["E9"] = True, True, True, True, True, True
    a["E7"].value = a["E7"].value                              # (no-op) lisibilité
    com = "E" if mode == "dropdown" else "F"
    a[f"{com}7"], a[f"{com}10"] = "Départ", "Doublon"
    syn = wb["Synthèse"]
    syn["A14"], syn["B14"], syn["C14"], syn["F14"] = "Jean Test", "jt", "Activites_etudes_Full", "Nouveau"
    syn["A16"], syn["F16"] = "Sans groupe", "précisez"
    src = tmp_path / "filled.xlsx"
    wb.save(src)
    w = _recalc(src, tmp_path)
    r = w["Récapitulatif"]
    got = [[c for c in row[:7]] for row in r.iter_rows(min_row=8, max_row=14, values_only=True)]
    assert (r["B3"].value, r["B4"].value) == (2, 3)
    assert got[0] == ["Ajout", "E.Test7", "Jean Test", "jt", "Modification", "Nouveau", "Activites_etudes_Full"]
    assert got[1] == ["Ajout", "E.Test7", "Sans groupe", None, None, "précisez", None]             # propriétaire déduit du classeur
    assert got[2] == ["Retrait", "E.Test7", "AAAA BBBB", "KKK", "Aucun (retrait)", "Départ", "Activites_etudes_Full"]
    assert got[3] == ["Retrait", "E.Test7", "CASA", "CASA", "Aucun (retrait)", "Doublon", "Activites_etudes_Read"]
    assert got[4][:5] == ["Retrait", "E.Test7", "AAAA BBBB", "KKK", "Aucun (retrait)"] and got[4][6] == "Documentation_INTERNE_Full"
    assert all(v in (None, "") for v in got[5] + got[6])
    assert w["En-tête"]["G11"].value == "Non signé"                                              # validation reprise de la signature
    treated = w["Synthèse"]["F5"].value                                                            # Activites_etudes : lignes 7, 8, 10 traitées
    assert treated == 3 and w["Synthèse"]["G5"].value == 4


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice absent")
def test_signature_status_live(ntftest, cfg, tmp_path):
    _, paths = build_all(ntftest, cfg, tmp_path / "b")
    wb = openpyxl.load_workbook(paths["E.Test7"])
    ws = wb["Instructions"]
    top = frow(ws, "Nom et prénom")
    ws.cell(top, 3).value = "Marie Dupont"
    ws.cell(top + 2, 3).value, ws.cell(top + 2, 4).value, ws.cell(top + 2, 5).value = 20, "11 - Novembre", 2026
    ws.cell(top + 3, 3).value = True
    src = tmp_path / "s.xlsx"
    wb.save(src)
    got = _recalc(src, tmp_path)
    assert got["Instructions"].cell(top + 5, 3).value == "Non signé"                      # signature manquante
    ws.cell(top + 4, 3).value = "Marie Dupont"
    wb.save(src)
    got = _recalc(src, tmp_path)
    assert got["Instructions"].cell(top + 5, 3).value == "Signé"
    ws.cell(top + 2, 3).value, ws.cell(top + 2, 4).value = 31, "02 - Février"             # date impossible
    wb.save(src)
    got = _recalc(src, tmp_path)
    assert got["Instructions"].cell(top + 5, 3).value == "Non signé"
    ws.cell(top + 2, 3).value, ws.cell(top + 2, 4).value, ws.cell(top + 2, 5).value = 20, "11 - Novembre", None   # année manquante
    wb.save(src)
    got = _recalc(src, tmp_path)
    assert got["Instructions"].cell(top + 5, 3).value == "Non signé"
    ws.cell(top + 2, 5).value = 2026
    wb.save(src)
    got = _recalc(src, tmp_path)
    assert got["En-tête"]["G11"].value == "Signé"                                          # repris dans le tableau « Mise à jour »


def test_sous_dossiers_distingues(cfg, tmp_path):
    """Des sous-dossiers (Alpha, Beta, Gamma…) dans un export donnent des onglets distincts, nommés d'après le sous-dossier."""
    body = '﻿"name,""distinguishedName"",""SamAccountName"""\n"AAAA BBBB,""CN=BBBB AAAA,OU=T,DC=TEST,DC=COM"",""KKK"""\n'
    root = tmp_path / "in"
    for sub in ("Delta", "Gamma", "Alpha", "Beta"):
        for lvl in ("Full", "Read"):
            f = root / "Export_E.Martin" / sub / f"Res1_{lvl}.csv"
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text(body, encoding="utf-8", newline="")
    (root / "Export_E.Martin" / "Autre_Full.csv").write_text(body, encoding="utf-8", newline="")
    _, paths = build_all(root, cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["E.Martin"])
    names = [n for n in wb.sheetnames if n not in ("En-tête", "Instructions", "Synthèse", "Récapitulatif", "_calc", "_meta")]
    assert sorted(names) == ["Autre", "Res1_Alpha", "Res1_Beta", "Res1_Delta", "Res1_Gamma"]
    assert wb["Res1_Delta"]["B2"].value == "Export_E.Martin / Delta"
    syn = [r for r in wb["Synthèse"].iter_rows(min_row=5, max_row=9, values_only=True) if r[0]]
    assert sorted(r[1] for r in syn if r[0] == "Res1") == ["Export_E.Martin / Alpha", "Export_E.Martin / Beta",
                                                          "Export_E.Martin / Delta", "Export_E.Martin / Gamma"]


def _unsuffixed_root(tmp_path):
    root = tmp_path / "in_ns"
    d = root / "Export_E.Un"
    write_export(d / "Delta.csv", [("A A", dn_user("A A"), "aa"), ("B B", dn_user("B B"), "bb")])
    write_export(d / "Beta.csv", [("C C", dn_user("C C"), "cc")])
    write_export(d / "Res_Full.csv", [("A A", dn_user("A A"), "aa")])
    return root


def test_no_unspecified_column_without_unsuffixed_files(books):
    syn = openpyxl.load_workbook(books["E.Test7"])["Synthèse"]
    assert [c.value for c in syn[4]][:8] == ["Ressource", "Dossier", "Membres Full", "Membres Read", "Lignes à traiter",
                                              "Traitées", "Reste", "Avancement"]


def test_unsuffixed_files_in_workbook(cfg, tmp_path):
    """Delta.csv / Beta.csv : onglets avec « Non précisé », colonne dédiée dans la Synthèse (absente sinon)."""
    _, paths = build_all(_unsuffixed_root(tmp_path), cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["E.Un"])
    assert wb["Delta"]["C7"].value == "Non précisé" and wb["Delta"]["C8"].value == "Non précisé"
    assert "Droit non précisé : Delta (2)" in wb["Delta"]["B3"].value
    syn = wb["Synthèse"]
    assert [c.value for c in syn[4]][:9] == ["Ressource", "Dossier", "Membres Full", "Membres Read",
                                              "Membres (droit non précisé)", "Lignes à traiter", "Traitées", "Reste", "Avancement"]
    rows = {r[0]: r for r in syn.iter_rows(min_row=5, max_row=7, values_only=True)}
    assert rows["Delta"][2:6] == (0, 0, 2, 2) and rows["Beta"][2:6] == (0, 0, 1, 1) and rows["Res"][2:6] == (1, 0, 0, 1)
    assert syn["E8"].value.startswith("=SUBTOTAL(109,E5:E7") and syn["F8"].value.startswith("=SUBTOTAL(109,F5:F7")


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice absent")
def test_unsuffixed_recap_live(cfg, tmp_path):
    _, paths = build_all(_unsuffixed_root(tmp_path), cfg, tmp_path / "out")
    wb = openpyxl.load_workbook(paths["E.Un"])
    q = wb["Delta"]
    q["D7"], q["E7"] = "Non", "Départ"
    syn = wb["Synthèse"]
    first_add = next(r for r in range(1, 60) if syn.cell(r, 1).value == "Nom") + 1
    syn.cell(first_add, 1).value, syn.cell(first_add, 2).value, syn.cell(first_add, 3).value = "Jean Test", "jt", "Beta"
    syn.cell(first_add, 6).value = "Besoin"
    src = tmp_path / "f.xlsx"
    wb.save(src)
    w = _recalc(src, tmp_path)
    r = w["Récapitulatif"]
    got = [list(row[:7]) for row in r.iter_rows(min_row=8, max_row=9, values_only=True)]
    assert got[0][:6] == ["Ajout", "E.Un", "Jean Test", "jt", "Non précisé", "Besoin"]
    assert got[1][:5] == ["Retrait", "E.Un", "A A", "aa", "Aucun (retrait)"]
    s = w["Synthèse"]
    assert [s.cell(r, 7).value for r in (5, 6, 7)] == [0, 1, 0]                          # « Traitées » : une ligne de Delta traitée
    assert [s.cell(8, c).value for c in (3, 4, 5, 6)] == [1, 0, 3, 4]


def test_cover_approval_from_config(cfg, ntftest, tmp_path):
    assert build_all(ntftest, cfg, tmp_path / "a")[1]["E.Test7"] and \
        openpyxl.load_workbook(tmp_path / "a" / "E.Test7.xlsx")["En-tête"]["F11"].value is None       # vide par défaut : à saisir
    cfg["review"]["approval"] = "J. Martin"
    paths = build_all(ntftest, cfg, tmp_path / "b")[1]
    ws = openpyxl.load_workbook(paths["E.Test7"])["En-tête"]
    assert ws["F11"].value == "J. Martin" and not ws["F11"].protection.locked
