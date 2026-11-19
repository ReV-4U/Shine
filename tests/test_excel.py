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
    ok_formulas = ('="Progression', "=COUNTIF", "=SUBTOTAL", "=E", "=IF", "=IFERROR", "=SUMPRODUCT")
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
    assert wb.sheetnames == ["Instructions", "Synthèse", "Activites_etudes", "Documentation_INTERNE", "TEST_CHINE", "_meta"]
    assert {ws.title: ws.sheet_state for ws in wb}["_meta"] == "hidden"
    assert all(len(n) <= 31 for n in wb.sheetnames)


def test_resource_sheet_layout(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    assert ws["B1"].value == "Activites_etudes" and ws["B2"].value == "Export_E.Test7"
    assert "Modification : Activites_etudes_Full (3)" in ws["B3"].value and "Lecture : Activites_etudes_Read (4)" in ws["B3"].value
    assert ws["B4"].value == dt.datetime(2026, 11, 19)
    header, rows = table_rows(ws, "Membres")
    assert header == ["Nom", "Identifiant", "Droit actuel", "Accès légitime ?", "Commentaire", "Clé"]   # ni service, ni type, ni DN
    # une ligne par ligne d'export : les 3 comptes apparaissent en Modification ET en Lecture, sans fusion
    assert [(r[1], r[2]) for r in rows] == [("KKK", "Modification"), ("MMM", "Modification"), ("GGG", "Modification"),
                                            ("CASA", "Lecture"), ("MMM", "Lecture"), ("GGG", "Lecture"), ("KKK", "Lecture")]
    assert [r[0] for r in rows][:3] == ["AAAA BBBB", "CCCC XXXX", "DDDD L?LLL"] and rows[6][0] == "ZZZZ BBBB"
    assert ws.freeze_panes == "A7"
    assert ws.page_setup.orientation == "landscape" and ws.page_setup.fitToHeight == 0
    assert ws.sheet_properties.pageSetUpPr.fitToPage and ws.print_title_rows == "$6:$6"
    assert sorted(str(m) for m in ws.merged_cells.ranges) == ["B3:D3", "B5:E5"]   # en-tête seulement : aucune fusion dans la zone de données
    assert get_table(ws, "Membres").autoFilter is not None
    assert ws.max_row == 13 and ws.column_dimensions["F"].hidden              # plus de bloc d'ajout sur les feuilles de ressource
    assert "Personnes à ajouter" not in " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None
                                                 and not str(c.value).startswith("↓"))
    assert ws["E3"].hyperlink.location == "'Synthèse'!A11" and ws["E3"].value.startswith("↓ Personnes à ajouter")


def test_default_state_is_untreated(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    assert all(ws.cell(r, 4).value is None for r in range(7, 14))      # liste vide = non traité


def test_validation_and_conditional_formats(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Activites_etudes"]
    dv = {str(d.sqref): d for d in ws.data_validations.dataValidation}
    assert list(dv) == ["D7:D13"] and dv["D7:D13"].formula1 == '"Oui,Non,À vérifier"' and dv["D7:D13"].errorStyle in (None, "stop")
    cf = [(str(rng.sqref), r.formula[0], r.stopIfTrue) for rng, rules in ws.conditional_formatting._cf_rules.items() for r in rules]
    assert [f for _, f, _ in cf] == ['$D7="Oui"', '$D7="Non"', '$D7="À vérifier"', "MOD(ROW(),2)=0", "MOD(ROW(),2)=1"]
    assert all(s == "A7:D13" and stop for s, _, stop in cf)            # vert / rouge / jaune / gris zébré


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
    assert header == ["Nom", "Identifiant", "DN (annuaire)", "Droit actuel", "Accès légitime ?", "Commentaire", "Clé"]
    assert rows[0][2] == "CN=BBBB AAAA,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"    # DN brut
    assert str(ws.data_validations.dataValidation[0].sqref) == "E7:E13"


def test_protection_without_password_and_unlocked_cells(books):
    z = zipfile.ZipFile(books["E.Test7"])
    xml = z.read("xl/worksheets/sheet3.xml").decode()
    prot = re.search(r"<sheetProtection[^>]*>", xml).group(0)
    assert "password" not in prot.lower()
    assert 'sort="0"' in prot and 'autoFilter="0"' in prot          # filtre et tri autorisés
    wb = openpyxl.load_workbook(books["E.Test7"])
    ws = wb["Activites_etudes"]
    assert ws.protection.sheet
    unlocked = {c.coordinate for row in ws.iter_rows() for c in row if not c.protection.locked}
    assert {f"{col}{r}" for col in "DE" for r in range(7, 14)} | {"B5"} <= unlocked
    for r in range(7, 14):                                           # nom, identifiant, droit, clé : toujours verrouillés
        assert all(ws.cell(r, c).protection.locked for c in (1, 2, 3, 6))
    assert wb["Synthèse"].protection.sheet and wb["Instructions"].protection.sheet


def test_synthese(books):
    ws = openpyxl.load_workbook(books["E.Test7"])["Synthèse"]
    assert ws.freeze_panes == "A5"                                    # seules les lignes d'en-tête sont figées, pas la colonne A
    header, rows = table_rows(ws, "Synthese")
    assert header == ["Ressource", "Dossier", "Membres Full", "Membres Read", "Lignes à traiter", "Traitées",
                      "À préciser", "Reste", "Avancement"]
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
    assert cf[:3] == ["AND($D7=TRUE,$E7<>TRUE)", "AND($E7=TRUE,$D7<>TRUE)", "AND($D7=TRUE,$E7=TRUE)"]
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
    assert rows[0][2:7] == [0, 0, 0, 0, 0]
    assert rows[0][7] == "=E5-F5-G5"                                  # reste : formule, valeur 0


def test_empty_read_group_is_shown_with_zero(books):
    ws = openpyxl.load_workbook(books["G.Test1"])["MEXIQUE"]            # MEXIQUE_Read : 5 octets, groupe vide
    assert "Lecture : MEXIQUE_Read (0)" in ws["B3"].value and "Modification : MEXIQUE_Full (3)" in ws["B3"].value
    assert ws.max_row >= 9 and sum(1 for r in range(7, 10) if ws.cell(r, 3).value == "Modification") == 3


# --------------------------------------------------------------------------- _meta (contrat phase 2)
def test_meta_contract(books):
    wb = openpyxl.load_workbook(books["E.Test7"])
    meta = wb["_meta"]
    kv = {r[0].value: r[1].value for r in meta[get_table(meta, "MetaKV").ref] if r[0].value}
    assert kv["schema_version"] == "3" and kv["folders"] == "Export_E.Test7" and kv["checkbox_mode"] == "dropdown"
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
        assert f"C{r}:E{r}" in merged and f"F{r}:I{r}" in merged
    assert len(merged) == 2 * 21                                   # en-tête + 20 lignes, deux fusions par ligne
    # liste déroulante « Groupe de ressource » : plage de l'onglet masqué _meta = noms des fichiers CSV (sans extension)
    dv = {str(d.sqref): d for d in ws.data_validations.dataValidation}[f"C{first}:C{last}"]
    meta = wb["_meta"]
    ref = get_table(meta, "MetaGroups").ref
    group_rows = [[c.value for c in r] for r in meta[ref]]
    assert group_rows[0] == ["group", "sheet", "folder", "subpath", "resource", "level", "file"]
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
    res_names = openpyxl.load_workbook(paths["A.Un"]).sheetnames[2:-1]
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
