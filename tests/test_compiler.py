from collections import Counter

import pytest

import compiler
from compiler import compile_all, group_books, row_key
from conftest import dn_user, write_export
from model import FULL, READ

U = lambda n, sam, ou="TEST_DIRECTION": (n, dn_user(n, ou), sam)  # noqa: E731


def res_of(result, scope, name):
    s = next(s for s in result.scopes if s.name == scope)
    return next(r for r in s.resources if r.name == name)


# --------------------------------------------------------------------------- NTFTEST
def test_ntftest_overview(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    assert [s.name for s in r.scopes] == ["Export_E.Test7", "Export_F.Test10", "Export_G.Test1"]
    assert [s.label for s in r.scopes] == ["E.Test7", "F.Test10", "G.Test1"]
    assert r.stats["resources"] == 10 and r.stats["files_read"] == 20 and r.stats["files_empty"] == 4
    assert r.stats["rows"] == 45                      # = toutes les lignes des 20 CSV, sans fusion ni dédoublonnage
    names = {s.name: [x.name for x in s.resources] for s in r.scopes}
    assert names["Export_E.Test7"] == ["Activites_etudes", "Documentation_INTERNE", "TEST_CHINE"]
    assert names["Export_F.Test10"] == ["CHINE", "Documentation_INTERNE", "Test_fichier"]
    assert names["Export_G.Test1"] == ["Documentation_INTERNE", "Formeld", "ME_fg-", "MEXIQUE"]  # « ME_fg- » garde son tiret
    assert r.problems == []                           # fichiers vides, mêmes personnes en Full et Read… : rien à signaler


def test_same_person_in_full_and_read_gives_two_rows(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    t = res_of(r, "Export_F.Test10", "Test_fichier")       # Test_fichier_Full + Test_fichier_read (casse différente)
    assert t.full.level == FULL and t.read.level == READ
    assert [(p.sam, p.level) for p in t.rows] == [
        ("KKK", FULL), ("MMM", FULL), ("GGG", FULL),            # Modification d'abord
        ("CASA", READ), ("MMM", READ), ("GGG", READ), ("KKK", READ)]   # puis Lecture, triée par nom affiché
    assert len({p.key for p in t.rows}) == 7                    # clés distinctes, même pour KKK en Full et en Read
    # pas de réconciliation des noms : chaque ligne garde le nom de SON fichier
    assert [p.name for p in t.rows if p.sam == "KKK"] == ["AAAA BBBB", "ZZZZ BBBB"]


def test_data_is_shown_as_exported(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    t = res_of(r, "Export_E.Test7", "Activites_etudes")
    p = next(p for p in t.rows if p.sam == "GGG" and p.level == FULL)
    assert p.name == "DDDD L?LLL"                      # caractère déjà perdu à l'export : non corrigé
    assert p.dn == "CN=DDDD L?LLL,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"   # DN brut, non interprété
    d = res_of(r, "Export_E.Test7", "Documentation_INTERNE")
    assert [(p.sam, p.level) for p in d.rows] == [("KKK", FULL), ("MMM", FULL), ("GGG", FULL), ("Utilisa. du domaine", READ)]


def test_never_merges_across_folders(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    docs = [x for s in r.scopes for x in s.resources if x.name == "Documentation_INTERNE"]
    assert len(docs) == 3
    keys = [{p.key for p in x.rows} for x in docs]
    assert all(not (keys[i] & keys[j]) for i in range(3) for j in range(i + 1, 3))


def test_empty_group(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    c = res_of(r, "Export_G.Test1", "MEXIQUE")
    assert c.read.empty and c.count(READ) == 0 and c.count(FULL) == 3 and len(c.rows) == 3


# --------------------------------------------------------------------------- cas limites
def test_no_dedup_no_merge(tmp_path, cfg):
    d = tmp_path / "Export_A.Un"
    write_export(d / "Res_Full.csv", [U("Bob Zed", "bz"), U("Ana Bee", "ab"), U("Ana Bee", "AB"), U("Cat Cee", "cc")])
    write_export(d / "Res_Read.csv", [U("Dan Dee", "dd"), U("Ana Bee", "ab")])
    r = compile_all(tmp_path, cfg)
    rows = r.scopes[0].resources[0].rows
    assert [(p.sam, p.level) for p in rows] == [("ab", FULL), ("AB", FULL), ("bz", FULL), ("cc", FULL),
                                                ("ab", READ), ("dd", READ)]          # doublon du Full conservé
    assert len({p.key for p in rows}) == 6
    assert r.problems == []


def test_missing_counterpart_is_not_a_problem(tmp_path, cfg):
    d = tmp_path / "Export_A.Un"
    write_export(d / "Seul_Full.csv", [U("A A", "aa")])
    write_export(d / "Autre_Read.csv", [U("B B", "bb")])
    r = compile_all(tmp_path, cfg)
    assert [(x.name, [(p.sam, p.level) for p in x.rows]) for x in r.scopes[0].resources] == [
        ("Autre", [("bb", READ)]), ("Seul", [("aa", FULL)])]
    assert r.problems == []


def test_unreadable_file_does_not_stop_everything(tmp_path, cfg, monkeypatch):
    d = tmp_path / "Export_A.Un"
    write_export(d / "Bon_Full.csv", [U("A A", "aa")])
    write_export(d / "Casse_Full.csv", [U("B B", "bb")])
    real = compiler.parse_file
    monkeypatch.setattr(compiler, "parse_file",
                        lambda p: (_ for _ in ()).throw(PermissionError("refusé")) if "Casse" in str(p) else real(p))
    r = compile_all(tmp_path, cfg)
    assert [p.code for p in r.problems] == ["FICHIER_ILLISIBLE"] and r.stats["files_unreadable"] == 1
    assert [x.name for x in r.scopes[0].resources] == ["Bon", "Casse"]
    assert res_of(r, "Export_A.Un", "Casse").count(FULL) is None
    assert res_of(r, "Export_A.Un", "Bon").rows[0].sam == "aa"


def test_broken_line_is_logged_and_processing_continues(tmp_path, cfg):
    p = write_export(tmp_path / "Export_A.Un" / "R_Full.csv", [U("A A", "aa")])
    p.write_bytes(p.read_bytes() + b"ligne cassee\r\n")
    r = compile_all(tmp_path, cfg)
    assert [x.code for x in r.problems] == ["LIGNE_NON_PARSABLE"] and len(r.scopes[0].resources[0].rows) == 1
    assert r.stats["lines_unparsable"] == 1


def test_unknown_suffix_ignored_and_reported(tmp_path, cfg):
    write_export(tmp_path / "Export_A.Un" / "Machin.csv", [U("A A", "aa")])
    r = compile_all(tmp_path, cfg)
    assert [p.code for p in r.problems] == ["SUFFIXE_INCONNU"] and not r.scopes


def test_recursive_walk_and_scope_detection(tmp_path, cfg):
    write_export(tmp_path / "x" / "y" / "Export_Z.Deep" / "sub" / "R_Full.csv", [U("A A", "aa")])
    write_export(tmp_path / "hors" / "Q_Full.csv", [U("A A", "aa")])
    r = compile_all(tmp_path, cfg)
    scopes = {s.name: s for s in r.scopes}
    assert scopes["Export_Z.Deep"].resources[0].subpath == "sub"
    assert scopes["hors"].resources[0].name == "Q" and r.problems == []


def test_only_filter_by_name_or_label(ntftest, cfg):
    assert [s.name for s in compile_all(ntftest, cfg, only=["export_f.test10"]).scopes] == ["Export_F.Test10"]
    assert [s.name for s in compile_all(ntftest, cfg, only=["G.Test1"]).scopes] == ["Export_G.Test1"]


def test_output_folder_inside_input_is_ignored(tmp_path, cfg):
    write_export(tmp_path / "Export_A.Un" / "R_Full.csv", [U("A A", "aa")])
    out = tmp_path / "sortie"
    out.mkdir()
    (out / "rapport_lecture.csv").write_text("code;dossier\r\n", encoding="utf-8")
    r = compile_all(tmp_path, cfg, exclude=out)
    assert r.problems == [] and r.stats["resources"] == 1


def test_row_key_is_stable_and_scoped():
    k = row_key("Export_A.Un", "", "Res", FULL, "ab")
    assert k == row_key("export_a.un", "", "RES", FULL, "AB") and len(k) == 16
    assert len({k, row_key("Export_B.Deux", "", "Res", FULL, "ab"), row_key("Export_A.Un", "", "Res", READ, "ab"),
                row_key("Export_A.Un", "", "Res", FULL, "ab", 2)}) == 4


def test_deterministic(ntftest, cfg):
    a, b = compile_all(ntftest, cfg), compile_all(ntftest, cfg)
    f = lambda r: [(p.key, p.level, p.name) for s in r.scopes for x in s.resources for p in x.rows]  # noqa: E731
    assert f(a) == f(b)


def test_group_books(ntftest, cfg):
    r = compile_all(ntftest, cfg)
    assert [b.label for b in group_books(r.scopes, cfg)] == ["E.Test7", "F.Test10", "G.Test1"]
    cfg["owners"] = {"export_e.test7": "Dupont", "Export_F.Test10": "Dupont"}
    assert [(b.label, len(b.scopes)) for b in group_books(r.scopes, cfg)] == [("Dupont", 2), ("G.Test1", 1)]


def test_workbook_label_template(cfg):
    cfg["input"]["workbook_label"] = "{name}"
    assert compiler.scope_label("Export_E.Test7", cfg) == ("Test7", "E")
    cfg["input"]["workbook_label"] = "{code}_{name}"
    assert compiler.scope_label("Export_E.Test7", cfg) == ("E_Test7", "E")
    assert compiler.scope_label("Autre dossier", cfg) == ("_Autre dossier", "")
