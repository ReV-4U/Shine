import csv
import datetime as dt
import time

import openpyxl

import revue_droits
from conftest import dn_user, write_export


RAPPORT = f"rapport_lecture_{dt.datetime.now():%d%m%y}.csv"


def run(args, capsys):
    rc = revue_droits.main(args)
    return rc, capsys.readouterr()


def test_end_to_end_ntftest(ntftest, tmp_path, capsys):
    out = tmp_path / "out"
    rc, cap = run(["--input", str(ntftest), "--output", str(out), "--period", "2026-11"], capsys)
    assert rc == 0
    day = dt.datetime.now().strftime("%d%m%y")
    assert sorted(p.name for p in out.iterdir()) == [f"Revue_NTFS_{c}_{day}" for c in ("E.Test7", "F.Test10", "G.Test1")]
    for c in ("E.Test7", "F.Test10", "G.Test1"):                     # un dossier par export, contenant son classeur
        assert [p.name for p in (out / f"Revue_NTFS_{c}_{day}").iterdir()] == [f"Revue_NTFS_{c}_{day}.xlsx"]
    assert "Dossiers traités   : 3" in cap.out and "Ressources         : 10" in cap.out
    assert "Lignes             : 45" in cap.out and "Problèmes de lecture : 0" in cap.out
    for forbidden in ("AAAA", "BBBB", "KKK", "CASA", "Utilisa"):     # aucune donnée personnelle en console
        assert forbidden not in cap.out + cap.err
    wb = openpyxl.load_workbook(out / f"Revue_NTFS_E.Test7_{day}" / f"Revue_NTFS_E.Test7_{day}.xlsx")
    assert sum(1 for _ in wb["Activites_etudes"].tables) == 1


def test_read_report_only_when_needed(tmp_path, capsys):
    root = tmp_path / "in"
    p = write_export(root / "Export_A.Un" / "R_Full.csv", [("A A", dn_user("A A"), "aa")])
    p.write_bytes(p.read_bytes() + b"ligne cassee\r\n")
    write_export(root / "Export_A.Un" / "Machin.csv", [("B B", dn_user("B B"), "bb")])
    out = tmp_path / "out"
    rc, cap = run(["--input", str(root), "--output", str(out)], capsys)
    raw = (out / RAPPORT).read_bytes()
    assert rc == 0 and raw.startswith(b"\xef\xbb\xbf")
    rows = list(csv.reader(raw.decode("utf-8-sig").splitlines(), delimiter=";"))
    assert rows[0] == ["code", "dossier", "ressource", "fichier", "detail"]
    assert sorted(r[0] for r in rows[1:]) == ["LIGNE_NON_PARSABLE", "SUFFIXE_INCONNU"]
    assert "Problèmes de lecture : 2" in cap.out and RAPPORT in cap.out
    p.write_bytes(p.read_bytes().replace(b"ligne cassee\r\n", b""))
    (root / "Export_A.Un" / "Machin.csv").unlink()
    run(["--input", str(root), "--output", str(out)], capsys)
    assert not (out / RAPPORT).exists()               # pas de rapport périmé


def test_output_inside_input_can_be_rerun(tmp_path, capsys):
    root = tmp_path / "in"
    write_export(root / "Export_A.Un" / "R_Full.csv", [("A A", dn_user("A A"), "aa")])
    p = write_export(root / "Export_A.Un" / "Machin.csv", [("B B", dn_user("B B"), "bb")])
    out = root / "sortie"
    run(["--input", str(root), "--output", str(out)], capsys)
    assert (out / RAPPORT).exists()
    p.unlink()
    rc, cap = run(["--input", str(root), "--output", str(out)], capsys)     # le rapport précédent n'est pas relu comme export
    assert rc == 0 and "Problèmes de lecture : 0" in cap.out and not (out / RAPPORT).exists()


def test_only_option_and_native_mode(ntftest, tmp_path, capsys):
    out = tmp_path / "out"
    rc, cap = run(["--input", str(ntftest), "--output", str(out), "--only", "Export_G.Test1", "--checkbox-mode", "native",
                   "--period", "2026-11"], capsys)
    assert rc == 0 and [p.name for p in out.iterdir()] == [f"Revue_NTFS_G.Test1_{dt.datetime.now():%d%m%y}"]
    assert "Mode de réponse    : native" in cap.out


def test_errors_return_code_2(tmp_path, capsys):
    assert run(["--input", str(tmp_path / "absent"), "--output", str(tmp_path / "o")], capsys)[0] == 2
    assert run(["--input", str(tmp_path), "--output", str(tmp_path / "o"), "--only", "Rien"], capsys)[0] == 2
    bad = tmp_path / "bad.yaml"
    bad.write_text("excel:\n  checkbox_mode: carré\n")
    assert run(["--input", str(tmp_path), "--output", str(tmp_path / "o"), "--config", str(bad)], capsys)[0] == 2


def test_user_config_overrides(ntftest, tmp_path, capsys):
    cfgf = tmp_path / "c.yaml"
    cfgf.write_text("input:\n  workbook_label: '{name}'\nreview:\n  period: '2027-01'\n  deadline: '31/01'\nexcel:\n  show_dn: true\n")
    out = tmp_path / "out"
    rc, _ = run(["--input", str(ntftest), "--output", str(out), "--config", str(cfgf)], capsys)
    f = out / f"Revue_NTFS_Test7_{dt.datetime.now():%d%m%y}" / f"Revue_NTFS_Test7_{dt.datetime.now():%d%m%y}.xlsx"
    assert rc == 0 and f.exists()
    ws = openpyxl.load_workbook(f)["Activites_etudes"]
    assert ws["C6"].value == "DN (annuaire)"


def test_csv_injection_in_read_report(tmp_path, capsys):
    root = tmp_path / "in"
    write_export(root / "Export_A.Un" / "=evil.csv", [("x", dn_user("x"), "x")])
    run(["--input", str(root), "--output", str(tmp_path / "o")], capsys)
    text = (tmp_path / "o" / RAPPORT).read_text("utf-8-sig")
    assert "'=evil" in text and ";=evil" not in text


def test_volume(tmp_path, capsys):
    """Charge réaliste : 6 dossiers × 10 ressources × (300 Full + 200 Read) ≈ 30 000 lignes."""
    root = tmp_path / "in"
    for s in range(6):
        for r in range(10):
            d = root / f"Export_{'ABCDEF'[s]}.S{s}"
            people = [(f"N{i} P{i}", dn_user(f"N{i} P{i}", f"SERV_{i % 9}"), f"u{r}x{i}") for i in range(300)]
            write_export(d / f"Res{r}_Full.csv", people)
            write_export(d / f"Res{r}_Read.csv", people[250:] + [(f"Z{i}", dn_user(f"Z{i}"), f"z{r}x{i}") for i in range(150)])
    t = time.time()
    rc, cap = run(["--input", str(root), "--output", str(tmp_path / "out"), "--period", "2026-11"], capsys)
    dur = time.time() - t
    assert rc == 0 and "Lignes             : 30000" in cap.out
    print(f"\nvolume : {dur:.1f}s")
    assert dur < 90


def test_default_output_root_and_pattern(ntftest, tmp_path, capsys):
    """Sans --output : dossier racine de la config (créé), nom de dossier et format de date paramétrables."""
    root = tmp_path / "Documents" / "shine"
    cfgf = tmp_path / "c.yaml"
    cfgf.write_text(f"output:\n  root: '{root.as_posix()}'\n  folder_pattern: 'R_{{label}}_{{date}}'\n  date_format: '%Y%m%d'\n")
    rc, cap = run(["--input", str(ntftest), "--config", str(cfgf), "--only", "Export_E.Test7"], capsys)
    d = f"{dt.datetime.now():%Y%m%d}"
    assert rc == 0 and (root / f"R_E.Test7_{d}" / f"R_E.Test7_{d}.xlsx").exists()


def test_output_root_defaults_to_documents_shine(monkeypatch):
    import config
    cfg = config.load_config()
    assert cfg["output"]["root"] == "~/Documents/shine" and cfg["output"]["date_format"] == "%d%m%y"
    assert cfg["output"]["folder_pattern"] == "Revue_NTFS_{label}_{date}"


def test_invalid_output_pattern_rejected(tmp_path, capsys):
    bad = tmp_path / "bad.yaml"
    bad.write_text("output:\n  folder_pattern: 'R/{label}'\n")
    assert run(["--input", str(tmp_path), "--output", str(tmp_path / "o"), "--config", str(bad)], capsys)[0] == 2
    bad.write_text("output:\n  folder_pattern: 'R_{inconnu}'\n")
    assert run(["--input", str(tmp_path), "--output", str(tmp_path / "o"), "--config", str(bad)], capsys)[0] == 2


def test_same_day_rerun_overwrites_and_owner_grouping_makes_one_folder(ntftest, tmp_path, capsys):
    out = tmp_path / "out"
    cfgf = tmp_path / "c.yaml"
    cfgf.write_text("owners:\n  Export_E.Test7: Dupont\n  Export_F.Test10: Dupont\n")
    for _ in range(2):
        assert run(["--input", str(ntftest), "--output", str(out), "--config", str(cfgf)], capsys)[0] == 0
    day = f"{dt.datetime.now():%d%m%y}"
    assert sorted(p.name for p in out.iterdir()) == [f"Revue_NTFS_Dupont_{day}", f"Revue_NTFS_G.Test1_{day}"]


def test_shine_process_folder_no_gui(ntftest, tmp_path):
    """Logique du lanceur (sans fenêtre) : résumé en compteurs, dossier de sortie renvoyé, rien d'écrit en console."""
    import shine
    out = tmp_path / "shine_out"
    rc, text, resolved = shine.process_folder(str(ntftest / "Nov"), str(out))
    assert rc == 0 and resolved == out and "Lignes             : 45" in text and "=== Résumé" not in text
    assert len(list(out.iterdir())) == 3
    for forbidden in ("AAAA", "KKK", "CASA"):
        assert forbidden not in text


def test_shine_process_folder_error(tmp_path):
    import shine
    rc, text, resolved = shine.process_folder(str(tmp_path / "absent"), str(tmp_path / "o"))
    assert rc == 2 and "introuvable" in text


def test_shine_main_dispatch(ntftest, tmp_path, capsys, monkeypatch):
    import shine
    monkeypatch.setattr(shine, "gui", lambda: 42)
    assert shine.main([]) == 42                                    # sans argument : fenêtre
    assert shine.main(["--input", str(ntftest), "--output", str(tmp_path / "o")]) == 0   # avec arguments : CLI


def test_user_config_next_to_tool_is_loaded(ntftest, tmp_path, monkeypatch, capsys):
    shine_yaml = tmp_path / "shine.yaml"
    shine_yaml.write_text("input:\n  workbook_label: '{name}'\n")
    monkeypatch.setattr(revue_droits, "app_dir", lambda: tmp_path)
    out = tmp_path / "o"
    assert run(["--input", str(ntftest), "--output", str(out)], capsys)[0] == 0
    assert any(p.name.startswith("Revue_NTFS_Test7_") for p in out.iterdir())


# --------------------------------------------------------------------------- date limite
import pytest  # noqa: E402


@pytest.mark.parametrize("txt,expected", [("30/11/2026", "30/11/2026"), ("5/12/2026", "05/12/2026"), ("2026-12-05", "05/12/2026"),
                                          ("05-12-2026", "05/12/2026"), ("05.12.2026", "05/12/2026"), (" 05/12/26 ", "05/12/2026")])
def test_parse_deadline_ok(txt, expected):
    assert revue_droits.parse_deadline(txt) == expected


@pytest.mark.parametrize("txt", ["", "demain", "31/02/2026", "2026/12", "[date limite à renseigner]"])
def test_parse_deadline_invalid(txt):
    with pytest.raises(ValueError):
        revue_droits.parse_deadline(txt)


def test_cli_deadline_shown_in_instructions(ntftest, tmp_path, capsys):
    out = tmp_path / "out"
    rc, cap = run(["--input", str(ntftest), "--output", str(out), "--only", "Export_G.Test1", "--deadline", "5/12/2026"], capsys)
    assert rc == 0 and "Date limite        : 05/12/2026" in cap.out
    f = next(out.glob("*/*.xlsx"))
    texts = [str(c.value) for row in openpyxl.load_workbook(f)["Instructions"].iter_rows() for c in row if c.value]
    assert any("05/12/2026" in t for t in texts) and not any("[date limite" in t for t in texts)


def test_cli_invalid_deadline_rejected(ntftest, tmp_path, capsys):
    with pytest.raises(SystemExit) as e:
        revue_droits.parse_args(["--input", str(ntftest), "--deadline", "bientôt"])
    assert e.value.code == 2


def test_shine_process_folder_with_deadline(ntftest, tmp_path):
    import shine
    rc, text, _ = shine.process_folder(str(ntftest / "Nov"), str(tmp_path / "o"), deadline="12/12/2026")
    assert rc == 0 and "Date limite        : 12/12/2026" in text


def test_shine_suggest_and_check_deadline():
    import shine
    today = dt.date(2026, 11, 19)
    assert shine.suggest_deadline("[date limite à renseigner]", today) == "03/12/2026"      # aujourd'hui + 14 jours
    assert shine.suggest_deadline(None, today) == "03/12/2026"
    assert shine.suggest_deadline("20/12/2026", today) == "20/12/2026"
    assert shine.check_deadline("19/11/2026", today) == "19/11/2026"                         # aujourd'hui accepté
    for bad in ("18/11/2026", "abc", ""):
        with pytest.raises(ValueError):
            shine.check_deadline(bad, today)


def test_warning_when_deadline_or_contact_left_as_placeholder(ntftest, tmp_path, capsys):
    rc, cap = run(["--input", str(ntftest), "--output", str(tmp_path / "o1"), "--only", "Export_G.Test1"], capsys)
    assert rc == 0 and "ATTENTION : date limite non renseigné(e)" in cap.out and "ATTENTION : contact non renseigné(e)" in cap.out
    cfgf = tmp_path / "c.yaml"
    cfgf.write_text("review:\n  contact: 'securite@exemple.fr'\n")
    rc, cap = run(["--input", str(ntftest), "--output", str(tmp_path / "o2"), "--only", "Export_G.Test1", "--config", str(cfgf),
                   "--deadline", "30/11/2026"], capsys)
    assert rc == 0 and "ATTENTION" not in cap.out
