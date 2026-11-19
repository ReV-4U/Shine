from conftest import dn_user, write_export
from parser import parse_file, parse_text


def E(ntftest):
    return ntftest / "Nov" / "Liaison" / "Export_E.Test7"


def test_ntftest_double_parsing(ntftest):
    r = parse_file(E(ntftest) / "Activites_etudes_Full.csv")
    assert r.fmt == "nested" and not r.issues and not r.empty
    assert [m.sam for m in r.members] == ["KKK", "MMM", "GGG"]
    m = r.members[0]
    assert m.name == "AAAA BBBB"
    assert m.dn == "CN=BBBB AAAA,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"  # le DN fait foi
    assert r.members[2].name == "DDDD L?LLL"      # caractère déjà perdu à l'export : non corrigé


def test_empty_group_file_is_not_an_error(ntftest):
    p = E(ntftest) / "TEST_CHINE_Read.csv"
    assert p.stat().st_size == 5
    r = parse_file(p)
    assert r.empty and not r.members and not r.issues and not r.failed


def test_header_only_is_empty():
    assert parse_text('"name,""distinguishedName"",""SamAccountName"""\r\n').empty


def test_normal_csv_autodetected():
    r = parse_text('name,distinguishedName,SamAccountName\r\nAA BB,"CN=AA BB,OU=X_Users,DC=T",aabb\r\n')
    assert r.fmt == "flat" and r.members[0].sam == "aabb" and r.members[0].dn.startswith("CN=AA BB,OU=X")


def test_semicolon_csv_autodetected():
    r = parse_text('name;distinguishedName;SamAccountName\r\nAA BB;CN=AA BB,OU=X;aabb\r\n')
    assert r.fmt == "flat" and r.members[0].sam == "aabb"


def test_nested_with_semicolon_inside():
    r = parse_text('"name;distinguishedName;SamAccountName"\r\n"AA;CN=AA,OU=X;aa"\r\n')
    assert r.fmt == "nested" and [m.sam for m in r.members] == ["aa"]


def test_broken_lines_are_reported_and_skipped():
    text = ('"name,""distinguishedName"",""SamAccountName"""\r\n'
            '"OK1,""CN=OK1,OU=X"",""ok1"""\r\n'
            'ligne cassée sans rien\r\n'
            '"SEUL,""CN=SEUL"""\r\n'
            '"OK2,""CN=OK2,OU=X"",""ok2"""\r\n')
    r = parse_text(text)
    assert [m.sam for m in r.members] == ["ok1", "ok2"]
    assert [i.code for i in r.issues] == ["LIGNE_NON_PARSABLE"] * 2
    assert [i.line for i in r.issues] == [3, 4]


def test_empty_identifier_is_reported():
    r = parse_text('name,distinguishedName,SamAccountName\r\nA,"CN=A",\r\nB,"CN=B",bb\r\n')
    assert [m.sam for m in r.members] == ["bb"] and r.issues[0].code == "LIGNE_NON_PARSABLE"


def test_issues_never_contain_personal_data():
    text = ('"name,""distinguishedName"",""SamAccountName"""\r\n"SECRETNAME,""CN=SECRETNAME"""\r\n')
    r = parse_text(text)
    assert r.issues and all("SECRET" not in i.detail for i in r.issues)


def test_missing_header_falls_back_to_position():
    r = parse_text('AA,"CN=AA,OU=X",aa\r\n')
    assert r.members[0].sam == "aa" and r.issues[0].code == "EN_TETE_ABSENT"


def test_unknown_format_fails_softly():
    r = parse_text("n'importe quoi\r\nautre chose\r\n")
    assert r.failed and r.issues[0].code == "FORMAT_INCONNU"


def test_cp1252_fallback(tmp_path):
    p = tmp_path / "x_Full.csv"
    p.write_bytes('name,distinguishedName,SamAccountName\r\nÉric,"CN=Éric,OU=X",eric\r\n'.encode("cp1252"))
    r = parse_file(p)
    assert r.members[0].name == "Éric" and r.encoding == "cp1252"
    assert any(i.code == "ENCODAGE" for i in r.issues)


def test_control_characters_are_stripped():
    r = parse_text('name,distinguishedName,SamAccountName\r\nA\x00B,"CN=A",a\x07a\r\n')
    assert r.members[0].name == "AB" and r.members[0].sam == "aa"


def test_helper_roundtrip(tmp_path):
    p = write_export(tmp_path / "a_Full.csv", [("AA BB", dn_user("AA BB"), "aabb")])
    r = parse_file(p)
    assert r.fmt == "nested" and r.members[0].sam == "aabb" and r.members[0].name == "AA BB"
