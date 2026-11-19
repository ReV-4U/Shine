import csv
import io
import shutil
from pathlib import Path

import pytest

from config import load_config

# Jeu NTFTEST reconstitué (mêmes octets que l'archive fournie : BOM, CRLF, format « double » mal formé).

C1 = b'\xef\xbb\xbf"name,""distinguishedName"",""SamAccountName"""\r\n"AAAA BBBB,""CN=BBBB AAAA,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""KKK"""\r\n"CCCC XXXX,""CN=XXXX CCCC,OU=METHODES_DUSTRIELLES,OU=TEST_DEVELOPPEMENT_&_INDUSTRIALISATION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""MMM"""\r\n"DDDD L?LLL,""CN=DDDD L?LLL,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""GGG"""\r\n'
C2 = b'\xef\xbb\xbf"name,""distinguishedName"",""SamAccountName"""\r\n"ZZZZ BBBB,""CN=BBBB AAAA,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""KKK"""\r\n"GGGG XXXX,""CN=XXXX CCCC,OU=METHODES_DUSTRIELLES,OU=TEST_DEVELOPPEMENT_&_INDUSTRIALISATION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""MMM"""\r\n"OOOO L?LLL,""CN=DDDD L?LLL,OU=TEST_DIRECTION,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""GGG"""\r\n"CASA,""CN=CASA,OU=TEST_01_GENERIQUES,OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"",""CASA"""\r\n'
C3 = b'\xef\xbb\xbf"name,""distinguishedName"",""SamAccountName"""\r\n"Utilisa. du domaine,""CN=Utilisa. du domaine,CN=Users,DC=TEST,DC=COM"",""Utilisa. du domaine"""\r\n'
C4 = b'\xef\xbb\xbf\r\n'

NTFTEST_FILES = {
    'Nov/Liaison/Export_E.Test7/Activites_etudes_Full.csv': C1,
    'Nov/Liaison/Export_E.Test7/Activites_etudes_Read.csv': C2,
    'Nov/Liaison/Export_E.Test7/Documentation_INTERNE_Full.csv': C1,
    'Nov/Liaison/Export_E.Test7/Documentation_INTERNE_Read.csv': C3,
    'Nov/Liaison/Export_E.Test7/TEST_CHINE_Full.csv': C1,
    'Nov/Liaison/Export_E.Test7/TEST_CHINE_Read.csv': C4,
    'Nov/Liaison/Export_F.Test10/CHINE_Full.csv': C1,
    'Nov/Liaison/Export_F.Test10/CHINE_Read.csv': C4,
    'Nov/Liaison/Export_F.Test10/Documentation_INTERNE_Full.csv': C1,
    'Nov/Liaison/Export_F.Test10/Documentation_INTERNE_Read.csv': C3,
    'Nov/Liaison/Export_F.Test10/Test_fichier_Full.csv': C1,
    'Nov/Liaison/Export_F.Test10/Test_fichier_read.csv': C2,
    'Nov/Liaison/Export_G.Test1/Documentation_INTERNE_Full.csv': C1,
    'Nov/Liaison/Export_G.Test1/Documentation_INTERNE_Read.csv': C3,
    'Nov/Liaison/Export_G.Test1/Formeld_Full.csv': C1,
    'Nov/Liaison/Export_G.Test1/Formeld_read.csv': C2,
    'Nov/Liaison/Export_G.Test1/MEXIQUE_Full.csv': C1,
    'Nov/Liaison/Export_G.Test1/MEXIQUE_Read.csv': C4,
    'Nov/Liaison/Export_G.Test1/ME_fg-_Full.csv': C1,
    'Nov/Liaison/Export_G.Test1/ME_fg-_Read.csv': C4,
}


@pytest.fixture
def cfg():
    return load_config()


@pytest.fixture
def ntftest(tmp_path):
    """Reconstitue l'arborescence NTFTEST fournie (20 CSV) dans un dossier temporaire."""
    dst = tmp_path / "NTFTEST"
    build_ntftest(dst)
    return dst


def build_ntftest(dst: Path) -> Path:
    for rel, content in NTFTEST_FILES.items():
        p = dst / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(content)
    return dst


def dn_user(cn, ou="TEST_DIRECTION"):
    return f"CN={cn},OU={ou},OU=TEST_Users,OU=001_TEST,DC=TEST,DC=COM"


def write_export(path: Path, members, nested=True, header=True, bom=True):
    """Écrit un export au format de l'AD (CSV « double », mal formé) ou en CSV normal.

    members : liste de (name, dn, sam). Fichier vide (BOM + CRLF, 5 octets) si la liste est vide.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    out = io.StringIO()
    if not members:
        out.write("\r\n")
    else:
        rows = ([("name", "distinguishedName", "SamAccountName")] if header else []) + list(members)
        for name, dn, sam in rows:
            if nested:  # un seul champ entre guillemets, qui contient du CSV avec guillemets doublés
                inner = 'name,"distinguishedName","SamAccountName"' if name == "name" else f'{name},"{dn}","{sam}"'
                csv.writer(out, lineterminator="\r\n", quoting=csv.QUOTE_ALL).writerow([inner])
            else:
                csv.writer(out, lineterminator="\r\n").writerow([name, dn, sam])
    path.write_bytes((b"\xef\xbb\xbf" if bom else b"") + out.getvalue().encode("utf-8"))
    return path
