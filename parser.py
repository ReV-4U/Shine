# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Lecture tolérante des exports CSV d'Active Directory.

Format principal (mal formé) : chaque ligne est UN champ CSV entre guillemets qui contient
lui-même du CSV (guillemets doublés) -> double parsing. Un CSV normal à 3 colonnes est
détecté automatiquement. Aucune donnée personnelle n'est écrite dans les messages d'erreur.
"""
from __future__ import annotations

import csv
import io
import re
from pathlib import Path

from model import Issue, ParseResult, RawMember

DELIMS = (",", ";", "\t")
_ALIASES = {
    "name": "name", "displayname": "name", "nom": "name",
    "distinguishedname": "dn", "dn": "dn",
    "samaccountname": "sam", "sam": "sam", "identifiant": "sam",
}
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_FIRST_LINE = re.compile(r"[^\r\n]*\S[^\r\n]*")


def clean(value: str) -> str:
    """Retire les caractères de contrôle (illégaux en XML/Excel) et les espaces de bord."""
    return _CTRL.sub("", value).strip()


def _split(text: str, delim: str) -> list[str]:
    for row in csv.reader([text], delimiter=delim):
        return row
    return []


def _detect(line: str) -> tuple[str, str | None, str] | None:
    """Renvoie (format, délimiteur externe, délimiteur interne) d'après la ligne d'en-tête."""
    for d in DELIMS:
        if len(_split(line, d)) >= 3:
            return "flat", d, d
    outer = _split(line, ",")
    if len(outer) == 1:
        for d in DELIMS:
            if len(_split(outer[0], d)) >= 3:
                return "nested", ",", d
    return None


def decode(raw: bytes) -> tuple[str, str]:
    """UTF-8 (avec ou sans BOM) ; à défaut cp1252, sans jamais échouer."""
    try:
        return raw.decode("utf-8-sig"), "utf-8-sig"
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace"), "cp1252"


def parse_text(text: str) -> ParseResult:
    res = ParseResult()
    m = _FIRST_LINE.search(text)
    if not m:  # BOM + CRLF seulement, ou rien du tout : groupe vide, ce n'est pas une erreur
        res.empty = True
        return res

    det = _detect(m.group(0))
    if det is None:
        res.failed = True
        res.fmt = "unknown"
        res.issues.append(Issue(1, "FORMAT_INCONNU", "première ligne non reconnue (moins de 3 champs)"))
        return res
    res.fmt, outer, inner = det

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=outer)
    idx: dict[str, int] | None = None
    while True:
        try:
            row = next(reader)
        except StopIteration:
            break
        except csv.Error as e:
            res.issues.append(Issue(reader.line_num, "LIGNE_NON_PARSABLE", f"erreur CSV ({e})"))
            continue
        if not row or all(not c.strip() for c in row):
            continue
        line_no = reader.line_num

        fields: list[str] | None = None
        if len(row) >= 3:
            fields = row
        elif len(row) == 1:
            try:
                inner_row = _split(row[0], inner)
            except csv.Error:
                inner_row = []
            if len(inner_row) >= 3:
                fields = inner_row
        if fields is None:
            res.issues.append(Issue(line_no, "LIGNE_NON_PARSABLE", f"{len(row)} champ(s) exploitable(s), 3 attendus"))
            continue

        if idx is None:  # première ligne exploitable : en-tête ?
            norm = {_ALIASES.get(re.sub(r"[\s_-]", "", c.lower())): i for i, c in enumerate(fields)}
            norm.pop(None, None)
            if "sam" in norm:
                idx = norm
                continue
            idx = {"name": 0, "dn": 1, "sam": 2}
            res.issues.append(Issue(line_no, "EN_TETE_ABSENT", "pas d'en-tête reconnu : colonnes lues par position (nom, DN, identifiant)"))

        res.data_rows += 1
        sam = clean(fields[idx["sam"]]) if idx["sam"] < len(fields) else ""
        if not sam:
            res.issues.append(Issue(line_no, "LIGNE_NON_PARSABLE", "identifiant (SamAccountName) vide"))
            continue
        dn = clean(fields[idx["dn"]]) if "dn" in idx and idx["dn"] < len(fields) else ""
        name = clean(fields[idx["name"]]) if "name" in idx and idx["name"] < len(fields) else ""
        res.members.append(RawMember(name or sam, dn, sam, line_no))

    if not res.members and res.data_rows == 0 and not res.issues:
        res.empty = True
    elif not res.members and not res.data_rows:
        res.empty = True
    return res


def parse_file(path: str | Path) -> ParseResult:
    """Lit un export. Lève OSError si le fichier est illisible (géré par l'appelant)."""
    raw = Path(path).read_bytes()
    text, enc = decode(raw)
    res = parse_text(text)
    res.encoding = enc
    if enc != "utf-8-sig":
        res.issues.append(Issue(0, "ENCODAGE", "fichier non UTF-8, relu en cp1252"))
    return res
