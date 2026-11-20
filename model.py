# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Structures de données partagées par le parseur, le compilateur et la génération Excel.

L'outil ne fait que de la MISE EN FORME : chaque ligne d'un CSV devient une ligne du classeur.
Aucune fusion entre groupes, aucun dédoublonnage, aucune interprétation du DN.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

FULL = "full"
READ = "read"
NONE = "none"          # fichier sans suffixe _Full / _Read : le droit n'est pas précisé
LEVELS = (FULL, READ, NONE)


@dataclass(frozen=True)
class Issue:
    """Problème de lecture d'un fichier (sans aucune donnée personnelle)."""
    line: int
    code: str
    detail: str


@dataclass
class RawMember:
    name: str
    dn: str
    sam: str
    line: int = 0


@dataclass
class ParseResult:
    members: list[RawMember] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    empty: bool = False          # fichier vide ou sans ligne de données
    failed: bool = False         # format inexploitable : le fichier est écarté
    fmt: str = "empty"           # 'nested' | 'flat' | 'empty' | 'unknown'
    encoding: str = "utf-8-sig"
    data_rows: int = 0


@dataclass
class Person:
    """Une ligne du classeur = une ligne d'un CSV (un membre d'UN groupe, avec son niveau)."""
    sam: str
    name: str
    dn: str
    level: str                   # FULL | READ | NONE
    key: str                     # clé stable (voir compiler.row_key)


@dataclass
class GroupFile:
    level: str                   # FULL | READ | NONE
    resource: str                # nom de la ressource tel qu'écrit dans le fichier
    rel_path: str                # chemin relatif à l'entrée
    stem: str                    # nom de fichier sans extension
    members: list[RawMember] = field(default_factory=list)
    empty: bool = False
    error: str | None = None     # fichier illisible / format inexploitable


@dataclass
class ResourceData:
    scope_name: str              # nom du dossier de périmètre
    subpath: str                 # sous-dossier relatif au périmètre ('' en général)
    name: str                    # nom de la ressource (suffixe retiré)
    full: GroupFile | None
    read: GroupFile | None
    rows: list[Person] = field(default_factory=list)
    other: GroupFile | None = None   # fichier sans suffixe de niveau (droit non précisé)

    def count(self, level: str) -> int | None:
        """Nombre de lignes du groupe ; None si le fichier est illisible."""
        g = self.full if level == FULL else self.read if level == READ else self.other
        if g is None:
            return 0
        return None if g.error else len(g.members)


@dataclass
class ScopeData:
    name: str                    # nom du dossier (ex. Export_E.Test7)
    label: str                   # libellé calculé (nom de classeur)
    code: str
    export_date: dt.date
    resources: list[ResourceData] = field(default_factory=list)


@dataclass
class Book:
    """Un classeur à produire : un dossier (défaut) ou plusieurs dossiers d'un ayant droit."""
    label: str
    scopes: list[ScopeData]


@dataclass
class ReadProblem:
    """Problème de LECTURE (fichier ou ligne qui n'a pas pu être pris en compte). Pas une analyse des données."""
    code: str
    folder: str
    resource: str
    file: str
    detail: str


@dataclass
class CompileResult:
    scopes: list[ScopeData]
    problems: list[ReadProblem]
    stats: dict
