# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Parcours des exports et mise en forme des lignes par ressource.

L'outil ne fait QUE de la mise en forme :
  * chaque ligne d'un CSV devient une ligne du classeur, avec le niveau de son groupe (Modification / Lecture) ;
  * pas de fusion Full/Read : une personne présente dans les deux groupes apparaît deux fois, c'est voulu ;
  * pas de dédoublonnage, pas d'interprétation du DN, pas de détection d'anomalies ;
  * jamais de fusion entre dossiers.
Seuls les problèmes de LECTURE sont consignés (fichier illisible, ligne non lisible, fichier ignoré),
pour qu'aucune ligne ne disparaisse sans que personne ne le sache.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import logging
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

from model import (FULL, LEVELS, NONE, READ, Book, CompileResult, GroupFile, Person, ReadProblem,
                   ResourceData, ScopeData)
from parser import parse_file

log = logging.getLogger("revue_droits")

_FS_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class FileRef:
    path: Path
    rel: str            # chemin relatif à l'entrée (posix)
    scope_dir: Path
    subpath: str        # chemin relatif au dossier de périmètre ('' le plus souvent)
    level: str
    resource: str
    stem: str


# --------------------------------------------------------------------------- découverte
def _suffix_regex(cfg: dict) -> re.Pattern:
    f, r = cfg["resources"]["suffix_full"], cfg["resources"]["suffix_read"]
    return re.compile(rf"^(?P<res>.+?)(?P<suf>{re.escape(f)}|{re.escape(r)})$", re.I)


def row_key(scope_name: str, subpath: str, resource: str, level: str, sam: str, occurrence: int = 1) -> str:
    """Clé stable d'une ligne : indépendante du chemin d'entrée, de la casse et de l'ordre des lignes.

    `occurrence` numérote les lignes qui répéteraient le même identifiant dans le même groupe
    (les doublons sont conservés tels quels) : 1 pour la première, 2 pour la suivante, etc.
    """
    raw = "|".join((scope_name.casefold(), subpath.casefold(), resource.casefold(), level, sam.casefold(), str(occurrence)))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def _find_scope(d: Path, root: Path, pattern: re.Pattern) -> Path:
    """Dossier de périmètre : l'ancêtre le plus proche dont le nom correspond au motif, sinon le dossier du fichier."""
    cur = d
    while True:
        if pattern.match(cur.name):
            return cur
        if cur == root or cur.parent == cur:
            return d
        cur = cur.parent


def discover(root: Path, cfg: dict, problems: list[ReadProblem], exclude: Path | None = None) -> list[FileRef]:
    pattern = re.compile(cfg["input"]["folder_pattern"])
    suffix = _suffix_regex(cfg)
    out: list[FileRef] = []
    seen: set[tuple] = set()
    excl = exclude.resolve() if exclude else None
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(x for x in dirnames if not (excl and (Path(dirpath) / x).resolve() == excl))
        d = Path(dirpath)
        for fn in sorted(filenames):
            if not fn.lower().endswith(".csv"):
                continue
            p = d / fn
            rel = p.relative_to(root).as_posix()
            scope_dir = _find_scope(d, root, pattern)
            stem = fn[:-4]
            m = suffix.match(stem)
            if m:
                level = FULL if m.group("suf").casefold() == cfg["resources"]["suffix_full"].casefold() else READ
                res = m.group("res")
            elif str(cfg["resources"].get("unsuffixed", "unknown")).lower() == "unknown":
                level, res = NONE, stem          # ex. Delta.csv : traité, avec le droit « Non précisé »
            else:
                problems.append(ReadProblem("SUFFIXE_INCONNU", scope_dir.name, stem, rel,
                                            "fichier ignoré : le nom ne se termine ni par _Full ni par _Read"))
                continue
            sub = d.relative_to(scope_dir).as_posix()
            sub = "" if sub == "." else sub
            ident = (str(scope_dir), sub, res.casefold(), level)
            if ident in seen:
                problems.append(ReadProblem("DOUBLON_FICHIER", scope_dir.name, res, rel,
                                            "deux fichiers pour le même groupe (casse différente) : le premier est conservé"))
                continue
            seen.add(ident)
            out.append(FileRef(p, rel, scope_dir, sub, level, res, stem))
    return out


def _fs_safe(text: str) -> str:
    return _FS_BAD.sub("_", text).strip(" .") or "_"


def scope_label(name: str, cfg: dict) -> tuple[str, str]:
    """(libellé du périmètre, code) d'après folder_pattern et workbook_label."""
    m = re.match(cfg["input"]["folder_pattern"], name)
    groups = m.groupdict() if m else {}
    folder = name[7:] if name.lower().startswith("export_") else name
    values = {"folder": folder, "full": name, "code": groups.get("code") or "", "name": groups.get("name") or folder}
    try:
        label = cfg["input"]["workbook_label"].format(**values)
    except (KeyError, IndexError):
        label = folder
    return _fs_safe(label or folder), values["code"]


# --------------------------------------------------------------------------- lecture
def _read_group(ref: FileRef, problems: list[ReadProblem], scope: str, resource: str, stats: Counter) -> GroupFile:
    g = GroupFile(level=ref.level, resource=ref.resource, rel_path=ref.rel, stem=ref.stem)
    try:
        res = parse_file(ref.path)
    except Exception as e:  # noqa: BLE001 - un fichier illisible ne doit jamais tout arrêter
        g.error = f"illisible ({type(e).__name__})"
        problems.append(ReadProblem("FICHIER_ILLISIBLE", scope, resource, ref.rel, g.error))
        stats["files_unreadable"] += 1
        return g
    for i in res.issues:
        if i.code == "ENCODAGE":
            continue            # information technique, sans effet sur le contenu affiché
        problems.append(ReadProblem(i.code if not res.failed else "FORMAT_INCONNU", scope, resource, ref.rel,
                                    f"ligne {i.line} : {i.detail}" if i.line else i.detail))
        if i.code == "LIGNE_NON_PARSABLE":
            stats["lines_unparsable"] += 1
    if res.failed:
        g.error = "format inexploitable"
        stats["files_unreadable"] += 1
        return g
    g.members = list(res.members)            # TOUTES les lignes, sans dédoublonner
    g.empty = not g.members
    if g.empty:
        stats["files_empty"] += 1             # groupe vide : normal, affiché « Aucun membre »
    stats["files_read"] += 1
    return g


def _build_rows(scope_name: str, subpath: str, name: str, groups: list[GroupFile | None]) -> list[Person]:
    rows: list[Person] = []
    for g in groups:                          # Full, Read puis niveau non précisé : aucune fusion entre eux
        if g is None or g.error:
            continue
        seen: Counter = Counter()
        for m in g.members:
            k = m.sam.casefold()
            seen[k] += 1
            rows.append(Person(m.sam, m.name, m.dn, g.level, row_key(scope_name, subpath, name, g.level, m.sam, seen[k])))
    # tri stable et déterministe : droit (Modification, Lecture, non précisé), nom, identifiant
    rank = {l: i for i, l in enumerate(LEVELS)}
    rows.sort(key=lambda p: (rank[p.level], p.name.casefold(), p.sam.casefold()))
    return rows


# --------------------------------------------------------------------------- orchestration
def _export_date(files: list[FileRef], override: dt.date | None) -> dt.date:
    if override:
        return override
    mt = max((f.path.stat().st_mtime for f in files), default=0)
    return dt.date.fromtimestamp(mt) if mt else dt.date.today()


def compile_all(root: str | Path, cfg: dict, only: list[str] | None = None,
                export_date: dt.date | None = None, exclude: str | Path | None = None) -> CompileResult:
    root = Path(root)
    if not root.is_dir():
        raise NotADirectoryError(f"dossier d'entrée introuvable : {root}")
    problems: list[ReadProblem] = []
    stats: Counter = Counter()

    refs = discover(root, cfg, problems, Path(exclude) if exclude else None)
    scopes_all: dict[Path, list[FileRef]] = defaultdict(list)
    for r in refs:
        scopes_all[r.scope_dir].append(r)

    def wanted(scope_dir: Path) -> bool:
        if not only:
            return True
        label, _code = scope_label(scope_dir.name, cfg)
        return any(o.casefold() in {scope_dir.name.casefold(), label.casefold()} for o in only)

    scopes: list[ScopeData] = []
    for scope_dir in sorted(scopes_all, key=lambda p: (p.name.casefold(), str(p))):
        if not wanted(scope_dir):
            continue
        files = scopes_all[scope_dir]
        label, code = scope_label(scope_dir.name, cfg)
        scope = ScopeData(scope_dir.name, label, code, _export_date(files, export_date))
        groups: dict[tuple, dict[str, FileRef]] = defaultdict(dict)
        for r in files:
            groups[(r.subpath, r.resource.casefold())][r.level] = r
        for (sub, _cf), lv in sorted(groups.items(), key=lambda kv: (kv[0][0].casefold(), kv[0][1])):
            name = (lv.get(FULL) or lv.get(READ) or lv.get(NONE)).resource
            gf = _read_group(lv[FULL], problems, scope.name, name, stats) if FULL in lv else None
            gr = _read_group(lv[READ], problems, scope.name, name, stats) if READ in lv else None
            go = _read_group(lv[NONE], problems, scope.name, name, stats) if NONE in lv else None
            res = ResourceData(scope.name, sub, name, gf, gr, _build_rows(scope.name, sub, name, [gf, gr, go]), go)
            scope.resources.append(res)
            stats["resources"] += 1
            stats["rows"] += len(res.rows)
        scopes.append(scope)
        stats["scopes"] += 1

    problems.sort(key=lambda p: (p.folder.casefold(), p.resource.casefold(), p.code, p.file))
    return CompileResult(scopes, problems, dict(stats))


def group_books(scopes: list[ScopeData], cfg: dict) -> list[Book]:
    """Un classeur par dossier ; ou par ayant droit si `owners` regroupe plusieurs dossiers."""
    owners = {k.casefold(): v for k, v in (cfg.get("owners") or {}).items()}
    books: dict[str, Book] = {}
    for s in scopes:
        owner = owners.get(s.name.casefold())
        key = _fs_safe(owner) if owner else s.label
        books.setdefault(key, Book(key, [])).scopes.append(s)
    return [books[k] for k in sorted(books, key=str.casefold)]
