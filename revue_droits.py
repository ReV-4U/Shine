#!/usr/bin/env python3
# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Revue périodique des droits d'accès : met en forme des exports AD (CSV) en classeurs Excel à valider.

Mise en forme uniquement : aucune fusion, aucun dédoublonnage, aucune analyse des données.

Usage :
    python revue_droits.py --input <chemin> [--output <dossier>] [--config config.yaml]
                           [--checkbox-mode native|dropdown] [--only <dossier>]... [--deadline JJ/MM/AAAA]

Sortie (par défaut ~/Documents/shine) : un dossier `Revue_NTFS_<X>.<NOM>_JJMMAA` par dossier d'export,
contenant le classeur compilé ; `rapport_lecture_JJMMAA.csv` à côté si des fichiers n'ont pas pu être lus.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import logging
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from compiler import compile_all, group_books  # noqa: E402
from config import DEFAULT_CONFIG_PATH, ConfigError, load_config  # noqa: E402
from excel import build_book, safe  # noqa: E402
from xlsxwriter.exceptions import FileCreateError  # noqa: E402

log = logging.getLogger("revue_droits")


def write_read_report(path: Path, problems, delimiter: str = ";") -> None:
    """rapport_lecture.csv : fichiers/lignes qui n'ont pas pu être lus (UTF-8 avec BOM, valeurs neutralisées)."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=delimiter)
        w.writerow(["code", "dossier", "ressource", "fichier", "detail"])
        for p in problems:
            w.writerow([safe(x) for x in (p.code, p.folder, p.resource, p.file, p.detail)])


def app_dir() -> Path:
    """Dossier de l'exécutable (version .exe) ou du script."""
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def default_user_config() -> str | None:
    """`config.yaml` posé à côté de l'exécutable : chargé automatiquement, il remplace la copie embarquée.

    Depuis les sources, c'est le fichier par défaut lui-même : inutile de le charger deux fois.
    """
    f = app_dir() / "config.yaml"
    return str(f) if f.is_file() and f.resolve() != DEFAULT_CONFIG_PATH.resolve() else None


DEADLINE_FORMATS = ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%Y-%m-%d")


def parse_deadline(text: str) -> str:
    """Normalise une date saisie (JJ/MM/AAAA, JJ-MM-AAAA, JJ.MM.AAAA, AAAA-MM-JJ) en JJ/MM/AAAA. ValueError sinon."""
    return _parse_date(text).strftime("%d/%m/%Y")


def _parse_date(text: str) -> dt.date:
    t = (text or "").strip()
    for fmt in DEADLINE_FORMATS:
        try:
            return dt.datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    raise ValueError("format attendu : JJ/MM/AAAA")


def _deadline_arg(s: str) -> str:
    try:
        return parse_deadline(s)
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e)) from e


def _date(s: str) -> dt.date:
    try:
        return dt.date.fromisoformat(s)
    except ValueError as e:
        raise argparse.ArgumentTypeError("format attendu : AAAA-MM-JJ") from e


def parse_args(argv=None):
    p = argparse.ArgumentParser(prog="revue_droits.py", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--input", required=True, help="dossier à parcourir récursivement")
    p.add_argument("--output", help="dossier racine de sortie (défaut : output.root, soit ~/Documents/shine ; créé si besoin)")
    p.add_argument("--config", help="fichier YAML fusionné avec la configuration par défaut")
    p.add_argument("--checkbox-mode", choices=("native", "dropdown"), help="cases natives (Excel 365) ou liste déroulante")
    p.add_argument("--only", action="append", default=[], metavar="DOSSIER",
                   help="ne traiter que ce dossier (nom ou libellé) ; option répétable")
    p.add_argument("--period", help="période AAAA-MM du nom de fichier (défaut : mois courant)")
    p.add_argument("--deadline", type=_deadline_arg, metavar="JJ/MM/AAAA",
                   help="date limite de retour des classeurs, affichée dans les onglets En-tête et Instructions (défaut : review.deadline)")
    p.add_argument("--export-date", type=_date, help="date de l'export AAAA-MM-JJ (défaut : date des fichiers CSV)")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def _os_error(e: BaseException) -> OSError | None:
    """L'OSError d'origine (xlsxwriter enveloppe l'erreur d'écriture dans FileCreateError)."""
    for x in (e, *e.args):
        if isinstance(x, OSError):
            return x
    return None


def _write_error_message(out: Path, e: BaseException) -> str:
    err = _os_error(e)
    reason = (err.strerror or type(err).__name__) if err else type(e).__name__
    name = Path(err.filename).name if err and err.filename else ""
    return (f"Impossible d'écrire les classeurs dans « {out} » ({reason}){f' : {name}' if name else ''}.\n"
            "Causes les plus probables :\n"
            "  1) Le classeur du même jour est déjà ouvert dans Excel : fermez-le puis relancez.\n"
            "  2) Windows bloque l'écriture dans ce dossier : « Accès contrôlé aux dossiers » (Sécurité Windows), antivirus "
            "ou synchronisation OneDrive. Autorisez Shine.exe ou choisissez un autre dossier de sortie (--output).\n"
            "  3) Chemin trop long : choisissez un dossier de sortie plus court.")


def run(args, now: dt.datetime | None = None) -> int:
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(message)s")
    try:
        cfg = load_config(args.config or default_user_config())
    except (ConfigError, OSError) as e:
        print(f"Configuration invalide : {e}", file=sys.stderr)
        return 2
    if args.deadline:
        cfg["review"]["deadline"] = args.deadline
    mode = args.checkbox_mode or cfg["excel"]["checkbox_mode"]
    now = now or dt.datetime.now()
    period = args.period or cfg["review"].get("period") or now.strftime("%Y-%m")

    out = Path(args.output or cfg["output"]["root"]).expanduser()
    args.resolved_output = out
    try:
        result = compile_all(args.input, cfg, only=args.only or None, export_date=args.export_date, exclude=out)
    except NotADirectoryError as e:
        print(str(e), file=sys.stderr)
        return 2
    if args.only and not result.scopes:
        print("Aucun dossier ne correspond à --only.", file=sys.stderr)
        return 2

    stamp = now.strftime(cfg["output"]["date_format"])
    books = group_books(result.scopes, cfg)
    reports, used = [], set()
    try:
        out.mkdir(parents=True, exist_ok=True)
        for book in books:
            stem = cfg["output"]["folder_pattern"].format(label=book.label, date=stamp)
            k = 2
            while stem.casefold() in used:
                stem, k = f"{cfg['output']['folder_pattern'].format(label=book.label, date=stamp)}_{k}", k + 1
            used.add(stem.casefold())
            folder = out / stem
            folder.mkdir(parents=True, exist_ok=True)
            reports.append(build_book(folder / f"{stem}.xlsx", book, cfg, mode, period, now))
    except (OSError, FileCreateError) as e:
        print(_write_error_message(out, e), file=sys.stderr)
        return 3
    report_path = out / f"rapport_lecture_{stamp}.csv"
    try:
        if result.problems:
            write_read_report(report_path, result.problems, cfg["output"]["report_delimiter"])
        elif report_path.exists():
            report_path.unlink()          # pas de problème cette fois : on ne laisse pas un ancien rapport trompeur
    except OSError as e:
        print(_write_error_message(out, e), file=sys.stderr)
        return 3

    # ---- résumé console : uniquement des compteurs, jamais de données personnelles
    st = result.stats
    by_code = Counter(p.code for p in result.problems)
    print("\n=== Résumé ===")
    print(f"Dossiers traités   : {st.get('scopes', 0)}")
    print(f"Ressources         : {st.get('resources', 0)}")
    print(f"Fichiers lus       : {st.get('files_read', 0)}  (groupes vides : {st.get('files_empty', 0)}, "
          f"illisibles : {st.get('files_unreadable', 0)})")
    print(f"Lignes             : {st.get('rows', 0)}  (une ligne par ligne d'export, sans fusion)")
    print(f"Problèmes de lecture : {len(result.problems)}")
    for code, n in sorted(by_code.items(), key=lambda kv: (-kv[1], kv[0])):
        print(f"   - {code:<22} {n}")
    print(f"Mode de réponse    : {mode}")
    print(f"Date limite        : {cfg['review']['deadline']}")
    for label, value, how in (("date limite", cfg["review"]["deadline"], "--deadline"),
                              ("contact", cfg["review"]["contact"], "review.contact dans config.yaml")):
        if str(value).strip().startswith("[") or not str(value).strip():
            print(f"ATTENTION : {label} non renseigné(e) : les classeurs afficheront un texte à compléter ({how}).")
    print(f"Dossier de sortie  : {out}")
    print("Fichiers produits  :")
    for r in reports:
        print(f"   - {Path(r.path).parent.name}{'/'}{Path(r.path).name}  ({len(r.sheets)} ressource(s), {r.rows} ligne(s))")
    if result.problems:
        print(f"   - {report_path.name}")
    return 0


def main(argv=None) -> int:
    return run(parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
